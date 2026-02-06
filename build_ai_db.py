"""
엠씨넥스 AI 분석용 통합 DB 생성
- mcnex_full.db: 구조화된 재무/임원/배당 등 정량 데이터
- mcnex_dart.db: 공시 원문에서 핵심 텍스트 추출
→ mcnex_ai.db: 하나로 통합
"""
import sqlite3
import os
import sys
import re
from html import unescape

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FULL_DB = os.path.join(BASE_DIR, "mcnex_full.db")
DART_DB = os.path.join(BASE_DIR, "mcnex_dart.db")
AI_DB = os.path.join(BASE_DIR, "mcnex_ai.db")

if os.path.exists(AI_DB):
    os.remove(AI_DB)


def clean_text(text):
    """텍스트 정리"""
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def extract_section(text, start_patterns, end_patterns, max_len=5000):
    """본문에서 특정 섹션 추출"""
    for sp in start_patterns:
        match = re.search(sp, text, re.IGNORECASE)
        if match:
            start = match.start()
            # 끝 패턴 찾기
            sub = text[start:]
            best_end = min(len(sub), max_len)
            for ep in end_patterns:
                end_match = re.search(ep, sub[100:])  # 시작 직후는 건너뛰기
                if end_match:
                    best_end = min(best_end, end_match.start() + 100)
                    break
            return clean_text(sub[:best_end])
    return ""


# ============================================================
print("=" * 60)
print("STEP 1: 구조화 데이터 복사 (mcnex_full.db → mcnex_ai.db)")
print("=" * 60)

# mcnex_full.db를 기반으로 복사
conn_full = sqlite3.connect(FULL_DB)
conn_ai = sqlite3.connect(AI_DB)

# full DB의 모든 테이블/뷰/인덱스를 AI DB로 복사
conn_full.backup(conn_ai)
conn_full.close()
print("  구조화 데이터 복사 완료")

# ============================================================
print()
print("=" * 60)
print("STEP 2: 공시 원문에서 핵심 텍스트 추출")
print("=" * 60)

conn_dart = sqlite3.connect(DART_DB)

# 새 테이블 생성
conn_ai.executescript("""
    -- 사업보고서 핵심 섹션 (회사개요, 사업내용, 제품, 리스크 등)
    CREATE TABLE IF NOT EXISTS business_report_sections (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bsns_year TEXT,
        rcept_no TEXT,
        section_name TEXT,
        section_text TEXT
    );

    -- 특허 공시
    CREATE TABLE IF NOT EXISTS patents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rcept_dt TEXT,
        rcept_no TEXT,
        report_nm TEXT,
        patent_name TEXT,
        patent_detail TEXT,
        patent_date TEXT,
        patent_plan TEXT
    );

    -- 주요 경영 이벤트 요약
    CREATE TABLE IF NOT EXISTS key_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rcept_dt TEXT,
        rcept_no TEXT,
        event_type TEXT,
        event_summary TEXT
    );

    -- 잠정실적 공시
    CREATE TABLE IF NOT EXISTS earnings_announcements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        rcept_dt TEXT,
        rcept_no TEXT,
        report_nm TEXT,
        content TEXT
    );

    CREATE INDEX IF NOT EXISTS idx_brs_year ON business_report_sections(bsns_year);
    CREATE INDEX IF NOT EXISTS idx_pat_dt ON patents(rcept_dt);
    CREATE INDEX IF NOT EXISTS idx_evt_dt ON key_events(rcept_dt);
    CREATE INDEX IF NOT EXISTS idx_earn_dt ON earnings_announcements(rcept_dt);
""")

# ---- 2-1. 사업보고서 핵심 섹션 추출 ----
print("  사업보고서 핵심 섹션 추출 중...")

annual_reports = conn_dart.execute("""
    SELECT disc.rcept_dt, disc.rcept_no, disc.report_nm
    FROM disclosures disc
    WHERE disc.report_nm LIKE '사업보고서%'
    ORDER BY disc.rcept_dt
""").fetchall()

section_defs = [
    ("회사개요", [r'I\.\s*회사의\s*개요', r'회사의\s*개요', r'1\.\s*회사의\s*개요'],
     [r'II\.\s*사업', r'Ⅱ\.\s*사업', r'\n\s*2\.\s*회사의\s*연혁']),
    ("회사연혁", [r'회사의\s*연혁', r'2\.\s*회사의\s*연혁'],
     [r'자본금\s*변동', r'3\.\s*자본금', r'II\.\s*사업']),
    ("사업내용", [r'II\.\s*사업의\s*내용', r'Ⅱ\.\s*사업의\s*내용', r'사업의\s*내용'],
     [r'III\.\s*재무', r'Ⅲ\.\s*재무', r'IV\.\s*이사']),
    ("주요제품_매출", [r'주요\s*제품\s*등의\s*현황', r'주요\s*제품.*매출'],
     [r'주요\s*원재료', r'생산\s*및\s*설비', r'III\.\s*재무']),
    ("연구개발", [r'연구개발활동', r'연구\s*개발\s*활동'],
     [r'그\s*밖에\s*투자', r'III\.\s*재무', r'재무에\s*관한']),
    ("위험관리_전망", [r'경영상의\s*주요\s*계약', r'파생상품', r'위험\s*관리'],
     [r'III\.\s*재무', r'IV\.\s*이사', r'감사인']),
    ("임원_보수", [r'IV\.\s*이사의\s*경영진단', r'Ⅳ\.\s*이사의\s*경영진단', r'이사의\s*경영진단'],
     [r'V\.\s*회계', r'Ⅴ\.\s*회계', r'감사인']),
    ("주주_배당", [r'주주에\s*관한\s*사항', r'배당에\s*관한\s*사항'],
     [r'임원\s*및\s*직원', r'이사회\s*등', r'V\.\s*회계']),
]

for rcept_dt, rcept_no, report_nm in annual_reports:
    # 연도 추출
    year_match = re.search(r'\((\d{4})\.\d{2}\)', report_nm)
    bsns_year = year_match.group(1) if year_match else rcept_dt[:4]

    # 해당 사업보고서의 메인 파일 (가장 큰 것)
    doc = conn_dart.execute("""
        SELECT body_text FROM documents
        WHERE rcept_no = ? ORDER BY file_size DESC LIMIT 1
    """, (rcept_no,)).fetchone()

    if not doc or not doc[0]:
        continue

    text = doc[0]
    extracted = 0
    for sec_name, start_pats, end_pats in section_defs:
        section_text = extract_section(text, start_pats, end_pats, max_len=8000)
        if section_text and len(section_text) > 50:
            conn_ai.execute("""
                INSERT INTO business_report_sections (bsns_year, rcept_no, section_name, section_text)
                VALUES (?, ?, ?, ?)
            """, (bsns_year, rcept_no, sec_name, section_text))
            extracted += 1

    print(f"    {bsns_year} ({report_nm}): {extracted}개 섹션 추출")

conn_ai.commit()

# ---- 2-2. 특허 공시 추출 ----
print()
print("  특허 공시 추출 중...")

patent_discs = conn_dart.execute("""
    SELECT disc.rcept_dt, disc.rcept_no, disc.report_nm, d.body_text
    FROM disclosures disc
    JOIN documents d ON disc.rcept_no = d.rcept_no
    WHERE disc.report_nm LIKE '%특허%' OR disc.report_nm LIKE '%투자판단%특허%'
    ORDER BY disc.rcept_dt
""").fetchall()

patent_count = 0
for rcept_dt, rcept_no, report_nm, body in patent_discs:
    if not body:
        continue
    # 특허명 추출
    name_match = re.search(r'특허명칭\s*[:\s]*(.+?)(?:\(2\)|특허\s*주요|$)', body)
    patent_name = clean_text(name_match.group(1)) if name_match else ""
    if not patent_name:
        # 제목에서 추출
        title_match = re.search(r'\(([^)]*특허[^)]*)\)', report_nm) or re.search(r'제목\s*(.+?)(?:2\.|주요)', body)
        patent_name = clean_text(title_match.group(1)) if title_match else report_nm

    # 주요 내용 추출
    detail_match = re.search(r'주요\s*내용\s*(.+?)(?:\(3\)|특허권자|특허취득|3\.\s*이사회)', body, re.DOTALL)
    patent_detail = clean_text(detail_match.group(1))[:3000] if detail_match else clean_text(body)[:2000]

    # 취득일
    date_match = re.search(r'특허취득\s*일자\s*[:\s]*([\d\-\.]+)', body)
    patent_date = date_match.group(1) if date_match else rcept_dt

    # 활용 계획
    plan_match = re.search(r'특허활용\s*계획\s*(.+?)(?:3\.\s*이사회|사외이사|$)', body, re.DOTALL)
    patent_plan = clean_text(plan_match.group(1))[:1000] if plan_match else ""

    conn_ai.execute("""
        INSERT INTO patents (rcept_dt, rcept_no, report_nm, patent_name, patent_detail, patent_date, patent_plan)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (rcept_dt, rcept_no, report_nm, patent_name, patent_detail, patent_date, patent_plan))
    patent_count += 1

conn_ai.commit()
print(f"    {patent_count}건 특허 추출 완료")

# ---- 2-3. 주요 경영 이벤트 추출 ----
print()
print("  주요 경영 이벤트 추출 중...")

event_types = [
    ('유상증자', '%유상증자%'),
    ('무상증자', '%무상증자%'),
    ('전환사채', '%전환사채%'),
    ('자기주식', '%자기주식%결정%'),
    ('타법인투자', '%타법인주식%취득%'),
    ('채무보증', '%채무보증%'),
    ('신규시설투자', '%신규시설투자%'),
    ('배당결정', '%배당결정%'),
    ('주식소각', '%주식소각%'),
    ('IR', '%IR%개최%'),
    ('실적전망', '%영업실적%전망%'),
    ('매출변동', '%매출액또는손익%변동%'),
    ('매출변동', '%매출액또는손익%변경%'),
    ('상장폐지', '%상장폐지%'),
    ('대량보유', '%대량보유%'),
]

event_count = 0
for etype, pattern in event_types:
    rows = conn_dart.execute(f"""
        SELECT disc.rcept_dt, disc.rcept_no, disc.report_nm, SUBSTR(d.body_text, 1, 3000)
        FROM disclosures disc
        LEFT JOIN documents d ON disc.rcept_no = d.rcept_no
        WHERE disc.report_nm LIKE '{pattern}'
        AND disc.report_nm NOT LIKE '%기재정정%'
        AND disc.report_nm NOT LIKE '%첨부정정%'
        ORDER BY disc.rcept_dt
    """).fetchall()

    for rcept_dt, rcept_no, report_nm, body in rows:
        summary = clean_text(body)[:2000] if body else report_nm
        conn_ai.execute("""
            INSERT INTO key_events (rcept_dt, rcept_no, event_type, event_summary)
            VALUES (?, ?, ?, ?)
        """, (rcept_dt, rcept_no, etype, summary))
        event_count += 1

conn_ai.commit()
print(f"    {event_count}건 이벤트 추출 완료")

# ---- 2-4. 잠정실적 공시 추출 ----
print()
print("  잠정실적 공시 추출 중...")

earnings = conn_dart.execute("""
    SELECT disc.rcept_dt, disc.rcept_no, disc.report_nm, SUBSTR(d.body_text, 1, 3000)
    FROM disclosures disc
    JOIN documents d ON disc.rcept_no = d.rcept_no
    WHERE disc.report_nm LIKE '%잠정%실적%'
    AND disc.report_nm NOT LIKE '%기재정정%'
    ORDER BY disc.rcept_dt
""").fetchall()

earn_count = 0
for rcept_dt, rcept_no, report_nm, body in earnings:
    content = clean_text(body)[:2000] if body else ""
    conn_ai.execute("""
        INSERT INTO earnings_announcements (rcept_dt, rcept_no, report_nm, content)
        VALUES (?, ?, ?, ?)
    """, (rcept_dt, rcept_no, report_nm, content))
    earn_count += 1

conn_ai.commit()
print(f"    {earn_count}건 실적공시 추출 완료")

conn_dart.close()

# ============================================================
print()
print("=" * 60)
print("STEP 3: 편의 뷰 추가")
print("=" * 60)

conn_ai.executescript("""
    DROP VIEW IF EXISTS v_annual_performance;
    CREATE VIEW v_annual_performance AS
    SELECT
        bsns_year,
        MAX(CASE WHEN account_nm LIKE '%매출액%' AND sj_div='CIS' THEN thstrm_amount END) AS revenue,
        MAX(CASE WHEN account_nm LIKE '%영업이익%' AND sj_div='CIS' THEN thstrm_amount END) AS operating_profit,
        MAX(CASE WHEN account_nm = '당기순이익(손실)' AND sj_div='CIS' THEN thstrm_amount END) AS net_income,
        MAX(CASE WHEN account_nm = '기본주당이익(손실)' AND sj_div='CIS' THEN thstrm_amount END) AS eps,
        MAX(CASE WHEN account_nm = '자산총계' AND sj_div='BS' THEN thstrm_amount END) AS total_assets,
        MAX(CASE WHEN account_nm = '부채총계' AND sj_div='BS' THEN thstrm_amount END) AS total_liabilities,
        MAX(CASE WHEN account_nm = '자본총계' AND sj_div='BS' THEN thstrm_amount END) AS total_equity
    FROM financial_statements
    WHERE reprt_code = '11011' AND reprt_nm LIKE '%연결%'
    GROUP BY bsns_year
    ORDER BY bsns_year;

    -- 사업보고서 섹션 요약 뷰
    CREATE VIEW IF NOT EXISTS v_business_sections AS
    SELECT bsns_year, section_name, LENGTH(section_text) as text_len,
           SUBSTR(section_text, 1, 200) as preview
    FROM business_report_sections
    ORDER BY bsns_year, section_name;

    -- 특허 이력 뷰
    CREATE VIEW IF NOT EXISTS v_patent_history AS
    SELECT rcept_dt, patent_name, patent_date,
           SUBSTR(patent_detail, 1, 200) as detail_preview,
           SUBSTR(patent_plan, 1, 200) as plan_preview
    FROM patents
    ORDER BY rcept_dt;

    -- 이벤트 타임라인 뷰
    CREATE VIEW IF NOT EXISTS v_event_timeline AS
    SELECT rcept_dt, event_type,
           SUBSTR(event_summary, 1, 300) as summary_preview
    FROM key_events
    ORDER BY rcept_dt DESC;

    -- DB 전체 요약 뷰 (갱신)
    DROP VIEW IF EXISTS v_db_summary;
    CREATE VIEW v_db_summary AS
    SELECT 'company_info' as tbl, '기업개황' as description, COUNT(*) as cnt FROM company_info
    UNION ALL SELECT 'disclosures', '전체공시목록', COUNT(*) FROM disclosures
    UNION ALL SELECT 'financial_statements', '전체재무제표(BS/IS/CF)', COUNT(*) FROM financial_statements
    UNION ALL SELECT 'financial_summary', '주요계정요약', COUNT(*) FROM financial_summary
    UNION ALL SELECT 'executives', '임원현황', COUNT(*) FROM executives
    UNION ALL SELECT 'employees', '직원현황', COUNT(*) FROM employees
    UNION ALL SELECT 'major_shareholders', '대주주현황', COUNT(*) FROM major_shareholders
    UNION ALL SELECT 'minority_shareholders', '소액주주현황', COUNT(*) FROM minority_shareholders
    UNION ALL SELECT 'dividends', '배당현황', COUNT(*) FROM dividends
    UNION ALL SELECT 'treasury_stock', '자기주식현황', COUNT(*) FROM treasury_stock
    UNION ALL SELECT 'capital_changes', '증자감자현황', COUNT(*) FROM capital_changes
    UNION ALL SELECT 'stock_total', '주식총수현황', COUNT(*) FROM stock_total
    UNION ALL SELECT 'investments', '타법인출자현황', COUNT(*) FROM investments
    UNION ALL SELECT 'individual_pay', '개인별보수(5억이상)', COUNT(*) FROM individual_pay
    UNION ALL SELECT 'business_report_sections', '사업보고서핵심섹션(텍스트)', COUNT(*) FROM business_report_sections
    UNION ALL SELECT 'patents', '특허공시', COUNT(*) FROM patents
    UNION ALL SELECT 'key_events', '주요경영이벤트', COUNT(*) FROM key_events
    UNION ALL SELECT 'earnings_announcements', '잠정실적공시', COUNT(*) FROM earnings_announcements;
""")
conn_ai.commit()

# ============================================================
print()
print("=" * 60)
print("DB 생성 완료!")
print("=" * 60)

db_size = os.path.getsize(AI_DB) / 1024 / 1024
print(f"  파일: {AI_DB}")
print(f"  크기: {db_size:.1f} MB")
print()
print("--- 전체 테이블/건수 ---")
rows = conn_ai.execute("SELECT * FROM v_db_summary").fetchall()
for r in rows:
    print(f"  {r[0]:35s} | {r[1]:25s} | {r[2]:>6}건")

print()
print("--- 연도별 실적 (연결, 억원) ---")
print(f"  {'연도':>6} | {'매출액':>8} | {'영업이익':>8} | {'순이익':>8} | {'총자산':>8} | {'총자본':>8}")
print("  " + "-" * 60)
rows = conn_ai.execute("SELECT * FROM v_annual_performance").fetchall()
for r in rows:
    def f(v):
        try: return f"{int(v)//100000000:>7}억"
        except: return f"{'':>8}"
    print(f"  {r[0]:>6} | {f(r[1])} | {f(r[2])} | {f(r[3])} | {f(r[5])} | {f(r[7])}")

print()
print("--- 추출된 사업보고서 섹션 ---")
rows = conn_ai.execute("SELECT bsns_year, section_name, text_len FROM v_business_sections").fetchall()
for r in rows:
    print(f"  {r[0]} | {r[1]:15s} | {r[2]:>6}자")

print()
print("--- 특허 이력 ---")
rows = conn_ai.execute("SELECT rcept_dt, patent_name FROM patents ORDER BY rcept_dt").fetchall()
for r in rows:
    print(f"  {r[0]} | {r[1][:60]}")

conn_ai.close()

print()
print("완료! AI에게 이 DB를 주면 회사 전체를 분석할 수 있습니다.")
