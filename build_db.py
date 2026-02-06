import sqlite3
import json
import os
import zipfile
import xml.etree.ElementTree as ET
import re
import sys
from html import unescape

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
DB_PATH = os.path.join(BASE_DIR, "mcnex_dart.db")
LIST_PATH = os.path.join(BASE_DIR, "disclosure_list.json")


def strip_xml_tags(text):
    """XML/HTML 태그 제거 후 텍스트만 추출"""
    text = re.sub(r'<[^>]+>', ' ', text)
    text = unescape(text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n', '\n', text)
    return text.strip()


def extract_text_from_xml(xml_bytes):
    """XML 바이트에서 본문 텍스트 추출"""
    # Try multiple encodings
    content = None
    for enc in ['utf-8', 'euc-kr', 'cp949', 'latin-1']:
        try:
            content = xml_bytes.decode(enc)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if content is None:
        content = xml_bytes.decode('utf-8', errors='replace')

    # Extract text from BODY section if present
    body_match = re.search(r'<BODY[^>]*>(.*?)</BODY>', content, re.DOTALL | re.IGNORECASE)
    if body_match:
        body_text = strip_xml_tags(body_match.group(1))
    else:
        body_text = strip_xml_tags(content)

    return content, body_text


def extract_summary_fields(xml_content):
    """XML SUMMARY 섹션에서 메타데이터 추출"""
    fields = {}
    for match in re.finditer(r'<EXTRACTION\s+ACODE="([^"]+)"[^>]*>([^<]*)</EXTRACTION>', xml_content):
        fields[match.group(1)] = match.group(2)
    return fields


def main():
    # Load disclosure list
    with open(LIST_PATH, "r", encoding="utf-8") as f:
        disclosures = json.load(f)

    print(f"공시 목록: {len(disclosures)}건")

    # Create database
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # ============ 테이블 생성 ============
    cur.executescript("""
        -- 공시 메타데이터 테이블
        CREATE TABLE disclosures (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            corp_code TEXT,
            corp_name TEXT,
            stock_code TEXT,
            corp_cls TEXT,
            report_nm TEXT,
            rcept_no TEXT UNIQUE,
            flr_nm TEXT,
            rcept_dt TEXT,
            rm TEXT
        );

        -- 문서 내용 테이블
        CREATE TABLE documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rcept_no TEXT,
            file_name TEXT,
            file_size INTEGER,
            xml_raw TEXT,
            body_text TEXT,
            FOREIGN KEY (rcept_no) REFERENCES disclosures(rcept_no)
        );

        -- XML 요약 메타데이터 테이블
        CREATE TABLE doc_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rcept_no TEXT,
            acode TEXT,
            value TEXT,
            FOREIGN KEY (rcept_no) REFERENCES disclosures(rcept_no)
        );

        -- 인덱스
        CREATE INDEX idx_disc_rcept_no ON disclosures(rcept_no);
        CREATE INDEX idx_disc_rcept_dt ON disclosures(rcept_dt);
        CREATE INDEX idx_disc_report_nm ON disclosures(report_nm);
        CREATE INDEX idx_doc_rcept_no ON documents(rcept_no);
        CREATE INDEX idx_summary_rcept_no ON doc_summary(rcept_no);
        CREATE INDEX idx_summary_acode ON doc_summary(acode);
    """)

    # ============ 공시 메타데이터 입력 ============
    print("공시 메타데이터 입력 중...")
    for disc in disclosures:
        cur.execute("""
            INSERT OR IGNORE INTO disclosures
            (corp_code, corp_name, stock_code, corp_cls, report_nm, rcept_no, flr_nm, rcept_dt, rm)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            disc['corp_code'], disc['corp_name'], disc['stock_code'],
            disc['corp_cls'], disc['report_nm'].strip(), disc['rcept_no'],
            disc['flr_nm'], disc['rcept_dt'], disc['rm']
        ))
    conn.commit()
    print(f"  {len(disclosures)}건 입력 완료")

    # ============ ZIP 파일 파싱 및 문서 내용 입력 ============
    print("문서 내용 추출 및 입력 중...")
    zip_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.zip')]
    success = 0
    fail = 0

    for i, zf_name in enumerate(sorted(zip_files)):
        zf_path = os.path.join(DOWNLOAD_DIR, zf_name)

        # Extract rcept_no from filename
        rcept_no_match = re.search(r'(\d{14})\.zip$', zf_name)
        if not rcept_no_match:
            fail += 1
            continue
        rcept_no = rcept_no_match.group(1)

        try:
            with zipfile.ZipFile(zf_path) as zf:
                for member in zf.namelist():
                    xml_bytes = zf.read(member)
                    xml_raw, body_text = extract_text_from_xml(xml_bytes)

                    # Insert document
                    cur.execute("""
                        INSERT INTO documents (rcept_no, file_name, file_size, xml_raw, body_text)
                        VALUES (?, ?, ?, ?, ?)
                    """, (rcept_no, member, len(xml_bytes), xml_raw, body_text))

                    # Extract and insert summary fields
                    summary = extract_summary_fields(xml_raw)
                    for acode, value in summary.items():
                        cur.execute("""
                            INSERT INTO doc_summary (rcept_no, acode, value)
                            VALUES (?, ?, ?)
                        """, (rcept_no, acode, value))

            success += 1
            if (i + 1) % 50 == 0:
                conn.commit()
                print(f"  [{i+1}/{len(zip_files)}] 처리 완료...")

        except Exception as e:
            fail += 1
            print(f"  [{i+1}/{len(zip_files)}] 오류: {zf_name} - {e}")

    conn.commit()

    # ============ FTS (전문 검색) 테이블 생성 ============
    print("전문 검색(FTS) 인덱스 생성 중...")
    cur.executescript("""
        CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts USING fts5(
            rcept_no,
            body_text,
            content='documents',
            content_rowid='id'
        );
        INSERT INTO documents_fts(rowid, rcept_no, body_text)
            SELECT id, rcept_no, body_text FROM documents;
    """)
    conn.commit()

    # ============ 통계 뷰 생성 ============
    cur.executescript("""
        -- 연도별 공시 수 뷰
        CREATE VIEW v_yearly_count AS
        SELECT substr(rcept_dt, 1, 4) AS year, COUNT(*) AS count
        FROM disclosures
        GROUP BY year ORDER BY year;

        -- 보고서 유형별 수 뷰
        CREATE VIEW v_report_type_count AS
        SELECT report_nm, COUNT(*) AS count
        FROM disclosures
        GROUP BY report_nm ORDER BY count DESC;

        -- 공시 + 문서 결합 뷰
        CREATE VIEW v_disclosure_docs AS
        SELECT
            d.rcept_no, d.report_nm, d.rcept_dt, d.flr_nm,
            doc.file_name, doc.file_size,
            LENGTH(doc.body_text) AS text_length
        FROM disclosures d
        LEFT JOIN documents doc ON d.rcept_no = doc.rcept_no;
    """)
    conn.commit()

    # ============ 결과 요약 ============
    cur.execute("SELECT COUNT(*) FROM disclosures")
    disc_count = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM documents")
    doc_count = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM doc_summary")
    summary_count = cur.fetchone()[0]

    db_size = os.path.getsize(DB_PATH) / 1024 / 1024

    print()
    print("=" * 60)
    print("DB 생성 완료!")
    print("=" * 60)
    print(f"  DB 파일: {DB_PATH}")
    print(f"  DB 크기: {db_size:.1f} MB")
    print(f"  공시 수: {disc_count}건")
    print(f"  문서 수: {doc_count}건")
    print(f"  요약 메타: {summary_count}건")
    print(f"  ZIP 성공: {success}건, 실패: {fail}건")
    print()

    # 연도별 통계
    print("--- 연도별 공시 수 ---")
    cur.execute("SELECT * FROM v_yearly_count")
    for row in cur.fetchall():
        print(f"  {row[0]}년: {row[1]}건")

    print()
    print("--- 보고서 유형 TOP 15 ---")
    cur.execute("SELECT * FROM v_report_type_count LIMIT 15")
    for row in cur.fetchall():
        print(f"  {row[0]}: {row[1]}건")

    conn.close()
    print()
    print("사용 예시:")
    print("  sqlite3 mcnex_dart.db")
    print("  SELECT * FROM disclosures WHERE report_nm LIKE '%사업보고서%';")
    print("  SELECT * FROM documents_fts WHERE body_text MATCH '매출';")


if __name__ == "__main__":
    main()
