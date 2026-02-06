import sqlite3
import requests
import json
import time
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

API_KEY = "fee81b8f1226ef15d145dbfa04d0569e34ac1656"
CORP_CODE = "00564562"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "mcnex_full.db")

# 보고서 코드
REPRT_CODES = {
    "11011": "사업보고서",
    "11012": "반기보고서",
    "11013": "1분기보고서",
    "11014": "3분기보고서",
}

# 사업연도 범위
YEARS = list(range(2007, 2026))

def api_call(endpoint, params, retries=2):
    """OpenDART API 호출"""
    url = f"https://opendart.fss.or.kr/api/{endpoint}"
    params["crtfc_key"] = API_KEY
    for attempt in range(retries + 1):
        try:
            resp = requests.get(url, params=params, timeout=30)
            data = resp.json()
            if data.get("status") == "000":
                return data.get("list", [])
            elif data.get("status") == "013":  # 조회된 데이터가 없음
                return []
            else:
                return []
        except Exception as e:
            if attempt < retries:
                time.sleep(2)
            else:
                return []
    return []


def create_tables(conn):
    conn.executescript("""
        -- ==============================
        -- 1. 기업개황
        -- ==============================
        CREATE TABLE IF NOT EXISTS company_info (
            corp_code TEXT, corp_name TEXT, corp_name_eng TEXT, stock_name TEXT,
            stock_code TEXT, ceo_nm TEXT, corp_cls TEXT, jurir_no TEXT,
            bizr_no TEXT, adres TEXT, hm_url TEXT, ir_url TEXT,
            phn_no TEXT, fax_no TEXT, induty_code TEXT, est_dt TEXT,
            acc_mt TEXT, fetched_at TEXT
        );

        -- ==============================
        -- 2. 재무제표 (전체 계정)
        -- ==============================
        CREATE TABLE IF NOT EXISTS financial_statements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            reprt_nm TEXT,
            rcept_no TEXT,
            sj_div TEXT,      -- 재무제표구분 (BS/IS/CIS/CF/SCE)
            sj_nm TEXT,       -- 재무제표명
            account_id TEXT,  -- 계정ID
            account_nm TEXT,  -- 계정명
            account_detail TEXT,
            thstrm_nm TEXT,   -- 당기명
            thstrm_amount TEXT, -- 당기금액
            frmtrm_nm TEXT,   -- 전기명
            frmtrm_amount TEXT, -- 전기금액
            bfefrmtrm_nm TEXT,  -- 전전기명
            bfefrmtrm_amount TEXT, -- 전전기금액
            ord TEXT,         -- 계정과목 정렬순서
            currency TEXT
        );

        -- ==============================
        -- 3. 주요계정 (요약)
        -- ==============================
        CREATE TABLE IF NOT EXISTS financial_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            reprt_nm TEXT,
            rcept_no TEXT,
            account_nm TEXT,
            thstrm_nm TEXT,
            thstrm_amount TEXT,
            frmtrm_nm TEXT,
            frmtrm_amount TEXT,
            bfefrmtrm_nm TEXT,
            bfefrmtrm_amount TEXT,
            ord TEXT
        );

        -- ==============================
        -- 4. 임원현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS executives (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            nm TEXT,           -- 성명
            sexdstn TEXT,      -- 성별
            birth_ym TEXT,     -- 출생년월
            ofcps TEXT,        -- 직위
            rgist_exctv_at TEXT, -- 등기임원여부
            fte_at TEXT,       -- 상근여부
            chrg_job TEXT,     -- 담당업무
            mxmm_shrholdr_relate TEXT, -- 최대주주와의 관계
            hffc_pd TEXT,      -- 재직기간
            tenure_end_on TEXT -- 임기만료일
        );

        -- ==============================
        -- 5. 직원현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            fo_bbm TEXT,       -- 사업부문
            sexdstn TEXT,      -- 성별
            reform_bfe_emp_co_rgllbr TEXT, -- 정규직
            reform_bfe_emp_co_cnttk TEXT,  -- 계약직
            reform_bfe_emp_co_etc TEXT,
            rgllbr_co TEXT,
            rgllbr_abacpt_labrr_co TEXT,
            cnttk_co TEXT,
            cnttk_abacpt_labrr_co TEXT,
            sm TEXT,           -- 합계
            avrg_cnwk_sdytrn TEXT, -- 평균근속연수
            fyer_salary_totamt TEXT, -- 연간급여총액
            jan_salary_am TEXT -- 1인평균급여액
        );

        -- ==============================
        -- 6. 대주주 현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS major_shareholders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            nm TEXT,           -- 성명
            relate TEXT,       -- 관계
            stock_knd TEXT,    -- 주식종류
            bsis_posesn_stock_co TEXT,  -- 기초소유주식수
            bsis_posesn_stock_qota_rt TEXT, -- 기초지분율
            trmend_posesn_stock_co TEXT,    -- 기말소유주식수
            trmend_posesn_stock_qota_rt TEXT, -- 기말지분율
            rm TEXT
        );

        -- ==============================
        -- 7. 소액주주 현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS minority_shareholders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            se TEXT,
            shrholdr_co TEXT,
            shrholdr_tot_co TEXT,
            shrholdr_rate TEXT,
            hold_stock_co TEXT,
            stock_tot_co TEXT,
            hold_stock_rate TEXT
        );

        -- ==============================
        -- 8. 배당 현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS dividends (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            se TEXT,           -- 구분
            stock_knd TEXT,    -- 주식종류
            thstrm TEXT,       -- 당기
            frmtrm TEXT,       -- 전기
            lwfr TEXT          -- 전전기
        );

        -- ==============================
        -- 9. 자기주식 취득/처분
        -- ==============================
        CREATE TABLE IF NOT EXISTS treasury_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            stock_knd TEXT,
            acqs_mth1 TEXT,
            acqs_mth2 TEXT,
            acqs_mth3 TEXT,
            bsis_qy TEXT,
            change_qy_acqs TEXT,
            change_qy_dsps TEXT,
            change_qy_incnr TEXT,
            trmend_qy TEXT,
            rm TEXT
        );

        -- ==============================
        -- 10. 증자/감자 현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS capital_changes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            isu_dcrs_de TEXT,
            isu_dcrs_stle TEXT,
            isu_dcrs_stock_knd TEXT,
            isu_dcrs_qy TEXT,
            isu_dcrs_mstvdv_fval_amount TEXT,
            isu_dcrs_mstvdv_amount TEXT,
            rm TEXT
        );

        -- ==============================
        -- 11. 주식 총수 현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS stock_total (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            se TEXT,
            isu_stock_totqy TEXT,
            now_to_isu_stock_totqy TEXT,
            now_to_dcrs_stock_totqy TEXT,
            redc TEXT,
            rdmstkdiv TEXT,
            istc_totqy TEXT,
            tesstk_co TEXT,
            distb_stock_co TEXT
        );

        -- ==============================
        -- 12. 타법인 출자 현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS investments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            inv_prm TEXT,      -- 법인명
            frst_acqs_de TEXT, -- 최초취득일
            invstmnt_purps TEXT, -- 투자목적
            frst_acqs_amount TEXT,
            bsis_blce_qy TEXT,
            bsis_blce_qota_rt TEXT,
            bsis_blce_acntbk_amount TEXT,
            incrs_dcrs_acqs_dsps_qy TEXT,
            incrs_dcrs_acqs_dsps_amount TEXT,
            incrs_dcrs_evl_lstmn TEXT,
            trmend_blce_qy TEXT,
            trmend_blce_qota_rt TEXT,
            trmend_blce_acntbk_amount TEXT,
            recent_bsns_year_fnnr_sttus_tot_assets TEXT,
            recent_bsns_year_fnnr_sttus_thstrm_ntpf TEXT
        );

        -- ==============================
        -- 13. 이사·감사 보수현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS exec_compensation (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            se TEXT,
            nmpr TEXT,
            mendng_totamt TEXT,
            jan_avrg_mendng_am TEXT
        );

        -- ==============================
        -- 14. 개인별 보수 (5억 이상)
        -- ==============================
        CREATE TABLE IF NOT EXISTS individual_pay (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            nm TEXT,
            ofcps TEXT,
            mendng_totamt TEXT,
            mendng_totamt_ct_incls_mendng TEXT
        );

        -- ==============================
        -- 15. 사외이사 현황
        -- ==============================
        CREATE TABLE IF NOT EXISTS outside_directors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bsns_year TEXT,
            reprt_code TEXT,
            rcept_no TEXT,
            nm TEXT,
            main_career TEXT,
            maxholder_relate TEXT,
            apntmt_dt TEXT,
            enddt TEXT,
            rmndt TEXT
        );

        -- 인덱스
        CREATE INDEX IF NOT EXISTS idx_fs_year ON financial_statements(bsns_year);
        CREATE INDEX IF NOT EXISTS idx_fs_sj ON financial_statements(sj_div);
        CREATE INDEX IF NOT EXISTS idx_fs_account ON financial_statements(account_nm);
        CREATE INDEX IF NOT EXISTS idx_fsum_year ON financial_summary(bsns_year);
        CREATE INDEX IF NOT EXISTS idx_exec_year ON executives(bsns_year);
        CREATE INDEX IF NOT EXISTS idx_ms_year ON major_shareholders(bsns_year);
        CREATE INDEX IF NOT EXISTS idx_div_year ON dividends(bsns_year);
        CREATE INDEX IF NOT EXISTS idx_inv_year ON investments(bsns_year);
    """)
    conn.commit()


def insert_rows(conn, table, rows, extra_fields=None):
    """API 결과를 테이블에 삽입"""
    if not rows:
        return 0
    count = 0
    for row in rows:
        if extra_fields:
            row.update(extra_fields)
        cols = [c for c in row.keys() if c not in ('status', 'message')]
        # 테이블 컬럼 정보 조회
        cur = conn.execute(f"PRAGMA table_info({table})")
        table_cols = {r[1] for r in cur.fetchall()}
        # 테이블에 있는 컬럼만 사용
        valid_cols = [c for c in cols if c in table_cols]
        if not valid_cols:
            continue
        placeholders = ','.join(['?'] * len(valid_cols))
        col_names = ','.join(valid_cols)
        values = [row.get(c, '') for c in valid_cols]
        try:
            conn.execute(f"INSERT INTO {table} ({col_names}) VALUES ({placeholders})", values)
            count += 1
        except Exception as e:
            pass
    conn.commit()
    return count


# API 엔드포인트 정의
API_ENDPOINTS = [
    # (엔드포인트, 테이블명, 설명, 사업보고서만 여부)
    ("fnlttSinglAcntAll.json", "financial_statements", "전체 재무제표", False),
    ("fnlttSinglAcnt.json", "financial_summary", "주요계정", False),
    ("hyslrSttus.json", "executives", "임원현황", True),
    ("hyslrChgSttus.json", "employees", "직원현황", True),
    ("lvlhldSttus.json", "major_shareholders", "대주주현황", True),
    ("mrhlSttus.json", "minority_shareholders", "소액주주현황", True),
    ("alotMatter.json", "dividends", "배당현황", True),
    ("tesstkAcqsDspsSttus.json", "treasury_stock", "자기주식현황", True),
    ("irdsSttus.json", "capital_changes", "증자감자현황", True),
    ("stockTotqySttus.json", "stock_total", "주식총수현황", True),
    ("otrCprInvstmntSttus.json", "investments", "타법인출자현황", True),
    ("drctrAdtAllMendngSttus.json", "exec_compensation", "이사감사보수현황", True),
    ("indvdlByPay.json", "individual_pay", "개인별보수현황", True),
    ("outcmpnyDrctrNdAudtSttus.json", "outside_directors", "사외이사현황", True),
]


def main():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    create_tables(conn)

    # ============ 1. 기업개황 ============
    print("=" * 60)
    print("1. 기업개황 조회")
    print("=" * 60)
    url = f"https://opendart.fss.or.kr/api/company.json"
    params = {"crtfc_key": API_KEY, "corp_code": CORP_CODE}
    resp = requests.get(url, params=params, timeout=30)
    info = resp.json()
    if info.get("status") == "000":
        cols = [c for c in info.keys() if c not in ('status', 'message')]
        info['fetched_at'] = time.strftime('%Y-%m-%d')
        cols.append('fetched_at')
        placeholders = ','.join(['?'] * len(cols))
        col_names = ','.join(cols)
        values = [info.get(c, '') for c in cols]
        conn.execute(f"INSERT INTO company_info ({col_names}) VALUES ({placeholders})", values)
        conn.commit()
        print(f"  회사명: {info.get('corp_name')}")
        print(f"  대표이사: {info.get('ceo_nm')}")
        print(f"  주소: {info.get('adres')}")
        print(f"  업종코드: {info.get('induty_code')}")
        print(f"  설립일: {info.get('est_dt')}")
    time.sleep(0.5)

    # ============ 2~15. 각종 구조화 데이터 수집 ============
    for endpoint, table, desc, annual_only in API_ENDPOINTS:
        print()
        print("=" * 60)
        print(f"수집 중: {desc} ({endpoint})")
        print("=" * 60)

        total = 0
        for year in YEARS:
            codes = {"11011": "사업보고서"} if annual_only else REPRT_CODES
            for reprt_code, reprt_nm in codes.items():
                params = {
                    "corp_code": CORP_CODE,
                    "bsns_year": str(year),
                    "reprt_code": reprt_code,
                }
                # fnlttSinglAcntAll needs fs_div
                if endpoint == "fnlttSinglAcntAll.json":
                    for fs_div in ["OFS", "CFS"]:  # 개별/연결
                        params["fs_div"] = fs_div
                        rows = api_call(endpoint, params)
                        extra = {"bsns_year": str(year), "reprt_code": reprt_code, "reprt_nm": f"{reprt_nm}({'연결' if fs_div=='CFS' else '개별'})"}
                        cnt = insert_rows(conn, table, rows, extra)
                        total += cnt
                        time.sleep(0.3)
                elif endpoint == "fnlttSinglAcnt.json":
                    for fs_div in ["OFS", "CFS"]:
                        params["fs_div"] = fs_div
                        rows = api_call(endpoint, params)
                        extra = {"bsns_year": str(year), "reprt_code": reprt_code, "reprt_nm": f"{reprt_nm}({'연결' if fs_div=='CFS' else '개별'})"}
                        cnt = insert_rows(conn, table, rows, extra)
                        total += cnt
                        time.sleep(0.3)
                else:
                    rows = api_call(endpoint, params)
                    extra = {"bsns_year": str(year), "reprt_code": reprt_code}
                    cnt = insert_rows(conn, table, rows, extra)
                    total += cnt
                    time.sleep(0.3)

            if year % 5 == 0:
                print(f"  ~ {year}년까지 누적 {total}건")

        print(f"  >> {desc} 총 {total}건 저장")

    # ============ 기존 공시 목록도 통합 ============
    print()
    print("=" * 60)
    print("기존 공시 목록 통합")
    print("=" * 60)
    list_path = os.path.join(BASE_DIR, "disclosure_list.json")
    if os.path.exists(list_path):
        conn.execute("""
            CREATE TABLE IF NOT EXISTS disclosures (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                corp_code TEXT, corp_name TEXT, stock_code TEXT,
                corp_cls TEXT, report_nm TEXT, rcept_no TEXT UNIQUE,
                flr_nm TEXT, rcept_dt TEXT, rm TEXT
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_disc_dt ON disclosures(rcept_dt)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_disc_rn ON disclosures(report_nm)")
        with open(list_path, "r", encoding="utf-8") as f:
            discs = json.load(f)
        for d in discs:
            try:
                conn.execute("""INSERT OR IGNORE INTO disclosures
                    (corp_code, corp_name, stock_code, corp_cls, report_nm, rcept_no, flr_nm, rcept_dt, rm)
                    VALUES (?,?,?,?,?,?,?,?,?)""",
                    (d['corp_code'], d['corp_name'], d['stock_code'], d['corp_cls'],
                     d['report_nm'].strip(), d['rcept_no'], d['flr_nm'], d['rcept_dt'], d['rm']))
            except:
                pass
        conn.commit()
        print(f"  {len(discs)}건 공시목록 입력")

    # ============ 편의 뷰 생성 ============
    print()
    print("편의 뷰 생성 중...")
    conn.executescript("""
        -- 연도별 매출/영업이익/당기순이익 (연결)
        CREATE VIEW IF NOT EXISTS v_annual_performance AS
        SELECT
            bsns_year,
            MAX(CASE WHEN account_nm LIKE '%매출액%' AND sj_div='IS' THEN thstrm_amount END) AS revenue,
            MAX(CASE WHEN account_nm LIKE '%영업이익%' AND sj_div='IS' THEN thstrm_amount END) AS operating_profit,
            MAX(CASE WHEN account_nm LIKE '%당기순이익%' AND sj_div='IS' THEN thstrm_amount END) AS net_income,
            MAX(CASE WHEN account_nm LIKE '%자산총계%' AND sj_div='BS' THEN thstrm_amount END) AS total_assets,
            MAX(CASE WHEN account_nm LIKE '%부채총계%' AND sj_div='BS' THEN thstrm_amount END) AS total_liabilities,
            MAX(CASE WHEN account_nm LIKE '%자본총계%' AND sj_div='BS' THEN thstrm_amount END) AS total_equity
        FROM financial_statements
        WHERE reprt_code = '11011' AND reprt_nm LIKE '%연결%'
        GROUP BY bsns_year
        ORDER BY bsns_year;

        -- 연도별 배당 현황
        CREATE VIEW IF NOT EXISTS v_annual_dividends AS
        SELECT bsns_year, se, stock_knd, thstrm, frmtrm, lwfr
        FROM dividends
        ORDER BY bsns_year;

        -- 최대주주 변동 이력
        CREATE VIEW IF NOT EXISTS v_major_shareholder_history AS
        SELECT bsns_year, nm, relate, stock_knd,
               bsis_posesn_stock_co, bsis_posesn_stock_qota_rt,
               trmend_posesn_stock_co, trmend_posesn_stock_qota_rt
        FROM major_shareholders
        ORDER BY bsns_year, trmend_posesn_stock_qota_rt DESC;

        -- 임원 변동 이력
        CREATE VIEW IF NOT EXISTS v_executive_history AS
        SELECT bsns_year, nm, ofcps, chrg_job, rgist_exctv_at,
               fte_at, mxmm_shrholdr_relate, hffc_pd, tenure_end_on
        FROM executives
        ORDER BY bsns_year, ofcps;

        -- 공시 타임라인
        CREATE VIEW IF NOT EXISTS v_disclosure_timeline AS
        SELECT rcept_dt, report_nm, flr_nm, rcept_no
        FROM disclosures
        ORDER BY rcept_dt DESC;

        -- DB 테이블별 건수 요약
        CREATE VIEW IF NOT EXISTS v_db_summary AS
        SELECT 'company_info' as tbl, COUNT(*) as cnt FROM company_info
        UNION ALL SELECT 'disclosures', COUNT(*) FROM disclosures
        UNION ALL SELECT 'financial_statements', COUNT(*) FROM financial_statements
        UNION ALL SELECT 'financial_summary', COUNT(*) FROM financial_summary
        UNION ALL SELECT 'executives', COUNT(*) FROM executives
        UNION ALL SELECT 'employees', COUNT(*) FROM employees
        UNION ALL SELECT 'major_shareholders', COUNT(*) FROM major_shareholders
        UNION ALL SELECT 'minority_shareholders', COUNT(*) FROM minority_shareholders
        UNION ALL SELECT 'dividends', COUNT(*) FROM dividends
        UNION ALL SELECT 'treasury_stock', COUNT(*) FROM treasury_stock
        UNION ALL SELECT 'capital_changes', COUNT(*) FROM capital_changes
        UNION ALL SELECT 'stock_total', COUNT(*) FROM stock_total
        UNION ALL SELECT 'investments', COUNT(*) FROM investments
        UNION ALL SELECT 'exec_compensation', COUNT(*) FROM exec_compensation
        UNION ALL SELECT 'individual_pay', COUNT(*) FROM individual_pay
        UNION ALL SELECT 'outside_directors', COUNT(*) FROM outside_directors;
    """)
    conn.commit()

    # ============ 최종 요약 ============
    print()
    print("=" * 60)
    print("DB 생성 완료!")
    print("=" * 60)
    print(f"  파일: {DB_PATH}")
    db_size = os.path.getsize(DB_PATH) / 1024 / 1024
    print(f"  크기: {db_size:.1f} MB")
    print()
    print("--- 테이블별 데이터 건수 ---")
    rows = conn.execute("SELECT * FROM v_db_summary").fetchall()
    for r in rows:
        print(f"  {r[0]}: {r[1]}건")

    print()
    print("--- 연도별 실적 (연결) ---")
    rows = conn.execute("SELECT * FROM v_annual_performance").fetchall()
    if rows:
        print(f"  {'연도':>6} | {'매출액':>15} | {'영업이익':>15} | {'당기순이익':>15}")
        print("  " + "-" * 65)
        for r in rows:
            rev = f"{int(r[1]):,}" if r[1] else "-"
            op = f"{int(r[2]):,}" if r[2] else "-"
            ni = f"{int(r[3]):,}" if r[3] else "-"
            print(f"  {r[0]:>6} | {rev:>15} | {op:>15} | {ni:>15}")

    conn.close()


if __name__ == "__main__":
    main()
