# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

OpenDART API에서 한국 상장기업의 공시 데이터를 수집하고, SQLite DB로 정제한 뒤, openpyxl로 엑셀 보고서를 생성하는 범용 기업분석 파이프라인.

종목코드만 입력하면 **어떤 기업이든** 동일한 방법으로 데이터 수집 → DB 구축 → AI 분석용 DB 생성이 가능.

## 분석 완료 기업

| 종목코드 | 기업명 | 업종 | 보고서 수 | 비고 |
|----------|--------|------|----------|------|
| 097520 | 엠씨넥스 | 카메라모듈 부품 | 5종 | 최초 구축 |
| 035250 | 강원랜드 | 카지노/리조트 | 6종 | 규제독점, 무차입경영 |

## 폴더 구조

```
mcnex-analysis/
  config.py              # API 키 + 공용 유틸리티 (get_company_dir, ensure_company_dir)
  run_pipeline.py        # 전체 파이프라인 한번에 실행
  download_all.py        # 1단계: 공시 다운로드
  build_db.py            # 2단계: 원문 DB 생성
  build_full_db.py       # 3단계: 구조화 DB 생성
  build_ai_db.py         # 4단계: 통합 AI DB 생성
  Method/                # 분석 프레임워크 문서 (Korean_Guru_Framework.md 포함)
  companies/
    097520_엠씨넥스/     # 종목코드_회사명 폴더
      create_report.py        # → 엠씨넥스_기업분석보고서.xlsx (9시트)
      create_valuation.py     # → 엠씨넥스_밸류에이션.xlsx (5시트)
      create_combined.py      # → 엠씨넥스_종합보고서.xlsx (12시트)
      create_mobile.py        # → 엠씨넥스_모바일용.xlsx (단일시트)
      create_guru_report.py   # → 엠씨넥스_투자구루분석.xlsx (7시트)
    035250_강원랜드/
      create_report.py        # → 강원랜드_기업분석보고서.xlsx (9시트)
      create_valuation.py     # → 강원랜드_밸류에이션.xlsx (5시트)
      create_combined.py      # → 강원랜드_종합보고서.xlsx (12시트)
      create_mobile.py        # → 강원랜드_모바일용.xlsx (단일시트)
      create_guru_report.py   # → 강원랜드_투자구루분석.xlsx (7시트)
      create_profit_analysis.py # → 강원랜드_이익역성장분석.xlsx (7시트)
```

각 회사 폴더에는 파이프라인 실행 시 자동 생성되는 파일도 포함:
- `company_info.json`, `disclosure_list.json` (메타데이터)
- `downloads/` (공시 ZIP 파일)
- `dart.db` (원문 DB, FTS), `full.db` (구조화 DB), `ai.db` (통합 분석용 DB)
- `*.xlsx` (보고서 출력)

## 실행 명령어

모든 스크립트는 Python 3. 의존성은 `requests`(데이터 수집)와 `openpyxl`(보고서 생성).

**Windows 인코딩 주의**: `run_pipeline.py`로 실행 시 UnicodeEncodeError 발생 가능. 아래처럼 개별 실행 권장:
```bash
set PYTHONIOENCODING=utf-8 && python -X utf8 download_all.py 035250
```

### 전체 파이프라인 (종목코드만 입력)

```bash
# 새 기업 분석 (전체 4단계 자동 실행)
python run_pipeline.py 097520    # 엠씨넥스
python run_pipeline.py 035250    # 강원랜드

# 개별 단계 실행 (Windows 인코딩 문제 시)
set PYTHONIOENCODING=utf-8 && python -X utf8 download_all.py 035250
set PYTHONIOENCODING=utf-8 && python -X utf8 build_db.py 035250
set PYTHONIOENCODING=utf-8 && python -X utf8 build_full_db.py 035250
set PYTHONIOENCODING=utf-8 && python -X utf8 build_ai_db.py 035250
```

### 보고서 생성 (회사 폴더에서 실행)

```bash
# 엠씨넥스
cd companies/097520_엠씨넥스
python create_report.py           # → 기업분석보고서 (9시트)
python create_valuation.py        # → 밸류에이션 (5시트, 독립형)
python create_combined.py         # → 종합보고서 (12시트)
python create_mobile.py           # → 모바일용 (단일시트, 3열 세로)
python create_guru_report.py      # → 투자구루분석 (7시트, Buffett/Munger)

# 강원랜드
cd companies/035250_강원랜드
set PYTHONIOENCODING=utf-8 && python -X utf8 create_report.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_valuation.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_combined.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_mobile.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_guru_report.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_profit_analysis.py  # 이익역성장 분석 (강원랜드 전용)
```

## 아키텍처

### 데이터 흐름

```
OpenDART API ─→ download_all.py ─→ companies/{code}_{name}/downloads/*.zip
                                  + disclosure_list.json + company_info.json
                                          │
                                    build_db.py ─→ dart.db (원문, FTS)
                                          │
OpenDART API ─→ build_full_db.py ─→ full.db (구조화 정량)
                                          │
                           dart.db ──┐    │
                                     ▼    ▼
                                build_ai_db.py ─→ ai.db (통합 분석용)
                                                      │
                            ┌──────────┬──────────┬───┘
                            ▼          ▼          ▼
                     create_report  create_combined  create_mobile
                     create_valuation  create_guru_report
                     create_profit_analysis (강원랜드 전용)
```

### 스크립트 역할

| 스크립트 | 입력 | 출력 | 핵심 로직 |
|----------|------|------|-----------|
| `config.py` | - | - | API_KEY, get_company_dir(), ensure_company_dir() |
| `run_pipeline.py` | 종목코드 | - | subprocess로 4단계 순차 실행 |
| `download_all.py` | 종목코드 | `downloads/`, `disclosure_list.json`, `company_info.json` | corpCode.xml에서 고유번호 조회 → 공시목록 페이징 → 문서 ZIP 다운로드 |
| `build_db.py` | 종목코드 | `dart.db` | ZIP 내 XML을 멀티인코딩으로 파싱, BODY 텍스트 추출, FTS5 인덱스 |
| `build_full_db.py` | 종목코드 | `full.db` | 15종 구조화 API를 연도×보고서구분 조합으로 호출 |
| `build_ai_db.py` | 종목코드 | `ai.db` | full_db + dart_db 통합, 사업보고서 8개 섹션 분리, 특허/이벤트/잠정실적 패턴매칭 |
| `create_*.py` | `ai.db` | xlsx | 회사별 보고서 생성 (회사 폴더 내 위치) |

### 보고서 종류 (6종)

| 보고서 | 시트 수 | 설명 |
|--------|---------|------|
| `create_report.py` | 9시트 | 기업분석보고서 (표지~모니터링) |
| `create_valuation.py` | 5시트 | 밸류에이션 (PER/PBR/EV_EBITDA/RIM/시나리오) |
| `create_combined.py` | 12시트 | 종합보고서 (report + valuation 통합) |
| `create_mobile.py` | 1시트 | 모바일용 (3열 세로, 큰 폰트) |
| `create_guru_report.py` | 7시트 | 투자구루분석 (Buffett/Munger Four Filters) |
| `create_profit_analysis.py` | 7시트 | 이익역성장분석 (강원랜드 전용, 비용구조·일시적요인·인건비) |

### 보고서 스크립트 패턴

모든 보고서 스크립트가 동일한 패턴:
1. openpyxl 스타일 상수 정의 (NAVY, DARK_BLUE 등 색상 + Font/Fill/Alignment/Border 객체)
2. 헬퍼 함수: `sw()`, `wh()`, `wr()`, `st()`, `fmt()`, `fw()`, `pct()`
3. 시트별 순차 생성: 데이터 배열 → 행 단위 write → 수식 삽입
4. 대부분의 보고서는 재무 데이터가 파이썬 상수로 하드코딩 (DB 의존 최소화)

### DB 스키마 (ai.db)

분석 시 `ai.db`만 사용하면 됨. 모든 기업 공통 스키마.

**정량 테이블**: `company_info`, `disclosures`, `financial_statements`, `financial_summary`, `executives`, `employees`, `dividends`, `treasury_stock`, `capital_changes`, `stock_total`, `investments`, `minority_shareholders`, `individual_pay`

**정성 테이블**: `business_report_sections` (연도×8섹션), `patents`, `key_events`, `earnings_announcements`

**편의 뷰**: `v_annual_performance`, `v_annual_dividends`, `v_major_shareholder_history`, `v_executive_history`, `v_disclosure_timeline`, `v_business_sections`, `v_patent_history`, `v_event_timeline`, `v_db_summary`

### 주요 컨벤션

- 재무제표 금액 단위: **원** (억원 변환 시 ÷ 100,000,000)
- `reprt_code`: 11011=사업보고서, 11012=반기, 11013=1분기, 11014=3분기
- `sj_div`: BS=재무상태표, CIS=포괄손익계산서, CF=현금흐름표
- 연결 재무제표: `reprt_nm LIKE '%연결%'` / 개별: `reprt_nm LIKE '%개별%'`
- `business_report_sections.section_name` 값: 회사개요, 회사연혁, 사업내용, 주요제품_매출, 연구개발, 위험관리_전망, 임원_보수, 주주_배당
- API 키는 `config.py`에 중앙 관리
- 회사 폴더: `companies/{종목코드}_{회사명}/` 형식으로 자동 생성
- 데이터 출처: OpenDART API (https://opendart.fss.or.kr)

### 기업별 데이터 특이사항

**강원랜드 (035250)**:
- 매출 계정명이 연도별로 다름: "매출" (2015-2018) → "수익(매출액)" (2019+)
- `v_annual_performance` 뷰가 잘못된 매출값 반환 → 직접 쿼리 필요
- 2021-2024 연결 영업이익이 사업보고서(11011) financial_statements에 누락 → 잠정실적으로 보완
- employees 테이블 데이터가 모두 NULL (API에서 미반환)
- 개별 보수 공시 미해당 (CEO 보수 비공개)
- 특허 없음 (카지노업)
- treasury_stock의 stock_knd가 연도별로 '보통주식'/'보통주'/None 혼재

## 자주 쓰는 쿼리

```sql
-- 연도별 핵심 실적 (연결) — 강원랜드는 이 뷰가 부정확할 수 있음
SELECT * FROM v_annual_performance;

-- 강원랜드 매출액 직접 쿼리 (계정명 변경 대응)
SELECT bsns_year, thstrm_amount FROM financial_statements
WHERE (account_nm='수익(매출액)' OR account_nm='매출')
AND sj_div='CIS' AND reprt_code='11011' AND reprt_nm LIKE '%연결%'
ORDER BY bsns_year;

-- 특정 연도 사업보고서 섹션 읽기
SELECT section_text FROM business_report_sections
WHERE bsns_year = '2024' AND section_name = '사업내용';

-- 특허 이력
SELECT rcept_dt, patent_name, patent_detail, patent_plan FROM patents ORDER BY rcept_dt;

-- 주요 이벤트 타임라인
SELECT rcept_dt, event_type, event_summary FROM key_events ORDER BY rcept_dt DESC;

-- 잠정실적 공시
SELECT rcept_dt, report_nm, content FROM earnings_announcements ORDER BY rcept_dt DESC;
```
