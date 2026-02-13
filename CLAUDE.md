# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

OpenDART API에서 한국 상장기업의 공시 데이터를 수집하고, SQLite DB로 정제한 뒤, openpyxl로 엑셀 보고서를 생성하는 범용 기업분석 파이프라인.

종목코드만 입력하면 **어떤 기업이든** 동일한 방법으로 데이터 수집 → DB 구축 → AI 분석용 DB 생성이 가능.

## 분석 완료 기업

| 종목코드 | 기업명 | 업종 | 보고서 수 | 비고 |
|----------|--------|------|----------|------|
| 097520 | 엠씨넥스 | 카메라모듈 부품 | 6종 | 최초 구축 |
| 035250 | 강원랜드 | 카지노/리조트 | 7종 | 규제독점, 무차입경영 |
| 003490 | 대한항공 | 항공운송 | 7종 | 아시아나 합병, 고레버리지 |

## 의존성

요구사항 파일 없음. 임포트에서 추론:
- `requests` — OpenDART API 호출
- `openpyxl` — 엑셀 보고서 생성
- `pywin32` (`win32com`) — `export_pdf.py` 전용 (Excel COM 자동화, Windows 한정)
- Python 3 표준 라이브러리: `sqlite3`, `json`, `xml.etree.ElementTree`, `zipfile`, `subprocess`

## 폴더 구조

```
mcnex-analysis/
  config.py              # API 키 + 공용 유틸리티 (get_company_dir, ensure_company_dir)
  run_pipeline.py        # 전체 파이프라인 한번에 실행
  download_all.py        # 1단계: 공시 다운로드
  build_db.py            # 2단계: 원문 DB 생성 (FTS5)
  build_full_db.py       # 3단계: 구조화 DB 생성 (15종 API)
  build_ai_db.py         # 4단계: 통합 AI DB 생성
  export_pdf.py          # xlsx → PDF 변환 (Excel COM, Windows 전용)
  Method/                # 분석 프레임워크 문서
    Korean_Guru_Framework.md     # Buffett/Munger Four Filters 한국시장 적용
    Corporate_Analysis_Framework # 종합 투자분석 프레임워크
    AI_Prompt_Templates          # LLM 분석용 프롬프트 템플릿
  companies/
    {종목코드}_{회사명}/
      ai.db, dart.db, full.db   # 파이프라인 생성 DB (gitignored)
      company_info.json          # 메타데이터 (gitignored)
      disclosure_list.json       # 공시목록 (gitignored)
      downloads/                 # 공시 ZIP 파일 (gitignored)
      create_report.py           # → 기업분석보고서 (9시트)
      create_valuation.py        # → 밸류에이션 (5시트)
      create_combined.py         # → 종합보고서 (12시트)
      create_mobile.py           # → 모바일용 (1시트)
      create_guru_report.py      # → 투자구루분석 (7시트)
      create_master.py           # → 마스터보고서 (통합, 종합+구루+α)
      create_profit_analysis.py  # → 이익역성장분석 (강원랜드 전용)
      create_debt_analysis.py    # → 항공업구조분석 (대한항공 전용)
      create_segment_analysis.py # → 사업부문별수익분석 (대한항공 전용)
```

**Git LFS**: `*.db` 파일과 `downloads/*.zip`은 `.gitattributes`에서 Git LFS로 추적.
**gitignored**: `*.db`, `*.xlsx`, `downloads/`, `disclosure_list.json`, `company_info.json`

## 실행 명령어

**Windows 인코딩 주의**: `run_pipeline.py`로 실행 시 UnicodeEncodeError 발생 가능. 개별 실행 권장:
```bash
set PYTHONIOENCODING=utf-8 && python -X utf8 <script>.py <종목코드>
```

### 전체 파이프라인 (종목코드만 입력)

```bash
# 새 기업 분석 (전체 4단계 자동 실행)
python run_pipeline.py 097520

# 개별 단계 실행 (Windows 인코딩 문제 시)
set PYTHONIOENCODING=utf-8 && python -X utf8 download_all.py 035250
set PYTHONIOENCODING=utf-8 && python -X utf8 build_db.py 035250
set PYTHONIOENCODING=utf-8 && python -X utf8 build_full_db.py 035250
set PYTHONIOENCODING=utf-8 && python -X utf8 build_ai_db.py 035250
```

### 보고서 생성 (회사 폴더에서 실행)

```bash
cd companies\035250_강원랜드
set PYTHONIOENCODING=utf-8 && python -X utf8 create_report.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_valuation.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_combined.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_mobile.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_guru_report.py
set PYTHONIOENCODING=utf-8 && python -X utf8 create_master.py
```

### PDF 변환 (프로젝트 루트에서, Windows 전용)

```bash
python export_pdf.py                                    # 기본 대상 모두 변환
python export_pdf.py companies\097520_엠씨넥스\엠씨넥스_마스터보고서.xlsx  # 특정 파일
```

## 아키텍처

### 데이터 흐름

```
OpenDART API ─→ download_all.py ─→ companies/{code}_{name}/downloads/*.zip
                                  + disclosure_list.json + company_info.json
                                          │
                                    build_db.py ─→ dart.db (원문, FTS)
                                          │
OpenDART API ─→ build_full_db.py ─→ full.db (구조화 정량, 15종 API)
                                          │
                           dart.db ──┐    │
                                     ▼    ▼
                                build_ai_db.py ─→ ai.db (통합 분석용)
                                                      │
                            ┌──────────┬──────────┬───┘
                            ▼          ▼          ▼
                     create_report  create_combined  create_mobile
                     create_valuation  create_guru_report  create_profit_analysis
                            │          │          │
                            └──────┬───┘──────────┘
                                   ▼
                            create_master.py ─→ 마스터보고서.xlsx
                                   │
                            export_pdf.py ─→ 마스터보고서.pdf
```

### 스크립트 역할

| 스크립트 | 입력 | 출력 | 핵심 로직 |
|----------|------|------|-----------|
| `config.py` | - | - | API_KEY, get_company_dir(), ensure_company_dir() |
| `run_pipeline.py` | 종목코드 | - | subprocess로 4단계 순차 실행 |
| `download_all.py` | 종목코드 | `downloads/`, `disclosure_list.json`, `company_info.json` | corpCode.xml에서 고유번호 조회 → 공시목록 페이징 → 문서 ZIP 다운로드 |
| `build_db.py` | 종목코드 | `dart.db` | ZIP 내 XML을 멀티인코딩으로 파싱(utf-8→euc-kr→cp949→latin-1), BODY 텍스트 추출, FTS5 인덱스 |
| `build_full_db.py` | 종목코드 | `full.db` | 15종 구조화 API를 연도×보고서구분 조합으로 호출, `time.sleep(0.3~1)` rate limiting |
| `build_ai_db.py` | 종목코드 | `ai.db` | full_db + dart_db 통합, 사업보고서 8개 섹션 분리, 특허/이벤트/잠정실적 패턴매칭 |
| `create_*.py` | `ai.db` | xlsx | 회사별 보고서 생성 (회사 폴더 내 위치) |
| `create_master.py` | 개별 xlsx 파일들 | 마스터보고서.xlsx | 종합+구루+α 시트를 하나의 파일로 병합 |
| `export_pdf.py` | 마스터보고서.xlsx | .pdf | win32com Excel COM 자동화, A4 가로 인쇄 |

**주의**: 모든 `build_*.py` 스크립트는 기존 DB를 삭제 후 재생성함 (`os.remove` → 새 DB 생성).

### 보고서 스크립트 패턴

모든 `create_*.py`가 동일한 패턴:
1. **스타일 상수**: `NAVY="1B2A4A"`, `DARK_BLUE="2C3E6B"` 등 색상 + Font/Fill/Alignment/Border 객체
2. **헬퍼 함수** (이름과 시그니처 통일):
   - `sw(ws, widths)` — 열 너비 설정
   - `wh(ws, row, height)` — 행 높이 설정
   - `wr(ws, row, data, ...)` — 데이터 행 쓰기
   - `st(ws, row, text, ...)` — 스타일 텍스트 쓰기
   - `fmt(value)` — 억원 포맷 (`÷ 100,000,000`)
   - `fw(value)` — 원 포맷
   - `pct(value)` — 퍼센트 포맷
3. **데이터**: 대부분 재무 데이터가 **파이썬 상수로 하드코딩** (DB 의존 최소화). 새 기업 보고서 작성 시 해당 기업 데이터로 교체 필요.
4. **인쇄 설정**: Letter 용지 가로 방향, `fitToWidth=1, fitToHeight=0`
5. **폰트**: 맑은 고딕 (Malgun Gothic) 통일

`create_master.py`는 openpyxl `load_workbook`으로 개별 보고서를 열고 시트를 복사하여 하나의 파일로 병합. 중복 시트명에는 접미사 추가.

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
- 회사 폴더: `companies/{종목코드}_{회사명}/` 형식으로 자동 생성 (`ensure_company_dir`)
- 데이터 출처: OpenDART API (https://opendart.fss.or.kr)

### 새 기업 추가 절차

1. 파이프라인 실행: `python run_pipeline.py {종목코드}` → 자동으로 `companies/{코드}_{이름}/` 생성 + DB 4종 구축
2. 기존 기업의 `create_*.py`를 새 폴더로 복사
3. 스크립트 내 하드코딩된 데이터(기업명, 주가, 주식수, 재무 상수 등)를 새 기업 데이터로 교체
4. `create_master.py`의 `SOURCES` 리스트와 `OUTPUT` 파일명 수정

### 기업별 데이터 특이사항

**강원랜드 (035250)**:
- 매출 계정명이 연도별로 다름: "매출" (2015-2018) → "수익(매출액)" (2019+)
- `v_annual_performance` 뷰가 잘못된 매출값 반환 → 직접 쿼리 필요
- 2021-2024 연결 영업이익이 사업보고서(11011) financial_statements에 누락 → 잠정실적으로 보완
- employees 테이블 데이터가 모두 NULL (API에서 미반환)
- 개별 보수 공시 미해당 (CEO 보수 비공개)
- 특허 없음 (카지노업)
- treasury_stock의 stock_knd가 연도별로 '보통주식'/'보통주'/None 혼재

**대한항공 (003490)**:
- 매출 계정명 변경: "매출" (2015-2023) → "영업수익" (2024)
- `v_annual_performance` 뷰가 빈 매출값 반환 → `financial_summary`에서 `account_nm='매출액'`으로 직접 쿼리 필요
- 2024년 아시아나항공 합병으로 자산/부채 급증 (자산 +16.6조, 부채 +15.5조)
- 리스부채 10.9조원 (IFRS 16) → 부채비율 329%이나 리스 제외 시 약 180%
- 특허 없음 (항공운송업), 항공우주사업부 R&D는 별도 집계
- 전용 보고서: create_debt_analysis.py (항공업 구조분석), create_segment_analysis.py (사업부문별 수익분석)

### Method/ 분석 프레임워크

- **Korean_Guru_Framework.md**: Buffett/Munger Four Filters를 한국시장에 적용. Korea Discount 반영 멀티플(PER 8-12x, PBR 0.8-1.5x), 100점 스코어카드, A-D 등급
- **Corporate_Analysis_Framework**: 종합 투자분석 프레임워크. 무형자산 평가, AI 통합, 공급망 회복력, 2트랙 분석(합리적+심리적)
- **AI_Prompt_Templates**: Four Filters 평가용 시스템 프롬프트, 산업별 체크리스트, 분기별 트래킹, 다기업 비교 템플릿

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

-- DB 전체 요약 (테이블별 건수)
SELECT * FROM v_db_summary;
```
