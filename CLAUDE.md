# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# 엠씨넥스(MCNEX) 기업분석 프로젝트

## 회사 개요

- **회사명**: (주)엠씨넥스 / MCNEX CO.,LTD
- **종목코드**: 097520 (코스닥)
- **DART 고유번호**: 00564562
- **대표이사**: 민동욱
- **설립일**: 2004년 12월 22일
- **본사**: 인천광역시 연수구 송도과학로16번길 13-39 엠씨넥스타워
- **업종코드**: 26519 (기타 영상기기 제조업)
- **홈페이지**: www.mcnex.co.kr
- **주요사업**: CCM(카메라모듈) 기술 기반 영상전문기업. 휴대폰용 카메라모듈, 자동차용 카메라모듈이 주력. 엑츄에이터, 생체인식모듈, 멀티카메라, 블랙박스, 로봇, CCTV, 3D 카메라모듈 등으로 영역 확장.
- **종속회사**: 엠씨넥스VINA(베트남, 카메라모듈 제조), 엠씨넥스상해전자무역유한공사(중국, 수출입/CS), 엠씨넥스에프앤비(음식점)

## 데이터 파이프라인

3단계 순서로 실행해야 함. 모든 스크립트는 Python 3, 주요 의존성은 `requests`뿐.

```bash
# 1단계: OpenDART에서 574건 공시 원본 ZIP 다운로드 → downloads/ + disclosure_list.json
python download_all.py

# 2단계: ZIP → XML 파싱 → mcnex_dart.db (원문 전체, 153MB, FTS 검색 지원)
python build_db.py

# 3단계-A: OpenDART 구조화 API 15종 호출 → mcnex_full.db (정량 데이터)
python build_full_db.py

# 3단계-B: mcnex_full.db + mcnex_dart.db 통합 → mcnex_ai.db (분석용 최종 DB)
python build_ai_db.py
```

| 스크립트 | 입력 | 출력 | 소요시간 |
|----------|------|------|----------|
| `download_all.py` | OpenDART API | `downloads/`, `disclosure_list.json` | ~10분 (API 제한) |
| `build_db.py` | `downloads/*.zip`, `disclosure_list.json` | `mcnex_dart.db` | ~1분 |
| `build_full_db.py` | OpenDART API | `mcnex_full.db` | ~30분 (API 호출 多) |
| `build_ai_db.py` | `mcnex_full.db`, `mcnex_dart.db` | `mcnex_ai.db` | ~10초 |

`build_ai_db.py`에서 사업보고서 원문(mcnex_dart.db)의 BODY 텍스트를 정규식으로 8개 섹션(회사개요/연혁/사업내용/제품/연구개발/위험관리/임원보수/주주배당)으로 분리 추출함. 특허/경영이벤트/잠정실적도 패턴 매칭으로 추출.

## DB 접근 방법

분석 시 SQLite CLI 또는 Python sqlite3 모듈로 `mcnex_ai.db`만 열면 됨.

```bash
sqlite3 mcnex_ai.db "SELECT * FROM v_annual_performance;"
```

```python
import sqlite3
conn = sqlite3.connect('mcnex_ai.db')
conn.execute("SELECT * FROM v_annual_performance").fetchall()
```

## 핵심 DB 파일

### `mcnex_ai.db` (5.5MB) — AI 분석용 통합 DB (이것만 사용하면 됨)

OpenDART에서 수집한 구조화 데이터 + 공시 원문에서 추출한 핵심 텍스트를 통합한 SQLite DB.

#### 정량 데이터 테이블

| 테이블 | 설명 | 건수 | 비고 |
|--------|------|------|------|
| `company_info` | 기업개황 | 1 | 대표, 주소, 설립일, 업종 |
| `disclosures` | 전체 공시 목록 | 574 | 2007~2026년 모든 공시 타임라인 |
| `financial_statements` | 전체 재무제표 | 11,764 | BS/CIS/CF 모든 계정, 연결+개별, 2012~2024 |
| `financial_summary` | 주요계정 요약 | 2,074 | 분기/반기/연간별 핵심 계정 |
| `executives` | 임원현황 | 70 | 성명, 직위, 담당업무, 재직기간, 최대주주관계 |
| `employees` | 직원현황 | 12 | 인원수, 평균근속, 급여총액 |
| `dividends` | 배당현황 | 149 | 배당성향, 주당배당금 등 |
| `treasury_stock` | 자기주식 | 216 | 취득/처분 이력 |
| `capital_changes` | 증자/감자 | 346 | 유상증자, 무상증자, 전환 등 이력 |
| `stock_total` | 주식 총수 | 32 | 발행주식, 자기주식, 유통주식 |
| `investments` | 타법인 출자 | 67 | 자회사/관계사 투자 현황 |
| `minority_shareholders` | 소액주주 | 12 | 소액주주 비율 |
| `individual_pay` | 고액보수 | 8 | 5억 이상 개인별 보수 |

#### 정성 데이터 테이블 (텍스트)

| 테이블 | 설명 | 건수 | 비고 |
|--------|------|------|------|
| `business_report_sections` | 사업보고서 핵심 섹션 | 102 | 13년치 × 8개 섹션(회사개요/사업내용/제품/연구개발/위험관리/임원보수/주주배당/연혁) |
| `patents` | 특허 공시 | 24 | 특허명, 상세내용, 취득일, 활용계획 |
| `key_events` | 주요 경영 이벤트 | 157 | 증자/투자/배당/자기주식/채무보증/IR/실적변동 등 |
| `earnings_announcements` | 잠정실적 공시 | 21 | 분기별 잠정 영업실적 |

#### 편의 뷰

| 뷰 | 설명 |
|----|------|
| `v_annual_performance` | 연도별 매출/영업이익/순이익/EPS/총자산/총부채/총자본 (연결) |
| `v_annual_dividends` | 연도별 배당 |
| `v_major_shareholder_history` | 대주주 변동 이력 |
| `v_executive_history` | 임원 변동 이력 |
| `v_disclosure_timeline` | 전체 공시 타임라인 (최신순) |
| `v_business_sections` | 사업보고서 섹션별 미리보기 |
| `v_patent_history` | 특허 이력 |
| `v_event_timeline` | 경영 이벤트 타임라인 |
| `v_db_summary` | DB 전체 테이블별 건수 |

### 기타 파일

| 파일 | 설명 |
|------|------|
| `mcnex_dart.db` (153MB) | 574건 공시 원문 XML 전체 (상세 텍스트 필요 시 참조) |
| `mcnex_full.db` (3.8MB) | 구조화 데이터만 (mcnex_ai.db에 포함됨) |
| `downloads/` | 574건 공시 원본 ZIP 파일 |
| `disclosure_list.json` | 공시 목록 JSON |

## 자주 쓰는 쿼리

```sql
-- 연도별 핵심 실적 (연결)
SELECT * FROM v_annual_performance;

-- 특정 연도 사업보고서 '사업내용' 읽기
SELECT section_text FROM business_report_sections
WHERE bsns_year = '2024' AND section_name = '사업내용';

-- 전체 특허 이력
SELECT rcept_dt, patent_name, patent_detail, patent_plan FROM patents ORDER BY rcept_dt;

-- 주요 이벤트 타임라인
SELECT rcept_dt, event_type, event_summary FROM key_events ORDER BY rcept_dt DESC;

-- 특정 재무 계정 연도별 추이 (예: 매출액)
SELECT bsns_year, thstrm_amount FROM financial_statements
WHERE account_nm LIKE '%매출액%' AND sj_div='CIS' AND reprt_code='11011' AND reprt_nm LIKE '%연결%'
ORDER BY bsns_year;

-- 임원 현황 (최신)
SELECT nm, ofcps, chrg_job, hffc_pd, tenure_end_on FROM executives
WHERE bsns_year = (SELECT MAX(bsns_year) FROM executives);

-- 배당 이력
SELECT * FROM dividends WHERE se LIKE '%주당%배당금%' ORDER BY bsns_year;

-- 자회사 투자 현황 (최신)
SELECT inv_prm, trmend_blce_qota_rt, trmend_blce_acntbk_amount
FROM investments WHERE bsns_year = (SELECT MAX(bsns_year) FROM investments);

-- 본문 전체 검색 (사업보고서 내)
SELECT bsns_year, section_name, section_text FROM business_report_sections
WHERE section_text LIKE '%자동차%카메라%';

-- 잠정실적 공시 이력
SELECT rcept_dt, report_nm, content FROM earnings_announcements ORDER BY rcept_dt DESC;
```

## 분석 시 참고사항

- 재무제표 금액 단위: **원** (억원 변환 시 ÷ 100,000,000)
- `reprt_code`: 11011=사업보고서, 11012=반기, 11013=1분기, 11014=3분기
- `sj_div`: BS=재무상태표, CIS=포괄손익계산서, CF=현금흐름표, SCE=자본변동표
- 연결 재무제표: `reprt_nm LIKE '%연결%'` / 개별: `reprt_nm LIKE '%개별%'`
- `business_report_sections.section_name` 값: 회사개요, 회사연혁, 사업내용, 주요제품_매출, 연구개발, 위험관리_전망, 임원_보수, 주주_배당
- 2012년부터 사업보고서 존재 (2007~2011은 감사보고서만 있음)
- 데이터 출처: OpenDART API (https://opendart.fss.or.kr)
- 수집일: 2026-02-06
