import requests
import zipfile
import io
import xml.etree.ElementTree as ET
import json
import os
import time

API_KEY = "fee81b8f1226ef15d145dbfa04d0569e34ac1656"
STOCK_CODE = "097520"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# ============ STEP 1: corpCode.xml에서 고유번호 찾기 ============
print("=" * 60)
print("STEP 1: 고유번호 조회 중...")
print("=" * 60)

url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={API_KEY}"
resp = requests.get(url, timeout=300, stream=True)

# Stream download with progress
content = b""
total = int(resp.headers.get("content-length", 0))
downloaded = 0
for chunk in resp.iter_content(chunk_size=1024 * 256):
    content += chunk
    downloaded += len(chunk)
    if total:
        pct = downloaded / total * 100
        print(f"\r  다운로드: {downloaded // 1024}KB / {total // 1024}KB ({pct:.0f}%)", end="", flush=True)
    else:
        print(f"\r  다운로드: {downloaded // 1024}KB", end="", flush=True)
print()

corp_code = None
corp_name = None
with zipfile.ZipFile(io.BytesIO(content)) as zf:
    xml_name = zf.namelist()[0]
    with zf.open(xml_name) as f:
        tree = ET.parse(f)
        root = tree.getroot()
        for corp in root.findall("list"):
            sc = corp.findtext("stock_code", "").strip()
            if sc == STOCK_CODE:
                corp_code = corp.findtext("corp_code", "")
                corp_name = corp.findtext("corp_name", "")
                break

if not corp_code:
    print(f"종목코드 {STOCK_CODE}에 해당하는 회사를 찾을 수 없습니다.")
    exit(1)

print(f"  회사명: {corp_name}")
print(f"  고유번호: {corp_code}")
print(f"  종목코드: {STOCK_CODE}")

# ============ STEP 2: 전체 공시 목록 수집 ============
print()
print("=" * 60)
print("STEP 2: 전체 공시 목록 수집 중...")
print("=" * 60)

all_disclosures = []
page_no = 1
page_count = 100

while True:
    url = "https://opendart.fss.or.kr/api/list.json"
    params = {
        "crtfc_key": API_KEY,
        "corp_code": corp_code,
        "bgn_de": "19990101",
        "end_de": "20261231",
        "page_no": page_no,
        "page_count": page_count,
    }
    resp = requests.get(url, params=params, timeout=30)
    data = resp.json()

    if data.get("status") != "000":
        if page_no == 1:
            print(f"  API 오류: {data.get('message')}")
        break

    items = data.get("list", [])
    all_disclosures.extend(items)
    total_count = data.get("total_count", 0)
    total_page = data.get("total_page", 0)
    print(f"  페이지 {page_no}/{total_page} 수집완료 (누적: {len(all_disclosures)}/{total_count}건)")

    if page_no >= total_page:
        break
    page_no += 1
    time.sleep(0.5)  # API 호출 간격

# 목록 저장
with open(os.path.join(BASE_DIR, "disclosure_list.json"), "w", encoding="utf-8") as f:
    json.dump(all_disclosures, f, ensure_ascii=False, indent=2)

print(f"\n  총 {len(all_disclosures)}건 공시 목록 저장 완료 (disclosure_list.json)")

# ============ STEP 3: 전체 문서 다운로드 ============
print()
print("=" * 60)
print(f"STEP 3: 전체 {len(all_disclosures)}건 문서 다운로드 중...")
print("=" * 60)

success = 0
fail = 0
skip = 0

for i, disc in enumerate(all_disclosures):
    rcept_no = disc["rcept_no"]
    report_nm = disc.get("report_nm", "unknown")
    rcept_dt = disc.get("rcept_dt", "unknown")

    # 파일명: 날짜_보고서명_접수번호.zip
    safe_name = report_nm.replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace('"', "_").replace("<", "_").replace(">", "_").replace("|", "_")
    filename = f"{rcept_dt}_{safe_name}_{rcept_no}.zip"
    filepath = os.path.join(DOWNLOAD_DIR, filename)

    if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
        skip += 1
        print(f"  [{i+1}/{len(all_disclosures)}] SKIP (이미 존재): {filename}")
        continue

    # 문서 다운로드 API
    doc_url = "https://opendart.fss.or.kr/api/document.xml"
    doc_params = {
        "crtfc_key": API_KEY,
        "rcept_no": rcept_no,
    }

    try:
        doc_resp = requests.get(doc_url, params=doc_params, timeout=60)

        # Check if it's an error JSON response
        content_type = doc_resp.headers.get("Content-Type", "")
        if "application/json" in content_type or doc_resp.content[:1] == b"{":
            try:
                err = doc_resp.json()
                print(f"  [{i+1}/{len(all_disclosures)}] FAIL: {report_nm} - {err.get('message', 'unknown error')}")
                fail += 1
            except:
                print(f"  [{i+1}/{len(all_disclosures)}] FAIL: {report_nm} - 응답 파싱 실패")
                fail += 1
            time.sleep(1)
            continue

        with open(filepath, "wb") as f:
            f.write(doc_resp.content)

        size_kb = len(doc_resp.content) / 1024
        print(f"  [{i+1}/{len(all_disclosures)}] OK: {filename} ({size_kb:.1f}KB)")
        success += 1

    except Exception as e:
        print(f"  [{i+1}/{len(all_disclosures)}] ERROR: {report_nm} - {e}")
        fail += 1

    time.sleep(1)  # API 호출 제한 방지 (분당 약 60회)

# ============ 결과 요약 ============
print()
print("=" * 60)
print("다운로드 완료!")
print("=" * 60)
print(f"  성공: {success}건")
print(f"  실패: {fail}건")
print(f"  스킵(이미존재): {skip}건")
print(f"  저장 위치: {DOWNLOAD_DIR}")
