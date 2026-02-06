import subprocess
import sys
import os
import time

sys.stdout.reconfigure(encoding='utf-8')

if len(sys.argv) < 2:
    print("사용법: python run_pipeline.py <종목코드>")
    print("예시:   python run_pipeline.py 097520")
    print("        python run_pipeline.py 005930")
    sys.exit(1)

STOCK_CODE = sys.argv[1]
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

STEPS = [
    ("1/4", "공시 다운로드", "download_all.py"),
    ("2/4", "원문 DB 생성", "build_db.py"),
    ("3/4", "구조화 DB 생성", "build_full_db.py"),
    ("4/4", "통합 AI DB 생성", "build_ai_db.py"),
]

print("=" * 60)
print(f"  기업분석 파이프라인 실행 - 종목코드: {STOCK_CODE}")
print("=" * 60)
print()

start_total = time.time()

for step_no, step_name, script in STEPS:
    print("=" * 60)
    print(f"  [{step_no}] {step_name} ({script})")
    print("=" * 60)

    script_path = os.path.join(BASE_DIR, script)
    start = time.time()

    result = subprocess.run(
        [sys.executable, script_path, STOCK_CODE],
        cwd=BASE_DIR,
    )

    elapsed = time.time() - start

    if result.returncode != 0:
        print(f"\n  [{step_no}] 실패! (종료코드: {result.returncode}, 소요: {elapsed:.1f}초)")
        print("  파이프라인을 중단합니다.")
        sys.exit(result.returncode)

    print(f"\n  [{step_no}] 완료 (소요: {elapsed:.1f}초)")
    print()

total_elapsed = time.time() - start_total

print("=" * 60)
print("  전체 파이프라인 완료!")
print("=" * 60)
print(f"  종목코드: {STOCK_CODE}")
print(f"  총 소요시간: {total_elapsed:.1f}초")

from config import get_company_dir
company_dir = get_company_dir(STOCK_CODE)
if company_dir:
    print(f"  회사 폴더: {company_dir}")
    print(f"  AI DB: {os.path.join(company_dir, 'ai.db')}")
print()
