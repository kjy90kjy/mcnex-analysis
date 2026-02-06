"""공용 설정: API 키, 경로 유틸리티"""
import os
import glob

API_KEY = "fee81b8f1226ef15d145dbfa04d0569e34ac1656"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
COMPANIES_DIR = os.path.join(BASE_DIR, "companies")
os.makedirs(COMPANIES_DIR, exist_ok=True)


def get_company_dir(stock_code):
    """종목코드로 회사 폴더 찾기. companies/{code}_{name}/ 패턴 매칭."""
    pattern = os.path.join(COMPANIES_DIR, f"{stock_code}_*")
    matches = glob.glob(pattern)
    if matches:
        return matches[0]
    return None


def ensure_company_dir(stock_code, corp_name):
    """회사 폴더가 없으면 생성, 있으면 기존 폴더 반환."""
    existing = get_company_dir(stock_code)
    if existing:
        return existing
    folder = os.path.join(COMPANIES_DIR, f"{stock_code}_{corp_name}")
    os.makedirs(folder, exist_ok=True)
    return folder
