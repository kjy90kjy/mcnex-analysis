"""
마스터보고서 xlsx → PDF 변환 (Excel COM 자동화)

사용법:
    python export_pdf.py                          # 두 보고서 모두 변환
    python export_pdf.py path/to/file.xlsx        # 특정 파일만 변환
"""

import os
import sys
import win32com.client


# 기본 대상 파일 (인자 없을 때)
DEFAULT_FILES = [
    r"companies\097520_엠씨넥스\엠씨넥스_마스터보고서.xlsx",
    r"companies\035250_강원랜드\강원랜드_마스터보고서.xlsx",
]


def convert_to_pdf(xlsx_path: str) -> str:
    """xlsx 파일을 같은 디렉토리에 PDF로 변환하고 PDF 경로를 반환한다."""
    abs_path = os.path.abspath(xlsx_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {abs_path}")

    pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
    print(f"변환 중: {os.path.basename(abs_path)}")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(abs_path)

        # 각 시트에 셀 크기 조정 + 인쇄 설정 적용
        sheet_count = wb.Worksheets.Count
        for ws in wb.Worksheets:
            # 컬럼 자동 맞춤 (내용이 잘리지 않도록)
            used = ws.UsedRange
            if used is not None:
                used.Columns.AutoFit()
                used.Rows.AutoFit()

            ps = ws.PageSetup
            ps.Orientation = 2               # xlLandscape (가로)
            ps.Zoom = False
            ps.FitToPagesWide = 1            # 너비 1페이지에 맞춤
            ps.FitToPagesTall = False         # 높이는 자동
            ps.PaperSize = 9                 # xlPaperA4
            ps.LeftMargin = excel.CentimetersToPoints(1.0)
            ps.RightMargin = excel.CentimetersToPoints(1.0)
            ps.TopMargin = excel.CentimetersToPoints(1.5)
            ps.BottomMargin = excel.CentimetersToPoints(1.5)

        # PDF 내보내기 (0 = xlTypePDF)
        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)

        print(f"  완료: {os.path.basename(pdf_path)} ({sheet_count}시트)")
        return pdf_path

    except Exception as e:
        print(f"  오류: {e}", file=sys.stderr)
        raise
    finally:
        excel.Quit()


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    if len(sys.argv) > 1:
        files = sys.argv[1:]
    else:
        files = [os.path.join(base_dir, f) for f in DEFAULT_FILES]

    results = []
    for f in files:
        pdf = convert_to_pdf(f)
        results.append(pdf)

    print(f"\n총 {len(results)}개 PDF 생성 완료:")
    for r in results:
        size_mb = os.path.getsize(r) / (1024 * 1024)
        print(f"  {r}  ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
