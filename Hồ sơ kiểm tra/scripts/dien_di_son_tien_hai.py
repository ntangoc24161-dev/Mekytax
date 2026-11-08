# -*- coding: utf-8 -*-
"""Dien placeholder «...» cho bo DI-01 - DI-07 (Cong ty Son Tien Hai, MST 3603859012).
Nguon phoi: temp_test/DI-01.docx .. DI-07.docx (mail-merge, giu nguyen format goc theo QD 1412/1161).
"""
import sys
from pathlib import Path
from docx import Document

BRAIN_DIR = Path(__file__).resolve().parent.parent
SRC_DIR = BRAIN_DIR / "temp_test"
OUT_DIR = BRAIN_DIR / "output" / "3603859012_DU_THAO-TRINH-KY_son-tien-hai"

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TEN_CONG_TY = "Công ty TNHH Thương mại Sản xuất Sơn Tiền Hải"
MST = "3603859012"
DIA_CHI = "Số 23, Tổ 1, KP 3A, Tỉnh Đồng Nai"
NNKD = "Bán buôn vật liệu xây dựng"
THOI_KY = "Từ năm 2022 đến tháng 10 năm 2026"
GD = "Phan Thế Hanh"

NOI_DUNG_KIEM_TRA = (
    "- Giá vốn hàng bán Từ năm 2022 đến tháng 10/2026.\n"
    "- Chi phí quản lý kinh doanh Từ năm 2022 đến tháng 10/2026.\n"
    "- Doanh thu bán hàng và cung cấp dịch vụ Từ năm 2022 đến tháng 10/2026.\n"
    "- Thuế GTGT phải nộp Từ năm 2022 đến tháng 10/2026."
)
# Đồng bộ phân công trên sổ JAVIS PRO (hàng MST 3603859012).
PHAN_CONG_TD = "Giá vốn hàng bán; Chi phí quản lý kinh doanh"
PHAN_CONG_TV = "Doanh thu bán hàng và cung cấp dịch vụ; Thuế GTGT phải nộp"

COMMON = {
    "«Tên_công_ty»": TEN_CONG_TY,
    "«MST»": MST,
    "«địa_chỉ»": DIA_CHI,
    "«NNKD»": NNKD,
    "«TĐ»": "Nguyễn Thị Ánh Ngọc",
    "«CVTĐ»": "Chuyên viên",
    "«TV»": "Nguyễn Thành Đồng",
    "«CVTV»": "Kiểm tra viên thuế",
    "«Thời_kỳ_kiểm_tra»": THOI_KY,
    "«Thời_hạn_kiểm_tra_»": "......",
    "«Nội_dung_kiểm_tra»": NOI_DUNG_KIEM_TRA,
    "«Phân_công_trưởng_đoàn»": PHAN_CONG_TD,
    "«Phân_công_thành_viên»": PHAN_CONG_TV,
    "«GĐ»": GD,
    "«CCCD»": "......................",
    "«ngày_cấp»": "...../...../.........",
    "«Nơi_cấp»": "......................",
    # Chưa có số thật: để trống theo mẫu trình ký, không dùng DU_THAO-TRINH-KY.
    "«QĐGS»": "              /QĐ-TCS3",
    "«QĐKT»": "              /QĐ-TCS3",
    "«ngày_QĐ»": "...../....../2026",
}

# Cac dong phan tich BCTC cu (so lieu khong thuoc Son Tien Hai) can xoa, de trong cho can bo dien sau
STALE_ANALYSIS_LINES = {
    "- Doanh nghiệp kê khai lãi từ năm 2022, 2024- 2025; kê khai lỗ năm 2023.",
    "- Tỷ lệ giá vốn hàng bán trên doanh thu năm 2022 là 88,75%, 2023 là 86,18%, 2024 là 85,66%, năm 2025 là 89,16%.",
    "- Tỷ lệ Chi phí quản lý kinh doanh trên doanh thu năm 2022 là 9,96%, 2023 là 14,15%, 2024 là 14,78%, năm 2025 là 8,49%.",
    "- Tỷ lệ giá vốn hàng bán và chi phí quản lý kinh doanh trên doanh thu biến động không đều qua các năm, không tương ứng biến động doanh thu.",
}


def replace_in_runs(paragraph, mapping):
    for run in paragraph.runs:
        t = run.text
        if not t:
            continue
        for ph, val in mapping.items():
            if ph in t:
                t = t.replace(ph, val)
        if t != run.text:
            run.text = t


def clear_stale_lines(paragraph):
    if paragraph.text.strip() in STALE_ANALYSIS_LINES:
        for run in paragraph.runs:
            run.text = ""


def fill_file(src: Path, dst: Path):
    doc = Document(str(src))
    for p in doc.paragraphs:
        clear_stale_lines(p)
        replace_in_runs(p, COMMON)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    clear_stale_lines(p)
                    replace_in_runs(p, COMMON)
    # DI-06: ngay ke hoach giam sat con luu lai tu lan dien truoc, blank lai cho dung thuc te chua duyet
    if src.name == "DI-06.docx":
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    if "Trấn Biên, ngày 28" in cell.text:
                        for p in cell.paragraphs:
                            if "Trấn Biên" in p.text:
                                for run in p.runs:
                                    if run.text.strip():
                                        run.text = "Trấn Biên, ngày      tháng     năm 2026" if "Trấn Biên" in run.text else ""
    doc.save(str(dst))


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for i in range(1, 8):
        name = f"DI-{i:02d}.docx"
        src = SRC_DIR / name
        dst = OUT_DIR / name
        fill_file(src, dst)
        print(f"Da dien: {dst}")
    print(f"\nThu muc output: {OUT_DIR}")


if __name__ == "__main__":
    main()
