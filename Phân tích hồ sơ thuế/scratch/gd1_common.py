"""Hằng số và hàm dùng chung cho MST 3603698904 (PHÚC NGUYÊN TIẾN)."""
import os
import openpyxl

MST = "3603698904"
CASE = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\PHUC_NGUYEN_TIEN_3603698904-chua so"
OUT_DIR = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\scratch\out_3603698904"


def src_path(rel: str) -> str:
    """rel dùng dấu '/' giữa các thư mục."""
    return os.path.join(CASE, *rel.split("/"))


def read_rows(rel: str, sheet_index: int = 0):
    wb = openpyxl.load_workbook(src_path(rel), read_only=True, data_only=True)
    ws = wb.worksheets[sheet_index]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows
