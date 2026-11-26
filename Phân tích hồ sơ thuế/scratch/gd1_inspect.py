"""Xem nhanh cấu trúc các file nguồn (đọc only) cho MST 3603698904."""
import sys
import openpyxl

CASE = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\PHUC_NGUYEN_TIEN_3603698904-chua so"


def show(rel, nrows=5, ncols=14):
    import os
    p = os.path.join(CASE, *rel.split("/"))
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    print("=== FILE", rel, "| sheets:", wb.sheetnames)
    for ws in wb.worksheets[:3]:
        print("--- SHEET", ws.title, "max_row", ws.max_row, "max_col", ws.max_column)
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= nrows:
                break
            print([c for c in row][:ncols])
    wb.close()


if __name__ == "__main__":
    for rel in sys.argv[1:]:
        show(rel)
