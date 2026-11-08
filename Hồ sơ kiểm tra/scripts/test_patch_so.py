import openpyxl
import shutil
from pathlib import Path

SO_SRC = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm")
SO_TEST = Path(r"temp_test\so_tested.xlsm")

shutil.copy2(SO_SRC, SO_TEST)

wb = openpyxl.load_workbook(SO_TEST, keep_vba=True)
ws = wb["01_DA_BAN_HANH"]

r = None
for i in range(5, ws.max_row + 1):
    if str(ws.cell(i, 4).value or "") == "3603859012":
        r = i
        break

print(f"Row for Son Tien Hai: {r}")

ws.cell(r, 16).value = "Lê Hoài Phương"
# Col 17 CCCD KT left blank
ws.cell(r, 21).value = "6722949824 tại Ngân hàng TMCP Đầu tư và Phát triển Việt Nam"
ws.cell(r, 27).value = "-Giá vốn hàng bán\n- Chi phí bán hàng\n- Chi phí quản lý doanh nghiệp"
ws.cell(r, 29).value = "7092/QĐ-TCS3"
ws.cell(r, 30).value = "7093/QĐ-TCS3"
ws.cell(r, 31).value = "26/10/2026"
ws.cell(r, 40).value = 13052423
ws.cell(r, 41).value = 0
ws.cell(r, 42).value = 13052423
ws.cell(r, 44).value = 662730
ws.cell(r, 45).value = 662730
ws.cell(r, 46).value = 2610485
ws.cell(r, 47).value = 0
ws.cell(r, 48).value = 3273215
ws.cell(r, 49).value = 16325638
ws.cell(r, 50).value = 809868861
ws.cell(r, 51).value = 0

wb.save(SO_TEST)
print("Saved SO_TEST successfully!")

# Verify
wb2 = openpyxl.load_workbook(SO_TEST, data_only=True)
ws2 = wb2["01_DA_BAN_HANH"]
for c in [16, 17, 21, 27, 29, 30, 31, 40, 42, 44, 46, 48, 49, 50]:
    print(f"Col {c} ({ws2.cell(4, c).value}): {repr(ws2.cell(r, c).value)}")
