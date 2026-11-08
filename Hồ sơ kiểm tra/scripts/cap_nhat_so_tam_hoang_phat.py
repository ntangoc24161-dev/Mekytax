"""Cap nhat so Danh sach kiem tra quyet toan pro.xlsm tu sheet
TINH PHAT + CHAM NOP (C T) cua Bang tinh.xlsx (Tam Hoang Phat).
Nguon: ban Drive 05/11/2026 04:13 (so) va 03:40 (Bang tinh).
Chi ghi cac o so lieu, khong sua cong thuc.
"""
import shutil
import openpyxl

D = "/brains/Hồ sơ kiểm tra/exports/tam-hoang-phat-3603721832/"
BT = D + "bangtinh-drive-20261105.xlsx"
SO_SRC = D + "so-drive-20261105-0413.xlsm"
SO_DST = D + "Danh sach kiem tra quyet toan pro.xlsm"
ROW = 14  # Cong Ty TNHH TM DV Tam Hoang Phat, MST 3603721832
TONG = 67

bt = openpyxl.load_workbook(BT, data_only=True)["TINH PHAT + CHAM NOP (C T)"]
moi = {
    "AN": round(bt["B%d" % TONG].value or 0),   # truy thu TNDN
    "AO": round(bt["C%d" % TONG].value or 0),   # truy thu GTGT
    "AQ": round(bt["I%d" % TONG].value or 0),   # cham nop GTGT
    "AR": round(bt["M%d" % TONG].value or 0),   # cham nop TNDN 0,03%
    "AT": round(bt["N%d" % TONG].value or 0),   # phat khai sai 20%
    "AJ": bt["D52"].value.strftime("%d/%m/%Y"), # ngay ky BBKT
}

shutil.copy2(SO_SRC, SO_DST)
wb = openpyxl.load_workbook(SO_DST, keep_vba=True)
ws = wb["01_DA_BAN_HANH"]
for col, gt in moi.items():
    o = "%s%d" % (col, ROW)
    cu = ws[o].value
    if cu != gt:
        ws[o] = gt
        print("sua %s: %r -> %r" % (o, cu, gt))
    else:
        print("giu %s: %r" % (o, cu))
wb.save(SO_DST)
print("tong phat sinh :", moi["AN"] + moi["AO"] + moi["AQ"] + moi["AR"] + moi["AT"])
print("Q67 bang tinh  :", bt["Q%d" % TONG].value)
