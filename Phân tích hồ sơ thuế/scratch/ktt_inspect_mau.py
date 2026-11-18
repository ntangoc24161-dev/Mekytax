# -*- coding: utf-8 -*-
import os
from openpyxl import load_workbook

MAU = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\thu-muc-mau-ho-so-doanh-nghiep\MAU_FILE_TONG_HOP"
KQ11 = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\KIEN_THANH_TAN_3603216935\KQ1.1_Kiem_ke_CQT_3603216935_2015-2025_v2.xlsx"
KQ12 = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\KIEN_THANH_TAN_3603216935\KQ1.2_Kiem_ke_ho_so_NNT_3603216935_2015-2025_v1.xlsx"

def lp(p):
    return "\\\\?\\" + os.path.abspath(p)

def dump_wb(path, max_sheets=20):
    print("\n====", os.path.basename(path))
    wb = load_workbook(lp(path), read_only=True, data_only=False)
    print("sheets", wb.sheetnames)
    for sn in wb.sheetnames[:max_sheets]:
        ws = wb[sn]
        print(f" -- {sn} max_row={ws.max_row} max_col={ws.max_column}")
        for r in range(1, min(12, (ws.max_row or 1)+1)):
            vals = []
            for c in range(1, min((ws.max_column or 1)+1, 20)):
                v = ws.cell(r,c).value
                if v is not None:
                    s = str(v).replace("\n"," ")[:60]
                    vals.append(f"{c}:{s}")
            if vals:
                print("  R"+str(r), vals)
    wb.close()

for fn in ["KQ2.1_Phan_tich_CQT.xlsx","KQ2.2_Phan_tich_ho_so_cty.xlsx","KQ3_Ket_luan_danh_gia.xlsx"]:
    dump_wb(os.path.join(MAU, fn))

# extra GĐ3 templates
root = MAU
for fn in os.listdir(root):
    if "Dau_hieu" in fn or "Giai_trinh" in fn or "KQ3_" in fn:
        print("found", fn)

print("\n==== KQ1.1 IN headers")
wb = load_workbook(lp(KQ11), read_only=True)
for sn in wb.sheetnames:
    ws = wb[sn]
    hdr=[]
    for c in range(1, min(20, (ws.max_column or 1)+1)):
        hdr.append(str(ws.cell(2,c).value or ws.cell(1,c).value)[:40])
    print(sn, "rows", ws.max_row, "hdr", hdr)
wb.close()
print("\n==== KQ1.2 IN headers")
wb = load_workbook(lp(KQ12), read_only=True)
for sn in wb.sheetnames:
    ws = wb[sn]
    hdr=[]
    for c in range(1, min(18, (ws.max_column or 1)+1)):
        hdr.append(str(ws.cell(2,c).value or ws.cell(1,c).value)[:40])
    print(sn, "rows", ws.max_row, "hdr", hdr)
wb.close()
