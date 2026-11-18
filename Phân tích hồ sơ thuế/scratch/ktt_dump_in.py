# -*- coding: utf-8 -*-
"""Dump IN summaries for KTT GĐ2/GĐ3."""
from __future__ import annotations
import os, sys
from collections import defaultdict
from datetime import date, datetime
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

KQ11 = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\KIEN_THANH_TAN_3603216935\KQ1.1_Kiem_ke_CQT_3603216935_2015-2025_v2.xlsx"
KQ12 = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\KIEN_THANH_TAN_3603216935\KQ1.2_Kiem_ke_ho_so_NNT_3603216935_2015-2025_v1.xlsx"
KQ13 = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\KIEN_THANH_TAN_3603216935\KQ1.3_Danh_gia_ho_so_3603216935_2015-2025_v2.xlsx"
MAU = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\thu-muc-mau-ho-so-doanh-nghiep\MAU_FILE_TONG_HOP"

def lp(p):
    s = os.path.abspath(p)
    return s if s.startswith("\\\\?\\") else "\\\\?\\" + s

def nnum(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    return 0.0

# IN01
print("=== IN01 ===")
wb = load_workbook(lp(KQ11), read_only=True, data_only=True)
ws = wb["IN01_HDDT"]
hdr = [ws.cell(2,c).value for c in range(1, 30)]
print("hdr", hdr)
mua = defaultdict(lambda: {"n":0,"gt":0,"thue":0})
ban = defaultdict(lambda: {"n":0,"gt":0,"thue":0})
n_mua=n_ban=0
miss_mst=0
chieu_vals=set()
n=0
for i, row in enumerate(ws.iter_rows(min_row=3, values_only=True)):
    if not row or not row[0]:
        continue
    n += 1
    ch = str(row[0] or "")
    chieu_vals.add(ch)
    nam = row[1]
    gt = nnum(row[16] if len(row)>16 else 0)
    th = nnum(row[18] if len(row)>18 else 0)
    mst = str(row[11] or "")
    if "MUA" in ch.upper() or ch.upper().startswith("M"):
        n_mua += 1
        mua[nam]["n"] += 1; mua[nam]["gt"] += gt; mua[nam]["thue"] += th
    else:
        n_ban += 1
        ban[nam]["n"] += 1; ban[nam]["gt"] += gt; ban[nam]["thue"] += th
    if not mst.strip() or mst in ("None","0"):
        miss_mst += 1
print("rows", n, "mua", n_mua, "ban", n_ban, "miss_mst", miss_mst, "chieu", chieu_vals)
print("mua by year", {k: {kk: round(vv,0) if isinstance(vv,float) else vv for kk,vv in v.items()} for k,v in sorted(mua.items(), key=lambda x: str(x[0]))})
print("ban by year", {k: {kk: round(vv,0) if isinstance(vv,float) else vv for kk,vv in v.items()} for k,v in sorted(ban.items(), key=lambda x: str(x[0]))})
print("sample row3", [ws.cell(3,c).value for c in range(1, 20)])

print("\n=== IN03 ===")
ws = wb["IN03_ToKhai_GTGT"]
hdr = [ws.cell(2,c).value for c in range(1, 25)]
print("hdr", hdr)
gtgt_ky=[]
tndn_y=[]
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    loai = str(row[2] or "")
    if "GTGT" in loai.upper():
        gtgt_ky.append((row[0], row[1], row[4], nnum(row[7]) if len(row)>7 else None))  # nam ky hieu luc + some ct
    if "TNDN" in loai.upper():
        tndn_y.append((row[0], row[1], loai, row[4]))
print("gtgt n", len(gtgt_ky), "sample", gtgt_ky[:8], "... last", gtgt_ky[-3:])
print("tndn", tndn_y)

print("\n=== IN04 ===")
ws = wb["IN04_BCTC_TNDN"]
hdr = [ws.cell(2,c).value for c in range(1, 10)]
print("hdr", hdr)
by = defaultdict(list)
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    by[(row[0], row[1], row[2])].append((row[3], row[4]))
print("keys", sorted(by.keys(), key=lambda x: str(x))[:30], "nkeys", len(by))
# print 01, 11, 20, 140, 141 for each year
for y in range(2015, 2026):
    hits=[]
    for (yy, bc, ma), items in by.items():
        if yy!=y: continue
        if str(ma) in ("01","02","11","20","21","22","50","60","110","131","140","141","211","311","331","411","152","181"):
            hits.append((bc, ma, items[0] if items else None))
    if hits:
        print(y, hits[:20])

print("\n=== IN05 ===")
ws = wb["IN05_NLD_TNCN"]
print("hdr", [ws.cell(2,c).value for c in range(1, 16)])
from collections import Counter
c=Counter()
qtt=[]
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    c[str(row[0])+"|"+str(row[1])] += 1
    if str(row[0])=="QTT":
        qtt.append((row[1], row[7] if len(row)>7 else None, row[8] if len(row)>8 else None, row[16] if len(row)>16 else None))
print("counts", dict(c))
print("qtt", qtt)
wb.close()

print("\n=== IN02 ===")
wb = load_workbook(lp(KQ12), read_only=True, data_only=True)
ws = wb["IN02_BangKe"]
print("hdr", [ws.cell(2,c).value for c in range(1, 18)])
mua=defaultdict(lambda: {"n":0,"gt":0,"thue":0})
ban=defaultdict(lambda: {"n":0,"gt":0,"thue":0})
n=0
chieu_vals=set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    n+=1
    ch=str(row[0] or "")
    chieu_vals.add(ch)
    nam=row[1]
    gt=nnum(row[13] if len(row)>13 else 0)
    th=nnum(row[15] if len(row)>15 else 0)
    if "MUA" in ch.upper() or ch.upper().startswith("M"):
        mua[nam]["n"]+=1; mua[nam]["gt"]+=gt; mua[nam]["thue"]+=th
    else:
        ban[nam]["n"]+=1; ban[nam]["gt"]+=gt; ban[nam]["thue"]+=th
print("rows", n, "chieu", chieu_vals)
print("mua", {k:{kk:round(vv,0) if isinstance(vv,float) else vv for kk,vv in v.items()} for k,v in sorted(mua.items(), key=lambda x:str(x[0]))})
print("ban", {k:{kk:round(vv,0) if isinstance(vv,float) else vv for kk,vv in v.items()} for k,v in sorted(ban.items(), key=lambda x:str(x[0]))})

print("\n=== IN08 ===")
ws = wb["IN08_So_KeToan"]
print("hdr", [ws.cell(2,c).value for c in range(1, 12)])
tk_y=defaultdict(lambda: {"no":0,"co":0,"n":0})
years=set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row:
        continue
    nam=row[0]; tk=str(row[2] or "")
    years.add(nam)
    if not tk: continue
    tk3=tk[:3]
    tk_y[(nam,tk3)]["no"] += nnum(row[3]); tk_y[(nam,tk3)]["co"] += nnum(row[4]); tk_y[(nam,tk3)]["n"] += 1
print("years", sorted(str(y) for y in years if y))
for tk in ["111","112","131","133","152","156","331","511","632","641","642"]:
    for y in sorted([yy for yy,_ in {(a,b) for (a,b) in tk_y} if str(b)==tk], key=lambda x: str(x)):
        print(" ", y, tk, tk_y[(y,tk)])

print("\n=== IN09 ===")
ws = wb["IN09_Kho_GiaThanh"]
print("max", ws.max_row, "hdr", [ws.cell(2,c).value for c in range(1, 16)])
for r in range(2, min(12, (ws.max_row or 2)+1)):
    print(" r", r, [ws.cell(r,c).value for c in range(1, 16)])

print("\n=== IN10 ===")
ws = wb["IN10_Chi_tien_mat"]
print("hdr", [ws.cell(2,c).value for c in range(1, 13)])
n=0; tot=0; yset=set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or row[7] is None and row[0] is None:
        continue
    n+=1; tot+=nnum(row[7]); yset.add(row[1])
print("n", n, "tot", tot, "years", yset)

print("\n=== IN11 IN12 ===")
for sn in ["IN11_Sao_ke_NH","IN12_GDLK"]:
    if sn in wb.sheetnames:
        ws=wb[sn]
        print(sn, "max", ws.max_row, [ws.cell(r,c).value for r in range(1,4) for c in range(1,8)])

print("\n=== IN13 ===")
ws = wb["IN13_So_511_632"]
print("hdr", [ws.cell(2,c).value for c in range(1, 18)])
print("sample r3", [ws.cell(3,c).value for c in range(1, 18)])
y511=defaultdict(lambda: {"n":0,"co":0,"no":0})
y632=defaultdict(lambda: {"n":0,"co":0,"no":0})
n=0
tks=set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row:
        continue
    n+=1
    nam=row[0]; tk=str(row[4] or "")
    tks.add(tk[:3] if tk else "")
    if tk.startswith("511"):
        y511[nam]["n"]+=1; y511[nam]["co"]+=nnum(row[11]); y511[nam]["no"]+=nnum(row[12])
    elif tk.startswith("632"):
        y632[nam]["n"]+=1; y632[nam]["co"]+=nnum(row[11]); y632[nam]["no"]+=nnum(row[12])
print("n", n, "tks", tks)
print("511", dict(y511))
print("632", dict(y632))
wb.close()

print("\n=== KQ1.3 rules ===")
wb = load_workbook(lp(KQ13), read_only=True, data_only=True)
print("sheets", wb.sheetnames)
ws = wb["KQ1_Kha_nang_phan_tich"]
print("hdr r9/10", [ws.cell(9,c).value for c in range(1,12)])
print("hdr r10", [ws.cell(10,c).value for c in range(1,12)])
n=0
for r in range(8, 95):
    ma=ws.cell(r,1).value
    if ma:
        n+=1
        if n<=5 or n>75:
            print(r, [ws.cell(r,c).value for c in range(1,11)])
print("n rules", n)
wb.close()

print("\n=== MAU KQ2.2 511 ===")
wb = load_workbook(lp(os.path.join(MAU,"KQ2.2_Phan_tich_ho_so_cty.xlsx")), data_only=False)
ws=wb["KQ2_511_632"]
for r in range(1, 20):
    vals=[]
    for c in range(1, 12):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:50]}")
    if vals:
        print("R", r, vals)
print("\n=== MAU KQ2.2 Kho ===")
ws=wb["KQ2_Kho_Gia_thanh"]
for r in range(1, 15):
    vals=[]
    for c in range(1, 14):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:40]}")
    if vals:
        print("R", r, vals)
print("\n=== MAU Cong_no ===")
ws=wb["KQ2_Cong_no_Tien_mat"]
for r in range(1, 15):
    vals=[]
    for c in range(1, 14):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:40]}")
    if vals:
        print("R", r, vals)
print("\n=== MAU Gia_thanh ===")
ws=wb["KQ2_Gia_thanh_va_Dinh_muc"]
for r in range(1, 15):
    vals=[]
    for c in range(1, 14):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:40]}")
    if vals:
        print("R", r, vals)
wb.close()

print("\n=== MAU KQ2.1 NCC/TNCN/GTGT row layout ===")
wb = load_workbook(lp(os.path.join(MAU,"KQ2.1_Phan_tich_CQT.xlsx")), data_only=False)
ws=wb["KQ2_NCC_mua_vao"]
for r in range(1, 12):
    vals=[]
    for c in range(1, 15):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:50]}")
    if vals:
        print("NCC R", r, vals)
ws=wb["KQ2_TNCN_LD2"]
for r in range(1, 12):
    vals=[]
    for c in range(1, 10):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:50]}")
    if vals:
        print("TNCN R", r, vals)
ws=wb["IN07_NCC_CanhBao"]
for r in range(1, 6):
    vals=[]
    for c in range(1, 19):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:40]}")
    if vals:
        print("IN07 R", r, vals)
ws=wb["KQ2_Doi_chieu_GTGT"]
print("GTGT max", ws.max_row, ws.max_column)
for r in [1,2,3,7,8,9,10,11,12,13,154,155,156,157,165,166]:
    vals=[]
    for c in range(1, 25):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:40]}")
    if vals:
        print("GTGT R", r, vals)
wb.close()

print("\n=== MAU KQ3 headers ===")
wb = load_workbook(lp(os.path.join(MAU,"KQ3_Ket_luan_danh_gia.xlsx")), data_only=False)
ws=wb["KQ3"]
print("r1-8")
for r in range(1, 11):
    vals=[]
    for c in range(1, 40):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:30]}")
    if vals:
        print("R", r, vals)
ws=wb["Danh_gia_28_tieu_chi_rui_ro"]
for r in range(1, 6):
    vals=[]
    for c in range(1, 10):
        v=ws.cell(r,c).value
        if v is not None:
            vals.append(f"{c}:{str(v)[:50]}")
    if vals:
        print("TC R", r, vals)
wb.close()
print("DONE")
