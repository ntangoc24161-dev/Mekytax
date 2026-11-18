# -*- coding: utf-8 -*-
from __future__ import annotations
import os, sys
from collections import defaultdict, Counter
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

wb = load_workbook(lp(KQ11), read_only=True, data_only=True)
print("IN03 sheets ok")
ws = wb["IN03_ToKhai_GTGT"]
print("max", ws.max_row)
loai=Counter(); ky=[]
for i, row in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True), 1):
    print("r", i, row[:12])
for row in ws.iter_rows(min_row=3, values_only=True):
    loai[str(row[2])] += 1
    if row[0]:
        ky.append((row[0], row[1], row[2], row[3], row[4], row[16] if len(row)>16 else None, row[7] if len(row)>7 else None))
print("loai", loai)
print("n data", len(ky))
print("first 12", ky[:12])
print("last 8", ky[-8:])

print("\nIN01 Tinh_vao_tong / unique")
ws = wb["IN01_HDDT"]
tinh=Counter(); keys_ban=Counter(); keys_mua=Counter()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    tinh[str(row[22])] += 1
    k = (str(row[4] or ""), str(row[5] or ""), str(row[11] or ""))
    if row[0]=="BAN":
        keys_ban[k]+=1
    else:
        keys_mua[k]+=1
print("tinh", tinh)
print("unique ban", len(keys_ban), "dup keys", sum(1 for n in keys_ban.values() if n>1), "max", max(keys_ban.values()) if keys_ban else 0)
print("unique mua", len(keys_mua), "dup keys", sum(1 for n in keys_mua.values() if n>1), "max", max(keys_mua.values()) if keys_mua else 0)
# sum unique BAN 2023-2025
from collections import defaultdict
ban_u=defaultdict(float); mua_u=defaultdict(float)
seen=set()
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or not row[0]:
        continue
    k=(row[0], str(row[4] or ""), str(row[5] or ""), str(row[11] or ""), row[1])
    if k in seen:
        continue
    seen.add(k)
    if row[0]=="BAN":
        ban_u[row[1]] += nnum(row[16])
    else:
        mua_u[row[1]] += nnum(row[16])
print("unique BAN gt", dict(ban_u))
print("unique MUA gt", dict(mua_u))
wb.close()

wb = load_workbook(lp(KQ12), read_only=True, data_only=True)
print("\nIN08 TK 511/632/133/111")
ws = wb["IN08_So_KeToan"]
tk_y=defaultdict(lambda: {"no":0.0,"co":0.0,"n":0,"dk":0.0,"ckn":0.0,"ckc":0.0})
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or row[2] is None:
        continue
    nam=row[0]; tk=str(row[2])[:3]
    rec=tk_y[(nam,tk)]
    rec["no"] += nnum(row[3]); rec["co"] += nnum(row[4]); rec["n"] += 1
    rec["dk"] += nnum(row[5]); rec["ckn"] += nnum(row[6]); rec["ckc"] += nnum(row[7])
for tk in ["111","112","131","133","152","156","331","511","632","641","642"]:
    print("---", tk)
    for (y,t),v in sorted(tk_y.items(), key=lambda x: (str(x[0][0]), x[0][1])):
        if t==tk:
            print(" ", y, {k: round(vv,0) if isinstance(vv,float) else vv for k,vv in v.items()})

print("\nIN09")
ws=wb["IN09_Kho_GiaThanh"]
print("max", ws.max_row)
for r in range(1, min(12,(ws.max_row or 1)+1)):
    print(r, [ws.cell(r,c).value for c in range(1,16)])

print("\nIN10 years")
ws=wb["IN10_Chi_tien_mat"]
n=0; tot=0; ys=Counter(); over=0
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row:
        continue
    amt=nnum(row[7])
    if row[0] is None and amt==0:
        continue
    n+=1; tot+=amt; ys[row[1]]+=1
    if amt>=20_000_000:
        over+=1
print("n", n, "tot", tot, "over20tr", over, "years", dict(ys))

print("\nIN11/12")
for sn in ["IN11_Sao_ke_NH","IN12_GDLK"]:
    ws=wb[sn]
    print(sn, "max", ws.max_row)
    for r in range(1, min(5,(ws.max_row or 1)+1)):
        print(" ", r, [ws.cell(r,c).value for c in range(1,8)])

print("\nIN13")
ws=wb["IN13_So_511_632"]
y511=defaultdict(lambda: {"n":0,"co":0,"no":0})
y632=defaultdict(lambda: {"n":0,"co":0,"no":0})
tks=Counter(); n=0
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or row[4] is None:
        continue
    n+=1
    nam=row[0]; tk=str(row[4])
    tks[tk[:3]] += 1
    if tk.startswith("511"):
        y511[nam]["n"]+=1; y511[nam]["co"]+=nnum(row[11]); y511[nam]["no"]+=nnum(row[12])
    elif tk.startswith("632"):
        y632[nam]["n"]+=1; y632[nam]["co"]+=nnum(row[11]); y632[nam]["no"]+=nnum(row[12])
print("n", n, "tks", dict(tks))
print("511", {k:dict(v) for k,v in y511.items()})
print("632", {k:dict(v) for k,v in y632.items()})
print("sample", [ws.cell(3,c).value for c in range(1,18)])
wb.close()

wb=load_workbook(lp(KQ13), read_only=True, data_only=True)
ws=wb["KQ1_Kha_nang_phan_tich"]
print("\nKQ1.3 rules header")
for r in range(7, 14):
    print(r, [ws.cell(r,c).value for c in range(1,12)])
n=0
for r in range(8, 100):
    if ws.cell(r,1).value:
        n+=1
print("n", n)
print("kiem ke C4", wb["KQ1_Kiem_ke"].cell(4,3).value)
wb.close()

wb=load_workbook(lp(os.path.join(MAU,"KQ2.2_Phan_tich_ho_so_cty.xlsx")))
print("\n511_632")
ws=wb["KQ2_511_632"]
for r in range(1,22):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:45]}" for c in range(1,12) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
print("\nKho")
ws=wb["KQ2_Kho_Gia_thanh"]
for r in range(1,16):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:40]}" for c in range(1,14) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
print("\nGia thanh")
ws=wb["KQ2_Gia_thanh_va_Dinh_muc"]
for r in range(1,16):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:40]}" for c in range(1,14) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
print("\nCong no")
ws=wb["KQ2_Cong_no_Tien_mat"]
for r in range(1,16):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:40]}" for c in range(1,14) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
print("Tong hop co KQ2.2")
ws=wb["KQ2_Tong_hop_co"]
for r in range(1,20):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:50]}" for c in range(1,10) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
wb.close()

wb=load_workbook(lp(os.path.join(MAU,"KQ2.1_Phan_tich_CQT.xlsx")))
print("\nNCC")
ws=wb["KQ2_NCC_mua_vao"]
for r in range(1,12):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:50]}" for c in range(1,15) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
print("\nTNCN")
ws=wb["KQ2_TNCN_LD2"]
for r in range(1,12):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:50]}" for c in range(1,10) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
print("\nIN07")
ws=wb["IN07_NCC_CanhBao"]
for r in range(1,6):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:40]}" for c in range(1,19) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
print("\nGTGT key rows")
ws=wb["KQ2_Doi_chieu_GTGT"]
print("max", ws.max_row, ws.max_column)
for r in [1,2,3,7,8,9,10,11,154,155,156,165,166]:
    vals=[f"{c}:{str(ws.cell(r,c).value)[:40]}" for c in range(1,26) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
print("\nTong hop co KQ2.1")
ws=wb["KQ2_Tong_hop_co"]
for r in range(1,25):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:55]}" for c in range(1,8) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
wb.close()

wb=load_workbook(lp(os.path.join(MAU,"KQ3_Ket_luan_danh_gia.xlsx")))
ws=wb["KQ3"]
print("\nKQ3 header r1-9")
for r in range(1,11):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:28]}" for c in range(1,40) if ws.cell(r,c).value is not None]
    if vals: print("R",r,vals)
ws=wb["Danh_gia_28_tieu_chi_rui_ro"]
for r in range(1,6):
    vals=[f"{c}:{str(ws.cell(r,c).value)[:50]}" for c in range(1,10) if ws.cell(r,c).value is not None]
    if vals: print("TC R",r,vals)
wb.close()
print("DONE")
