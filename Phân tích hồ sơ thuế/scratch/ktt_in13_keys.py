# -*- coding: utf-8 -*-
import os, sys
from collections import Counter, defaultdict
from openpyxl import load_workbook
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
p=r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\KIEN_THANH_TAN_3603216935\KQ1.2_Kiem_ke_ho_so_NNT_3603216935_2015-2025_v1.xlsx"
s=os.path.abspath(p)
wb=load_workbook("\\\\?\\"+s, read_only=True, data_only=True)
ws=wb["IN13_So_511_632"]
rows=[]
for row in ws.iter_rows(min_row=3, values_only=True):
    if not row or row[4] is None: continue
    nam=row[0]
    if nam==2026: continue
    rows.append(row)
print("n", len(rows))
print("hdr sample 511")
n511=n632=0
for row in rows:
    tk=str(row[4])
    if tk.startswith("511") and n511<5:
        print("511", row[:15]); n511+=1
    if tk.startswith("632") and n632<5:
        print("632", row[:15]); n632+=1
# unique by all
uniq=set()
for row in rows:
    uniq.add(tuple(str(x) for x in row[:15]))
print("unique tuples", len(uniq), "raw", len(rows))
# keys
c511=Counter(); c632=Counter()
for row in rows:
    k=(row[0], str(row[3] or ""), str(row[2] or "")[:10], str(row[7] or "")[:40])
    if str(row[4]).startswith("511"): c511[k]+=1
    else: c632[k]+=1
print("511 keys", len(c511), "dup", sum(1 for n in c511.values() if n>1), "max", max(c511.values()))
print("632 keys", len(c632), "dup", sum(1 for n in c632.values() if n>1), "max", max(c632.values()))
inter=set(c511)&set(c632)
print("intersect so+ngay+dg", len(inter))
k511=set((r[0], str(r[3] or "")) for r in rows if str(r[4]).startswith("511"))
k632=set((r[0], str(r[3] or "")) for r in rows if str(r[4]).startswith("632"))
print("intersect nam+so_ct", len(k511&k632), "only511", len(k511-k632), "only632", len(k632-k511))
print("sample 511 so", list(k511)[:8])
print("sample 632 so", list(k632)[:8])
wb.close()
