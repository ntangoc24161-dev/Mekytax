import sys
from gd1_common import read_rows

FILES = [
    "2. HO SO NNT CUNG CAP/cty gui/NĂM 2021/Bảng kê mua vào 2021.xlsx",
    "2. HO SO NNT CUNG CAP/cty gui/SỔ TỔNG HỢP/Bảng kê bán ra năm 2024.xlsx",
    "2. HO SO NNT CUNG CAP/2.5. Bang ke mua vao ban ra/BẢNG KÊ MUA VÀO QIV.25.xlsx",
]
for rel in FILES:
    rows = read_rows(rel)
    print("==", rel, "rows:", len(rows))
    for i, r in enumerate(rows[:14]):
        print(i + 1, [c for c in r][:15])
    print("... tail")
    for i, r in enumerate(rows[-5:], start=len(rows) - 4):
        print(i, [c for c in r][:15])
