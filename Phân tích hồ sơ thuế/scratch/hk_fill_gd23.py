# -*- coding: utf-8 -*-
"""GĐ2/GĐ3 fill for HUY KHANG - executed in run_huy_khang_v1.py namespace."""
FONT_RED = Font(name="Times New Roman", size=11, color="C00000")
FONT_ORG = Font(name="Times New Roman", size=11, color="B25900")
FONT_GRN = Font(name="Times New Roman", size=11, color="385723")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
P03, P13, P14, P23 = 0.02, 0.05, 0.50, 0.10
YEARS = [2020, 2021, 2022, 2023, 2024, 2025]


def ncc_flag(base: str):
    if not base:
        return "", "", "", ""
    if base in ca_set:
        rec = ca_map.get(base) or {}
        return "F0", "CA", rec.get("cv", ""), rec.get("nd", "")
    if base in f0_set:
        rec = f0_map.get(base) or {}
        return "F0", "CANH_BAO", rec.get("cv", ""), rec.get("nd", "")
    if base in f1_set:
        rec = f1_map.get(base) or {}
        return "F1", "CANH_BAO", rec.get("cv", ""), rec.get("nd", "")
    return "", "", "", ""


def cash_nguong(ngay):
    if ngay is None:
        return 20_000_000
    if ngay >= date(2025, 7, 1):
        return 5_000_000
    return 20_000_000


mua_ok = [x for x in invoices if x["chieu"] == "MUA" and x["tinh"] == "Có"]
ban_ok = [x for x in invoices if x["chieu"] == "BAN" and x["tinh"] == "Có"]

ncc_agg = defaultdict(lambda: {"ten": "", "gt": 0.0, "thue": 0.0, "n": 0, "n_hd": set()})
for inv in mua_ok:
    k = mst_base(inv["mst_nb"])
    if not k:
        continue
    a = ncc_agg[k]
    a["ten"] = a["ten"] or inv["ten_nb"]
    a["gt"] += inv["gt"]
    a["thue"] += inv["thue"]
    a["n"] += 1
    a["n_hd"].add((inv["kyhieu"], inv["so"]))

tot_mua_gt = sum(a["gt"] for a in ncc_agg.values())
ranked = sorted(ncc_agg.items(), key=lambda x: x[1]["gt"], reverse=True)
top5 = ranked[:5]
top5_gt = sum(a["gt"] for _, a in top5)
top5_pct = (top5_gt / tot_mua_gt) if tot_mua_gt else 0.0

f0_ncc = [(k, a) for k, a in ranked if ncc_flag(k)[0] == "F0"]
f1_ncc = [(k, a) for k, a in ranked if ncc_flag(k)[0] == "F1"]
f0_gt = sum(a["gt"] for _, a in f0_ncc)
f0_thue = sum(a["thue"] for _, a in f0_ncc)
f1_gt = sum(a["gt"] for _, a in f1_ncc)
risk_gt = f0_gt + f1_gt
hd03_pct = (risk_gt / tot_mua_gt) if tot_mua_gt else 0.0

cash_hits = []
for inv in mua_ok:
    nguong = cash_nguong(inv["ngay"])
    amt = inv["tong"] if inv["tong"] else (inv["gt"] + inv["thue"])
    if amt >= nguong:
        cash_hits.append(inv)

# MST format
def mst_ok(s):
    d = re.sub(r"\D", "", s or "")
    return len(d) in (10, 13)

n_mst_sai = sum(1 for x in invoices if x["tinh"] == "Có" and not mst_ok(x["mst_nb"] if x["chieu"] == "MUA" else x["mst_nm"]))
n_thieu_kh = sum(1 for x in invoices if x["tinh"] == "Có" and not x["kyhieu"])
dup_keys = Counter((x["kyhieu"], x["so"], x["chieu"], mst_base(x["mst_nb"] if x["chieu"] == "MUA" else x["mst_nm"])) for x in invoices if x["tinh"] == "Có")
n_dup = sum(1 for k, n in dup_keys.items() if n > 1 and k[1])

print("NCC unique", len(ncc_agg), "F0", len(f0_ncc), "F1", len(f1_ncc), "top5", f"{top5_pct:.2%}", "cash>=nguong", len(cash_hits))


def tndn_dt(nam):
    t = tndn_by.get(nam)
    if not t:
        return None
    d = t.chi_tieu or {}
    for c in ("A1", "01", "1", "a1"):
        if c in d and isinstance(d[c], (int, float)):
            return float(d[c])
    return None


def gtgt_ct_year(nam, code):
    v = gtgt_year.get(nam, {}).get(f"ct{code}")
    return v if v else None


# ===========================================================================
# KQ2.1
# ===========================================================================
print("FILL KQ2.1")
wb = load_workbook(lp(paths["kq21"]))

ws = wb["KQ2_Tong_hop_co"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Phiên bản KQ: {VER} | Dữ liệu công ty: Chưa có", bold=True)

ws = wb["KQ2_Doi_chieu_GTGT"]
put(ws, 2, 1, f"Người nộp thuế: {TEN} - MST: {MST} - Kỳ: {KY}", bold=True)
put(ws, 3, 1, f"Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Phiên bản KQ: {VER} | Dữ liệu công ty: Chưa có | Nguồn: tờ khai GTGT IN03 (CQT), HĐĐT IN01 (CQT), BCTC IN04 (CQT). Bảng kê IN02 và sổ IN08 chưa nhận.")

period_off = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}


def year_header_row(y):
    return 9 + (y - 2016) * 14


for k, t in sorted(eff_gtgt.items(), key=lambda x: (x[1].nam, x[0])):
    y = t.nam
    if y < 2016 or y > 2025:
        continue
    ky = ky_norm(t.ky, t.nam).upper()
    mq = re.search(r"Q([1-4])", ky)
    if not mq:
        continue
    rr = year_header_row(y) + int(mq.group(1))
    dt32 = ct(t, "32")
    dt34 = ct(t, "34")
    th33 = ct(t, "33")
    th35 = ct(t, "35")
    # DT 8-10% col E; thuế 8-10% col I; mua K; thuế ĐV L
    if dt32:
        put(ws, rr, 5, dt32, NUM)
    elif dt34:
        put(ws, rr, 5, dt34, NUM)
    if th33:
        put(ws, rr, 9, th33, NUM)
    elif th35:
        put(ws, rr, 9, th35, NUM)
    put(ws, rr, 11, ct(t, "23"), NUM)
    put(ws, rr, 12, ct(t, "24"), NUM)
    if ct(t, "37"):
        put(ws, rr, 13, ct(t, "37"), NUM)
    if ct(t, "38"):
        put(ws, rr, 14, ct(t, "38"), NUM)
    if ct(t, "40"):
        put(ws, rr, 17, ct(t, "40"), NUM)
    if ct(t, "43"):
        put(ws, rr, 18, ct(t, "43"), NUM)

# Phần II/III yearly 2016-2025 at rows 156-165. Do NOT delete (Tong_hop_co hardcoded).
# Leave 2016-2019 empty (ngoài kỳ). Do not fill 0 for missing BK/NKC.
for y in YEARS:
    rr = 156 + (y - 2016)
    hy = hddt_year.get(y)
    if hy and hy["ban_n"]:
        put(ws, rr, 2, hy["ban_gt"], NUM)
        put(ws, rr, 3, hy["ban_thue"], NUM)
    # D/E bảng kê: trống
    gy = gtgt_year.get(y)
    if gy:
        if gy.get("ct34") or gy.get("ct32"):
            put(ws, rr, 6, gy.get("ct34") or gy.get("ct32"), NUM)
        if gy.get("ct35") or gy.get("ct33"):
            put(ws, rr, 7, gy.get("ct35") or gy.get("ct33"), NUM)
        if gy.get("ct23"):
            put(ws, rr, 15, gy.get("ct23"), NUM)  # O = GT TK [23]
        if gy.get("ct24"):
            put(ws, rr, 21, gy.get("ct24"), NUM)  # U = thuế TK [24]
    td = tndn_dt(y)
    if td is not None:
        put(ws, rr, 8, td, NUM)
    b01 = bctc_val(y, "01", which="kqkd")
    if b01 is None:
        b01 = bctc_val(y, "01", which="any")
    if b01 is not None:
        put(ws, rr, 9, b01, NUM)
    put(ws, rr, 12, y)  # L = năm Phần III
    if hy and hy["mua_n"]:
        put(ws, rr, 13, hy["mua_gt"], NUM)  # M
        put(ws, rr, 19, hy["mua_thue"], NUM)  # S
    # N BK, T BK, Y 133: trống

notes = [
    "Nhận xét / nguyên nhân trọng yếu (module GTGT, không gán mức rủi ro hồ sơ)",
    "NNT kê khai GTGT theo quý. Phần I lấy tờ khai hiệu lực (ưu tiên bổ sung/soLan cao) 2020-2025; Q1-Q2/2020 không có XML - để trống, không điền 0.",
    "Phần II/III: HĐĐT bán ra từ 21/04/2022, mua vào từ 16/12/2021. Bảng kê IN02 và sổ 133 IN08 chưa nhận nên cột BK/NKC trống, cờ khớp 3 nguồn = Chưa đủ nguồn.",
    "XML GTGT Q1-Q3/2026 có trong hồ sơ nhưng mẫu Phần I chỉ đến 2025 - không dựng khối năm mới. Đối chiếu 2026 ghi ở KQ3 nếu cần.",
]
for y in YEARS:
    tk34 = (gtgt_year.get(y) or {}).get("ct34") or (gtgt_year.get(y) or {}).get("ct32")
    hd = (hddt_year.get(y) or {}).get("ban_gt")
    mua_tk = (gtgt_year.get(y) or {}).get("ct23")
    mua_hd = (hddt_year.get(y) or {}).get("mua_gt")
    b01 = bctc_val(y, "01", which="any")
    bits = [str(y)]
    if tk34 is not None and hd:
        bits.append(f"TK[34]={tk34:,.0f} HĐĐT bán={hd:,.0f} lệch {tk34-hd:,.0f}")
    elif tk34 is not None:
        bits.append(f"TK[34]={tk34:,.0f} HĐĐT bán trống")
    if mua_tk is not None and mua_hd:
        bits.append(f"TK[23]={mua_tk:,.0f} HĐĐT mua={mua_hd:,.0f} lệch {mua_tk-mua_hd:,.0f}")
    if b01 is not None and tk34 is not None:
        bits.append(f"BCTC[01]={b01:,.0f} lệch GTGT {tk34-b01:,.0f}")
    if len(bits) > 1:
        notes.append(" | ".join(bits))
r_note = (ws.max_row or 169) + 2
for i, line in enumerate(notes):
    put(ws, r_note + i, 1, line, bold=(i == 0))
    try:
        ws.merge_cells(start_row=r_note + i, start_column=1, end_row=r_note + i, end_column=10)
    except Exception:
        pass
    if i == 0:
        ws.cell(r_note + i, 1).fill = FILL_SEC

# IN07
ws = wb["IN07_NCC_CanhBao"]
clear_example_from(ws, 3, file_col=18)
r = 3
in07_n = 0
# NNT itself
for base, nhom_force, nguon_force in [
    (MST, "F0" if nnt_on_f0 or nnt_on_ca else ("F1" if nnt_on_f1 else ""), "CA" if nnt_on_ca else ("CANH_BAO" if (nnt_on_f0 or nnt_on_f1) else "")),
]:
    if nhom_force:
        rec = ca_map.get(base) or f0_map.get(base) or f1_map.get(base) or {}
        put(ws, r, 1, MST)
        put(ws, r, 2, TEN)
        put(ws, r, 3, "NNT đang kiểm tra")
        put(ws, r, 4, nhom_force)
        put(ws, r, 5, nguon_force)
        put(ws, r, 6, rec.get("cv", ""))
        put(ws, r, 7, rec.get("ngay_cv"))
        put(ws, r, 8, "Cục Thuế / Công an" if nguon_force == "CA" else "Chi cục Thuế")
        put(ws, r, 9, rec.get("nd", ""))
        put(ws, r, 16, "Sổ CVCBRR / CA")
        put(ws, r, 18, CVCBRR.name if nguon_force != "CA" else CA.name)
        if nhom_force == "F0":
            for c in range(1, 19):
                ws.cell(r, c).fill = FILL_RED
        r += 1
        in07_n += 1
seen = {MST}
for k, a in ranked:
    nhom, nguon, cv, nd = ncc_flag(k)
    if not nhom:
        continue
    if k in seen:
        continue
    seen.add(k)
    rec = ca_map.get(k) or f0_map.get(k) or f1_map.get(k) or {}
    put(ws, r, 1, k)
    put(ws, r, 2, a["ten"] or rec.get("ten", ""))
    put(ws, r, 3, "NCC mua vào")
    put(ws, r, 4, nhom)
    put(ws, r, 5, nguon)
    put(ws, r, 6, cv or rec.get("cv", ""))
    put(ws, r, 7, rec.get("ngay_cv"))
    put(ws, r, 8, "Công an" if nguon == "CA" else "Chi cục Thuế")
    put(ws, r, 9, nd or rec.get("nd", ""))
    put(ws, r, 16, "HĐĐT + sổ CVCBRR/CA")
    put(ws, r, 17, f"GT mua {a['gt']:,.0f}")
    put(ws, r, 18, CVCBRR.name if nguon != "CA" else CA.name)
    if nhom == "F0":
        for c in range(1, 19):
            ws.cell(r, c).fill = FILL_RED
            ws.cell(r, c).font = FONT_RED
    elif nhom == "F1":
        for c in range(1, 19):
            ws.cell(r, c).fill = FILL_YEL
            ws.cell(r, c).font = FONT_ORG
    r += 1
    in07_n += 1
print("IN07 rows", in07_n)
autowidth(ws, max_row=40)

# NCC 100% HĐ mua hợp lệ
ws = wb["KQ2_NCC_mua_vao"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Phiên bản KQ: {VER} | Dữ liệu công ty: Chưa có", bold=True)
if ws.max_row > 8:
    ws.delete_rows(8, ws.max_row - 7)
r = 8
for i, inv in enumerate(mua_ok, 1):
    base = mst_base(inv["mst_nb"])
    nhom, nguon, cv, nd = ncc_flag(base)
    put(ws, r, 1, i)
    put(ws, r, 2, inv["mst_nb"])
    put(ws, r, 3, inv["ten_nb"])
    put(ws, r, 4, inv["so_goc"] if inv["so_goc"] is not None else inv["so"])
    put(ws, r, 5, inv["ngay"])
    put(ws, r, 6, inv["mh"])
    put(ws, r, 7, inv["gt"], NUM)
    put(ws, r, 8, inv["thue"], NUM)
    put(ws, r, 9, "HĐĐT")
    put(ws, r, 10, f"Có - {cv}" if cv else ("Có" if nhom else "Không"))
    put(ws, r, 11, nhom or "Không")
    put(ws, r, 12, "Chưa tra DKKD (thiếu danh bạ 3.2)")
    if nhom == "F0":
        for c in range(1, 13):
            ws.cell(r, c).fill = FILL_RED
            ws.cell(r, c).font = FONT_RED
    elif nhom == "F1":
        for c in range(1, 13):
            ws.cell(r, c).fill = FILL_YEL
            ws.cell(r, c).font = FONT_ORG
    r += 1
last_inv = r - 1
# HD03/HD04 flags in row 5
put(ws, 5, 2, hd03_pct, PCT)
put(ws, 5, 5, top5_pct, PCT)
rr = last_inv + 2
put(ws, rr, 1, "Nhận xét module (3.4) - không phải kết luận hồ sơ", bold=True)
ws.cell(rr, 1).fill = FILL_SEC
top5_txt = "; ".join(f"{k} {a['ten'][:40]} {a['gt']:,.0f}" for k, a in top5)
put(
    ws,
    rr + 1,
    1,
    f"HD04 Top-5 / tổng mua = {top5_pct:.2%} (ngưỡng P14=50%, tạm). Số NCC={len(ncc_agg)}. "
    f"F0={len(f0_ncc)} doanh số {f0_gt:,.0f} thuế {f0_thue:,.0f}. F1={len(f1_ncc)} doanh số {f1_gt:,.0f}. "
    f"HD03 tỷ trọng F0+F1 = {hd03_pct:.2%} (P13=5%). NNT trên F0: {'Có' if nnt_on_f0 else 'Không'}; F1: {'Có' if nnt_on_f1 else 'Không'}; CA: {'Có' if nnt_on_ca else 'Không'}. "
    f"Top-5: {top5_txt}. Nguồn NCC = HĐĐT (IN02 chưa nhận). Danh bạ ĐKKD 3.2 chưa có nên HD02/HD05 chưa chốt.",
)
try:
    ws.merge_cells(start_row=rr + 1, start_column=1, end_row=rr + 1, end_column=12)
except Exception:
    pass
ws.auto_filter.ref = f"A7:L{max(7, last_inv)}"
autowidth(ws, max_row=20)
print("NCC invoice rows", last_inv - 7)

# TNCN
ws = wb["KQ2_TNCN_LD2"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Phiên bản KQ: {VER} | Dữ liệu công ty: Chưa có", bold=True)
# clear example years 8+
for rr in range(8, 20):
    if ws.cell(rr, 1).value in {2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, None}:
        for c in range(1, 8):
            if rr >= 8:
                pass
# write QTT years starting row 8
# delete leftover example
max_clear = 8
for rr in range(8, 30):
    v = ws.cell(rr, 1).value
    if isinstance(v, (int, float)) or v in {"(ví dụ)", "ví dụ"}:
        for c in range(1, 8):
            ws.cell(rr, c).value = None
rr = 8
for y in YEARS:
    t = tncn_by.get(y)
    nld_y = [x for x in nld_rows if x.get("nam") == y]
    nld_52 = [x for x in nld_y if x.get("nguon") == "05-2"]
    cccd = [x.get("cccd") for x in nld_y if x.get("cccd")]
    dup = len(cccd) - len(set(cccd)) if cccd else 0
    no10 = sum(1 for x in nld_52 if (x.get("thue") or 0) == 0 and (x.get("tnct") or 0) > 0)
    put(ws, rr, 1, y)
    if t:
        put(ws, rr, 2, int(ct(t, "21") or 0) or (len(nld_y) or None), NUM)
        put(ws, rr, 4, ct(t, "23") or ct(t, "26") or ct(t, "22"), NUM)
        put(ws, rr, 5, ct(t, "32") or ct(t, "26") or ct(t, "33") or ct(t, "36"), NUM)
    put(ws, rr, 3, len(nld_52) or None, NUM)
    put(ws, rr, 6, no10 or None, NUM)
    put(ws, rr, 7, dup or None, NUM)
    rr += 1
put(ws, rr + 1, 1, "Nhận xét module TNCN (không ghi mức rủi ro hồ sơ)", bold=True)
ws.cell(rr + 1, 1).fill = FILL_SEC
put(
    ws,
    rr + 2,
    1,
    f"QTT TNCN XML các năm {sorted(tncn_by)}. Dòng NLĐ bóc từ phụ lục 05-1/05-2: {len(nld_rows)} dòng. "
    f"Thiếu file tra cứu đa nguồn thư mục 3.3 nên LD04 ngoài QTT chưa chốt. Tờ khai TNCN quý không có trong 1.1 - LD01 (QTT vs tạm tính quý) = KHÔNG_ĐỦ_DỮ_LIỆU.",
)

# EBITDA
if "KQ2_EBITDA_Phan_tich" not in wb.sheetnames:
    ews = wb.create_sheet("KQ2_EBITDA_Phan_tich", 1)
else:
    ews = wb["KQ2_EBITDA_Phan_tich"]
ews.sheet_view.showGridLines = True
put(ews, 1, 1, "KQ2_EBITDA_Phan_tich - NĐ 132/2020 (30% EBITDA) - số liệu BCTC CQT, chưa có sổ 635/khấu hao", bold=True)
ews.cell(1, 1).fill = FILL_NAVY
ews.cell(1, 1).font = FONT_W
try:
    ews.merge_cells("A1:H1")
except Exception:
    pass
headers = ["Chỉ tiêu"] + [str(y) for y in YEARS] + ["Ghi chú"]
for i, h in enumerate(headers, 1):
    put(ews, 3, i, h, bold=True)
    ews.cell(3, i).fill = FILL_NAVY
    ews.cell(3, i).font = FONT_W
rows_eb = [
    "Lợi nhuận trước thuế [50]",
    "Chi phí lãi vay ([23] / 635)",
    "Khấu hao TSCĐ",
    "EBITDA",
    "Trần 30% EBITDA",
    "Vượt ngưỡng lãi vay",
    "Có GDLK (IN12)",
]
for i, name in enumerate(rows_eb):
    put(ews, 4 + i, 1, name)
    for j, y in enumerate(YEARS):
        if i == 0:
            v = bctc_val(y, "50", which="kqkd") or bctc_val(y, "50", which="any")
            put(ews, 4 + i, 2 + j, v, NUM)
        elif i == 1:
            v = bctc_val(y, "23", which="kqkd") or bctc_val(y, "23", which="any")
            put(ews, 4 + i, 2 + j, v, NUM)
        elif i == 6:
            put(ews, 4 + i, 2 + j, "Chưa đủ nguồn IN12")
put(ews, 12, 1, "Nhận xét: lãi vay lấy BCTC mã [23] nếu có; thiếu khấu hao chi tiết nên EBITDA chưa chốt. IN12 chưa nhận. Không kết luận 30% ở KQ2.")
apply_font_sheet(wb["KQ2_Doi_chieu_GTGT"])
wb.save(lp(paths["kq21"]))
wb.close()
print("saved KQ2.1")

# ===========================================================================
# KQ2.2 - company empty (T10)
# ===========================================================================
print("FILL KQ2.2")
wb = load_workbook(lp(paths["kq22"]))
ws = wb["KQ2_Tong_hop_co"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Phiên bản KQ: {VER} | Dữ liệu công ty: Chưa có", bold=True)

ws = wb["KQ2_511_632"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | Phiên bản biểu mẫu: 1.4 | Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Phiên bản KQ: {VER} | Dữ liệu công ty: Chưa có", bold=True)
# wipe example zeros (Q18) - leave structure, empty numbers
for r in range(6, 17):
    for c in range(2, 8):
        cell = ws.cell(r, c)
        if isinstance(cell, MergedCell):
            continue
        val = cell.value
        if val in {0, 0.0, "(ví dụ - xóa khi chạy hồ sơ thật)"} or (isinstance(val, str) and "ví dụ" in val.lower()):
            cell.value = None
put(ws, 6, 7, "Chưa đủ nguồn IN13 - thư mục 2. HO SO NNT CUNG CAP trống. Không điền 0.")
put(ws, 16, 1, "Nhận xét module: chưa có sổ 511/632. DT01/DT02/DT03 = KHÔNG_ĐỦ_DỮ_LIỆU. Không ước lượng lãi gộp từ BCTC thay sổ.")

if "KQ2_511_632_Chi_tiet" in wb.sheetnames:
    cws = wb["KQ2_511_632_Chi_tiet"]
    clear_example_from(cws, 2)
    put(cws, 2, 1, "Chưa đủ nguồn IN13")
    put(cws, 2, 9, "KHÔNG_ĐỦ_DỮ_LIỆU")

ws = wb["KQ2_Kho_Gia_thanh"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | {VER} | Dữ liệu công ty: Chưa có", bold=True)
put(ws, 7, 12, "Chưa nhận IN09 (XNT 152/155/156). Không điền 0. TC3/KH04 chờ sổ kho; BCTC ct140/ct141 ghi ở KQ3 khi đọc được.")
put(ws, 20, 1, "Nhận xét TC3: chưa có XNT. Mẫu BCTC trên XML ghi sơ bộ, chưa đọc chữ in B01-DN/B01-DNN.")

ws = wb["KQ2_Gia_thanh_va_Dinh_muc"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | {VER} | Dữ liệu công ty: Chưa có", bold=True)
for j, y in enumerate(YEARS):
    put(ws, 7 + j, 1, y)
    dt = bctc_val(y, "01", which="any")
    gv = bctc_val(y, "02", which="any")
    if dt is not None or gv is not None:
        put(ws, 7 + j, 11, f"BCTC [01]={dt} [02]={gv} (CQT, không thay sổ 632)")
put(ws, 20, 1, "Nhận xét: thiếu sổ 621/622/627/154 và 632. Tỷ lệ GV/DT BCTC chỉ tham khảo, kết luận TC28 ở KQ3 nhãn sơ bộ.")

ws = wb["KQ2_Cong_no_Tien_mat"]
put(ws, 2, 1, f"MST: {MST} | Kỳ: {KY} | Giai đoạn: 2 | {VER} | Dữ liệu công ty: Chưa có", bold=True)
# BCTC [01] and [331] are CQT - fill those only
for j, y in enumerate(YEARS):
    rr = 8 + j
    put(ws, rr, 1, y)
    b01 = bctc_val(y, "01", which="any")
    v331 = bctc_val(y, "331", which="any")
    if b01 is not None:
        put(ws, rr, 11, b01, NUM)
    if v331 is not None:
        put(ws, rr, 9, v331, NUM)
put(ws, 40, 1, f"TT01 sơ bộ từ HĐĐT: {len(cash_hits)} HĐ mua đạt ngưỡng K5 (>=20tr đến 30/06/2025, >=5tr từ 01/07/2025). HĐĐT không có cột hình thức TT - không kết luận đã thanh toán tiền mặt. IN10/IN11 chưa nhận: TT02/GT16 = KHÔNG_ĐỦ_DỮ_LIỆU, không kết luận thiếu UNC.")
put(ws, 41, 1, "IN08/IN10/IN11 trống. Dư 331 lấy BCTC [331] nếu XML có mã; biến động 131/331 để trống.")

wb.save(lp(paths["kq22"]))
wb.close()
print("saved KQ2.2")

# ===========================================================================
# KQ3
# ===========================================================================
print("FILL KQ3")
wb = load_workbook(lp(paths["kq3"]))
ws = wb["KQ3"]
put(ws, 2, 2, MST, bold=True)
put(ws, 2, 7, TEN, bold=True)
put(ws, 3, 2, KY)
put(ws, 3, 7, "A", bold=True)
put(ws, 4, 2, VER)
put(ws, 4, 7, "1.4")
put(ws, 5, 2, TODAY)
put(ws, 5, 7, f"KQ1.3 {VER}; KQ2.1 {VER}; KQ2.2 {VER} (công ty chưa có)")
put(ws, 7, 7, "Chưa có")

# catalog from KQ1.3
cat = []
wb13 = load_workbook(lp(paths["kq13"]), data_only=True)
wsk = wb13["KQ1_Kha_nang_phan_tich"]
for r in range(10, 90):
    ma = wsk.cell(r, 1).value
    if not ma:
        continue
    cat.append(
        {
            "ma": str(ma),
            "nhom": wsk.cell(r, 2).value,
            "ten": wsk.cell(r, 3).value,
            "need": str(wsk.cell(r, 4).value or ""),
            "cty": str(wsk.cell(r, 5).value or ""),
            "tep": wsk.cell(r, 6).value,
            "sheet": wsk.cell(r, 7).value,
            "kn": wsk.cell(r, 8).value,
            "thieu": wsk.cell(r, 10).value,
        }
    )
wb13.close()

HAVE = {"IN01", "IN03", "IN04", "IN05", "IN07", "IN14"}
GD3_ONLY = {"DQ01", "GT06", "GT16", "TN04", "TN17", "TN18", "TT02", "TT03", "TT04", "TT05", "TT06", "KH01", "KH02", "KH03", "KH04", "KH05", "KH06", "KH07", "DT01", "DT02", "DT03", "LD02", "LD05"}


def year_diffs(get_a, get_b, full_years=None):
    out = []
    for y in YEARS:
        a, b = get_a(y), get_b(y)
        if a is None or b is None:
            continue
        full = True if full_years is None else y in full_years
        out.append((y, a, b, a - b, full))
    return out


def pick_diff(diffs, p03=P03):
    """Largest relative gap on fully covered years; else largest on partial (so bo)."""
    full = [x for x in diffs if x[4]]
    pool = full or diffs
    if not pool:
        return None
    y, a, b, d, full = max(pool, key=lambda x: abs(x[3]))
    base = max(abs(a), abs(b), 1)
    vuot = abs(d) / base > p03
    return y, a, b, d, full, vuot


def hd_ban(y):
    h = hddt_year.get(y)
    return h["ban_gt"] if h and h["ban_n"] else None


def hd_mua(y):
    h = hddt_year.get(y)
    return h["mua_gt"] if h and h["mua_n"] else None


def hd_ban_thue(y):
    h = hddt_year.get(y)
    return h["ban_thue"] if h and h["ban_n"] else None


def hd_mua_thue(y):
    h = hddt_year.get(y)
    return h["mua_thue"] if h and h["mua_n"] else None


def tk34(y):
    g = gtgt_year.get(y) or {}
    return g.get("ct34") or g.get("ct32") or None


def tk35(y):
    g = gtgt_year.get(y) or {}
    return g.get("ct35") or g.get("ct33") or None


def tk23(y):
    return (gtgt_year.get(y) or {}).get("ct23") or None


def tk24(y):
    return (gtgt_year.get(y) or {}).get("ct24") or None


# HĐĐT bán 21/04/2022-31/12/2025; mua 16/12/2021-31/12/2025 - năm phủ đủ
BAN_FULL, MUA_FULL = {2023, 2024, 2025}, {2022, 2023, 2024, 2025}
d_gt01 = year_diffs(tk34, hd_ban, BAN_FULL)
d_gt03 = year_diffs(tk35, hd_ban_thue, BAN_FULL)
d_gt04 = year_diffs(tk23, hd_mua, MUA_FULL)
d_gt05 = year_diffs(tk24, hd_mua_thue, MUA_FULL)
d_tn01 = year_diffs(tk34, lambda y: bctc_val(y, "01", which="any"))
d_tn02 = year_diffs(tndn_dt, lambda y: bctc_val(y, "01", which="any"))

rows_kq3 = []
findings = []  # (ma, muc, text, so)

for rd in cat:
    ma = rd["ma"]
    need = rd["need"]
    ins = set(re.findall(r"IN\d{2}", need.replace("–", "-").replace("—", "-")))
    if "IN01-IN08" in need.replace("–", "-") or "IN01–IN08" in need:
        ins |= {f"IN{i:02d}" for i in range(1, 9)}
    missing = sorted(i for i in ins if i not in HAVE)
    # IN02 chưa nhận: HĐĐT vẫn chạy được các quy tắc đối chiếu CQT / NCC
    if "IN02" in missing and ma in {"HD01", "HD03", "HD04", "HD06", "HD08", "HD09", "GT01", "GT03", "GT04", "GT05", "TT01", "DQ03", "DQ04", "DQ06"}:
        missing = [x for x in missing if x != "IN02"]
    status = "ĐÃ_CHẠY"
    ket = "Không"
    nhan = "Chắc"
    muc_rr = "Thấp"
    diem = 0
    va = vb = do = None
    ca_a = cb_b = ""
    v_tien = w_tien = None
    sac = ""
    thieu_txt = ""
    nguon = rd["tep"] or "Trực tiếp"
    ct_ref = rd["sheet"] or ""
    ma_tc = ""
    x_chu = ""
    hoi = ""
    chuoi = ""
    do_vung = ""

    if ma in GD3_ONLY:
        status = "KHÔNG_ĐỦ_DỮ_LIỆU"
        ket = "Chưa kết luận"
        nhan = ""
        muc_rr = ""
        thieu_txt = ",".join(missing) if missing else "IN08-IN13 (sổ DN chưa nhận)"
    elif ma == "HD07":
        status = "TREO"
        ket = "Chưa kết luận"
        nhan = ""
        muc_rr = ""
        thieu_txt = "Thiếu số liệu tham chiếu ngành (K13)"
    elif ma == "TN20":
        status = "CHƯA_TRIỂN_KHAI"
        ket = "Chưa kết luận"
        nhan = ""
        muc_rr = ""
        thieu_txt = "Khấu hao chi tiết chưa triển khai"
    elif missing:
        status = "KHÔNG_ĐỦ_DỮ_LIỆU"
        ket = "Chưa kết luận"
        nhan = ""
        muc_rr = ""
        thieu_txt = ",".join(missing)

    # --- compute ---
    if ma == "DQ02":
        status, ket, nhan, muc_rr, diem = "ĐÃ_CHẠY", "Không", "Chắc", "Thấp", 0
        ca_a, va = "Dòng Tổng cộng trong IN02", 0
        thieu_txt = "IN02 chưa nhận - không có BK để lẫn dòng Tổng cộng"
        # still ran on HDDT: filtered Tong
        ket = "Không"
    if ma == "DQ03" and status == "ĐÃ_CHẠY":
        ca_a, va = "HĐĐT thiếu/sai MST đối tác", n_mst_sai + missing_mst
        do = n_mst_sai + missing_mst
        nguon = "KQ1.1 IN01"
        if (n_mst_sai + missing_mst) > 0:
            ket, muc_rr, diem = "Có phát hiện", "Trung bình", 5
            findings.append(("DQ03", "Trung bình", f"{n_mst_sai + missing_mst} HĐ thiếu/sai MST", n_mst_sai + missing_mst))
    if ma == "DQ04" and status == "ĐÃ_CHẠY":
        ca_a, va = "HĐĐT thiếu ký hiệu", n_thieu_kh
        if n_thieu_kh:
            ket, muc_rr, diem = "Có phát hiện", "Thấp", 3
    if ma == "DQ05":
        status, ket = "KHÔNG_ĐỦ_DỮ_LIỆU", "Chưa kết luận"
        thieu_txt = "IN02"
        nhan = muc_rr = ""
    if ma == "DQ06" and status == "ĐÃ_CHẠY":
        ca_a, va = "Khóa HĐĐT trùng (KH|So|chieu|MST)", n_dup
        if n_dup:
            ket, muc_rr, diem = "Có phát hiện", "Trung bình", 5
    if ma == "GT01" and "IN01" in HAVE and "IN03" in HAVE:
        status = "ĐÃ_CHẠY"
        picked = pick_diff(d_gt01)
        if picked:
            y, a, b, d, full, vuot = picked
            ca_a, va, cb_b, vb, do = f"TK[34] {y}", a, f"HĐĐT bán {y}", b, d
            nguon, ct_ref = "KQ2.1 KQ2_Doi_chieu_GTGT", f"Phần II năm {y}"
            ma_tc = "TC27"
            if vuot:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", ("Chắc" if full else "Sơ bộ"), "Cao", 9, "GTGT"
                v_tien = abs(d)
                w_tien = abs(d) if full else 0
                if not full:
                    thieu_txt = "HĐĐT bán chưa phủ đủ năm (từ 21/04/2022) - không tính W"
                hoi = f"Vì sao DT tờ khai GTGT năm {y} lệch HĐĐT bán ra {d:,.0f} đồng?"
                chuoi = "HĐĐT portal - tờ khai XML - (thiếu BK) - (thiếu sổ 511) - (thiếu chứng từ gốc)"
                do_vung = "Không đủ dữ liệu"
                findings.append(("GT01", "Cao", f"{y} lệch {d:,.0f} {'đủ năm' if full else 'HĐĐT chưa phủ'}", abs(d)))
            else:
                ket, nhan, muc_rr = "Không", "Chắc", "Thấp"
                thieu_txt = f"Lệch lớn nhất {y}={d:,.0f} dưới P03=2% tạm"
        else:
            ket, nhan, muc_rr = "Không", "Chắc", "Thấp"
            thieu_txt = "Một số năm HĐĐT chưa phủ (2020-2021 bán ra)"
    if ma == "GT02":
        status, ket, nhan, muc_rr = "KHÔNG_ĐỦ_DỮ_LIỆU", "Chưa kết luận", "", ""
        thieu_txt = "IN02"
    if ma == "GT03" and status == "ĐÃ_CHẠY":
        picked = pick_diff(d_gt03)
        if picked:
            y, a, b, d, full, vuot = picked
            ca_a, va, cb_b, vb, do = f"TK[35] {y}", a, f"Thuế HĐĐT bán {y}", b, d
            if vuot:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", ("Chắc" if full else "Sơ bộ"), "Cao", 8, "GTGT"
                v_tien = abs(d)
                w_tien = 0  # T2: chủ là GT01
                x_chu = f"GT01|{MST}|{KY}"
                findings.append(("GT03", "Cao", f"{y} lệch thuế bán {d:,.0f}", abs(d)))
    if ma == "GT04" and "IN01" in HAVE and "IN03" in HAVE:
        status = "ĐÃ_CHẠY"
        picked = pick_diff(d_gt04)
        if picked:
            y, a, b, d, full, vuot = picked
            ca_a, va, cb_b, vb, do = f"TK[23] {y}", a, f"HĐĐT mua {y}", b, d
            nguon = "KQ2.1 KQ2_Doi_chieu_GTGT"
            if vuot:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", ("Chắc" if full else "Sơ bộ"), "Cao", 8, "GTGT"
                v_tien = abs(d)
                w_tien = abs(d) if full else 0
                if not full:
                    thieu_txt = "HĐĐT mua chưa phủ đủ năm (từ 16/12/2021) - không tính W"
                hoi = f"Vì sao GT mua tờ khai năm {y} lệch HĐĐT mua {d:,.0f} đồng?"
                findings.append(("GT04", "Cao", f"{y} lệch mua {d:,.0f} {'đủ năm' if full else 'HĐĐT chưa phủ'}", abs(d)))
    if ma == "GT05" and status == "ĐÃ_CHẠY":
        picked = pick_diff(d_gt05)
        if picked:
            y, a, b, d, full, vuot = picked
            ca_a, va, cb_b, vb, do = f"TK[24] {y}", a, f"Thuế HĐĐT mua {y}", b, d
            if vuot:
                ket, nhan, muc_rr, diem, sac = "Có phát hiện", ("Chắc" if full else "Sơ bộ"), "Trung bình", 6, "GTGT"
                v_tien = abs(d)
                w_tien = 0
                x_chu = f"GT04|{MST}|{KY}"
    if ma == "GT11" and status == "ĐÃ_CHẠY":
        # [43] vs [152]
        diffs = []
        for y in YEARS:
            a = (gtgt_year.get(y) or {}).get("ct43")
            b = bctc_val(y, "152", which="any") or bctc_val(y, "151", which="any")
            if a is None or b is None:
                continue
            diffs.append((y, a, b, a - b))
        if diffs:
            y, a, b, d = max(diffs, key=lambda x: abs(x[3]))
            ca_a, va, cb_b, vb, do = f"GTGT[43] {y}", a, f"BCTC[152/151] {y}", b, d
            ma_tc = "TC9"
            nhan = "Sơ bộ"
            if abs(d) > 1:
                ket, muc_rr, diem = "Có phát hiện", "Trung bình", 5
        else:
            thieu_txt = "BCTC XML chưa bóc được [152]/[151] một số năm"
            ket, nhan = "Chưa kết luận", "Sơ bộ"
    if ma == "GT14" and status == "ĐÃ_CHẠY":
        am = [(y, (gtgt_year.get(y) or {}).get("ct40")) for y in YEARS if (gtgt_year.get(y) or {}).get("ct40", 0) < 0]
        ca_a, va = "Số năm [40] âm", len(am)
        if am:
            ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Trung bình", 5
    if ma == "TN01" and "IN03" in HAVE and "IN04" in HAVE:
        status = "ĐÃ_CHẠY"
        nhan = "Sơ bộ"
        ma_tc = "TC1"
        picked = pick_diff(d_tn01)
        if picked:
            y, a, b, d, full, vuot = picked
            ca_a, va, cb_b, vb, do = f"GTGT[34] {y}", a, f"BCTC[01] {y}", b, d
            if vuot:
                ket, muc_rr, diem, sac = "Có phát hiện", "Cao", 7, "TNDN"
                v_tien = abs(d)
                w_tien = 0  # chờ sổ 511
                findings.append(("TN01", "Cao", f"{y} GTGT vs BCTC lệch {d:,.0f}", abs(d)))
            else:
                ket, muc_rr, diem = "Không", "Thấp", 0
                thieu_txt = f"Lệch lớn nhất {y}={d:,.0f} dưới P03=2% tạm (K8)"
        else:
            ket = "Chưa kết luận"
            thieu_txt = "BCTC [01] chưa bóc được hoặc HĐĐT/TK thiếu năm"
    if ma == "TN02" and status == "ĐÃ_CHẠY":
        nhan = "Sơ bộ"
        picked = pick_diff(d_tn02)
        if picked:
            y, a, b, d, full, vuot = picked
            ca_a, va, cb_b, vb, do = f"TNDN[A1/01] {y}", a, f"BCTC[01] {y}", b, d
            if vuot:
                ket, muc_rr, diem = "Có phát hiện", "Trung bình", 6
                v_tien = abs(d)
                w_tien = 0
    if ma == "TN03" and status == "ĐÃ_CHẠY":
        nhan = "Sơ bộ - chưa so ngành"
        ma_tc = "TC5"
        hit = False
        for y in YEARS:
            dt = bctc_val(y, "01", which="any")
            gv = bctc_val(y, "02", which="any")
            if dt is not None and gv is not None and dt - gv < 0:
                ket, muc_rr, diem = "Có phát hiện", "Cao", 7
                ca_a, va, cb_b, vb, do = f"DT {y}", dt, "GV", gv, dt - gv
                hit = True
                findings.append(("TN03", "Cao", f"{y} lãi gộp âm", abs(dt - gv)))
                break
        if not hit:
            ket, muc_rr = "Không", "Thấp"
    if ma == "TN05" and status == "ĐÃ_CHẠY":
        nhan = "Sơ bộ"
        ma_tc = "TC3"
        hit = False
        for y in YEARS:
            dt = bctc_val(y, "01", which="any")
            htk = bctc_val(y, "141", which="any")
            if htk is None or htk == 0:
                htk = bctc_val(y, "140", which="any")
            if dt and htk is not None and dt and htk / dt > 0.30:
                ket, muc_rr, diem = "Có phát hiện", "Trung bình", 6
                ca_a, va, cb_b, vb, do = f"HTK {y}", htk, "30% DT", 0.3 * dt, htk / dt
                hit = True
                break
        if not hit:
            ket = "Không" if any(bctc_val(y, "140", which="any") is not None or bctc_val(y, "141", which="any") is not None for y in YEARS) else "Chưa kết luận"
            if ket == "Chưa kết luận":
                thieu_txt = "BCTC chưa bóc HTK ct140/ct141 (mẫu chưa xác định K6)"
                nhan = "Sơ bộ"
                muc_rr = ""
    if ma == "TN06" and status == "ĐÃ_CHẠY":
        nhan = "Sơ bộ"
        ma_tc = "TC2"
        hit = False
        for y in YEARS:
            dt = bctc_val(y, "01", which="any")
            nm = bctc_val(y, "312", which="any")
            if dt and nm is not None and nm / dt > 0.05:
                ket, muc_rr, diem = "Có phát hiện", "Trung bình", 5
                ca_a, va, do = f"NM trả trước [312] {y}", nm, nm / dt
                hit = True
                break
        if not hit:
            ket = "Không"
    if ma == "TN14" and status == "ĐÃ_CHẠY":
        nhan = "Sơ bộ"
        ma_tc = "TC17"
        ket = "Chưa kết luận"
        thieu_txt = "Thiếu khấu hao chi tiết để chốt EBITDA; IN12 chưa nhận"
        muc_rr = ""
    if ma == "TN22" and status == "ĐÃ_CHẠY":
        nhan = "Sơ bộ - chưa so ngành"
        ma_tc = "TC28"
        # GV/DT from BCTC
        ratios = []
        for y in YEARS:
            dt = bctc_val(y, "01", which="any")
            gv = bctc_val(y, "02", which="any")
            if dt and gv is not None:
                ratios.append((y, gv / dt if dt else None, dt, gv))
        if len(ratios) >= 2:
            y0, r0, _, _ = ratios[0]
            y1, r1, dt, gv = ratios[-1]
            ca_a, va, cb_b, vb = f"GV/DT {y0}", r0, f"GV/DT {y1}", r1
            do = (r1 or 0) - (r0 or 0)
            ket, muc_rr, diem = "Có phát hiện" if abs(do or 0) > 0.02 else "Không", "Trung bình" if abs(do or 0) > 0.02 else "Thấp", 5 if abs(do or 0) > 0.02 else 0
        else:
            ket = "Chưa kết luận"
            thieu_txt = "BCTC [01]/[02] chưa đủ năm"
            muc_rr = ""
    if ma == "HD01" and status == "ĐÃ_CHẠY":
        ma_tc = "TC26"
        ca_a, va, cb_b, vb, do = "Số NCC F0/CA", len(f0_ncc), "Doanh số F0", f0_gt, f0_gt
        nguon = "KQ2.1 IN07 / KQ2_NCC_mua_vao"
        if f0_ncc or nnt_on_f0 or nnt_on_ca or nnt_on_f1:
            ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 9, "GTGT"
            v_tien = f0_thue
            w_tien = f0_thue
            hoi = "NNT giải trình chuỗi mua từ NCC F0/CA: hợp đồng, thanh toán NH, nhập kho, sử dụng."
            chuoi = "HĐĐT mua - (thiếu BK) - (thiếu PNK) - (thiếu UNC) - (thiếu sử dụng)"
            do_vung = "Không đủ dữ liệu"
            findings.append(("HD01", "Cao", f"{len(f0_ncc)} NCC F0/CA, DS {f0_gt:,.0f}, thuế {f0_thue:,.0f}", f0_thue))
        else:
            ket, muc_rr = "Không", "Thấp"
    if ma == "HD02":
        status, ket, nhan, muc_rr = "KHÔNG_ĐỦ_DỮ_LIỆU", "Chưa kết luận", "", ""
        thieu_txt = "Danh bạ ĐKKD 3.2"
    if ma == "HD03" and status == "ĐÃ_CHẠY":
        ca_a, va, cb_b, vb, do = "Tỷ trọng F0+F1 / tổng mua", hd03_pct, "P13", P13, hd03_pct - P13
        if hd03_pct > P13:
            ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Cao", 8
            findings.append(("HD03", "Cao", f"{hd03_pct:.2%}", hd03_pct))
        else:
            ket, muc_rr = "Không", "Thấp"
    if ma == "HD04" and status == "ĐÃ_CHẠY":
        ca_a, va, cb_b, vb, do = "Top-5 / tổng mua (gộp MST bán)", top5_pct, "P14", P14, top5_pct - P14
        if top5_pct > P14:
            ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Trung bình", 6
            findings.append(("HD04", "Trung bình", f"{top5_pct:.2%}", top5_pct))
        else:
            ket, muc_rr = "Không", "Thấp"
    if ma == "HD05":
        status, ket, nhan, muc_rr = "KHÔNG_ĐỦ_DỮ_LIỆU", "Chưa kết luận", "", ""
        thieu_txt = "Ngày cấp MST (danh bạ 3.2)"
    if ma == "HD08" and status == "ĐÃ_CHẠY":
        ca_a, va = "NNT trên F1/F0/CA", int(nnt_on_f1 or nnt_on_f0 or nnt_on_ca)
        if nnt_on_f1 or nnt_on_f0 or nnt_on_ca:
            ket, nhan, muc_rr, diem, sac = "Có phát hiện", "Chắc", "Cao", 9, "GTGT"
            findings.append(("HD08", "Cao", "NNT có trên sổ CVCBRR/CA", 1))
        else:
            ket, muc_rr = "Không", "Thấp"
    if ma == "HD09" and status == "ĐÃ_CHẠY":
        # mua-ban qua lai: MST vừa là người bán mua-vao vừa là người mua ban-ra
        sell = {mst_base(x["mst_nm"]) for x in ban_ok if x["mst_nm"]}
        buy = {mst_base(x["mst_nb"]) for x in mua_ok if x["mst_nb"]}
        both = (sell & buy) - {MST, ""}
        ca_a, va = "MST vừa mua vừa bán với NNT", len(both)
        if both:
            ket, nhan, muc_rr, diem = "Có phát hiện", "Sơ bộ", "Trung bình", 5
            findings.append(("HD09", "Trung bình", f"{len(both)} MST 2 chiều", len(both)))
        else:
            ket, muc_rr = "Không", "Thấp"
    if ma == "TT01":
        # GĐ2: HĐ đạt ngưỡng - Chắc về ngưỡng, Sơ bộ về tiền mặt
        status = "ĐÃ_CHẠY"
        ca_a, va = "HĐ mua đạt ngưỡng K5", len(cash_hits)
        do = len(cash_hits)
        nhan = "Sơ bộ"
        nguon = "KQ2.1 HĐĐT mua (không có cột HTTT)"
        if cash_hits:
            ket, muc_rr, diem = "Có phát hiện", "Trung bình", 5
            thieu_txt = "IN10/IN11 - chưa kết luận đã trả tiền mặt; không kết luận thiếu UNC"
            findings.append(("TT01", "Trung bình", f"{len(cash_hits)} HĐ >= ngưỡng K5", len(cash_hits)))
        else:
            ket, muc_rr = "Không", "Thấp"
    if ma == "LD01":
        status, ket, nhan, muc_rr = "KHÔNG_ĐỦ_DỮ_LIỆU", "Chưa kết luận", "", ""
        thieu_txt = "Tờ khai TNCN quý (chỉ có QTT năm)"
    if ma == "LD03" and status == "ĐÃ_CHẠY":
        n52 = sum(1 for x in nld_rows if x.get("nguon") == "05-2" and (x.get("thue") or 0) == 0 and (x.get("tnct") or 0) > 0)
        ca_a, va = "NLĐ 05-2 không trích 10%", n52
        if n_nld == 0:
            ket, nhan = "Chưa kết luận", "Sơ bộ"
            thieu_txt = "Phụ lục 05-1/05-2 XML chưa bóc được dòng NLĐ"
            muc_rr = ""
        elif n52:
            ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Trung bình", 5
        else:
            ket, muc_rr = "Không", "Thấp"
    if ma == "LD04" and status == "ĐÃ_CHẠY":
        cccd = [x.get("cccd") for x in nld_rows if x.get("cccd")]
        dup = len(cccd) - len(set(cccd)) if cccd else 0
        ca_a, va = "Trùng CCCD trên QTT đã bóc", dup
        thieu_txt = "Thiếu tra cứu đa nguồn 3.3"
        if n_nld == 0:
            ket, nhan, muc_rr = "Chưa kết luận", "Sơ bộ", ""
        elif dup:
            ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Trung bình", 5
        else:
            ket, muc_rr = "Không", "Thấp"
    if ma in {"LK01", "LK02"}:
        status, ket, nhan, muc_rr = "KHÔNG_ĐỦ_DỮ_LIỆU", "Chưa kết luận", "", ""
        thieu_txt = "IN12"
    if ma in {"LS01", "LS02"}:
        status = "ĐÃ_CHẠY" if n_in14 else "KHÔNG_ĐỦ_DỮ_LIỆU"
        ket = "Chưa kết luận"
        nhan = "Sơ bộ"
        muc_rr = ""
        thieu_txt = "IN14 có nghĩa vụ kê khai/SPS nhưng chưa có bảng nợ thuế/chậm nộp dòng đủ cột LS"
        ca_a, va = "Dòng IN14", n_in14
    if ma == "DQ01":
        # already GD3 - but DQ01 is "thieu nguon dau vao" which IS a finding at muc A
        status = "ĐÃ_CHẠY"
        ket, nhan, muc_rr, diem = "Có phát hiện", "Chắc", "Trung bình", 4
        ca_a, va = "IN còn thiếu", 7
        cb_b, vb = "IN02, IN08-IN13", 7
        thieu_txt = ""
        findings.append(("DQ01", "Trung bình", "Mức A - thiếu toàn bộ sổ DN", 7))

    if ket != "Có phát hiện":
        if status == "ĐÃ_CHẠY":
            w_tien = 0 if w_tien is None else w_tien
        else:
            w_tien = None
    hs = {"Cao": 3, "Trung bình": 2, "Thấp": 1}.get(muc_rr, 0)
    rows_kq3.append(
        {
            "Mã_PH": f"{ma}|{MST}|{KY}",
            "Tệp phân tích": rd["tep"] or "Trực tiếp",
            "Nhóm": rd["nhom"],
            "Mã quy tắc": ma,
            "Tên quy tắc": rd["ten"],
            "Mã TC": ma_tc,
            "Đối tượng": TEN,
            "Kỳ": KY,
            "Chỉ tiêu A": ca_a,
            "Giá trị A": va,
            "Chỉ tiêu B": cb_b,
            "Giá trị B": vb,
            "Giá trị đo": do,
            "Ngưỡng": "P03=2% tạm (mặc định K8)" if ma in {"TN01", "TN22"} else (rd.get("need") or ""),
            "Trạng thái chạy": status,
            "Kết quả": ket,
            "Nhãn": nhan,
            "Mức rủi ro": muc_rr,
            "Hệ số trọng yếu": hs if ket == "Có phát hiện" else (0 if status == "ĐÃ_CHẠY" else None),
            "Điểm": diem if ket == "Có phát hiện" else (0 if status == "ĐÃ_CHẠY" else None),
            "Sắc thuế": sac,
            "Số tiền liên quan (đồng)": v_tien,
            "Số tiền tính vào tổng (đồng)": w_tien if ket == "Có phát hiện" else (0 if status == "ĐÃ_CHẠY" else None),
            "Mã_PH chủ": x_chu,
            "Nguồn": nguon,
            "Tham chiếu chứng cứ": ct_ref,
            "Còn thiếu để kết luận": thieu_txt,
            "Nội dung kiểm tra tại trụ sở": "Xem chứng từ gốc" if ket == "Có phát hiện" else "",
            "Chứng từ cần xem": "HĐ, BK, sổ 511/632, UNC" if ket == "Có phát hiện" else "",
            "Câu hỏi chất vấn NNT": hoi,
            "Chuỗi bằng chứng 5 mắt xích": chuoi,
            "Độ vững chuỗi bằng chứng": do_vung,
            "Ghi chú": "",
        }
    )

hm = {}
for c in range(1, 40):
    h = ws.cell(9, c).value
    if h:
        hm[str(h).strip()] = c
if ws.max_row > 10:
    ws.delete_rows(10, ws.max_row - 9)
for i, rec in enumerate(rows_kq3):
    rr = 10 + i
    for k, v in rec.items():
        if k not in hm:
            continue
        c = hm[k]
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            put(ws, rr, c, v, NUM if abs(float(v)) >= 1 or float(v) == 0 else PCT)
        else:
            put(ws, rr, c, v)
    mucv = rec.get("Mức rủi ro")
    if mucv == "Cao" and rec.get("Kết quả") == "Có phát hiện":
        ws.cell(rr, hm["Mức rủi ro"]).fill = FILL_RED
        ws.cell(rr, hm["Mức rủi ro"]).font = FONT_RED
    elif mucv == "Trung bình" and rec.get("Kết quả") == "Có phát hiện":
        ws.cell(rr, hm["Mức rủi ro"]).fill = FILL_YEL
        ws.cell(rr, hm["Mức rủi ro"]).font = FONT_ORG
    if rec.get("Trạng thái chạy") == "KHÔNG_ĐỦ_DỮ_LIỆU":
        ws.cell(rr, hm["Trạng thái chạy"]).fill = FILL_YEL

w_gtgt = sum((r.get("Số tiền tính vào tổng (đồng)") or 0) for r in rows_kq3 if r.get("Sắc thuế") == "GTGT")
w_tndn = sum((r.get("Số tiền tính vào tổng (đồng)") or 0) for r in rows_kq3 if r.get("Sắc thuế") == "TNDN")
put(ws, 7, 2, f"GTGT {w_gtgt:,.0f} | TNDN {w_tndn:,.0f} | TNCN 0")
n_cao = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Cao" and r.get("Kết quả") == "Có phát hiện")
n_tb = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Trung bình" and r.get("Kết quả") == "Có phát hiện")
n_thap = sum(1 for r in rows_kq3 if r.get("Mức rủi ro") == "Thấp" and r.get("Kết quả") == "Có phát hiện")
put(ws, 6, 8, n_cao)
put(ws, 6, 9, n_tb)
print("KQ3 rules", len(rows_kq3), "Cao", n_cao, "TB", n_tb, "W_GTGT", w_gtgt)

# 28 TC
ws = wb["Danh_gia_28_tieu_chi_rui_ro"]
put(ws, 1, 1, f"Danh_gia_28_tieu_chi_rui_ro - MST: {MST} | Kỳ: {KY} | Phiên bản KQ: {VER} | Mức A", bold=True)
if str(ws.cell(4, 1).value or "").startswith("TC"):
    ws.delete_rows(4, 1)


def tc_fill(tid, kq, muc, kn):
    return kq, muc, kn


tc_rows = []
# TC1
_p = pick_diff(d_tn01)
if _p:
    y, a, b, d, full, vuot = _p
    tc_rows.append(("TC1", TC28[0][1], "Doanh thu", "lệch ≠ 0 (P03=2% tạm)", f"Năm {y}: GTGT[34]={a:,.0f} vs BCTC[01]={b:,.0f}, lệch {d:,.0f} ({'vượt' if vuot else 'dưới'} P03). Nhãn sơ bộ (chưa có sổ 511).", "Cao" if vuot else "Thấp", "Đối chiếu sổ 511, HĐ bán, cut-off từng năm."))
else:
    tc_rows.append(("TC1", TC28[0][1], "Doanh thu", "lệch ≠ 0", "Chưa đủ dữ liệu để đánh giá (BCTC [01] hoặc GTGT [34] thiếu năm).", "Thấp", "Bóc lại XML BCTC; đối chiếu khi có sổ 511."))
# TC2
hit_tc2 = False
txt2 = []
for y in YEARS:
    dt = bctc_val(y, "01", which="any")
    nm = bctc_val(y, "312", which="any")
    if dt and nm is not None:
        txt2.append(f"{y}: [312]={nm:,.0f} / DT={nm/dt:.2%}")
        if nm / dt > 0.05:
            hit_tc2 = True
tc_rows.append(("TC2", TC28[1][1], "BCTC", "5%", "; ".join(txt2) if txt2 else "Chưa đủ dữ liệu để đánh giá (thiếu [312]).", "Trung bình" if hit_tc2 else "Thấp", "Chọn mẫu hợp đồng, phiếu thu, HĐ chưa xuất."))
# TC3
txt3 = []
hit_tc3 = False
for y in YEARS:
    dt = bctc_val(y, "01", which="any")
    htk = bctc_val(y, "141", which="any")
    if htk is None or htk == 0:
        htk = bctc_val(y, "140", which="any")
    if dt and htk is not None:
        txt3.append(f"{y}: HTK={htk:,.0f} / DT={htk/dt:.2%}")
        if htk / dt > 0.30:
            hit_tc3 = True
tc_rows.append(("TC3", TC28[2][1], "Kho", "30% - đọc mẫu BCTC K6", ("; ".join(txt3) + ". Chưa có XNT. Mẫu B01 chưa đọc chữ in trên XML, nhãn sơ bộ.") if txt3 else "Chưa đủ dữ liệu để đánh giá (HTK ct140/ct141 chưa bóc / chưa xác định mẫu).", "Trung bình" if hit_tc3 else "Thấp", "Xin XNT 152+155+156; xác định B01-DN vs B01-DNN trên bản in BCTC."))
# TC4
tc_rows.append(("TC4", TC28[3][1], "BCTC", "> 0", "Chưa đủ dữ liệu để đánh giá (mã DT chưa thực hiện / TT99 319 chưa xác minh trên XML).", "Thấp", "Đọc thuyết minh BCTC (file PDF B09 có trong 1.1)."))
# TC5
hit_tc5 = False
txt5 = []
for y in YEARS:
    dt = bctc_val(y, "01", which="any")
    gv = bctc_val(y, "02", which="any")
    if dt is not None and gv is not None:
        txt5.append(f"{y}: DT {dt:,.0f} GV {gv:,.0f} gộp {dt-gv:,.0f}")
        if dt - gv < 0:
            hit_tc5 = True
tc_rows.append(("TC5", TC28[4][1], "Doanh thu - giá vốn", "< 0", "; ".join(txt5) if txt5 else "Chưa đủ dữ liệu để đánh giá.", "Cao" if hit_tc5 else "Thấp", "Khi có sổ 511/632 rà dòng P<=0; hiện chưa có sổ DN."))
# TC6-8
for i, tid in enumerate(["TC6", "TC7", "TC8"]):
    meta = TC28[5 + i]
    tc_rows.append((tid, meta[1], meta[2], meta[3], "Chưa đủ dữ liệu để đánh giá (thiếu sổ / chỉ tiêu XML chưa bóc hoặc chưa xác minh).", "Thấp", "Bổ sung sổ cái và thuyết minh."))
# TC9
if any(bctc_val(y, "152", which="any") is not None or bctc_val(y, "151", which="any") is not None for y in YEARS):
    bits = []
    for y in YEARS:
        a = (gtgt_year.get(y) or {}).get("ct43")
        b = bctc_val(y, "152", which="any") or bctc_val(y, "151", which="any")
        if a is not None and b is not None:
            bits.append(f"{y}: [43]={a:,.0f} [152/151]={b:,.0f}")
    tc_rows.append(("TC9", TC28[8][1], "GTGT", "[43] vs [152]", "; ".join(bits) if bits else "Có mã BCTC nhưng chưa khớp đủ năm.", "Trung bình" if bits else "Thấp", "Đối chiếu khấu trừ cuối kỳ vs CDKT."))
else:
    tc_rows.append(("TC9", TC28[8][1], "GTGT", "[43] vs [152]", "Chưa đủ dữ liệu để đánh giá (BCTC chưa bóc [152]).", "Thấp", "Bóc CDKT XML."))
# TC10
tc_rows.append(("TC10", TC28[9][1], "GTGT", "×10%", "Chưa đủ dữ liệu để đánh giá (thiếu HTK+TSCĐ đủ năm để so [43]).", "Thấp", "Lấy [140/141]+[211] khi bóc đủ CDKT."))
# TC11
hit_tc11 = False
txt11 = []
for y in YEARS:
    g = gtgt_year.get(y) or {}
    if g.get("ct26", 0) > 0:
        txt11.append(f"{y} [26]={g['ct26']:,.0f}")
        hit_tc11 = True
tc_rows.append(("TC11", TC28[10][1], "GTGT", "[26]>0", ("Có DT KCT: " + "; ".join(txt11) + ". Cần kiểm tra phân bổ thuế ĐV.") if hit_tc11 else "Không thấy [26]>0 trên tờ khai hiệu lực đã cộng năm.", "Trung bình" if hit_tc11 else "Thấp", "Rà phụ lục phân bổ nếu [26]>0."))
# TC12-16,18,19,21,24
for tid, meta in [
    ("TC12", TC28[11]),
    ("TC13", TC28[12]),
    ("TC14", TC28[13]),
    ("TC15", TC28[14]),
    ("TC16", TC28[15]),
    ("TC18", TC28[17]),
    ("TC19", TC28[18]),
    ("TC21", TC28[20]),
    ("TC24", TC28[23]),
    ("TC25", TC28[24]),
]:
    extra = ""
    if tid == "TC18":
        extra = " IN12 chưa nhận - KHÔNG_ĐỦ_DỮ_LIỆU."
    if tid == "TC24":
        extra = " IN14 chưa có cột KHBS vs hạch toán đủ để chạy."
    if tid == "TC16":
        extra = " Thiếu PL GDLK."
    tc_rows.append((tid, meta[1], meta[2], meta[3] or "", "Chưa đủ dữ liệu để đánh giá." + extra, "Thấp", "Bổ sung nguồn tương ứng."))
# TC17
tc_rows.append(("TC17", TC28[16][1], "TNDN", "30% EBITDA", "EBITDA chưa chốt: có [50]/[23] một phần, thiếu khấu hao chi tiết và IN12.", "Thấp", "Lấy sổ 635 và khấu hao TSCĐ."))
# TC20
txt20 = []
hit20 = False
for y in YEARS:
    dt = bctc_val(y, "01", which="any")
    v = bctc_val(y, "331", which="any")
    if dt and v is not None:
        txt20.append(f"{y}: [331]={v:,.0f} /DT={v/dt:.2%}")
        if v / dt > P23:
            hit20 = True
tc_rows.append(("TC20", TC28[19][1], "Công nợ", "[331] / P23=10%", "; ".join(txt20) if txt20 else "Chưa đủ dữ liệu để đánh giá (thiếu [331]).", "Trung bình" if hit20 else "Thấp", "Đối chiếu sổ 331 khi NNT nộp."))
# TC22
tc_rows.append(("TC22", TC28[21][1], "TNCN", "", f"Đã đọc XML QTT TNCN các năm {sorted(tncn_by)}; {len(nld_rows)} dòng NLĐ. Thiếu tờ khai quý và tra cứu 3.3.", "Thấp", "Đối chiếu bảng lương 2.4 vs QTT khi có."))
# TC23
tc_rows.append(("TC23", TC28[22][1], "TNCN", "20%", "Chưa đủ dữ liệu để đánh giá (thiếu chi phí lương BCTC chi tiết vs TNCT).", "Thấp", "Lấy [622]/chi phí lương thuyết minh vs QTT."))
# TC26
tc_rows.append(("TC26", TC28[25][1], "Hóa đơn", "F0/F1", f"NCC unique {len(ncc_agg)}; F0={len(f0_ncc)} DS {f0_gt:,.0f}; F1={len(f1_ncc)}; tỷ trọng rủi ro {hd03_pct:.2%}. NNT F0/F1/CA = {nnt_on_f0}/{nnt_on_f1}/{nnt_on_ca}.", "Cao" if (f0_ncc or nnt_on_f1 or nnt_on_f0 or nnt_on_ca) else "Thấp", "Chọn mẫu NCC F0: HĐ, UNC, PNK, sử dụng."))
# TC27
_p = pick_diff(d_gt01)
if _p:
    y, a, b, d, full, vuot = _p
    tc_rows.append(("TC27", TC28[26][1], "GTGT", "lệch ≠ 0 (P03=2% tạm)", f"Năm {y}: TK[34]={a:,.0f} vs HĐĐT bán={b:,.0f}, lệch {d:,.0f}. HĐĐT bán từ 21/04/2022, năm đủ phủ 2023-2025.", "Cao" if vuot else "Thấp", "Chọn mẫu HĐ điều chỉnh/hủy/sai kỳ; xin bảng kê NNT."))
else:
    tc_rows.append(("TC27", TC28[26][1], "GTGT", "lệch ≠ 0", "Chưa đủ cặp năm TK vs HĐĐT bán để chốt.", "Thấp", "Bổ sung HĐĐT 2020-2021 nếu thuộc kỳ."))
# TC28
if len([y for y in YEARS if bctc_val(y, "01", which="any") and bctc_val(y, "02", which="any")]) >= 2:
    bits = []
    for y in YEARS:
        dt = bctc_val(y, "01", which="any")
        gv = bctc_val(y, "02", which="any")
        if dt and gv is not None:
            bits.append(f"{y} GV/DT={gv/dt:.2%}")
    tc_rows.append(("TC28", TC28[27][1], "Giá thành", "ngang/dọc BCTC", "BCTC: " + "; ".join(bits) + ". Chưa có sổ 632. Nhãn sơ bộ, chưa so ngành (K13).", "Trung bình", "Xin sổ 511/632; tách hoạt động thương mại."))
else:
    tc_rows.append(("TC28", TC28[27][1], "Giá thành", "ngang/dọc BCTC", "Chưa đủ dữ liệu để đánh giá.", "Thấp", "Bóc [01]/[02] đủ năm."))

for i, (id_, ten, phan, nguong, kq, muc_tc, kn) in enumerate(tc_rows):
    put(ws, 4 + i, 1, id_, bold=True)
    put(ws, 4 + i, 2, ten)
    put(ws, 4 + i, 3, phan)
    put(ws, 4 + i, 4, nguong)
    put(ws, 4 + i, 5, kq)
    put(ws, 4 + i, 6, muc_tc, bold=True)
    put(ws, 4 + i, 7, kn)
    if muc_tc == "Cao":
        ws.cell(4 + i, 6).fill = FILL_RED
        ws.cell(4 + i, 6).font = FONT_RED
    elif muc_tc == "Trung bình":
        ws.cell(4 + i, 6).fill = FILL_YEL
        ws.cell(4 + i, 6).font = FONT_ORG
autowidth(ws, cap=60, max_row=40)
apply_font_sheet(wb["KQ3"])
wb.save(lp(paths["kq3"]))
wb.close()
print("saved KQ3")

# ===========================================================================
# Extra GĐ3
# ===========================================================================
print("FILL extra GĐ3")
FONT_H = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")

p_dau = OUT / f"KQ3_Danh_gia_Dau_hieu_Bat_thuong_{MST}_{KY}.xlsx"
wb = Workbook()
ws = wb.active
ws.title = "Dau_hieu"
heads = ["STT", "Mã", "Dấu hiệu", "Mức", "Số liệu", "Sheet KQ", "Hướng kiểm tra", "Tài liệu cần", "Trạng thái"]
for i, h in enumerate(heads, 1):
    put(ws, 1, i, h, bold=True)
    ws.cell(1, i).fill = FILL_NAVY
    ws.cell(1, i).font = FONT_H
dau = [
    (1, "GT01", "Lệch DT tờ khai GTGT vs HĐĐT bán ra", "Cao" if (pick_diff(d_gt01) and pick_diff(d_gt01)[5]) else "Thấp",
     "; ".join(f"{y}: {d:,.0f}{'*' if not f else ''}" for y, a, b, d, f in d_gt01[:6]) or "Chưa đủ cặp năm",
     "KQ2.1/KQ3", "1. HĐ điều chỉnh 2. Sai kỳ 3. HĐ hủy 4. Xin BK bán ra", "HĐĐT portal, BK bán ra, sổ 511", "Chưa kiểm tra tại trụ sở"),
    (2, "GT04", "Lệch GT mua vào tờ khai vs HĐĐT", "Cao" if (pick_diff(d_gt04) and pick_diff(d_gt04)[5]) else "Thấp",
     "; ".join(f"{y}: {d:,.0f}{'*' if not f else ''}" for y, a, b, d, f in d_gt04[:6]) or "Chưa đủ cặp năm",
     "KQ2.1/KQ3", "1. HĐ mua không kê 2. Kê không có HĐ 3. Hủy/thay thế", "HĐĐT mua, BK mua, sổ 133", "Chưa kiểm tra tại trụ sở"),
    (3, "HD01", f"NCC F0/CA: {len(f0_ncc)} MST, DS {f0_gt:,.0f}", "Cao" if f0_ncc else "Thấp",
     f"Thuế F0 {f0_thue:,.0f}. NNT F0/F1/CA={nnt_on_f0}/{nnt_on_f1}/{nnt_on_ca}",
     "KQ2.1 NCC", "1. Hợp đồng 2. Thanh toán NH 3. Nhập kho 4. Sử dụng", "IN07, UNC, PNK", "Thiếu IN11"),
    (4, "HD04", f"Top-5 NCC {top5_pct:.2%}", "Trung bình" if top5_pct > 0.5 else "Thấp",
     f"{len(ncc_agg)} NCC; P14=50%", "KQ2.1", "Chọn mẫu hợp đồng Top-5", "Hợp đồng, BK", "Chưa kiểm tra"),
    (5, "TN01", "Lệch GTGT [34] vs BCTC [01]", "Cao" if (pick_diff(d_tn01) and pick_diff(d_tn01)[5]) else "Thấp",
     "; ".join(f"{y}: {d:,.0f}" for y, a, b, d, f in d_tn01[:6]) or "Thiếu [01]",
     "KQ2.1/KQ3", "Cut-off, HĐ chưa xuất, sổ 511", "BCTC, tờ khai, sổ 511", "Sơ bộ - chưa có sổ DN"),
    (6, "TT01", f"{len(cash_hits)} HĐ mua đạt ngưỡng K5", "Trung bình" if cash_hits else "Thấp",
     "P10/P11; HĐĐT không có HTTT", "KQ2.2 Công nợ", "Đối UNC; không kết luận thiếu chứng từ khi chưa có sao kê", "Sao kê, UNC", "Thiếu IN11"),
    (7, "DQ01", "Hồ sơ mức A - thiếu toàn bộ sổ DN", "Trung bình",
     "IN02, IN08-IN13 chưa nhận", "KQ1.3", "Đề nghị NNT bổ sung theo KQ1.3", "Sổ KT, XNT, 511/632, sao kê", "Chưa nhận"),
]
for i, row in enumerate(dau):
    for c, v in enumerate(row, 1):
        put(ws, 2 + i, c, v)
    if row[3] == "Cao":
        ws.cell(2 + i, 4).fill = FILL_RED
        ws.cell(2 + i, 4).font = FONT_RED
    elif row[3] == "Trung bình":
        ws.cell(2 + i, 4).fill = FILL_YEL
        ws.cell(2 + i, 4).font = FONT_ORG
autowidth(ws)
apply_font_sheet(ws)
wb.save(lp(p_dau))
wb.close()

p_giai = OUT / f"KQ3_NNT_Giai_trinh_Chi_tiet_{MST}_{KY}.xlsx"
wb = Workbook()
sheets = [
    "00_Tong_hop_chi_tiet",
    "1_NCC_F0_CanhBao",
    "2_NCC_DKKD_bat_thuong",
    "3_Ty_le_gia_von",
    "4_HD_ban_duoi_gia",
    "5_Khong_vay_no_EBITDA",
    "6_Top5_NCC",
    "7_Lech_HDDT_ToKhai",
    "8_Cong_no_TK131_TK331",
    "9_TNCN_NLD_nhieu_nguon",
]
wb.active.title = sheets[0]
for s in sheets[1:]:
    wb.create_sheet(s)

ws = wb["00_Tong_hop_chi_tiet"]
put(ws, 1, 1, f"{TEN} - {MST} - {KY} - Tổng hợp chi tiết dấu hiệu (số liệu thật từ KQ2)", bold=True)
ws.cell(1, 1).fill = FILL_NAVY
ws.cell(1, 1).font = FONT_H
try:
    ws.merge_cells("A1:D1")
except Exception:
    pass
for i, h in enumerate(["Phần", "Chỉ tiêu", "Giá trị", "Nguồn"], 1):
    put(ws, 3, i, h, bold=True)
    ws.cell(3, i).fill = FILL_NAVY
    ws.cell(3, i).font = FONT_H
summary = [
    ("I", "Số HĐĐT mua (tính vào tổng)", len(mua_ok), "IN01"),
    ("I", "Số HĐĐT bán (tính vào tổng)", len(ban_ok), "IN01"),
    ("I", "Số dòng bảng kê mua", "KHÔNG_ĐỦ_DỮ_LIỆU", "IN02"),
    ("I", "Số NCC mua unique MST", len(ncc_agg), "IN01"),
    ("II", "NCC F0/CA", len(f0_ncc), "IN07"),
    ("II", "Doanh số F0", f0_gt, "IN07"),
    ("II", "Thuế F0", f0_thue, "IN07"),
    ("II", "Top5 %", top5_pct, "HD04"),
    ("III", "Tổng DT 511", "KHÔNG_ĐỦ_DỮ_LIỆU", "IN13"),
    ("III", "Tổng GV 632", "KHÔNG_ĐỦ_DỮ_LIỆU", "IN13"),
    ("III", "Dòng P<=0", "KHÔNG_ĐỦ_DỮ_LIỆU", "DT01"),
    ("IV", "HĐ mua đạt ngưỡng TM", len(cash_hits), "TT01"),
    ("IV", "Sao kê NH", "Chưa nhận", "IN11"),
    ("V", "NNT trên F0", "Có" if nnt_on_f0 else "Không", "HD08"),
    ("V", "NNT trên F1", "Có" if nnt_on_f1 else "Không", "HD08"),
    ("V", "NNT trên CA", "Có" if nnt_on_ca else "Không", "HD08"),
    ("VI", "Tổng W GTGT (KQ3)", w_gtgt, "KQ3"),
    ("VI", "Phát hiện Cao", n_cao, "KQ3"),
]
for i, row in enumerate(summary):
    for c, v in enumerate(row, 1):
        put(ws, 4 + i, c, v, NUM if isinstance(v, (int, float)) else None)

ws = wb["1_NCC_F0_CanhBao"]
for i, h in enumerate(["MST", "Tên", "DS", "Thuế", "Nhóm", "Nguồn", "CV"], 1):
    put(ws, 1, i, h, bold=True)
    ws.cell(1, i).fill = FILL_NAVY
    ws.cell(1, i).font = FONT_H
rr = 2
for k, a in f0_ncc:
    nhom, nguon, cv, tt = ncc_flag(k)
    put(ws, rr, 1, k)
    put(ws, rr, 2, a["ten"])
    put(ws, rr, 3, a["gt"], NUM)
    put(ws, rr, 4, a["thue"], NUM)
    put(ws, rr, 5, nhom)
    put(ws, rr, 6, nguon)
    put(ws, rr, 7, cv)
    for c in range(1, 8):
        ws.cell(rr, c).fill = FILL_RED
        ws.cell(rr, c).font = FONT_RED
    rr += 1
if rr == 2:
    put(ws, 2, 1, "Không có NCC F0/CA trong kỳ (theo sổ đã khớp MST).")

ws = wb["2_NCC_DKKD_bat_thuong"]
put(ws, 1, 1, "KHÔNG_ĐỦ_DỮ_LIỆU - chưa có danh bạ ĐKKD thư mục 3.2", bold=True)

ws = wb["3_Ty_le_gia_von"]
for i, h in enumerate(["Năm", "DT_511", "GV_632", "Lãi gộp", "Tỷ lệ", "BCTC[01]", "BCTC[02]"], 1):
    put(ws, 1, i, h, bold=True)
    ws.cell(1, i).fill = FILL_NAVY
    ws.cell(1, i).font = FONT_H
for j, y in enumerate(YEARS):
    put(ws, 2 + j, 1, y)
    put(ws, 2 + j, 2, "KHÔNG_ĐỦ_DỮ_LIỆU")
    put(ws, 2 + j, 3, "KHÔNG_ĐỦ_DỮ_LIỆU")
    put(ws, 2 + j, 6, bctc_val(y, "01", which="any"), NUM)
    put(ws, 2 + j, 7, bctc_val(y, "02", which="any"), NUM)

ws = wb["4_HD_ban_duoi_gia"]
put(ws, 1, 1, "KHÔNG_ĐỦ_DỮ_LIỆU - chưa có sổ 511/632 (IN13)", bold=True)

ws = wb["5_Khong_vay_no_EBITDA"]
put(ws, 1, 1, "EBITDA từ BCTC [50]/[23]; khấu hao chưa tách. IN12 chưa nhận.", bold=True)
ws.cell(1, 1).fill = FILL_NAVY
ws.cell(1, 1).font = FONT_H
for j, y in enumerate(YEARS):
    put(ws, 3, 2 + j, y, bold=True)
put(ws, 4, 1, "LN trước thuế [50]")
put(ws, 5, 1, "CP lãi vay [23]")
for j, y in enumerate(YEARS):
    put(ws, 4, 2 + j, bctc_val(y, "50", which="any"), NUM)
    put(ws, 5, 2 + j, bctc_val(y, "23", which="any"), NUM)

ws = wb["6_Top5_NCC"]
for i, h in enumerate(["Hạng", "MST", "Tên", "GT mua", "Thuế", "Tỷ trọng"], 1):
    put(ws, 1, i, h, bold=True)
    ws.cell(1, i).fill = FILL_NAVY
    ws.cell(1, i).font = FONT_H
for i, (k, a) in enumerate(top5, 1):
    put(ws, 1 + i, 1, i)
    put(ws, 1 + i, 2, k)
    put(ws, 1 + i, 3, a["ten"])
    put(ws, 1 + i, 4, a["gt"], NUM)
    put(ws, 1 + i, 5, a["thue"], NUM)
    put(ws, 1 + i, 6, a["gt"] / tot_mua_gt if tot_mua_gt else None, PCT)
put(ws, 8, 1, f"Tổng Top-5 {top5_gt:,.0f} / tổng mua {tot_mua_gt:,.0f} = {top5_pct:.2%}")

ws = wb["7_Lech_HDDT_ToKhai"]
for i, h in enumerate(["Năm", "TK[34]", "HĐĐT bán", "CL bán", "TK[23]", "HĐĐT mua", "CL mua", "BCTC[01]"], 1):
    put(ws, 1, i, h, bold=True)
    ws.cell(1, i).fill = FILL_NAVY
    ws.cell(1, i).font = FONT_H
for j, y in enumerate(YEARS):
    put(ws, 2 + j, 1, y)
    a, b = tk34(y), hd_ban(y)
    c, d = tk23(y), hd_mua(y)
    put(ws, 2 + j, 2, a, NUM)
    put(ws, 2 + j, 3, b, NUM)
    if a is not None and b is not None:
        put(ws, 2 + j, 4, a - b, NUM)
        if abs(a - b) > 1:
            ws.cell(2 + j, 4).fill = FILL_RED
            ws.cell(2 + j, 4).font = FONT_RED
    put(ws, 2 + j, 5, c, NUM)
    put(ws, 2 + j, 6, d, NUM)
    if c is not None and d is not None:
        put(ws, 2 + j, 7, c - d, NUM)
        if abs(c - d) > 1:
            ws.cell(2 + j, 7).fill = FILL_RED
            ws.cell(2 + j, 7).font = FONT_RED
    put(ws, 2 + j, 8, bctc_val(y, "01", which="any"), NUM)

ws = wb["8_Cong_no_TK131_TK331"]
put(ws, 1, 1, "KHÔNG_ĐỦ_DỮ_LIỆU biến động 131/331 (IN08). Cột BCTC [331]/[01] nếu có:", bold=True)
for i, h in enumerate(["Năm", "BCTC[331]", "BCTC[01]", "331/DT"], 1):
    put(ws, 3, i, h, bold=True)
    ws.cell(3, i).fill = FILL_NAVY
    ws.cell(3, i).font = FONT_H
for j, y in enumerate(YEARS):
    put(ws, 4 + j, 1, y)
    v = bctc_val(y, "331", which="any")
    dt = bctc_val(y, "01", which="any")
    put(ws, 4 + j, 2, v, NUM)
    put(ws, 4 + j, 3, dt, NUM)
    if v is not None and dt:
        put(ws, 4 + j, 4, v / dt, PCT)

ws = wb["9_TNCN_NLD_nhieu_nguon"]
put(ws, 1, 1, f"QTT TNCN năm {sorted(tncn_by)}; NLĐ bóc được {len(nld_rows)} dòng. Tra cứu 3.3 chưa có.", bold=True)
for i, h in enumerate(["Năm", "Số dòng NLĐ", "Nguồn 05-2", "Trùng CCCD"], 1):
    put(ws, 3, i, h, bold=True)
    ws.cell(3, i).fill = FILL_NAVY
    ws.cell(3, i).font = FONT_H
for j, y in enumerate(YEARS):
    nld_y = [x for x in nld_rows if x.get("nam") == y]
    cccd = [x.get("cccd") for x in nld_y if x.get("cccd")]
    put(ws, 4 + j, 1, y)
    put(ws, 4 + j, 2, len(nld_y), NUM)
    put(ws, 4 + j, 3, sum(1 for x in nld_y if x.get("nguon") == "05-2"), NUM)
    put(ws, 4 + j, 4, (len(cccd) - len(set(cccd))) if cccd else 0, NUM)

for s in wb.sheetnames:
    apply_font_sheet(wb[s])
    autowidth(wb[s], max_row=40)
wb.save(lp(p_giai))
wb.close()
print("saved extra", p_dau.name, p_giai.name)

# summary json
import json
summary_path = OUT / f"SUMMARY_{MST}_{KY}_{VER}.json"
payload = {
    "mst": MST,
    "ten": TEN,
    "ky": KY,
    "ver": VER,
    "muc": "A",
    "n_in01": n_in01,
    "n_in03": n_in03,
    "n_in04": n_in04,
    "n_nld": n_nld,
    "n_qtt": n_qtt,
    "n_ncc": len(ncc_agg),
    "n_f0": len(f0_ncc),
    "n_f1": len(f1_ncc),
    "f0_gt": f0_gt,
    "f0_thue": f0_thue,
    "top5_pct": top5_pct,
    "hd03_pct": hd03_pct,
    "cash_hits": len(cash_hits),
    "nnt_f0": nnt_on_f0,
    "nnt_f1": nnt_on_f1,
    "nnt_ca": nnt_on_ca,
    "n_cao": n_cao,
    "n_tb": n_tb,
    "w_gtgt": w_gtgt,
    "w_tndn": w_tndn,
    "xml_fail": xml_fail[:20],
    "gtgt_ky": sorted(eff_gtgt),
    "bctc_years": sorted(bctc_by),
    "tndn_years": sorted(tndn_by),
    "tncn_years": sorted(tncn_by),
    "files": {k: str(v) for k, v in paths.items()},
    "dau": str(p_dau),
    "giai": str(p_giai),
}
summary_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
print("SUMMARY", json.dumps({k: payload[k] for k in ["muc", "n_in01", "n_ncc", "n_f0", "nnt_f0", "nnt_f1", "nnt_ca", "n_cao", "w_gtgt", "top5_pct"]}, ensure_ascii=False))
print("GĐ2 GĐ3 done")
