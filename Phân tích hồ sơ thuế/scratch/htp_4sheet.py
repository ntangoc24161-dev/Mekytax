# -*- coding: utf-8 -*-
"""Skill tong-hop-hddt-bang-ke-4-sheet, du 3 giai doan, HUY THANH PHAT 3603820015."""
from __future__ import annotations

import datetime as dt
import os
import re
import shutil
import sys
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603820015"
TEN = "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ VÀ SẢN XUẤT HUY THÀNH PHÁT"
KY = "2021-2025"
D_FROM = dt.date(2021, 1, 1)
D_TO = dt.date(2025, 12, 31)

HS = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Dang kiem tra\KE HOACH\HUY_THANH_PHAT_3603820015"
)
HDDT_PATH = HS / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu" / "HDDT.xlsx"
if not HDDT_PATH.exists():
    HDDT_PATH = HS / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu & Bang ke" / "HDDT.xlsx"
BK_DIR = HS / "2. HO SO NNT CUNG CAP" / "2.5. Bang ke mua vao ban ra"
DANHBA = HS / "3. DU LIEU BO SUNG & TRA CUU QLT" / "3.2. Danh ba trang thai NNT" / "Danh ba NNT.xls"
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")

FNAME = f"Tong_hop_HDDT_BangKe_{MST}_{KY}.xlsx"
OUT_BRAIN = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\ket-qua\cau-thanh") / FNAME
OUTS = [
    OUT_BRAIN,
    HS / "4. KET QUA PHAN TICH" / FNAME,
    HS / "4. KET QUA PHAN TICH" / "cau-thanh" / FNAME,
    HS / "7. Sao luu tong hop" / FNAME,
]

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
TNR = "Times New Roman"
thin = Side(style="thin", color="000000")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
NUM = "#,##0"
DATE_FMT = "dd/mm/yyyy"


def chuan(s) -> str:
    if s is None:
        return ""
    s = str(s).strip().replace(" ", "")
    if not s or s.lower() == "none":
        return ""
    if len(s) == 14 and "-" in s:
        return s
    digits = re.sub(r"\D", "", s)
    if len(digits) == 10:
        return digits
    if len(digits) == 9:
        return "0" + digits
    if len(digits) in (12, 13):
        base, br = digits[:-3], digits[-3:]
        return base.zfill(10) + "-" + br
    return s


def mst_base(s: str) -> str:
    s = chuan(s)
    return s.split("-")[0] if s else ""


def nnum(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("%", "")
    if s.startswith("=") or s.startswith("#"):
        return 0.0
    s = s.replace(".", "").replace(",", ".") if s.count(",") == 1 and s.count(".") >= 1 else s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def ndate(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip().replace("00:00:00", "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return dt.datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return dt.date(y, mo, d)
        except ValueError:
            return None
    return None


def in_ky(d) -> bool:
    return d is not None and D_FROM <= d <= D_TO


def so_hd_key(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    s = re.sub(r"\s+", "", s)
    return s.lstrip("0") or "0"


def is_skip_hd_status(tt: str, kq: str = "") -> bool:
    t = (tt or "").casefold()
    k = (kq or "").casefold()
    keys = ["bị xóa", "bi xoa", "bị hủy", "bi huy", "bị thay thế", "bi thay the", "hủy bỏ", "huy bo"]
    if any(x in t for x in keys):
        return True
    if "không đủ điều kiện" in k:
        return True
    return False


def read_ssxml(path: Path):
    tree = ET.parse(str(path))
    out = {}
    for ws in tree.getroot().findall("ss:Worksheet", NS):
        name = ws.get("{urn:schemas-microsoft-com:office:spreadsheet}Name")
        rows = []
        for r in ws.iter(f"{{{NS['ss']}}}Row"):
            vals = []
            col = 0
            for c in r.findall("ss:Cell", NS):
                idx = c.get(f"{{{NS['ss']}}}Index")
                if idx:
                    while col < int(idx) - 1:
                        vals.append(None)
                        col += 1
                d = c.find("ss:Data", NS)
                vals.append(d.text if d is not None else None)
                col += 1
            rows.append(vals)
        out[name] = rows
    return out


# ---------------------------------------------------------------------------
# HDDT
# ---------------------------------------------------------------------------
print("READ HDDT", HDDT_PATH)
wb_hd = openpyxl.load_workbook(str(HDDT_PATH), read_only=True, data_only=True)
hd_mua, hd_ban = [], []
hd_skip = {"mua": 0, "ban": 0}

for ws in wb_hd.worksheets:
    name = ws.title.casefold()
    if "mua" in name:
        chieu = "mua"
    elif "ban" in name:
        chieu = "ban"
    else:
        continue
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c or "").strip() for c in rows[0]]
    ix = {h: i for i, h in enumerate(hdr)}

    def col(*cands):
        for c in cands:
            if c in ix:
                return ix[c]
        low = {h.casefold(): i for h, i in ix.items()}
        for c in cands:
            if c.casefold() in low:
                return low[c.casefold()]
        for h, i in ix.items():
            hl = h.casefold()
            if all(k.casefold() in hl for k in cands[0].split()):
                return i
        return None

    i_kh = col("Ký hiệu hóa đơn")
    i_so = col("Số hóa đơn")
    i_ng = col("Ngày lập")
    i_mst_b = None
    i_ten_b = None
    i_mst_m = None
    i_ten_m = None
    i_gt = col("Tổng tiền chưa thuế")
    i_thue = col("Tổng tiền thuế")
    i_tt = col("Tổng tiền thanh toán")
    i_tt_hd = col("Trạng thái hóa đơn")
    i_kq = col("Kết quả kiểm tra hóa đơn")
    for i, h in enumerate(hdr):
        hl = h.casefold()
        if "mst người bán" in hl or "mst nguoi ban" in hl:
            i_mst_b = i
        elif "tên người bán" in hl or "ten nguoi ban" in hl:
            i_ten_b = i
        elif "mst người mua" in hl or "mst nguoi mua" in hl:
            i_mst_m = i
        elif "tên người mua" in hl or "ten nguoi mua" in hl:
            i_ten_m = i
    for r in rows[1:]:
        if not r or all(c is None or str(c).strip() == "" for c in r):
            continue
        so = so_hd_key(r[i_so] if i_so is not None else None)
        if not so or so in {"0", "STT"}:
            continue
        tt = str(r[i_tt_hd] or "") if i_tt_hd is not None else ""
        kq = str(r[i_kq] or "") if i_kq is not None else ""
        if is_skip_hd_status(tt, kq):
            hd_skip[chieu] += 1
            continue
        ngay = ndate(r[i_ng] if i_ng is not None else None)
        if not in_ky(ngay):
            continue
        rec = dict(
            kh=str(r[i_kh] or "").strip() if i_kh is not None else "",
            so=str(r[i_so] or "").strip() if i_so is not None else "",
            ngay=ngay,
            mst_ban=str(r[i_mst_b] or "").strip() if i_mst_b is not None else "",
            ten_ban=str(r[i_ten_b] or "").strip() if i_ten_b is not None else "",
            mst_mua=str(r[i_mst_m] or "").strip() if i_mst_m is not None else "",
            ten_mua=str(r[i_ten_m] or "").strip() if i_ten_m is not None else "",
            chua=nnum(r[i_gt] if i_gt is not None else 0),
            thue=nnum(r[i_thue] if i_thue is not None else 0),
            tong=nnum(r[i_tt] if i_tt is not None else 0),
            tt=tt,
        )
        if chieu == "mua":
            if rec["mst_mua"].replace("-", "")[:10] not in ("", MST):
                continue
            rec["partner_raw"] = rec["mst_ban"]
            rec["partner_name"] = rec["ten_ban"]
            hd_mua.append(rec)
        else:
            if rec["mst_ban"].replace("-", "")[:10] not in ("", MST):
                continue
            rec["partner_raw"] = rec["mst_mua"]
            rec["partner_name"] = rec["ten_mua"]
            hd_ban.append(rec)
wb_hd.close()
print("HDDT mua", len(hd_mua), "skip", hd_skip["mua"], "| ban", len(hd_ban), "skip", hd_skip["ban"])


# ---------------------------------------------------------------------------
# Bang ke NNT (file goc theo ky, bo file tong hop HDDT)
# ---------------------------------------------------------------------------
def read_bk_file(path: Path, chieu: str):
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hdr_i = None
    for i, row in enumerate(rows[:8]):
        joined = " ".join(str(c or "").casefold() for c in row)
        if "số hóa đơn" in joined or "so hoa don" in joined:
            hdr_i = i
            break
    if hdr_i is None:
        return [], 0
    out, skipped = [], 0
    for row in rows[hdr_i + 1 :]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        so_raw = row[1] if len(row) > 1 else None
        so_s = str(so_raw or "").strip()
        joined = " ".join(str(c or "") for c in row[:8]).casefold()
        if (
            so_s.casefold().startswith("số dòng")
            or "nhóm hhdv" in joined
            or "tong cong" in joined
            or "tổng cộng" in joined
            or so_s.casefold() in {"số hóa đơn", "so hoa don"}
        ):
            skipped += 1
            continue
        so = so_hd_key(so_raw)
        if not so or so in {"0"}:
            skipped += 1
            continue
        ngay = ndate(row[2] if len(row) > 2 else None)
        if not in_ky(ngay):
            skipped += 1
            continue
        if chieu == "mua":
            name = str(row[3] or "").strip() if len(row) > 3 else ""
            mst = str(row[4] or "").strip() if len(row) > 4 else ""
            hang = str(row[5] or "").strip() if len(row) > 5 else ""
            chua = nnum(row[6] if len(row) > 6 else 0)
            ts = str(row[7] or "").strip() if len(row) > 7 else ""
            thue = nnum(row[8] if len(row) > 8 else 0)
        else:
            name = str(row[3] or "").strip() if len(row) > 3 else ""
            mst = str(row[4] or "").strip() if len(row) > 4 else ""
            hang = str(row[5] or "").strip() if len(row) > 5 else ""
            chua = nnum(row[6] if len(row) > 6 else 0)
            ts = ""
            thue = nnum(row[7] if len(row) > 7 else 0)
        out.append(
            dict(
                so=so_s,
                ngay=ngay,
                partner_name=name,
                partner_raw=mst,
                hang=hang,
                chua=chua,
                ts=ts,
                thue=thue,
                file=path.name,
            )
        )
    return out, skipped


bk_mua, bk_ban = [], []
bk_skip_m, bk_skip_b = 0, 0
bk_files_used = []
for p in sorted(BK_DIR.rglob("*.xlsx")):
    if p.name.startswith("Bang_ke_mua_vao_ban_ra_tong_hop"):
        continue
    nl = p.name.casefold()
    if "mua" in nl:
        recs, sk = read_bk_file(p, "mua")
        bk_mua.extend(recs)
        bk_skip_m += sk
        bk_files_used.append(p.name)
    elif "bán" in nl or "ban ra" in nl:
        recs, sk = read_bk_file(p, "ban")
        bk_ban.extend(recs)
        bk_skip_b += sk
        bk_files_used.append(p.name)
print("BK mua", len(bk_mua), "skip", bk_skip_m, "| BK ban", len(bk_ban), "skip", bk_skip_b)
print("BK files", len(bk_files_used))


# ---------------------------------------------------------------------------
# Danh ba
# ---------------------------------------------------------------------------
print("READ DANH BA", DANHBA)
ss = read_ssxml(DANHBA)
first_sheet = next(iter(ss.values()))
DB = {}
for r in first_sheet[2:]:
    if len(r) > 4 and r[3]:
        key = chuan(r[3])
        if not key:
            continue
        rec = dict(
            status=(r[44] if len(r) > 44 else "") or (r[43] if len(r) > 43 else "") or "",
            cqt=(r[1] if len(r) > 1 else "") or "",
            ten=(r[4] if len(r) > 4 else "") or "",
        )
        DB[key] = rec
        base = mst_base(key)
        if base and base not in DB:
            DB[base] = rec
print("DB keys", len(DB))


def db_status(key: str) -> str:
    hit = DB.get(key) or DB.get(mst_base(key))
    if hit:
        return hit["status"] or "(trống)"
    return "Không có trong danh bạ"


def db_cqt(key: str) -> str:
    hit = DB.get(key) or DB.get(mst_base(key))
    return hit["cqt"] if hit else ""


# ---------------------------------------------------------------------------
# CVCBRR F0 / F1 + Cong an (NNT)
# ---------------------------------------------------------------------------
print("READ CVCBRR", CVCBRR)
wbcv = openpyxl.load_workbook(str(CVCBRR), read_only=True, data_only=True)
F0 = defaultdict(list)
F1_SELF = []
F0_SELF = []
if "F0" in wbcv.sheetnames:
    for r in list(wbcv["F0"].iter_rows(values_only=True))[6:]:
        if not r or r[1] in (None, ""):
            continue
        key = chuan(r[1])
        item = dict(socv=r[3] if len(r) > 3 else "", ngaycv=r[4] if len(r) > 4 else "", dvbh=r[5] if len(r) > 5 else "", noidung=r[6] if len(r) > 6 else "")
        F0[key].append(item)
        b = mst_base(key)
        if b and b != key:
            F0[b].append(item)
        if b == MST or key == MST:
            F0_SELF.append(r)
if "F1" in wbcv.sheetnames:
    for r in list(wbcv["F1"].iter_rows(values_only=True))[6:]:
        if not r:
            continue
        mua = chuan(r[8] if len(r) > 8 else "")
        if mst_base(mua) == MST or mua == MST:
            F1_SELF.append(r)
wbcv.close()
print("F0 keys", len(F0), "F0 self", len(F0_SELF), "F1 self", len(F1_SELF))


def f0_text(key: str) -> str:
    items = F0.get(key) or F0.get(mst_base(key)) or []
    if not items:
        return ""
    seen, parts = set(), []
    for it in items:
        ngay = it["ngaycv"]
        if isinstance(ngay, (dt.datetime, dt.date)):
            ngay_s = ngay.strftime("%d/%m/%Y") if hasattr(ngay, "strftime") else str(ngay)
        else:
            ngay_s = str(ngay)[:10] if ngay else ""
        part = f"CV {it['socv'] or '?'} ngày {ngay_s or '?'} - {it['dvbh'] or ''}".strip(" -")
        if part not in seen:
            seen.add(part)
            parts.append(part)
    return "; ".join(parts)


def has_f0(key: str) -> bool:
    return bool(F0.get(key) or F0.get(mst_base(key)))


print("READ CA", CA)
ca_nnt = []
if CA.exists():
    wbca = openpyxl.load_workbook(str(CA), read_only=True, data_only=True)
    for sname in wbca.sheetnames:
        ws = wbca[sname]
        for r in ws.iter_rows(values_only=True):
            if not r:
                continue
            for val in r:
                m = re.sub(r"\D", "", str(val or ""))
                if m == MST or m[:10] == MST:
                    ca_nnt.append((sname, r[:8]))
                    break
    wbca.close()
print("CA hits NNT", len(ca_nnt))


# ---------------------------------------------------------------------------
# Partners
# ---------------------------------------------------------------------------
partners = {}


def add(chieu, key, name, src, amount, thue=0.0, cnt=1):
    if not key:
        return
    k = (chieu, key)
    p = partners.setdefault(k, dict(name="", hd_n=0, hd_v=0.0, hd_t=0.0, bk_n=0, bk_v=0.0, bk_t=0.0))
    if name and len(str(name)) > len(p["name"]):
        p["name"] = str(name)
    p[src + "_n"] += cnt
    p[src + "_v"] += amount
    p[src + "_t"] += thue


for d in hd_mua:
    add("Mua vào", chuan(d["partner_raw"]), d["partner_name"], "hd", d["chua"], d["thue"])
for d in hd_ban:
    add("Bán ra", chuan(d["partner_raw"]), d["partner_name"], "hd", d["chua"], d["thue"])
for d in bk_mua:
    add("Mua vào", chuan(d["partner_raw"]), d["partner_name"], "bk", d["chua"], d["thue"])
for d in bk_ban:
    add("Bán ra", chuan(d["partner_raw"]), d["partner_name"], "bk", d["chua"], d["thue"])


def ghi_chu(p, chieu):
    hd, bk = p["hd_n"] > 0, p["bk_n"] > 0
    if hd and bk:
        return "Chung_Mua" if chieu == "Mua vào" else "Chung_Ban"
    if hd:
        return "HDDT_Mua" if chieu == "Mua vào" else "HDDT_Ban"
    return "Bang_ke_Mua" if chieu == "Mua vào" else "Bang_ke_Ban"


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------
def table_sheet(ws, title, headers, rows, num_cols=(), date_cols=(), total_cols=(), widths=None, center_cols=()):
    n = len(headers)
    last = get_column_letter(n)
    ws["A1"] = title
    ws["A2"] = f"Đơn vị: {TEN}"
    ws["A3"] = f"Mã số thuế: {MST}"
    ws["A4"] = f"Kỳ: 01/01/2021 - 31/12/2025"
    ws["A5"] = f"Số dòng: {len(rows)}"
    for r in range(1, 6):
        ws.merge_cells(f"A{r}:{last}{r}")
        ws[f"A{r}"].font = Font(name=TNR, size=12 if r == 1 else 11, bold=r <= 2)
        ws[f"A{r}"].alignment = Alignment(horizontal="center" if r == 1 else "left")
    hr = 7
    for c, h in enumerate(headers, 1):
        cell = ws.cell(hr, c, h)
        cell.font = Font(name=TNR, bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[hr].height = 42
    first = hr + 1
    for i, row in enumerate(rows, 1):
        r = hr + i
        for c, v in enumerate(row, 1):
            if c == 1 and headers[0] == "STT":
                v = i
            if isinstance(v, str) and v.startswith("="):
                v = "'" + v
            cell = ws.cell(r, c, v)
            cell.font = Font(name=TNR, size=10)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=isinstance(v, str) and len(v) > 40)
            if c in date_cols:
                cell.number_format = DATE_FMT
            if c in num_cols:
                cell.number_format = NUM
            if c in center_cols or c == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
    last_row = hr + len(rows)
    if total_cols and rows:
        tr = last_row + 1
        ws.cell(tr, 1, "TỔNG CỘNG")
        ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=min(total_cols) - 1)
        for c in total_cols:
            L = get_column_letter(c)
            cell = ws.cell(tr, c, f"=SUM({L}{first}:{L}{last_row})")
            cell.number_format = NUM
        for c in range(1, n + 1):
            cell = ws.cell(tr, c)
            cell.font = Font(name=TNR, bold=True, size=10)
            cell.border = BORDER
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = f"A{hr + 1}"
    ws.print_title_rows = f"{hr}:{hr}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    return last_row


wb = openpyxl.Workbook()
wb.remove(wb.active)


def hd_rows(lst, side):
    out = []
    for d in lst:
        key = chuan(d["partner_raw"])
        warn = f0_text(key) if side == "mua" else ""
        out.append(
            [
                None,
                d["kh"],
                d["so"],
                d["ngay"],
                d["partner_raw"],
                key,
                d["partner_name"],
                d["chua"],
                d["thue"],
                d["tong"],
                d["tt"],
                db_status(key),
                db_cqt(key),
                warn,
            ]
        )
    if side == "ban":
        return [r[:-1] for r in out]
    return out


hd_mua_sorted = sorted(hd_mua, key=lambda d: (d["ngay"] or dt.date.min, d["kh"], so_hd_key(d["so"])))
hd_ban_sorted = sorted(hd_ban, key=lambda d: (d["ngay"] or dt.date.min, d["kh"], so_hd_key(d["so"])))
H = [
    "STT",
    "Ký hiệu HĐ",
    "Số HĐ",
    "Ngày lập",
    "MST đối tác (gốc)",
    "MST_chuan (BT6)",
    "Tên đối tác",
    "Tiền chưa thuế",
    "Tiền thuế GTGT",
    "Tổng thanh toán",
    "Tình trạng HĐ",
    "Tình trạng hoạt động (danh bạ)",
    "Cơ quan thuế quản lý (danh bạ)",
    "Cảnh báo rủi ro",
]
W = [6, 11, 10, 11, 15, 15, 42, 16, 15, 16, 16, 32, 28, 45]
ws = wb.create_sheet("HDDT_Mua_vao")
table_sheet(
    ws,
    "HĐĐT MUA VÀO",
    H,
    hd_rows(hd_mua_sorted, "mua"),
    num_cols=(8, 9, 10),
    date_cols=(4,),
    total_cols=(8, 9, 10),
    widths=W,
    center_cols=(2, 3, 11),
)
ws = wb.create_sheet("HDDT_Ban_ra")
Hb = H[:-1]
Hb[4] = "MST người mua (gốc)"
Hb[6] = "Tên người mua"
table_sheet(
    ws,
    "HĐĐT BÁN RA",
    Hb,
    hd_rows(hd_ban_sorted, "ban"),
    num_cols=(8, 9, 10),
    date_cols=(4,),
    total_cols=(8, 9, 10),
    widths=W[:-1],
    center_cols=(2, 3, 11),
)


def bk_rows(lst, side):
    out = []
    for d in sorted(lst, key=lambda d: (d["ngay"] or dt.date.min, so_hd_key(d["so"]))):
        key = chuan(d["partner_raw"])
        warn = f0_text(key) if side == "mua" else ""
        out.append(
            [
                None,
                d["so"],
                d["ngay"],
                d["partner_raw"],
                key,
                d["partner_name"],
                d["hang"],
                d["chua"],
                d["ts"],
                d["thue"],
                db_status(key),
                db_cqt(key),
                warn,
            ]
        )
    if side == "ban":
        return [r[:-1] for r in out]
    return out


HBK = [
    "STT",
    "Số HĐ",
    "Ngày HĐ",
    "MST đối tác (gốc)",
    "MST_chuan (BT6)",
    "Tên đối tác",
    "Mặt hàng",
    "Giá trị chưa thuế",
    "Thuế suất",
    "Thuế GTGT",
    "Tình trạng hoạt động (danh bạ)",
    "Cơ quan thuế quản lý (danh bạ)",
    "Cảnh báo rủi ro",
]
WB = [6, 12, 11, 15, 15, 42, 36, 16, 10, 15, 32, 28, 45]
ws = wb.create_sheet("Bang_ke_Mua_vao")
table_sheet(
    ws,
    "BẢNG KÊ NNT - MUA VÀO",
    HBK,
    bk_rows(bk_mua, "mua"),
    num_cols=(8, 10),
    date_cols=(3,),
    total_cols=(8, 10),
    widths=WB,
)
ws = wb.create_sheet("Bang_ke_Ban_ra")
HBKb = HBK[:-1]
HBKb[3] = "MST người mua (gốc)"
HBKb[5] = "Tên người mua"
table_sheet(
    ws,
    "BẢNG KÊ NNT - BÁN RA",
    HBKb,
    bk_rows(bk_ban, "ban"),
    num_cols=(8, 10),
    date_cols=(3,),
    total_cols=(8, 10),
    widths=WB[:-1],
)

rows_dt = []
for (chieu, key), p in sorted(partners.items(), key=lambda kv: (kv[0][0], -(kv[1]["hd_v"] + kv[1]["bk_v"]))):
    lech = p["hd_v"] - p["bk_v"] if (p["hd_n"] and p["bk_n"]) else None
    warn = f0_text(key) if chieu == "Mua vào" else "Không áp dụng (chỉ lọc F0 cho NCC mua vào)"
    rows_dt.append(
        [
            None,
            chieu,
            key,
            p["name"],
            p["hd_n"],
            p["hd_v"],
            p["bk_n"],
            p["bk_v"],
            lech,
            ghi_chu(p, chieu),
            db_status(key),
            db_cqt(key),
            warn,
        ]
    )
HDT = [
    "STT",
    "Chiều",
    "MST_chuan (BT6)",
    "Tên đối tác",
    "Số HĐ (HĐĐT)",
    "Chưa thuế (HĐĐT)",
    "Số dòng (bảng kê)",
    "Chưa thuế (bảng kê)",
    "Lệch (HĐĐT - BK)",
    "Ghi chú nguồn",
    "Tình trạng hoạt động (danh bạ)",
    "Cơ quan thuế quản lý (danh bạ)",
    "Cảnh báo rủi ro",
]
ws = wb.create_sheet("Tong_hop_doi_tac")
table_sheet(
    ws,
    "TỔNG HỢP ĐỐI TÁC (HĐĐT + BẢNG KÊ)",
    HDT,
    rows_dt,
    num_cols=(6, 8, 9),
    total_cols=(6, 8, 9),
    widths=[6, 11, 15, 42, 10, 16, 10, 16, 16, 14, 32, 28, 45],
    center_cols=(5, 7),
)

# Sheet 6
ws = wb.create_sheet("Tong_hop_trang_thai_rui_ro")
ws["A1"] = "TỔNG HỢP TRẠNG THÁI & RỦI RO"
ws.merge_cells("A1:F1")
ws["A1"].font = Font(name=TNR, bold=True, size=12)
ws["A1"].alignment = Alignment(horizontal="center")
r = 3


def put(vals, bold=False):
    global r
    for c, v in enumerate(vals, 1):
        cell = ws.cell(r, c, v)
        cell.font = Font(name=TNR, size=10, bold=bold)
        cell.border = BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            cell.number_format = NUM
    r += 1


nnt_db = DB.get(MST) or DB.get(mst_base(MST))
nnt_db_txt = nnt_db["status"] if nnt_db else f"Không có MST {MST} trong danh bạ"
f1_txt = f"{len(F1_SELF)} dòng có MST {MST}" if F1_SELF else "Không có"
f0_self_txt = f"{len(F0_SELF)} dòng có MST {MST}" if F0_SELF else "Không có"
ca_txt = f"{len(ca_nnt)} dòng có MST {MST}" if ca_nnt else f"Không có MST {MST}"

put([f"A. NNT ĐANG KIỂM TRA: {TEN} (MST {MST})"], bold=True)
put(["Nguồn", "Kết quả"], bold=True)
put(["Danh bạ NNT", nnt_db_txt])
put(["Sổ CVCBRR F0", f0_self_txt])
put(["Sổ CVCBRR F1 (bên mua rà soát)", f1_txt])
put(["Sổ Công an (Bao_cao_cong_van.xlsx)", ca_txt])
r += 1

put(["B. THỐNG KÊ TRẠNG THÁI HOẠT ĐỘNG ĐỐI TÁC"], bold=True)
put(["Chiều", "Tình trạng hoạt động (danh bạ)", "Số đối tác", "Chưa thuế (HĐĐT + bảng kê, gộp)"], bold=True)
agg = defaultdict(lambda: [0, 0.0])
for (chieu, key), p in partners.items():
    st = db_status(key)
    agg[(chieu, st)][0] += 1
    agg[(chieu, st)][1] += max(p["hd_v"], p["bk_v"])
for (chieu, st), (n, v) in sorted(agg.items()):
    put([chieu, st, n, float(v)])
r += 1

f0_hits = [(k, p) for (c, k), p in partners.items() if c == "Mua vào" and has_f0(k)]
put([f"C. NCC MUA VÀO CÓ CỜ F0 (CVCBRR): {len(f0_hits)} đối tác"], bold=True)
put(["MST_chuan", "Tên đối tác", "Chưa thuế (gộp)", "Cảnh báo rủi ro"], bold=True)
for k, p in sorted(f0_hits, key=lambda kv: -max(kv[1]["hd_v"], kv[1]["bk_v"])):
    put([k, p["name"], float(max(p["hd_v"], p["bk_v"])), f0_text(k)])
if not f0_hits:
    put(["(không có)", "", "", ""])
r += 1

ACTIVE = {"nnt đang hoạt động", "đang hoạt động"}
focus = []
for (chieu, k), p in partners.items():
    st = db_status(k)
    flag = []
    if chieu == "Mua vào" and has_f0(k):
        flag.append("F0")
    st_l = (st or "").casefold()
    if st_l not in ACTIVE and st_l not in {"(trống)"}:
        flag.append("trạng thái: " + st)
    if flag:
        focus.append((chieu, k, p["name"], max(p["hd_v"], p["bk_v"]), "; ".join(flag)))
put(
    [f"D. DANH SÁCH TRỌNG TÂM ({len(focus)} đối tác: có cờ F0 hoặc trạng thái không hoạt động / không có trong danh bạ)"],
    bold=True,
)
put(["Chiều", "MST_chuan", "Tên đối tác", "Chưa thuế (gộp)", "Lý do"], bold=True)
for row in sorted(focus, key=lambda x: -x[3]):
    put(list(row))
if not focus:
    put(["(không có)", "", "", "", ""])

r += 2
put(["E. HẠN CHẾ NGUỒN DỮ LIỆU"], bold=True)
put(["Bảng kê mua vào thiếu kỳ: 2023.Q1, 2023.Q2 (thư mục 2.5 không có file)."])
put(["Bảng kê bán ra thiếu kỳ: 2022 (cả năm), 2023.Q1-Q3, 2025.Q3."])
put(["File Bang_ke_mua_vao_ban_ra_tong_hop_2022-2025.xlsx là bản gộp từ HĐĐT, không dùng làm bảng kê NNT."])
put(["NCC mua vào chỉ gắn cờ F0 từ sổ CVCBRR; không gắn F1/Công an lên dòng hóa đơn mua vào."])

for c, w in enumerate([36, 42, 42, 22, 48, 20], 1):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

saved = []
for dest in OUTS:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest == OUT_BRAIN:
        wb.save(str(dest))
    else:
        shutil.copy2(str(OUT_BRAIN), str(dest))
    saved.append((str(dest), dest.stat().st_size))
    print("saved", dest, dest.stat().st_size)


def s(lst, k):
    return sum(float(d[k]) for d in lst)


print("--- STATS ---")
print("HD mua", len(hd_mua), round(s(hd_mua, "chua")), round(s(hd_mua, "thue")))
print("HD ban", len(hd_ban), round(s(hd_ban, "chua")), round(s(hd_ban, "thue")))
print("BK mua", len(bk_mua), round(s(bk_mua, "chua")), round(s(bk_mua, "thue")), "skip", bk_skip_m)
print("BK ban", len(bk_ban), round(s(bk_ban, "chua")), round(s(bk_ban, "thue")), "skip", bk_skip_b)
print("partners mua", sum(1 for k in partners if k[0] == "Mua vào"), "ban", sum(1 for k in partners if k[0] == "Bán ra"))
print("ghi chu", Counter(ghi_chu(p, c) for (c, k), p in partners.items()))
print("status mua", Counter(db_status(k) for (c, k) in partners if c == "Mua vào"))
print("status ban", Counter(db_status(k) for (c, k) in partners if c == "Bán ra"))
print("F0 hits mua", len(f0_hits), [(k, p["name"], max(p["hd_v"], p["bk_v"])) for k, p in f0_hits][:15])
print("focus", len(focus))
print("NNT db", nnt_db_txt)
print("sheets", wb.sheetnames)
