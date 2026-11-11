# -*- coding: utf-8 -*-
"""Extract HUY THANH PHAT 3603820015 sources to JSON cache."""
from __future__ import annotations

import hashlib
import json
import re
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603820015"
TEN = "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ VÀ SẢN XUẤT HUY THÀNH PHÁT"
KY = "2021-2025"
YEARS = [2021, 2022, 2023, 2024, 2025]
TODAY = "2026-11-09"

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Dang kiem tra\KE HOACH\HUY_THANH_PHAT_3603820015"
)
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
ENGINE = VAULT / "skills" / "phan-tich-ho-so-thue" / "engine"
sys.path.insert(0, str(ENGINE))
from ingest.xml_tkhai import parse_xml  # noqa: E402

CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
CACHE = Path(r"C:\Users\NGOC\AppData\Local\Temp\htp_cache.json")


def lp(p) -> str:
    s = str(p)
    return s if s.startswith("\\\\?\\") else "\\\\?\\" + s


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(lp(path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def nnum(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("%", "")
    s = s.replace(".", "").replace(",", ".") if s.count(",") == 1 and s.count(".") >= 1 else s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def ndate(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    s = s.replace("00:00:00", "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(s[:19], fmt).date().isoformat()
        except ValueError:
            continue
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, d).isoformat()
        except ValueError:
            return None
    return None


def norm_mst(v) -> str:
    if v is None:
        return ""
    s = str(v).strip().replace(" ", "").replace(".", "").upper()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        s = s[:-2]
    s = s.replace("_", "-")
    if "-" in s:
        a, b = s.split("-", 1)
        a = re.sub(r"\D", "", a)
        b = re.sub(r"\D", "", b)
        if a:
            return f"{a}-{b}" if b else a
        return ""
    return re.sub(r"\D", "", s)


def mst_base(s: str) -> str:
    s = norm_mst(s)
    return s.split("-")[0] if s else ""


def so_hd_key(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    s = re.sub(r"\s+", "", s)
    return s.lstrip("0") or "0"


def mst_valid(s: str) -> bool:
    base = mst_base(s)
    if len(base) != 10 or not base.isdigit():
        return False
    w = [31, 27, 23, 19, 17, 13, 9, 7, 5]
    t = sum(int(base[i]) * w[i] for i in range(9))
    return int(base[9]) == (10 - t % 10) % 10


def quy_from_date(iso: str | None) -> int | None:
    if not iso:
        return None
    try:
        m = int(iso[5:7])
        return (m - 1) // 3 + 1
    except Exception:
        return None


def year_from_date(iso: str | None) -> int | None:
    if not iso:
        return None
    try:
        return int(iso[:4])
    except Exception:
        return None


def iter_rows(path: Path, data_only=True):
    wb = load_workbook(lp(path), read_only=True, data_only=data_only)
    sheets = []
    for ws in wb.worksheets:
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            rows.append(list(row))
        sheets.append({"name": ws.title, "rows": rows})
    wb.close()
    return sheets


def classify_file(rel: str) -> tuple[str, str]:
    u = rel.replace("/", "\\").upper()
    if "1. DU LIEU THO NGANH THUE" in u or u.startswith("1."):
        src = "CQT"
    elif "2. HO SO NNT" in u or u.startswith("2."):
        src = "DN"
    elif "3. DU LIEU BO SUNG" in u or u.startswith("3."):
        src = "TRA_CUU"
    elif "0. BIEN BAN" in u:
        src = "BB"
    else:
        src = "KHAC"
    if "1.1" in u or "TO KHAI" in u:
        inn = "IN03" if src == "CQT" else "KHAC"
        if "TNDN" in u or "BCTC" in u:
            inn = "IN04"
        if "TNCN" in u or "QTT" in u:
            inn = "IN05"
    elif "1.2" in u or "HDDT" in u or "HOA DON DIEN TU" in u:
        inn = "IN01"
    elif "1.3" in u or "NGHIA VU" in u:
        inn = "IN14"
    elif "1.4" in u:
        inn = "IN05"
    elif "2.5" in u or "BANG KE" in u:
        inn = "IN02"
    elif "2.1" in u or "SO KE TOAN" in u or "NHAT KY" in u or "CHI TIET CAC TAI KHOAN" in u:
        inn = "IN08"
        if "CONG NO" in u:
            inn = "IN08"
    elif "2.2" in u or "TON KHO" in u or "GIA THANH" in u:
        inn = "IN09"
    elif "2.3" in u or "NGAN HANG" in u or "SAO KE" in u:
        inn = "IN11"
    elif "2.4" in u:
        inn = "IN08"
    elif "3.1" in u:
        inn = "IN07"
    elif "3.2" in u:
        inn = "IN07"
    elif "3.3" in u:
        inn = "IN05"
    else:
        inn = "KHAC"
    return src, inn


# ---------------------------------------------------------------------------
# 1. Manifest
# ---------------------------------------------------------------------------
print("SCAN files")
manifest = []
skip_top = {"4. KET QUA PHAN TICH", "7. Sao luu tong hop", "ket-qua"}
n = 0
for p in ROOT.rglob("*"):
    if not p.is_file():
        continue
    rel = str(p.relative_to(ROOT))
    top = rel.split("\\")[0]
    if top in skip_top:
        continue
    n += 1
    src, inn = classify_file(rel)
    try:
        h = sha256(p)
    except Exception:
        h = ""
    manifest.append(
        {
            "rel": rel,
            "name": p.name,
            "ext": p.suffix.lower(),
            "size": p.stat().st_size,
            "mtime": datetime.fromtimestamp(p.stat().st_mtime).date().isoformat(),
            "sha256": h,
            "nguon": src,
            "ma_in": inn,
        }
    )
print(" files", len(manifest))

# ---------------------------------------------------------------------------
# 2. XML
# ---------------------------------------------------------------------------
print("PARSE XML")
xml_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.1. To khai XML"
xml_rows = []
for p in sorted(xml_dir.glob("*.xml")):
    tk = parse_xml(p)
    rec = {
        "file": p.name,
        "rel": str(p.relative_to(ROOT)),
        "size": p.stat().st_size,
        "parsed": tk is not None,
    }
    if tk:
        rec.update(
            {
                "loai": tk.loai,
                "ma_tkhai": tk.ma_tkhai,
                "loai_tkhai": tk.loai_tkhai,
                "so_lan": tk.so_lan,
                "ky": tk.ky,
                "nam": tk.nam,
                "mst": tk.mst,
                "ten": tk.ten_nnt,
                "chi_tieu": tk.chi_tieu,
                "cdkt_cuoi": tk.bctc_cdkt_cuoi,
                "cdkt_dau": tk.bctc_cdkt_dau,
                "kqkd": tk.bctc_kqkd,
            }
        )
    else:
        try:
            tree = ET.parse(p)
            root = tree.getroot()

            def _ln(t):
                return t.split("}", 1)[-1] if t else ""

            def _txt(el):
                return (el.text or "").strip() if el is not None else ""

            def _find(name):
                for ch in root.iter():
                    if _ln(ch.tag) == name:
                        return ch
                return None

            rec.update(
                {
                    "loai": None,
                    "ma_tkhai": _txt(_find("maTKhai")),
                    "loai_tkhai": _txt(_find("loaiTKhai")),
                    "so_lan": _txt(_find("soLan")),
                    "ky": "",
                    "nam": 0,
                }
            )
            kyel = _find("KyKKhaiThue")
            if kyel is not None:
                for ch in list(kyel):
                    if _ln(ch.tag) == "kyKKhai":
                        rec["ky"] = (ch.text or "").strip()
        except Exception as e:
            rec["err"] = str(e)
    xml_rows.append(rec)
print(" xml", len(xml_rows), "parsed", sum(1 for x in xml_rows if x.get("parsed")))


def latest_xml(loai: str):
    """Pick effective return per ky: supplement with max lan, else C."""
    groups = defaultdict(list)
    for x in xml_rows:
        if not x.get("parsed") or x.get("loai") != loai:
            continue
        groups[(x.get("nam"), str(x.get("ky")))].append(x)
    out = []
    for key, rows in groups.items():
        rows_sorted = sorted(rows, key=lambda r: (0 if str(r.get("loai_tkhai")).upper() in {"B", "2"} else 1, -int(r.get("so_lan") or 0), -r.get("size", 0)))
        pick = rows_sorted[0]
        pick = dict(pick)
        pick["la_hieu_luc"] = True
        pick["so_ban"] = len(rows)
        out.append(pick)
    return sorted(out, key=lambda r: (r.get("nam") or 0, str(r.get("ky"))))


gtgt_eff = latest_xml("gtgt")
tndn_eff = latest_xml("tndn")
bctc_eff = latest_xml("bctc")
tncn_qtt = latest_xml("tncn")
print(" GTGT ky", len(gtgt_eff), "TNDN", len(tndn_eff), "BCTC", len(bctc_eff), "TNCN", len(tncn_qtt))

# ---------------------------------------------------------------------------
# 3. HDDT
# ---------------------------------------------------------------------------
print("PARSE HDDT")
hddt_path = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu" / "HDDT.xlsx"
if not hddt_path.exists():
    hddt_path = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu & Bang ke" / "HDDT.xlsx"
hddt = []
wb = load_workbook(lp(hddt_path), read_only=True, data_only=True)
for ws in wb.worksheets:
    name = ws.title.lower()
    if "mua" in name:
        chieu = "MUA"
    elif "ban" in name:
        chieu = "BAN"
    else:
        continue
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        continue
    hdr = [str(c or "").strip() for c in rows[0]]

    def col(*keys):
        for i, h in enumerate(hdr):
            hl = h.lower()
            if all(k.lower() in hl for k in keys):
                return i
        return None

    i_kh = col("ký hiệu hóa đơn") or col("ky hieu")
    i_so = col("số hóa đơn") or col("so hoa don")
    i_ng = col("ngày lập") or col("ngay lap")
    i_mst_b = None
    i_ten_b = None
    i_mst_m = None
    i_ten_m = None
    for i, h in enumerate(hdr):
        hl = h.lower()
        if "mst người bán" in hl or "mst nguoi ban" in hl or ("người bán" in hl and "mst" in hl):
            i_mst_b = i
        elif "tên người bán" in hl or "ten nguoi ban" in hl:
            i_ten_b = i
        elif "mst người mua" in hl or "mst nguoi mua" in hl:
            i_mst_m = i
        elif "tên người mua" in hl or "ten nguoi mua" in hl:
            i_ten_m = i
    i_gt = col("tổng tiền chưa thuế") or col("chua thue")
    i_thue = col("tổng tiền thuế") or col("tien thue")
    i_tt = col("tổng tiền thanh toán") or col("thanh toan")
    for ri, row in enumerate(rows[1:], 2):
        so = so_hd_key(row[i_so] if i_so is not None else None)
        if not so or so in {"0", "STT"}:
            continue
        ngay = ndate(row[i_ng] if i_ng is not None else None)
        nam = year_from_date(ngay)
        quy = quy_from_date(ngay)
        mst_b = norm_mst(row[i_mst_b] if i_mst_b is not None else None)
        mst_m = norm_mst(row[i_mst_m] if i_mst_m is not None else None)
        dt = mst_b if chieu == "MUA" else mst_m
        hddt.append(
            {
                "chieu": chieu,
                "nam": nam,
                "quy": quy,
                "thang": int(ngay[5:7]) if ngay else None,
                "ky_hieu": str(row[i_kh] or "").strip() if i_kh is not None else "",
                "so_hd": str(row[i_so] or "").strip() if i_so is not None else "",
                "so_key": so,
                "ngay": ngay,
                "mst_nb": mst_b,
                "ten_nb": str(row[i_ten_b] or "").strip() if i_ten_b is not None else "",
                "mst_nm": mst_m,
                "ten_nm": str(row[i_ten_m] or "").strip() if i_ten_m is not None else "",
                "mst_dt": dt,
                "gt": nnum(row[i_gt] if i_gt is not None else 0),
                "thue": nnum(row[i_thue] if i_thue is not None else 0),
                "tt": nnum(row[i_tt] if i_tt is not None else 0),
                "tinh_vao_tong": "Có",
                "file": hddt_path.name,
            }
        )
wb.close()
print(" HDDT", len(hddt), "mua", sum(1 for x in hddt if x["chieu"] == "MUA"), "ban", sum(1 for x in hddt if x["chieu"] == "BAN"))

# ---------------------------------------------------------------------------
# 4. Bang ke original (not the HDDT-derived tong hop)
# ---------------------------------------------------------------------------
print("PARSE BANG KE")
bk_root = ROOT / "2. HO SO NNT CUNG CAP" / "2.5. Bang ke mua vao ban ra"
bangke = []
for p in sorted(bk_root.rglob("*.xlsx")):
    if p.name.startswith("Bang_ke_mua_vao_ban_ra_tong_hop"):
        continue
    m = re.search(r"(20\d{2})\.Q(\d)", p.name)
    nam_f = int(m.group(1)) if m else None
    quy_f = int(m.group(2)) if m else None
    chieu = "MUA" if "mua" in p.name.lower() else ("BAN" if "bán" in p.name.lower() or "ban ra" in p.name.lower() else None)
    if chieu is None:
        continue
    sheets = iter_rows(p)
    for sh in sheets:
        rows = sh["rows"]
        if len(rows) < 4:
            continue
        hdr_i = None
        for i, row in enumerate(rows[:8]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "số hóa đơn" in joined or "so hoa don" in joined:
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = [str(c or "").strip().lower() for c in rows[hdr_i]]

        def icontains(*keys):
            for i, h in enumerate(hdr):
                if all(k in h for k in keys):
                    return i
            return None

        i_so = icontains("số hóa đơn") or icontains("số hoá đơn") or 1
        i_ng = icontains("ngày") or 2
        i_ten = None
        i_mst = None
        i_hang = icontains("mặt hàng") or 5
        i_gt = None
        i_ts = icontains("thuế suất")
        i_thue = None
        for i, h in enumerate(hdr):
            if "mã số thuế" in h or h.startswith("mst"):
                i_mst = i
            elif "tên người" in h or "ten nguoi" in h:
                i_ten = i
            elif "giá trị" in h or "doanh số" in h or "chua" in h and "thuế" in h:
                i_gt = i
            elif h in {"thuế gtgt", "thue gtgt"} or ("thuế gtgt" in h and "suất" not in h):
                i_thue = i
        if i_gt is None:
            i_gt = 6
        if i_thue is None:
            i_thue = 8 if chieu == "MUA" else 7
        if i_ten is None:
            i_ten = 3
        if i_mst is None:
            i_mst = 4
        for row in rows[hdr_i + 1 :]:
            so_raw = row[i_so] if i_so < len(row) else None
            ten = str(row[i_ten] or "").strip() if i_ten < len(row) else ""
            if so_raw is None and not ten:
                continue
            joined = " ".join(str(c or "") for c in row[:8]).lower()
            if "nhóm hhdv" in joined or "tong cong" in joined or "tổng cộng" in joined:
                continue
            so = so_hd_key(so_raw)
            if not so or so in {"0"}:
                continue
            ngay = ndate(row[i_ng] if i_ng < len(row) else None)
            nam = year_from_date(ngay) or nam_f
            quy = quy_from_date(ngay) or quy_f
            mst_dt = norm_mst(row[i_mst] if i_mst < len(row) else None)
            ts_raw = row[i_ts] if i_ts is not None and i_ts < len(row) else None
            bangke.append(
                {
                    "chieu": chieu,
                    "nam": nam,
                    "quy": quy,
                    "thang": int(ngay[5:7]) if ngay else None,
                    "so_hd": str(so_raw or "").strip(),
                    "so_key": so,
                    "ngay": ngay,
                    "mst_goc": str(row[i_mst] or "").strip() if i_mst < len(row) else "",
                    "mst_dt": mst_dt,
                    "ten": ten,
                    "hang": str(row[i_hang] or "").strip() if i_hang < len(row) else "",
                    "gt": nnum(row[i_gt] if i_gt < len(row) else 0),
                    "thue_suat": str(ts_raw or "").strip(),
                    "thue": nnum(row[i_thue] if i_thue < len(row) else 0),
                    "file": p.name,
                }
            )
print(" bangke", len(bangke), "mua", sum(1 for x in bangke if x["chieu"] == "MUA"), "ban", sum(1 for x in bangke if x["chieu"] == "BAN"))

# ---------------------------------------------------------------------------
# 5. So chi tiet
# ---------------------------------------------------------------------------
print("PARSE SO CHI TIET")
soct = []
so_dir = ROOT / "2. HO SO NNT CUNG CAP" / "2.1. So ke toan tong hop va chi tiet"
for p in so_dir.rglob("SO_CHI_TIET_CAC_TAI_KHOAN *.xlsx"):
    m = re.search(r"(20\d{2})", p.name)
    nam = int(m.group(1)) if m else None
    sheets = iter_rows(p)
    for sh in sheets:
        rows = sh["rows"]
        hdr_i = None
        for i, row in enumerate(rows[:8]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "tài khoản" in joined and "phát sinh" in joined:
                hdr_i = i
                break
        if hdr_i is None:
            continue
        for row in rows[hdr_i + 1 :]:
            if not row or len(row) < 8:
                continue
            tk = str(row[4] or "").strip()
            if not tk or tk.lower() in {"tài khoản", "tk"}:
                continue
            dg = str(row[3] or "").strip()
            if dg.lower().startswith("số dư đầu") or dg.lower().startswith("cộng"):
                soct.append(
                    {
                        "nam": nam,
                        "kind": "SODU" if "đầu" in dg.lower() or "dau" in dg.lower() else "CONG",
                        "ngay_ht": ndate(row[0]),
                        "ngay_ct": ndate(row[1]),
                        "so_ct": str(row[2] or "").strip(),
                        "dien_giai": dg,
                        "tk": tk,
                        "tk_du": str(row[5] or "").strip(),
                        "ps_no": nnum(row[6]),
                        "ps_co": nnum(row[7]),
                        "du_no": nnum(row[8]) if len(row) > 8 else 0,
                        "du_co": nnum(row[9]) if len(row) > 9 else 0,
                        "file": p.name,
                    }
                )
                continue
            soct.append(
                {
                    "nam": nam,
                    "kind": "PS",
                    "ngay_ht": ndate(row[0]),
                    "ngay_ct": ndate(row[1]),
                    "so_ct": str(row[2] or "").strip(),
                    "dien_giai": dg,
                    "tk": tk,
                    "tk_du": str(row[5] or "").strip(),
                    "ps_no": nnum(row[6]),
                    "ps_co": nnum(row[7]),
                    "du_no": nnum(row[8]) if len(row) > 8 else 0,
                    "du_co": nnum(row[9]) if len(row) > 9 else 0,
                    "file": p.name,
                }
            )
print(" soct", len(soct))

# ---------------------------------------------------------------------------
# 6. Ton kho / gia thanh
# ---------------------------------------------------------------------------
print("PARSE KHO")
kho = []
gt_rows = []
kho_dir = ROOT / "2. HO SO NNT CUNG CAP" / "2.2. Kho - Chi phi - Gia thanh"
for p in kho_dir.rglob("TONG_HOP_TON_KHO *.xlsx"):
    m = re.search(r"(20\d{2})", p.name)
    nam = int(m.group(1)) if m else None
    sheets = iter_rows(p)
    tot = {"nam": nam, "dk_sl": 0, "dk_tt": 0, "nhap_sl": 0, "nhap_tt": 0, "xuat_sl": 0, "xuat_tt": 0, "ck_sl": 0, "ck_tt": 0, "file": p.name, "n_hang": 0}
    for sh in sheets:
        rows = sh["rows"]
        start = 0
        for i, row in enumerate(rows[:8]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "mã hàng" in joined or "ma hang" in joined:
                start = i + 2
                break
        for row in rows[start:]:
            if not row or len(row) < 12:
                continue
            ma = str(row[1] or "").strip()
            tenh = str(row[2] or "").strip()
            if not ma and not tenh:
                continue
            if "tổng" in tenh.lower() or "cong" in tenh.lower():
                continue
            tot["n_hang"] += 1
            tot["dk_sl"] += nnum(row[4])
            tot["dk_tt"] += nnum(row[5])
            tot["nhap_sl"] += nnum(row[6])
            tot["nhap_tt"] += nnum(row[7])
            tot["xuat_sl"] += nnum(row[8])
            tot["xuat_tt"] += nnum(row[9])
            tot["ck_sl"] += nnum(row[10])
            tot["ck_tt"] += nnum(row[11])
    kho.append(tot)

for p in kho_dir.rglob("BANG_TINH_GIA_THANH *.xlsx"):
    m = re.search(r"(20\d{2})", p.name)
    nam = int(m.group(1)) if m else None
    sheets = iter_rows(p)
    sample = []
    for sh in sheets:
        for i, row in enumerate(sh["rows"][:25]):
            sample.append([None if c is None else (c if not isinstance(c, str) else c[:80]) for c in row[:12]])
    gt_rows.append({"nam": nam, "file": p.name, "sample": sample})
print(" kho years", [k["nam"] for k in kho], "giathanh", [g["nam"] for g in gt_rows])

# ---------------------------------------------------------------------------
# 7. Bank
# ---------------------------------------------------------------------------
print("PARSE BANK")
bank = []
bank_dir = ROOT / "2. HO SO NNT CUNG CAP" / "2.3. Ngan hang & Dong tien"
for p in bank_dir.rglob("LichSuTaiKhoan_*.xlsx"):
    sheets = iter_rows(p)
    rec = {"file": p.name, "stk": "", "ten_kh": "", "n": 0, "ghi_no": 0.0, "ghi_co": 0.0, "lines": []}
    for sh in sheets:
        rows = sh["rows"]
        hdr_i = None
        for i, row in enumerate(rows[:20]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "tên khách" in joined or "customer's name" in joined:
                rec["ten_kh"] = str(row[1] or row[0] or "")
            if "số tài khoản" in joined or "account no" in joined:
                rec["stk"] = str(row[1] or "")
                if not rec["stk"] and len(row) > 2:
                    rec["stk"] = str(row[2] or "")
            if "stt" in joined and ("ngày" in joined or "date" in joined or i >= 5):
                hdr_i = i
        if hdr_i is None:
            for i, row in enumerate(rows[:15]):
                joined = " ".join(str(c or "").lower() for c in row)
                if "stt" in joined:
                    hdr_i = i
                    break
        start = (hdr_i or 6) + 1
        for row in rows[start:]:
            if not row:
                continue
            vals = [c for c in row if c not in (None, "")]
            if len(vals) < 2:
                continue
            rec["n"] += 1
            # heuristic: last numeric cols as amounts
            nums = [nnum(c) for c in row if isinstance(c, (int, float)) or (isinstance(c, str) and re.search(r"\d", c or ""))]
            if len(nums) >= 2:
                rec["ghi_no"] += abs(nums[-2]) if False else 0
            rec["lines"].append({"raw": [None if c is None else (str(c)[:60] if not isinstance(c, (int, float)) else c) for c in row[:10]]})
            if rec["n"] >= 8:
                break
    bank.append(rec)
print(" bank files", len(bank))

# ---------------------------------------------------------------------------
# 8. Cong no tong hop
# ---------------------------------------------------------------------------
print("PARSE CONG NO")
cn = []
for p in so_dir.rglob("TONG_HOP_CONG_NO_*.xlsx"):
    m = re.search(r"(20\d{2})", p.name)
    nam = int(m.group(1)) if m else None
    loai = "THU" if "PHAI_THU" in p.name.upper() else "TRA"
    sheets = iter_rows(p)
    tot_ck = 0
    n_obj = 0
    for sh in sheets:
        rows = sh["rows"]
        for row in rows[3:]:
            if not row:
                continue
            nums = [nnum(c) for c in row if c not in (None, "")]
            ten = str(row[1] or row[0] or "")
            if "tổng" in ten.lower() or "cong" in ten.lower():
                continue
            if any(nums):
                n_obj += 1
                tot_ck += nums[-1] if nums else 0
    cn.append({"nam": nam, "loai": loai, "file": p.name, "n": n_obj, "ck_guess": tot_ck})
print(" cn", cn)

# ---------------------------------------------------------------------------
# 9. CVCBRR / CA / danh ba
# ---------------------------------------------------------------------------
print("PARSE CVCBRR")
f0, f1, ca_list, danhba = [], [], [], []


def load_mst_sheet(path, sheet_hint=None):
    out = []
    if not path.exists():
        return out
    wb = load_workbook(lp(path), read_only=True, data_only=True)
    names = wb.sheetnames
    for name in names:
        if sheet_hint and name.upper() != sheet_hint.upper():
            continue
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = 0
        for i, row in enumerate(rows[:8]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "mst" in joined or "mã số thuế" in joined:
                hdr_i = i
                break
        hdr = [str(c or "").strip().lower() for c in rows[hdr_i]]
        i_mst = next((i for i, h in enumerate(hdr) if "mst" in h or "mã số thuế" in h), 0)
        i_ten = next((i for i, h in enumerate(hdr) if "tên" in h or "ten nnt" in h or "tên nnt" in h), 1)
        i_cv = next((i for i, h in enumerate(hdr) if "công văn" in h or "so cv" in h or "số cv" in h or h == "cv"), None)
        i_ngay = next((i for i, h in enumerate(hdr) if "ngày" in h and "cv" in h or "ngay cv" in h), None)
        i_nd = next((i for i, h in enumerate(hdr) if "nội dung" in h or "noi dung" in h or "cảnh báo" in h), None)
        for row in rows[hdr_i + 1 :]:
            if not row:
                continue
            mst = norm_mst(row[i_mst] if i_mst < len(row) else None)
            if not mst:
                continue
            out.append(
                {
                    "mst": mst,
                    "base": mst_base(mst),
                    "ten": str(row[i_ten] or "").strip() if i_ten < len(row) else "",
                    "cv": str(row[i_cv] or "").strip() if i_cv is not None and i_cv < len(row) else "",
                    "ngay_cv": ndate(row[i_ngay] if i_ngay is not None and i_ngay < len(row) else None),
                    "nd": str(row[i_nd] or "").strip()[:200] if i_nd is not None and i_nd < len(row) else "",
                    "sheet": name,
                }
            )
    wb.close()
    return out


if CVCBRR.exists():
    f0 = load_mst_sheet(CVCBRR, "F0")
    f1 = load_mst_sheet(CVCBRR, "F1")
print(" F0", len(f0), "F1", len(f1))
if CA.exists():
    ca_list = load_mst_sheet(CA, None)
print(" CA", len(ca_list))

db_files = list((ROOT / "3. DU LIEU BO SUNG & TRA CUU QLT" / "3.2. Danh ba trang thai NNT").glob("*.xlsx"))
for p in db_files:
    danhba.extend(load_mst_sheet(p, None))
print(" danhba", len(danhba))

# ---------------------------------------------------------------------------
# 10. NLD from 1.4
# ---------------------------------------------------------------------------
print("PARSE NLD 1.4")
nld_path = ROOT / "1. DU LIEU THO NGANH THUE" / "1.4. Tong hop danh sach NLD theo QTT TNCN" / "Tong_hop_danh_sach_NLD_theo_QTT_TNCN.xlsx"
nld_qtt = []
nld_ds = []
if nld_path.exists():
    sheets = iter_rows(nld_path)
    for sh in sheets:
        name = sh["name"]
        rows = sh["rows"]
        if "tong hop qtt" in name.lower() or name.startswith("1."):
            for row in rows:
                y = None
                if row and row[0] not in (None, "") and str(row[0]).replace(".0", "").isdigit():
                    y = int(float(row[0]))
                if y and y >= 2015:
                    nld_qtt.append(
                        {
                            "nam": y,
                            "mau": str(row[1] or ""),
                            "so_ld": nnum(row[2]),
                            "tnct": nnum(row[4]) if len(row) > 4 else 0,
                            "thue_kt": nnum(row[10]) if len(row) > 10 else 0,
                            "thue_phai": nnum(row[11]) if len(row) > 11 else 0,
                            "file_xml": str(row[13] or "") if len(row) > 13 else "",
                        }
                    )
        if "05-1" in name.lower() or "danh sach" in name.lower():
            hdr_i = 0
            for i, row in enumerate(rows[:8]):
                joined = " ".join(str(c or "").lower() for c in row)
                if "họ" in joined or "ho ten" in joined or "họ tên" in joined or "cccd" in joined:
                    hdr_i = i
                    break
            hdr = [str(c or "").strip().lower() for c in rows[hdr_i]]
            for row in rows[hdr_i + 1 :]:
                if not row or not any(row):
                    continue
                nld_ds.append({"nam": nnum(row[0]) if row else None, "raw": [None if c is None else str(c)[:80] for c in row[:16]]})
print(" nld_qtt", nld_qtt, "nld_ds", len(nld_ds))

# ---------------------------------------------------------------------------
# 11. TMS
# ---------------------------------------------------------------------------
print("PARSE TMS")
tms = []
tms_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.3. Nghia vu nop thue & Thong bao"
for p in tms_dir.glob("*.xlsx"):
    sheets = iter_rows(p)
    rec = {"file": p.name, "kind": "", "ky": "", "rows_n": 0, "sample": []}
    for sh in sheets:
        rows = sh["rows"]
        rec["rows_n"] = len(rows)
        for i, row in enumerate(rows[:40]):
            joined = " ".join(str(c or "") for c in row)
            if "PHÂN TÍCH BẢNG CÂN ĐỐI" in joined:
                rec["kind"] = "CDKT"
            if "KẾT QUẢ HĐKD" in joined or "KET QUA HDKD" in joined.upper() or "HĐKD" in joined:
                rec["kind"] = rec["kind"] or "KQKD"
            if "Từ năm" in joined or "Tu nam" in joined:
                rec["ky"] = joined
            rec["sample"].append([None if c is None else (str(c)[:50] if not isinstance(c, (int, float)) else c) for c in row[:12]])
    tms.append(rec)
print(" tms", [(t["file"], t["kind"], t["rows_n"]) for t in tms])

cache = {
    "meta": {"mst": MST, "ten": TEN, "ky": KY, "years": YEARS, "today": TODAY, "folder": str(ROOT)},
    "manifest": manifest,
    "xml": xml_rows,
    "gtgt_eff": gtgt_eff,
    "tndn_eff": tndn_eff,
    "bctc_eff": bctc_eff,
    "tncn_qtt": tncn_qtt,
    "hddt": hddt,
    "bangke": bangke,
    "soct": soct,
    "kho": kho,
    "giathanh_sample": gt_rows,
    "bank": bank,
    "congno": cn,
    "f0": f0,
    "f1": f1,
    "ca": ca_list,
    "danhba": danhba,
    "nld_qtt": nld_qtt,
    "nld_ds": nld_ds[:400],
    "tms": tms,
    "cvcbrr_mtime": datetime.fromtimestamp(CVCBRR.stat().st_mtime).isoformat() if CVCBRR.exists() else None,
    "ca_mtime": datetime.fromtimestamp(CA.stat().st_mtime).isoformat() if CA.exists() else None,
}
CACHE.write_text(json.dumps(cache, ensure_ascii=False, default=str), encoding="utf-8")
print("WROTE", CACHE, "bytes", CACHE.stat().st_size)
