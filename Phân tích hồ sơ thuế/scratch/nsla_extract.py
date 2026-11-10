# -*- coding: utf-8 -*-
"""Extract NONG SAN LUU HOAI ANH 3603057410 sources to JSON cache."""
from __future__ import annotations

import hashlib
import io
import json
import os
import re
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

MST = "3603057410"
TEN = "CÔNG TY TNHH NÔNG SẢN LƯU HOÀI ANH"
KY = "2022-2025"
YEARS = [2022, 2023, 2024, 2025]
TODAY = "2026-11-09"

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Ke hoach nam\Bo sung\NONG_SAN_LUU_HOAI_ANH_3603057410"
)
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
ENGINE = VAULT / "skills" / "phan-tich-ho-so-thue" / "engine"
sys.path.insert(0, str(ENGINE))
from ingest import xml_tkhai as xt  # noqa: E402

# BCTC TT133 B01a-DNN uses maTKhai 683
xt.MA_LOAI["683"] = "bctc"
xt.MA_LOAI["402"] = "bctc"

CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
CACHE = Path(r"C:\Users\NGOC\AppData\Local\Temp\nsla_cache.json")


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
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
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
    s = str(v).strip().replace("00:00:00", "").strip()
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
    if "-" in s:
        a, b = s.split("-", 1)
        a = re.sub(r"\D", "", a)
        b = re.sub(r"\D", "", b)
        return f"{a}-{b}" if a and b else a
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


def classify_file(rel: str) -> tuple[str, str]:
    u = rel.replace("/", "\\").upper()
    if "1. DU LIEU THO NGANH THUE" in u:
        src = "CQT"
    elif "2. HO SO NNT" in u:
        src = "DN"
    elif "3. DU LIEU BO SUNG" in u:
        src = "TRA_CUU"
    elif "0. BIEN BAN" in u:
        src = "BB"
    else:
        src = "KHAC"
    if "1.1" in u or "TO KHAI" in u:
        inn = "IN03"
        if "TNDN" in u or "BCTC" in u:
            inn = "IN04"
        if "TNCN" in u or "QTT" in u:
            inn = "IN05"
        if "DANH_SACH_TO_KHAI" in u:
            inn = "IN00"
    elif "1.2" in u or "HDDT" in u:
        inn = "IN01"
    elif "1.3" in u or "NGHIA VU" in u:
        inn = "IN14"
    elif "1.4" in u:
        inn = "IN05"
    elif "2.5" in u or "BANG KE" in u:
        inn = "IN02"
    elif "2.1" in u:
        inn = "IN08"
    elif "2.2" in u:
        inn = "IN09"
    elif "2.3" in u:
        inn = "IN11"
    else:
        inn = "KHAC"
    return src, inn


def tk_to_dict(tk) -> dict:
    return {
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


def parse_any_xml(path: Path):
    tk = xt.parse_xml(path)
    if tk:
        return tk
    try:
        tree = ET.parse(lp(path))
        root = tree.getroot()
    except Exception:
        return None
    ma = xt._text(xt._find(root, "maTKhai"))
    ten_tk = xt._text(xt._find(root, "tenTKhai"))
    loai = xt.MA_LOAI.get(ma)
    if not loai:
        if "BCTC" in (ten_tk or "") or "B01" in (ten_tk or ""):
            loai = "bctc"
        elif "GTGT" in (ten_tk or ""):
            loai = "gtgt"
        elif "TNDN" in (ten_tk or ""):
            loai = "tndn"
        elif "TNCN" in (ten_tk or ""):
            loai = "tncn"
        else:
            return None
    ky_el = xt._find(root, "KyKKhaiThue")
    ky = xt._text(xt._child(ky_el, "kyKKhai")) if ky_el is not None else ""
    nam = xt._nam_from_ky(ky)
    nnt = xt._find(root, "NNT")
    tk = xt.ToKhai(
        loai=loai,
        ma_tkhai=ma,
        loai_tkhai=xt._text(xt._find(root, "loaiTKhai")) or "C",
        so_lan=int(xt._num(xt._text(xt._find(root, "soLan")))),
        ky=ky,
        nam=nam,
        tu_ngay=xt._text(xt._child(ky_el, "kyKKhaiTuNgay")) if ky_el is not None else "",
        den_ngay=xt._text(xt._child(ky_el, "kyKKhaiDenNgay")) if ky_el is not None else "",
        mst=xt._text(xt._child(nnt, "mst")) if nnt is not None else "",
        ten_nnt=xt._text(xt._child(nnt, "tenNNT")) if nnt is not None else "",
        chi_tieu={},
        file=str(path),
    )
    if loai == "bctc":
        tk.bctc_cdkt_cuoi = xt._ct_map(xt._find(root, "SoCuoiNam"))
        tk.bctc_cdkt_dau = xt._ct_map(xt._find(root, "SoDauNam"))
        kq = xt._find(root, "PL_KQHDSXKD")
        if kq is not None:
            tk.bctc_kqkd = xt._ct_map(xt._child(kq, "NamNay"))
            tk.chi_tieu = dict(tk.bctc_kqkd)
    elif loai == "gtgt":
        tk.chi_tieu = xt._ct_map(xt._find(root, "CTieuTKhaiChinh"))
        for code, val in xt._deep_ct(root).items():
            tk.chi_tieu.setdefault(code, val)
    elif loai == "tndn":
        pl = xt._find(root, "PLuc_03_1A_TNDN") or xt._find(root, "CTieuTKhaiChinh")
        tk.chi_tieu = xt._ct_map(pl) if pl is not None else xt._deep_ct(root)
        for code, val in xt._deep_ct(root).items():
            tk.chi_tieu.setdefault(code, val)
    elif loai == "tncn":
        tk.chi_tieu = xt._deep_ct(xt._find(root, "CTieuTKhaiChinh") or root)
    return tk


def xml_from_zip(zpath: Path):
    out = []
    with zipfile.ZipFile(lp(zpath)) as z:
        for name in z.namelist():
            if not name.lower().endswith(".xml"):
                continue
            data = z.read(name)
            tmp = Path(tempfile.gettempdir()) / ("nsla_" + hashlib.md5(name.encode()).hexdigest() + ".xml")
            tmp.write_bytes(data)
            tk = parse_any_xml(tmp)
            try:
                tmp.unlink()
            except OSError:
                pass
            out.append((name, tk))
    return out


print("SCAN files")
manifest = []
skip_top = {"4. KET QUA PHAN TICH"}
for dirpath, dirs, files in os.walk(lp(ROOT)):
    dirs[:] = [d for d in dirs if d not in skip_top]
    for fn in files:
        full = os.path.join(dirpath, fn)
        try:
            st = os.stat(full)
        except OSError:
            continue
        rel = os.path.relpath(full, lp(ROOT))
        src, inn = classify_file(rel)
        try:
            h = sha256(Path(full))
        except Exception:
            h = ""
        manifest.append(
            {
                "rel": rel,
                "name": fn,
                "ext": os.path.splitext(fn)[1].lower(),
                "size": st.st_size,
                "mtime": datetime.fromtimestamp(st.st_mtime).date().isoformat(),
                "sha256": h,
                "nguon": src,
                "ma_in": inn,
            }
        )
print(" files", len(manifest))

print("PARSE XML")
xml_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.1. To khai XML"
xml_rows = []
for fn in os.listdir(lp(xml_dir)):
    p = xml_dir / fn
    ext = p.suffix.lower()
    if ext not in {".xml", ".zip"}:
        continue
    recs = []
    if ext == ".zip":
        try:
            recs = xml_from_zip(p)
        except Exception as e:
            xml_rows.append({"file": fn, "parsed": False, "err": str(e)})
            continue
    else:
        recs = [(fn, parse_any_xml(p))]
    for inner, tk in recs:
        rec = {
            "file": fn,
            "inner": inner,
            "rel": str(Path("1. DU LIEU THO NGANH THUE") / "1.1. To khai XML" / fn),
            "size": os.path.getsize(lp(p)),
            "parsed": tk is not None,
        }
        if tk:
            rec.update(tk_to_dict(tk))
            rec["file"] = fn
        xml_rows.append(rec)
print(" xml", len(xml_rows), "parsed", sum(1 for x in xml_rows if x.get("parsed")))
print(" loai", sorted({x.get("loai") for x in xml_rows if x.get("parsed")}))


def ky_norm(x: dict) -> str:
    ky = str(x.get("ky") or "").strip()
    nam = x.get("nam") or 0
    if re.fullmatch(r"[1-4]/20\d{2}", ky):
        q, y = ky.split("/")
        return f"Q{q}/{y}"
    if re.fullmatch(r"Q[1-4]/20\d{2}", ky, re.I):
        return ky.upper()
    if re.fullmatch(r"20\d{2}", ky):
        return ky
    if nam:
        return f"{ky}|{nam}"
    return ky


def latest_xml(loai: str):
    groups = defaultdict(list)
    for x in xml_rows:
        if not x.get("parsed") or x.get("loai") != loai:
            continue
        groups[(x.get("nam"), ky_norm(x))].append(x)
    out = []
    for key, rows in groups.items():
        rows_sorted = sorted(
            rows,
            key=lambda r: (
                0 if str(r.get("loai_tkhai")).upper() in {"B", "2"} else 1,
                -int(r.get("so_lan") or 0),
                -int(r.get("size") or 0),
            ),
        )
        pick = dict(rows_sorted[0])
        pick["la_hieu_luc"] = True
        pick["so_ban"] = len(rows)
        pick["ky_norm"] = key[1]
        out.append(pick)
    return sorted(out, key=lambda r: (r.get("nam") or 0, str(r.get("ky_norm"))))


gtgt_eff = latest_xml("gtgt")
tndn_eff = latest_xml("tndn")
bctc_eff = latest_xml("bctc")
tncn_qtt = latest_xml("tncn")
print(" GTGT ky", [g.get("ky_norm") for g in gtgt_eff])
print(" TNDN", [g.get("nam") for g in tndn_eff], "BCTC", [g.get("nam") for g in bctc_eff], "TNCN", [g.get("nam") for g in tncn_qtt])

print("PARSE HDDT")
hddt_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu"
hddt = []
for fn in os.listdir(lp(hddt_dir)):
    if not fn.lower().endswith(".xlsx"):
        continue
    p = hddt_dir / fn
    chieu = "MUA" if "MUA" in fn.upper() else ("BAN" if "BAN" in fn.upper() else None)
    if chieu is None:
        continue
    wb = load_workbook(lp(p), read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, row in enumerate(rows[:12]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "số hóa đơn" in joined and "ngày lập" in joined:
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = [str(c or "").strip() for c in rows[hdr_i]]

        def col(*keys):
            for i, h in enumerate(hdr):
                hl = h.lower()
                if all(k.lower() in hl for k in keys):
                    return i
            return None

        i_kh = col("ký hiệu hđ") or col("ky hieu")
        i_so = col("số hóa đơn")
        i_ng = col("ngày lập")
        i_mst_b = col("mst người bán")
        i_ten_b = col("tên người bán")
        i_mst_m = col("mst người mua")
        i_ten_m = col("tên người mua")
        i_mh = col("mặt hàng")
        i_dvt = col("đơn vị tính")
        i_sl = col("số lượng")
        i_dg = col("đơn giá")
        i_gt = col("tiền chưa thuế")
        i_ts = col("thuế suất")
        i_thue = col("tiền thuế")
        i_tt = col("tổng thanh toán")
        i_ttien = col("đơn vị tiền")
        for row in rows[hdr_i + 1 :]:
            so_raw = row[i_so] if i_so is not None and i_so < len(row) else None
            so = so_hd_key(so_raw)
            if not so or so in {"0", "STT"}:
                continue
            ngay = ndate(row[i_ng] if i_ng is not None else None)
            nam = year_from_date(ngay)
            quy = quy_from_date(ngay)
            mst_b = norm_mst(row[i_mst_b] if i_mst_b is not None else None)
            mst_m = norm_mst(row[i_mst_m] if i_mst_m is not None else None)
            hddt.append(
                {
                    "chieu": chieu,
                    "nam": nam,
                    "quy": quy,
                    "thang": int(ngay[5:7]) if ngay else None,
                    "ky_hieu": str(row[i_kh] or "").strip() if i_kh is not None else "",
                    "so_hd": str(so_raw or "").strip(),
                    "so_key": so,
                    "ngay": ngay,
                    "mst_nb": mst_b,
                    "ten_nb": str(row[i_ten_b] or "").strip() if i_ten_b is not None else "",
                    "mst_nm": mst_m,
                    "ten_nm": str(row[i_ten_m] or "").strip() if i_ten_m is not None else "",
                    "mst_dt": mst_b if chieu == "MUA" else mst_m,
                    "mat_hang": str(row[i_mh] or "").strip() if i_mh is not None else "",
                    "dvt": str(row[i_dvt] or "").strip() if i_dvt is not None else "",
                    "sl": nnum(row[i_sl] if i_sl is not None else 0),
                    "don_gia": nnum(row[i_dg] if i_dg is not None else 0),
                    "gt": nnum(row[i_gt] if i_gt is not None else 0),
                    "thue_suat": row[i_ts] if i_ts is not None else None,
                    "thue": nnum(row[i_thue] if i_thue is not None else 0),
                    "tt": nnum(row[i_tt] if i_tt is not None else 0),
                    "tinh_vao_tong": "Có",
                    "file": fn,
                }
            )
    wb.close()
print(" HDDT", len(hddt), "mua", sum(1 for x in hddt if x["chieu"] == "MUA"), "ban", sum(1 for x in hddt if x["chieu"] == "BAN"))
print(" HDDT years mua", sorted({x["nam"] for x in hddt if x["chieu"] == "MUA"}))
print(" HDDT years ban", sorted({x["nam"] for x in hddt if x["chieu"] == "BAN"}))

print("PARSE CVCBRR")


def load_f0():
    out = []
    if not CVCBRR.exists():
        return out
    wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
    ws = wb["F0"]
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = 6
    for row in rows[hdr_i:]:
        mst = norm_mst(row[1] if len(row) > 1 else None)
        if not mst or not mst[0].isdigit():
            continue
        out.append(
            {
                "mst": mst,
                "base": mst_base(mst),
                "ten": str(row[2] or "").strip() if len(row) > 2 else "",
                "cv": str(row[3] or "").strip() if len(row) > 3 else "",
                "ngay_cv": ndate(row[4] if len(row) > 4 else None),
                "dvbh": str(row[5] or "").strip() if len(row) > 5 else "",
                "nd": str(row[6] or "").strip()[:240] if len(row) > 6 else "",
                "nhom": str(row[8] or "").strip() if len(row) > 8 else "",
                "sheet": "F0",
            }
        )
    wb.close()
    return out


def load_f1():
    out = []
    if not CVCBRR.exists():
        return out
    wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
    ws = wb["F1"]
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = 6
    for row in rows[hdr_i:]:
        mst_ban = norm_mst(row[1] if len(row) > 1 else None)
        mst_mua = norm_mst(row[8] if len(row) > 8 else None)
        if not mst_mua and not mst_ban:
            continue
        out.append(
            {
                "mst_ban": mst_ban,
                "ten_ban": str(row[2] or "").strip() if len(row) > 2 else "",
                "cv": str(row[3] or "").strip() if len(row) > 3 else "",
                "ngay_cv": ndate(row[4] if len(row) > 4 else None),
                "nd": str(row[6] or "").strip()[:240] if len(row) > 6 else "",
                "ten_mua": str(row[7] or "").strip() if len(row) > 7 else "",
                "mst_mua": mst_mua,
                "base_mua": mst_base(mst_mua),
                "sheet": "F1",
            }
        )
    wb.close()
    return out


def load_ca():
    out = []
    if not CA.exists():
        return out
    wb = load_workbook(lp(CA), read_only=True, data_only=True)
    for name in wb.sheetnames:
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, row in enumerate(rows[:12]):
            joined = " ".join(str(c or "").lower() for c in row)
            if "mst" in joined:
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = [str(c or "").strip().lower() for c in rows[hdr_i]]
        i_mst = next((i for i, h in enumerate(hdr) if h in {"mst", "mã số thuế"} or h == "mst"), None)
        if i_mst is None:
            i_mst = next((i for i, h in enumerate(hdr) if "mst" in h), None)
        i_cv = next((i for i, h in enumerate(hdr) if "cv" in h or "công văn" in h), None)
        i_ten = next((i for i, h in enumerate(hdr) if "tên" in h), None)
        if i_mst is None:
            continue
        for row in rows[hdr_i + 1 :]:
            mst = norm_mst(row[i_mst] if i_mst < len(row) else None)
            if not mst or not mst[0].isdigit():
                continue
            out.append(
                {
                    "mst": mst,
                    "base": mst_base(mst),
                    "ten": str(row[i_ten] or "").strip() if i_ten is not None and i_ten < len(row) else "",
                    "cv": str(row[i_cv] or "").strip() if i_cv is not None and i_cv < len(row) else "",
                    "sheet": name,
                }
            )
    wb.close()
    return out


f0 = load_f0()
f1 = load_f1()
ca_list = load_ca()
print(" F0", len(f0), "F1", len(f1), "CA", len(ca_list))
nnt_f1 = [x for x in f1 if x.get("base_mua") == MST]
nnt_ca = [x for x in ca_list if x.get("base") == MST]
print(" NNT on F1", len(nnt_f1), "NNT on CA", len(nnt_ca))

print("PARSE TMS")
tms = []
tms_dir = ROOT / "1. DU LIEU THO NGANH THUE" / "1.3. Nghia vu nop thue & Thong bao"
for fn in os.listdir(lp(tms_dir)):
    p = tms_dir / fn
    ext = os.path.splitext(fn)[1].lower()
    rec = {"file": fn, "ext": ext, "rows_n": 0, "sheets": [], "sample": []}
    if ext not in {".xlsx", ".xls"}:
        rec["size"] = os.path.getsize(lp(p))
        tms.append(rec)
        continue
    try:
        wb = load_workbook(lp(p), read_only=True, data_only=True)
    except Exception as e:
        rec["err"] = str(e)
        tms.append(rec)
        continue
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        rec["sheets"].append({"name": ws.title, "n": len(rows)})
        rec["rows_n"] += len(rows)
        for row in rows[:25]:
            rec["sample"].append(
                [None if c is None else (c if isinstance(c, (int, float)) else str(c)[:80]) for c in row[:12]]
            )
    wb.close()
    tms.append(rec)
print(" tms files", [(t["file"], t.get("rows_n"), t.get("err")) for t in tms])

# seller risk
f0_map = {}
for r in f0:
    b = r["base"]
    if b and b not in f0_map:
        f0_map[b] = r
ca_map = {}
for r in ca_list:
    b = r["base"]
    if b and b not in ca_map:
        ca_map[b] = r

mua = [x for x in hddt if x["chieu"] == "MUA" and x.get("nam") in YEARS]
ncc_hit = []
for x in mua:
    b = mst_base(x["mst_dt"])
    if b in f0_map or b in ca_map:
        ncc_hit.append(x["mst_dt"])
print(" HDDT mua in ky", len(mua), "NCC F0/CA hits invoices", len(ncc_hit), "unique", len(set(ncc_hit)))

cache = {
    "meta": {
        "mst": MST,
        "ten": TEN,
        "ky": KY,
        "years": YEARS,
        "today": TODAY,
        "folder": str(ROOT),
        "ky_suy_ra": "HĐĐT 18/02/2022-31/12/2025 + BCTC/TNDN 2022-2025; GTGT 2026 Q1-Q3 giữ ở IN03 ngoài kỳ",
    },
    "manifest": manifest,
    "xml": xml_rows,
    "gtgt_eff": gtgt_eff,
    "tndn_eff": tndn_eff,
    "bctc_eff": bctc_eff,
    "tncn_qtt": tncn_qtt,
    "hddt": hddt,
    "f1_nnt": nnt_f1[:80],
    "f1_count": len(f1),
    "ca_nnt": nnt_ca[:40],
    "ca_count": len(ca_list),
    "tms": tms,
    "cvcbrr_mtime": datetime.fromtimestamp(CVCBRR.stat().st_mtime).isoformat() if CVCBRR.exists() else None,
    "ca_mtime": datetime.fromtimestamp(CA.stat().st_mtime).isoformat() if CA.exists() else None,
}

# keep only F0/CA of sellers actually in HDDT + NNT to shrink JSON
used = {mst_base(x["mst_dt"]) for x in mua}
used.add(MST)
cache["f0_used"] = [f0_map[b] for b in used if b in f0_map]
cache["ca_used"] = [ca_map[b] for b in used if b in ca_map]
cache["f1_nnt"] = nnt_f1[:50]

CACHE.write_text(json.dumps(cache, ensure_ascii=False, default=str), encoding="utf-8")
print("WROTE", CACHE, "bytes", CACHE.stat().st_size)
