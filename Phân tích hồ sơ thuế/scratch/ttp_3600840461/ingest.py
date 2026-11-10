# -*- coding: utf-8 -*-
"""Ingest hồ sơ THANH THANH PHUONG MST 3600840461 -> cache JSON."""
from __future__ import annotations

import hashlib
import json
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import xlrd
from openpyxl import load_workbook

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH"
    r"\THANH_THANH_PHUONG_3600840461"
)
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
ENGINE = VAULT / "skills" / "phan-tich-ho-so-thue" / "engine"
CACHE = VAULT / "scratch" / "ttp_3600840461" / "cache"
MST = "3600840461"
TEN = "CÔNG TY TNHH THANH THANH PHƯƠNG"
KY = "2020-2025"
YEARS = [2020, 2021, 2022, 2023, 2024, 2025]
TODAY = "2026-11-09"

sys.path.insert(0, str(ENGINE))
from ingest.xml_tkhai import parse_xml, sniff_xml  # noqa: E402

CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON"
    r"\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR"
    r"\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")

SKIP_TOP = {
    "4. KET QUA PHAN TICH",
    "7. Sao luu tong hop",
    "ket-qua",
    "0. BIEN BAN",
}


def lp(p: Path | str) -> str:
    s = str(p)
    if not s.startswith("\\\\?\\"):
        s = "\\\\?\\" + s
    return s


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(lp(path), "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def to_num(v) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\xa0", "")
    if not s or s in {"-", "None"}:
        return None
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s[:10], fmt).date().isoformat()
        except ValueError:
            continue
    return None


def mst_chuan(raw: str) -> str:
    s = re.sub(r"[^0-9\-]", "", str(raw or "").strip())
    if not s:
        return ""
    if "-" in s:
        a, b = s.split("-", 1)
        a = re.sub(r"\D", "", a).zfill(10)[-10:]
        b = re.sub(r"\D", "", b)[:3]
        return f"{a}-{b}" if b else a
    d = re.sub(r"\D", "", s)
    if len(d) == 13:
        return f"{d[:10]}-{d[10:]}"
    if len(d) >= 10:
        return d[:10].zfill(10)
    if d:
        return d.zfill(10)
    return ""


def mst_core(m: str) -> str:
    return (m or "").split("-")[0]


def classify(rel: str) -> tuple[str, str, str]:
    r = rel.replace("/", "\\")
    if r.startswith("1."):
        goc = "CQT"
    elif r.startswith("2."):
        goc = "DN"
    elif r.startswith("3."):
        goc = "TRA_CUU"
    else:
        goc = "KHAC"
    low = r.lower()
    name = Path(r).name.lower()
    if "1.2" in r and "hddt" in name:
        return "IN01_HDDT", goc, "Excel"
    if "2.5" in r or ("bang ke" in low and "01_gtgt" in name):
        return "IN02_BangKe", goc, "Excel"
    if name.endswith(".xml") and any(x in name for x in ["gtgt", "tndn", "tncn", "bctc", "bc26", "tb01"]):
        if "gtgt" in name:
            return "IN03_ToKhai_GTGT", goc, "XML"
        if "bctc" in name:
            return "IN04_BCTC_TNDN", goc, "XML"
        if "tndn" in name:
            return "IN04_BCTC_TNDN", goc, "XML"
        if "tncn" in name:
            return "IN05_NLD_TNCN", goc, "XML"
        return "IN03_ToKhai_GTGT", goc, "XML"
    if "1.3" in r:
        return "IN14_TMS_Nghia_vu", goc, "Excel"
    if "1.4" in r or "nld" in name or "tncn" in name:
        return "IN05_NLD_TNCN", goc, "Excel"
    if "2.1" in r:
        n = name
        if "nkc" in n or "so cai" in n or "sc " in n:
            return "IN08_So_KeToan", goc, "Excel"
        if "131" in n or "331" in n:
            return "IN08_So_KeToan", goc, "Excel"
        if "cdkt" in n or "bcdtk" in n:
            return "IN08_So_KeToan", goc, "Excel"
        return "IN08_So_KeToan", goc, Path(r).suffix.upper().lstrip(".") or "Khac"
    if "2.2" in r or "nxt" in name or "156" in name:
        return "IN09_Kho_GiaThanh", goc, "Excel"
    if "2.3" in r:
        return "IN11_Sao_ke_NH", goc, "Khac"
    if "2.4" in r:
        return "IN08_So_KeToan", goc, "Excel"
    if "3.1" in r:
        return "IN07_NCC_CanhBao", goc, "Excel"
    if "3.2" in r:
        return "IN07_NCC_CanhBao", goc, "Excel"
    if "3.3" in r:
        return "IN05_NLD_TNCN", goc, "Excel"
    return "—", goc, Path(r).suffix.upper().lstrip(".") or "Khac"


def iter_files(root: Path):
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        try:
            rel = str(p.relative_to(root))
        except ValueError:
            continue
        top = rel.split("\\")[0]
        yield p, rel, top


def scan_manifest() -> list[dict]:
    rows = []
    n = 0
    for p, rel, top in iter_files(ROOT):
        n += 1
        if n % 50 == 0:
            print(f"  hash {n} {rel[:80]}", flush=True)
        ma, goc, fmt = classify(rel)
        if top in SKIP_TOP:
            ma = "—"
            goc = "KET_QUA_CU" if top != "0. BIEN BAN" else "BIEN_BAN"
        try:
            h = sha256(p)
            sz = p.stat().st_size
            mtime = datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d")
        except OSError as e:
            h, sz, mtime = "", 0, ""
            print("  ERR", rel, e)
        rows.append(
            {
                "rel": rel,
                "name": p.name,
                "ma_in": ma,
                "goc": goc,
                "fmt": fmt,
                "ext": p.suffix.lower(),
                "size": sz,
                "mtime": mtime,
                "sha256": h,
                "top": top,
            }
        )
    return rows


def parse_all_xml(manifest: list[dict]) -> list[dict]:
    seen_hash = {}
    out = []
    xmls = [m for m in manifest if m["ext"] == ".xml" and m["goc"] == "CQT"]
    print(f"XML CQT files {len(xmls)}", flush=True)
    for m in xmls:
        p = ROOT / m["rel"]
        if m["sha256"] in seen_hash:
            rec = dict(seen_hash[m["sha256"]])
            rec["file"] = m["rel"]
            rec["duplicate_of"] = seen_hash[m["sha256"]]["file"]
            out.append(rec)
            continue
        if not sniff_xml(p):
            out.append({"file": m["rel"], "error": "not HSoThueDTu", "sha256": m["sha256"]})
            continue
        tk = parse_xml(p)
        if tk is None:
            out.append({"file": m["rel"], "error": "parse fail", "sha256": m["sha256"]})
            continue
        rec = {
            "file": m["rel"],
            "sha256": m["sha256"],
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
        seen_hash[m["sha256"]] = rec
        out.append(rec)
    return out


def effective_xml(xmls: list[dict], loai: str) -> dict[str, dict]:
    groups: dict[str, list] = defaultdict(list)
    for x in xmls:
        if x.get("loai") != loai or x.get("error"):
            continue
        if x.get("duplicate_of"):
            continue
        ky = str(x.get("ky") or "")
        groups[ky].append(x)
    eff = {}
    for ky, items in groups.items():
        items = sorted(
            items,
            key=lambda z: (
                1 if str(z.get("loai_tkhai") or "").upper().startswith("B") else 0,
                int(z.get("so_lan") or 0),
            ),
        )
        best = items[-1]
        best = dict(best)
        best["la_hieu_luc"] = "Có"
        eff[ky] = best
    return eff


def load_hddt() -> list[dict]:
    p = ROOT / r"1. DU LIEU THO NGANH THUE\1.2. Hoa don dien tu\HDDT.xlsx"
    wb = load_workbook(lp(p), read_only=True, data_only=True)
    rows = []
    for sheet, chieu in (("Mua vao", "MUA"), ("Ban ra", "BAN")):
        ws = wb[sheet]
        header = [str(c or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        idx = {h: i for i, h in enumerate(header)}

        def col(*names):
            for n in names:
                for h, i in idx.items():
                    if n.lower() in h.lower():
                        return i
            return None

        i_kyhieu = col("Ký hiệu hóa đơn")
        i_so = col("Số hóa đơn")
        i_ngay = col("Ngày lập")
        i_mst_ban = col("MST người bán", "MST người xuất")
        i_ten_ban = col("Tên người bán", "Tên người xuất")
        i_mst_mua = col("MST người mua", "MST người nhận")
        i_ten_mua = col("Tên người mua", "Tên người nhận")
        i_gt = col("Tổng tiền chưa thuế")
        i_thue = col("Tổng tiền thuế")
        i_tt = col("Tổng tiền thanh toán")
        i_tt_hd = col("Trạng thái hóa đơn")
        i_kq = col("Kết quả kiểm tra")
        for r in ws.iter_rows(min_row=2, values_only=True):
            so = str(r[i_so] if i_so is not None else "").strip()
            if not so or so.lower() in {"none", "stt"}:
                continue
            ngay = parse_date(r[i_ngay] if i_ngay is not None else None)
            nam = int(ngay[:4]) if ngay else None
            thang = int(ngay[5:7]) if ngay else None
            st = str(r[i_tt_hd] if i_tt_hd is not None else "")
            tinh = "Không"
            if any(k in st for k in ("mới", "thay thế", "điều chỉnh")) and "đã bị" not in st.lower():
                tinh = "Có"
            if st.strip() in {"Hóa đơn mới", "Hóa đơn thay thế", "Hóa đơn điều chỉnh"}:
                tinh = "Có"
            mst_ban = mst_chuan(r[i_mst_ban] if i_mst_ban is not None else "")
            mst_mua = mst_chuan(r[i_mst_mua] if i_mst_mua is not None else "")
            dt = mst_ban if chieu == "MUA" else mst_mua
            rows.append(
                {
                    "chieu": chieu,
                    "nam": nam,
                    "thang": thang,
                    "ky_hieu": str(r[i_kyhieu] or "").strip() if i_kyhieu is not None else "",
                    "so": so,
                    "ngay": ngay,
                    "mst_ban": mst_ban,
                    "ten_ban": str(r[i_ten_ban] or "").strip() if i_ten_ban is not None else "",
                    "mst_mua": mst_mua,
                    "ten_mua": str(r[i_ten_mua] or "").strip() if i_ten_mua is not None else "",
                    "mst_dt": dt,
                    "gt": to_num(r[i_gt] if i_gt is not None else None) or 0.0,
                    "thue": to_num(r[i_thue] if i_thue is not None else None) or 0.0,
                    "tt": to_num(r[i_tt] if i_tt is not None else None) or 0.0,
                    "trang_thai": st,
                    "kq_kt": str(r[i_kq] or "") if i_kq is not None else "",
                    "tinh": tinh,
                    "file": "1. DU LIEU THO NGANH THUE\\1.2. Hoa don dien tu\\HDDT.xlsx",
                }
            )
    wb.close()
    return rows


def load_bangke() -> list[dict]:
    p = ROOT / r"2. HO SO NNT CUNG CAP\2.5. Bang ke mua vao ban ra\Bang_ke_mua_vao_ban_ra_tong_hop_2020-2025.xlsx"
    wb = load_workbook(lp(p), read_only=True, data_only=True)
    rows = []
    dummy = {"(2)", "(3)", "(4)", "5", "(5)"}
    for sheet, chieu in (("mua vao", "MUA"), ("ban ra", "BAN")):
        ws = wb[sheet]
        header_row = None
        headers = []
        for i, r in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
            vals = [str(c or "").strip() for c in r]
            if "So_HD" in vals or "Số HĐ" in vals or "STT" in vals and "Nam" in vals:
                header_row = i
                headers = vals
                break
        if header_row is None:
            continue
        idx = {h: i for i, h in enumerate(headers) if h}

        def g(r, *names):
            for n in names:
                if n in idx:
                    return r[idx[n]]
            return None

        for r in ws.iter_rows(min_row=header_row + 1, values_only=True):
            so = str(g(r, "So_HD") or "").strip()
            if not so or so in dummy:
                continue
            ten = str(g(r, "Ten_doi_tac") or "").strip()
            if ten in dummy:
                continue
            ngay = parse_date(g(r, "Ngay_HD"))
            nam = to_num(g(r, "Nam"))
            nam = int(nam) if nam else (int(ngay[:4]) if ngay else None)
            gt = to_num(g(r, "Doanh_so_chua_thue")) or 0.0
            thue = to_num(g(r, "Thue_GTGT")) or 0.0
            ts = to_num(g(r, "Thue_suat"))
            mst = mst_chuan(g(r, "MST_doi_tac") or "")
            rows.append(
                {
                    "chieu": chieu,
                    "nam": nam,
                    "thang": int(ngay[5:7]) if ngay else None,
                    "ky_hieu": str(g(r, "Seri_HD") or "").strip(),
                    "so": so,
                    "ngay": ngay,
                    "mst_goc": str(g(r, "MST_doi_tac") or "").strip(),
                    "mst": mst,
                    "ten": ten,
                    "hang": str(g(r, "Mat_hang") or "").strip(),
                    "gt": gt,
                    "thue": thue,
                    "ts": ts,
                    "file": str(g(r, "Nguon_file") or p.name),
                }
            )
    wb.close()
    return rows


def _nkc_iter_xlsx(path: Path):
    wb = load_workbook(lp(path), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        yield i, r
    wb.close()


def _nkc_iter_xls(path: Path):
    book = xlrd.open_workbook(lp(path))
    sh = book.sheet_by_index(0)
    for i in range(sh.nrows):
        row = []
        for j in range(sh.ncols):
            c = sh.cell(i, j)
            v = c.value
            if c.ctype == xlrd.XL_CELL_DATE:
                try:
                    v = xlrd.xldate_as_datetime(c.value, book.datemode)
                except Exception:
                    pass
            row.append(v)
        yield i + 1, tuple(row)


def parse_nkc(path: Path, nam_hint: int | None) -> dict:
    it = _nkc_iter_xlsx(path) if path.suffix.lower() == ".xlsx" else _nkc_iter_xls(path)
    by_tk = defaultdict(lambda: {"no": 0.0, "co": 0.0, "n": 0})
    by_tk_thang = defaultdict(lambda: {"no": 0.0, "co": 0.0})
    lines_511 = []
    lines_632 = []
    cash_chi = []
    vouchers = []
    cur = None
    n_line = 0
    for i, r in it:
        if not r or all(c is None or str(c).strip() == "" for c in r):
            continue
        a, b, c = r[0] if len(r) > 0 else None, r[1] if len(r) > 1 else None, r[2] if len(r) > 2 else None
        d = r[3] if len(r) > 3 else None
        g = r[6] if len(r) > 6 else None
        h = r[7] if len(r) > 7 else None
        ii = r[8] if len(r) > 8 else None
        ngay = parse_date(a) or parse_date(c)
        so_ct = str(b or "").strip()
        tk = str(g or "").strip().replace(" ", "")
        ps_no = to_num(h) or 0.0
        ps_co = to_num(ii) or 0.0
        if ngay and so_ct and so_ct not in {"Số hiệu", "Chứng từ"}:
            cur = {"ngay": ngay, "so": so_ct, "diengiai": str(d or ""), "lines": []}
            vouchers.append(cur)
        if not tk or not re.match(r"^\d", tk):
            continue
        n_line += 1
        tk_main = re.match(r"^(\d{3,4})", tk)
        tk3 = tk_main.group(1)[:3] if tk_main else tk[:3]
        by_tk[tk]["no"] += ps_no
        by_tk[tk]["co"] += ps_co
        by_tk[tk]["n"] += 1
        thang = int(ngay[5:7]) if ngay else 0
        nam = int(ngay[:4]) if ngay else nam_hint
        by_tk_thang[(nam, thang, tk3)]["no"] += ps_no
        by_tk_thang[(nam, thang, tk3)]["co"] += ps_co
        if cur:
            cur["lines"].append({"tk": tk, "no": ps_no, "co": ps_co, "dg": str(d or "")})
        if tk.startswith("511"):
            lines_511.append(
                {
                    "nam": nam,
                    "thang": thang,
                    "ngay": ngay,
                    "so": cur["so"] if cur else so_ct,
                    "tk": tk,
                    "dg": str(d or ""),
                    "co": ps_co,
                    "no": ps_no,
                    "file": path.name,
                }
            )
        if tk.startswith("632"):
            lines_632.append(
                {
                    "nam": nam,
                    "thang": thang,
                    "ngay": ngay,
                    "so": cur["so"] if cur else so_ct,
                    "tk": tk,
                    "dg": str(d or ""),
                    "co": ps_co,
                    "no": ps_no,
                    "file": path.name,
                }
            )
    # chi tiền mặt: phiếu PC, TK 111 PS Có cùng chứng từ với 331
    for v in vouchers:
        so = (v.get("so") or "").upper()
        tks = {ln["tk"][:3] for ln in v["lines"]}
        chi_111 = sum(ln["co"] for ln in v["lines"] if ln["tk"].startswith("111"))
        if chi_111 > 0 and (so.startswith("PC") or "331" in tks or "156" in tks):
            mst = ""
            ten = v.get("diengiai") or ""
            cash_chi.append(
                {
                    "ngay": v["ngay"],
                    "so": v["so"],
                    "tien": chi_111,
                    "dg": ten,
                    "co_331": "331" in tks,
                    "file": path.name,
                }
            )
    tk_sum = {k: v for k, v in by_tk.items()}
    thang_sum = {f"{k[0]}-{k[1]:02d}-{k[2]}": v for k, v in by_tk_thang.items()}
    return {
        "file": path.name,
        "nam": nam_hint,
        "n_line": n_line,
        "tk": tk_sum,
        "thang": thang_sum,
        "n_511": len(lines_511),
        "n_632": len(lines_632),
        "n_cash": len(cash_chi),
        "lines_511": lines_511,
        "lines_632": lines_632,
        "cash_chi": cash_chi,
    }


def load_all_nkc() -> list[dict]:
    d = ROOT / r"2. HO SO NNT CUNG CAP\2.1. So ke toan tong hop va chi tiet"
    out = []
    for p in sorted(d.iterdir()):
        if p.is_file() and "nkc" in p.name.lower():
            m = re.search(r"(20\d{2})", p.name)
            nam = int(m.group(1)) if m else None
            print(f"  NKC {p.name}", flush=True)
            out.append(parse_nkc(p, nam))
    return out


def load_cn_131_331() -> list[dict]:
    d = ROOT / r"2. HO SO NNT CUNG CAP\2.1. So ke toan tong hop va chi tiet"
    out = []
    for p in sorted(d.iterdir(), key=lambda x: x.name):
        if not p.is_file():
            continue
        low = p.name.lower()
        if "131" not in low and "331" not in low:
            continue
        tk = "131" if "131" in low else "331"
        m = re.search(r"(20\d{2}|0225)", p.name)
        nam = 2025 if m and m.group(1) == "0225" else (int(m.group(1)) if m else None)
        tot = {"dk_no": 0.0, "dk_co": 0.0, "ps_no": 0.0, "ps_co": 0.0, "ck_no": 0.0, "ck_co": 0.0, "n": 0}
        try:
            if p.suffix.lower() == ".xlsx":
                wb = load_workbook(lp(p), read_only=True, data_only=True)
                ws = wb[wb.sheetnames[0]]
                rows = list(ws.iter_rows(values_only=True))
                wb.close()
            else:
                book = xlrd.open_workbook(lp(p))
                sh = book.sheet_by_index(0)
                rows = [tuple(sh.cell_value(i, j) for j in range(sh.ncols)) for i in range(sh.nrows)]
        except Exception as e:
            out.append({"file": p.name, "tk": tk, "nam": nam, "error": str(e)})
            continue
        start = 0
        for i, r in enumerate(rows):
            vals = [str(c or "") for c in r]
            if any("Dư đầu" in v or "Du dau" in v or "Phát sinh" in v for v in vals):
                start = i + 2
                break
        for r in rows[start:]:
            if not r or all(c is None or str(c).strip() == "" for c in r[:8]):
                continue
            code = str(r[0] or "").strip()
            if not code or code.lower() in {"cộng", "tong", "tổng", "total"}:
                if any(to_num(x) for x in r[2:8]):
                    pass
                else:
                    continue
            nums = [to_num(x) or 0.0 for x in r[2:8]]
            if len(nums) < 6:
                continue
            tot["dk_no"] += nums[0]
            tot["dk_co"] += nums[1]
            tot["ps_no"] += nums[2]
            tot["ps_co"] += nums[3]
            tot["ck_no"] += nums[4]
            tot["ck_co"] += nums[5]
            tot["n"] += 1
        tot.update({"file": p.name, "tk": tk, "nam": nam})
        out.append(tot)
    return out


def load_cvcbrr() -> tuple[dict, dict]:
    f0, f1 = {}, {}
    wb = load_workbook(lp(CVCBRR), read_only=True, data_only=True)
    ws = wb["F0"]
    for i, r in enumerate(ws.iter_rows(min_row=8, values_only=True), 8):
        mst = mst_chuan(r[1] if len(r) > 1 else "")
        if not mst:
            continue
        rec = {
            "mst": mst,
            "ten": str(r[2] or "").strip() if len(r) > 2 else "",
            "so_cv": str(r[3] or "").strip() if len(r) > 3 else "",
            "ngay_cv": parse_date(r[4] if len(r) > 4 else None),
            "dv": str(r[5] or "").strip() if len(r) > 5 else "",
            "nd": str(r[6] or "").strip() if len(r) > 6 else "",
            "nhom": str(r[8] or "").strip() if len(r) > 8 else "",
            "nguon": "CANH_BAO",
        }
        f0[mst_core(mst)] = rec
        f0[mst] = rec
    ws = wb["F1"]
    f1_nnt = []
    for i, r in enumerate(ws.iter_rows(min_row=8, values_only=True), 8):
        mst_ban = mst_chuan(r[1] if len(r) > 1 else "")
        mst_mua = mst_chuan(r[8] if len(r) > 8 else "")
        rec = {
            "mst_ban": mst_ban,
            "ten_ban": str(r[2] or "").strip() if len(r) > 2 else "",
            "so_cv": str(r[3] or "").strip() if len(r) > 3 else "",
            "ngay_cv": parse_date(r[4] if len(r) > 4 else None),
            "mst_mua": mst_mua,
            "ten_mua": str(r[7] or "").strip() if len(r) > 7 else "",
        }
        if mst_mua:
            f1.setdefault(mst_core(mst_mua), []).append(rec)
            if mst_core(mst_mua) == MST:
                f1_nnt.append(rec)
    wb.close()
    return {"f0": {k: v for k, v in f0.items()}, "f1_nnt": f1_nnt, "f1_index": {k: len(v) for k, v in f1.items()}}, f0


def load_ca(f0: dict) -> dict:
    ca = {}
    if not CA.exists():
        return {"ok": False, "reason": "missing"}
    try:
        wb = load_workbook(lp(CA), read_only=True, data_only=True)
    except Exception as e:
        return {"ok": False, "reason": str(e)}
    hits_nnt = []
    for name in wb.sheetnames:
        ws = wb[name]
        rows = ws.iter_rows(min_row=1, max_row=8, values_only=True)
        headers = []
        header_i = 1
        for i, r in enumerate(rows, 1):
            vals = [str(c or "").strip() for c in r]
            if any("mst" in v.lower() or "mã số thuế" in v.lower() for v in vals):
                headers = vals
                header_i = i
                break
        if not headers:
            continue
        idx = {h.lower(): i for i, h in enumerate(headers)}
        i_mst = None
        for k, i in idx.items():
            if "mst" in k or "mã số thuế" in k:
                i_mst = i
                break
        if i_mst is None:
            continue
        n = 0
        for r in ws.iter_rows(min_row=header_i + 1, values_only=True):
            n += 1
            mst = mst_chuan(r[i_mst] if i_mst < len(r) else "")
            if not mst:
                continue
            core = mst_core(mst)
            rec = {"mst": mst, "sheet": name, "row": n}
            ca[core] = rec
            if core == MST:
                hits_nnt.append(rec)
            if core not in f0:
                f0[core] = {
                    "mst": mst,
                    "ten": "",
                    "so_cv": "",
                    "ngay_cv": None,
                    "dv": "Cong an",
                    "nd": "CA",
                    "nhom": "CA",
                    "nguon": "CA",
                }
        print(f"  CA sheet {name} rows~{n} mst_map {len(ca)}", flush=True)
    wb.close()
    return {"ok": True, "n_mst": len(ca), "nnt_hits": hits_nnt, "mtime": datetime.fromtimestamp(CA.stat().st_mtime).strftime("%Y-%m-%d")}


def load_danhba() -> dict:
    p = ROOT / r"3. DU LIEU BO SUNG & TRA CUU QLT\3.2. Danh ba trang thai NNT\Danh_sach_doi_tac_HDDT_Bang_ke.xlsx"
    out = {}
    if not p.exists():
        return out
    wb = load_workbook(lp(p), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), 1):
        vals = [str(c or "").strip() for c in r]
        if "Mã số thuế" in vals or "MST" in vals:
            header = vals
            start = i
            break
    if not header:
        wb.close()
        return out
    idx = {h: i for i, h in enumerate(header)}
    for r in ws.iter_rows(min_row=start + 1, values_only=True):
        mst = mst_chuan(r[idx.get("Mã số thuế", 1)] if len(r) > 1 else "")
        if not mst:
            continue
        st = ""
        for h, i in idx.items():
            if "trạng thái" in h.lower() or "dkkd" in h.lower():
                st = str(r[i] or "")
        out[mst_core(mst)] = {
            "mst": mst,
            "ten": str(r[idx.get("Tên đối tác / Doanh nghiệp", 2)] or "") if len(r) > 2 else "",
            "tt": st,
        }
    wb.close()
    return out


def load_nld() -> list[dict]:
    p = ROOT / r"1. DU LIEU THO NGANH THUE\1.4. Tong hop danh sach NLD theo QTT TNCN\Tong_hop_NLD_QTT_TNCN.xlsx"
    if not p.exists():
        p = ROOT / r"1. DU LIEU THO NGANH THUE\1.4. Tong hop danh sach NLD theo QTT TNCN\Tong_hop_TNCN_va_danh_sach_NLD_3600840461.xlsx"
    if not p.exists():
        return []
    wb = load_workbook(lp(p), read_only=True, data_only=True)
    rows = []
    for name in wb.sheetnames:
        ws = wb[name]
        headers = None
        start = 1
        for i, r in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True), 1):
            vals = [str(c or "").strip() for c in r]
            if any("họ" in v.lower() or "ho_ten" in v.lower() or "năm" in v.lower() or "Nam" == v for v in vals):
                headers = vals
                start = i
                break
        if not headers:
            continue
        idx = {h: i for i, h in enumerate(headers)}
        for r in ws.iter_rows(min_row=start + 1, values_only=True):
            rec = {headers[i]: r[i] if i < len(r) else None for i in range(len(headers))}
            if any(rec.values()):
                rows.append({"sheet": name, **{k: (v if not isinstance(v, datetime) else v.date().isoformat()) for k, v in rec.items()}})
    wb.close()
    return rows


def year_sum(rows, chieu, nam, field="gt", tinh_only=False):
    s = 0.0
    n = 0
    for r in rows:
        if r.get("chieu") != chieu:
            continue
        if r.get("nam") != nam:
            continue
        if tinh_only and r.get("tinh") != "Có":
            continue
        s += float(r.get(field) or 0)
        n += 1
    return s, n


def dump(obj, name):
    CACHE.mkdir(parents=True, exist_ok=True)
    p = CACHE / name
    p.write_text(json.dumps(obj, ensure_ascii=False, default=str), encoding="utf-8")
    print("wrote", p, "bytes", p.stat().st_size, flush=True)


def main():
    print("== SCAN ==", flush=True)
    manifest = scan_manifest()
    dump(manifest, "manifest.json")
    print("files", len(manifest), flush=True)

    print("== XML ==", flush=True)
    xmls = parse_all_xml(manifest)
    dump(xmls, "xml.json")
    eff_gtgt = effective_xml(xmls, "gtgt")
    eff_tndn = effective_xml(xmls, "tndn")
    eff_bctc = effective_xml(xmls, "bctc")
    eff_tncn = effective_xml(xmls, "tncn")
    print("eff GTGT", len(eff_gtgt), "TNDN", len(eff_tndn), "BCTC", len(eff_bctc), "TNCN", len(eff_tncn), flush=True)

    print("== HDDT ==", flush=True)
    hddt = load_hddt()
    dump(hddt, "hddt.json")
    print("hddt", len(hddt), flush=True)

    print("== BANG KE ==", flush=True)
    bk = load_bangke()
    dump(bk, "bangke.json")
    print("bangke", len(bk), flush=True)

    print("== NKC ==", flush=True)
    nkc = load_all_nkc()
    # shrink lines for cache? keep them, needed for IN13
    dump(nkc, "nkc.json")

    print("== 131/331 ==", flush=True)
    cn = load_cn_131_331()
    dump(cn, "congno.json")
    print(cn, flush=True)

    print("== CVCBRR ==", flush=True)
    meta, f0 = load_cvcbrr()
    ca = load_ca(f0)
    danhba = load_danhba()
    dump({"cvcbrr_mtime": datetime.fromtimestamp(CVCBRR.stat().st_mtime).strftime("%Y-%m-%d"),
          "f0_n": len(set(mst_core(k) for k in f0)),
          "f1_nnt": meta["f1_nnt"],
          "ca": ca,
          "danhba_n": len(danhba)}, "nguon_rui_ro.json")
    dump(f0, "f0.json")
    dump(danhba, "danhba.json")

    print("== NLD ==", flush=True)
    nld = load_nld()
    dump(nld, "nld.json")
    print("nld rows", len(nld), flush=True)

    # yearly summary
    years = {}
    for y in YEARS:
        hd_ban_gt, hd_ban_n = year_sum(hddt, "BAN", y, "gt", True)
        hd_ban_th, _ = year_sum(hddt, "BAN", y, "thue", True)
        hd_mua_gt, hd_mua_n = year_sum(hddt, "MUA", y, "gt", True)
        hd_mua_th, _ = year_sum(hddt, "MUA", y, "thue", True)
        bk_ban_gt, bk_ban_n = year_sum(bk, "BAN", y, "gt")
        bk_ban_th, _ = year_sum(bk, "BAN", y, "thue")
        bk_mua_gt, bk_mua_n = year_sum(bk, "MUA", y, "gt")
        bk_mua_th, _ = year_sum(bk, "MUA", y, "thue")
        gtgt_34 = gtgt_35 = gtgt_23 = gtgt_24 = gtgt_40 = gtgt_43 = 0.0
        n_ky = 0
        for ky, rec in eff_gtgt.items():
            if rec.get("nam") != y:
                continue
            n_ky += 1
            ct = rec.get("chi_tieu") or {}
            gtgt_34 += float(ct.get("34") or 0)
            gtgt_35 += float(ct.get("35") or 0)
            gtgt_23 += float(ct.get("23") or 0)
            gtgt_24 += float(ct.get("24") or 0)
            gtgt_40 += float(ct.get("40") or 0)
            gtgt_43 += float(ct.get("43") or 0)
        tndn = eff_tndn.get(str(y), {})
        bctc = eff_bctc.get(str(y), {})
        years[str(y)] = {
            "hddt_ban_gt": hd_ban_gt, "hddt_ban_th": hd_ban_th, "hddt_ban_n": hd_ban_n,
            "hddt_mua_gt": hd_mua_gt, "hddt_mua_th": hd_mua_th, "hddt_mua_n": hd_mua_n,
            "bk_ban_gt": bk_ban_gt, "bk_ban_th": bk_ban_th, "bk_ban_n": bk_ban_n,
            "bk_mua_gt": bk_mua_gt, "bk_mua_th": bk_mua_th, "bk_mua_n": bk_mua_n,
            "gtgt_nky": n_ky,
            "gtgt_34": gtgt_34, "gtgt_35": gtgt_35, "gtgt_23": gtgt_23, "gtgt_24": gtgt_24,
            "gtgt_40": gtgt_40, "gtgt_43": gtgt_43,
            "tndn_ct04": (tndn.get("chi_tieu") or {}).get("04"),
            "tndn_ct12": (tndn.get("chi_tieu") or {}).get("12"),
            "tndn_ct13": (tndn.get("chi_tieu") or {}).get("13"),
            "tndn_ct16": (tndn.get("chi_tieu") or {}).get("16"),
            "tndn_ct18": (tndn.get("chi_tieu") or {}).get("18"),
            "bctc_01": (bctc.get("kqkd") or {}).get("01"),
            "bctc_10": (bctc.get("kqkd") or {}).get("10"),
            "bctc_11": (bctc.get("kqkd") or {}).get("11"),
            "bctc_20": (bctc.get("kqkd") or {}).get("20"),
            "bctc_22": (bctc.get("kqkd") or {}).get("22"),
            "bctc_30": (bctc.get("kqkd") or {}).get("30"),
            "bctc_50": (bctc.get("kqkd") or {}).get("50"),
            "bctc_140": (bctc.get("cdkt_cuoi") or {}).get("140"),
            "bctc_141": (bctc.get("cdkt_cuoi") or {}).get("141"),
            "bctc_311": (bctc.get("cdkt_cuoi") or {}).get("311"),
            "bctc_312": (bctc.get("cdkt_cuoi") or {}).get("312"),
            "bctc_411": (bctc.get("cdkt_cuoi") or {}).get("411"),
            "bctc_421": (bctc.get("cdkt_cuoi") or {}).get("421"),
            "che_do": "TT200",
        }
    dump({"mst": MST, "ten": TEN, "ky": KY, "years": years,
          "eff_gtgt_keys": list(eff_gtgt.keys()),
          "eff_tndn_keys": list(eff_tndn.keys()),
          "eff_bctc_keys": list(eff_bctc.keys()),
          "eff_tncn_keys": list(eff_tncn.keys())}, "summary.json")
    dump(eff_gtgt, "eff_gtgt.json")
    dump(eff_tndn, "eff_tndn.json")
    dump(eff_bctc, "eff_bctc.json")
    dump(eff_tncn, "eff_tncn.json")
    print("DONE ingest", flush=True)
    for y, rec in years.items():
        print(y, "GTGT34", rec["gtgt_34"], "BK_ban", rec["bk_ban_gt"], "HDDT_ban", rec["hddt_ban_gt"],
              "BCTC01", rec["bctc_01"], "nky", rec["gtgt_nky"])


if __name__ == "__main__":
    main()
