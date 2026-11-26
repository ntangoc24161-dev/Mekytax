# -*- coding: utf-8 -*-
"""Tổng hợp HĐĐT + bảng kê 4 sheet (3 giai đoạn) - CÔNG TY TNHH MTV NHẤT VINH MST 3601946069."""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MST = "3601946069"
TEN = "CÔNG TY TNHH MTV NHẤT VINH"
KY = "2018-2025"
TODAY = date(2026, 11, 25)

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\NHAT_VINH_3601946069"
)
HDDT_MUA = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu" / "HDDT_MUA VAO.xlsx"
HDDT_BAN = ROOT / "1. DU LIEU THO NGANH THUE" / "1.2. Hoa don dien tu" / "HDDT _BAN RA.xlsx"
BK_DIR = ROOT / "2. HO SO NNT CUNG CAP" / "2.5. Bang ke mua vao ban ra"
DANHBA = ROOT / "3. DU LIEU BO SUNG & TRA CUU QLT" / "3.2. Danh ba trang thai NNT" / "Danh ba NNT.xls"
CVCBRR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
CA = Path(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
OUT_G = ROOT / "4. KET QUA PHAN TICH" / "cau-thanh" / f"Tong_hop_HDDT_BangKe_{MST}_{KY}.xlsx"
OUT_VAULT = (
    Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
    / "ket-qua"
    / "cau-thanh"
    / f"Tong_hop_HDDT_BangKe_{MST}_{KY}.xlsx"
)

FONT = Font(name="Times New Roman", size=10)
FONT_B = Font(name="Times New Roman", size=10, bold=True)
FONT_H = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_T = Font(name="Times New Roman", size=12, bold=True, color="1F497D")
FONT_S = Font(name="Times New Roman", size=11, bold=True)
FILL_H = PatternFill("solid", fgColor="1F497D")
FILL_T = PatternFill("solid", fgColor="DCE6F1")
FILL_R = PatternFill("solid", fgColor="FCE4D6")
FILL_Y = PatternFill("solid", fgColor="FFF2CC")
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
TOTB = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="double", color="000000"),
)
NUM = "#,##0"
DATE_FMT = "dd/mm/yyyy"


def lp(p: Path | str) -> str:
    s = str(p)
    return s if s.startswith("\\\\?\\") else "\\\\?\\" + s


def nnum(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("\xa0", "")
    if not s or s in {"-", "None"}:
        return 0.0
    if re.match(r"^-?\d{1,3}(\.\d{3})+(,\d+)?$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def ndate(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def chuan_hoa_mst(val) -> str:
    if val is None:
        return ""
    s = str(val).strip().replace(" ", "")
    s = re.sub(r"[^0-9\-]", "", s)
    if not s or s.lower() == "none":
        return ""
    if len(s) == 10 and s.isdigit():
        return s
    if len(s) == 9 and s.isdigit():
        return "0" + s
    if len(s) == 12 and s.isdigit():
        return "0" + s[:9] + "-" + s[9:]
    if len(s) == 13 and s.isdigit():
        return s[:10] + "-" + s[10:]
    if len(s) == 14 and "-" in s:
        return s
    return s


def mst_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def valid_mst(s: str) -> bool:
    d = mst_digits(s)
    if len(d) not in (10, 13):
        return False
    if d.startswith("000"):
        return False
    return d[0] in "01356789"


def sanitize(val):
    if val is None:
        return None
    if isinstance(val, str):
        s = val.strip()
        return "'" + s if s.startswith("=") else s
    return val


def is_cancel(status: str) -> bool:
    t = (status or "").casefold()
    keys = ["bị xóa", "bi xoa", "bị hủy", "bi huy", "bị thay thế", "bi thay the", "hủy bỏ", "huy bo"]
    return any(k in t for k in keys)


def is_total_text(row) -> bool:
    texts = " ".join(str(c or "") for c in row).lower()
    if "tổng cộng" in texts or "tong cong" in texts:
        if "tổng công ty" in texts or "tong cong ty" in texts:
            return False
        return True
    stt = str(row[0] if row else "").strip().upper()
    return stt in {"TONG", "TỔNG", "TỔNG CỘNG", "TOTAL"}


def load_danhba(path: Path) -> dict:
    if not path.exists():
        print("danhba missing", path)
        return {}
    tree = ET.parse(lp(path))
    root = tree.getroot()

    def ss_index(cell, default):
        for key, val in cell.attrib.items():
            if key == "Index" or key.endswith("}Index"):
                return int(val)
        return default

    def ss_text(cell):
        for child in list(cell):
            if child.tag.endswith("}Data") or child.tag == "Data":
                return (child.text or "").strip()
        return ""

    out = {}
    for row in root.iter():
        if not (row.tag.endswith("}Row") or row.tag == "Row"):
            continue
        cells = {}
        idx = 1
        for cell in list(row):
            if not (cell.tag.endswith("}Cell") or cell.tag == "Cell"):
                continue
            idx = ss_index(cell, idx)
            cells[idx] = ss_text(cell)
            idx += 1
        mst = cells.get(4, "").strip()
        if not mst or mst == "Mã số thuế":
            continue
        rec = {
            "tt": cells.get(45, "").strip() or cells.get(44, "").strip() or cells.get(53, "").strip(),
            "cqt": cells.get(2, "").strip(),
            "ten": cells.get(5, "").strip(),
        }
        clean = mst.replace("-", "")
        out[mst] = rec
        out[clean] = rec
        if len(clean) >= 10:
            out[clean[:10]] = rec
    print("danhba", len(out))
    return out


def danhba_info(mp: dict, mst: str):
    if not mst:
        return "", ""
    clean = str(mst).replace("-", "").strip()
    hit = mp.get(str(mst).strip()) or mp.get(clean)
    if not hit and len(clean) >= 10:
        hit = mp.get(clean[:10])
    if hit:
        return hit.get("tt") or "", hit.get("cqt") or ""
    return "", ""


def load_f0(path: Path) -> dict:
    out = {}
    if not path.exists():
        print("cvcbrr missing")
        return out
    wb = load_workbook(lp(path), data_only=True, read_only=True)
    if "F0" not in wb.sheetnames:
        wb.close()
        return out
    for r in wb["F0"].iter_rows(min_row=8, max_col=12, values_only=True):
        m = chuan_hoa_mst(r[1] if r and len(r) > 1 else "")
        if not valid_mst(m):
            continue
        key = mst_digits(m)[:10]
        if key in out:
            continue
        ngay = r[4] if len(r) > 4 else None
        if isinstance(ngay, datetime):
            ngay = ngay.date()
        out[key] = {
            "mst": m,
            "ten": str(r[2] or "").strip() if len(r) > 2 else "",
            "cv": str(r[3] or "").strip() if len(r) > 3 else "",
            "ngay": ngay.strftime("%d/%m/%Y") if isinstance(ngay, date) else str(ngay or "")[:10],
            "co_quan": str(r[5] or "").strip() if len(r) > 5 else "",
        }
    wb.close()
    print("F0", len(out))
    return out


def load_f1_nnt(path: Path) -> list:
    hits = []
    if not path.exists():
        return hits
    wb = load_workbook(lp(path), data_only=True, read_only=True)
    if "F1" not in wb.sheetnames:
        wb.close()
        return hits
    for r in wb["F1"].iter_rows(min_row=8, max_col=12, values_only=True):
        m_mua = mst_digits(chuan_hoa_mst(r[8] if r and len(r) > 8 else ""))
        if m_mua == MST:
            hits.append(
                {
                    "cv": str(r[3] or "").strip() if len(r) > 3 else "",
                    "ngay": str(r[4] or "")[:10] if len(r) > 4 else "",
                    "mst_ban": chuan_hoa_mst(r[1] if len(r) > 1 else ""),
                    "ten_ban": str(r[2] or "").strip() if len(r) > 2 else "",
                }
            )
    wb.close()
    print("F1 NNT hits", len(hits))
    return hits


def load_ca(path: Path) -> set:
    out = set()
    if not path.exists():
        print("ca missing")
        return out
    wb = load_workbook(lp(path), data_only=True, read_only=True)
    if "NHAP_HO_SO" not in wb.sheetnames:
        wb.close()
        return out
    ws = wb["NHAP_HO_SO"]
    hdr = None
    for i, row in enumerate(ws.iter_rows(max_row=8, max_col=20, values_only=True), 1):
        vals = [str(c or "").lower() for c in row]
        if any("mst" in v for v in vals):
            hdr = (i, vals)
            break
    start = 5
    mst_i = 9
    if hdr:
        cand = [i for i, v in enumerate(hdr[1]) if "mst" in v and "mẫu" not in v]
        if cand:
            mst_i = cand[-1]
        start = hdr[0] + 1
    for row in ws.iter_rows(min_row=start, max_col=20, values_only=True):
        raw = row[mst_i] if mst_i is not None and len(row) > mst_i else None
        m = chuan_hoa_mst(raw)
        if valid_mst(m):
            out.add(mst_digits(m)[:10])
    if not out:
        for row in ws.iter_rows(min_row=5, max_col=20, values_only=True):
            for val in row:
                m = chuan_hoa_mst(val)
                if valid_mst(m) and (str(m).startswith("0") or str(m).startswith("3") or mst_digits(m).startswith("0") or mst_digits(m).startswith("3")):
                    out.add(mst_digits(m)[:10])
    wb.close()
    print("CA mst", len(out), "NNT in CA", MST in out)
    return out


def f0_label(f0: dict, mst: str) -> str:
    key = mst_digits(mst)[:10]
    rec = f0.get(key)
    if not rec:
        return ""
    parts = ["F0"]
    if rec["cv"]:
        parts.append(f"CV {rec['cv']}")
    if rec["ngay"]:
        parts.append(f"ngày {rec['ngay']}")
    if rec["co_quan"]:
        parts.append(f"({rec['co_quan']})")
    return " ".join(parts)


def load_hddt(path: Path, chieu: str):
    wb = load_workbook(lp(path), data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], []
    header = [str(c or "").strip() for c in rows[0]]
    data = []
    skip = 0
    st_i = next((i for i, h in enumerate(header) if "trạng thái" in h.lower()), 20)
    for row in rows[1:]:
        if not any(row):
            continue
        st = str(row[st_i] if st_i < len(row) else "" or "")
        if is_cancel(st):
            skip += 1
            continue
        data.append(list(row))
    print(f"HDDT {chieu} keep {len(data)} skip {skip}")
    return header, data


def hddt_mst_idx(header, chieu: str) -> int:
    hl = [h.lower() for h in header]
    if chieu == "MUA":
        for i, h in enumerate(hl):
            if "mst người bán" in h or "mst người xuất" in h:
                return i
        return 5
    for i, h in enumerate(hl):
        if "mst người mua" in h or "mst người nhận" in h:
            return i
    return 7


def hddt_ten_idx(header, mst_i: int) -> int:
    return mst_i + 1 if mst_i + 1 < len(header) else mst_i


def money_idx_from_header(header) -> list[int]:
    out = []
    for i, h in enumerate(header):
        t = h.lower()
        if any(k in t for k in ["chưa thuế", "tiền thuế", "thanh toán", "chiết khấu", "tiền phí"]):
            out.append(i)
    return out


def parse_bk_file(path: Path, year_hint: int | None):
    recs = []
    try:
        wb = load_workbook(lp(path), data_only=True)
    except Exception as e:
        print("bk fail", path.name, e)
        return recs
    for s in wb.sheetnames:
        sl = s.lower()
        if sl in {"vao", "mua vao", "mv"} or "vao" in sl:
            chieu = "MUA"
        elif sl in {"ra", "ban ra", "br"} or sl.endswith("ra") or "ban" in sl:
            chieu = "BAN"
        else:
            continue
        ws = wb[s]
        header_r = None
        for r in range(1, 16):
            a = str(ws.cell(r, 1).value or "").strip().upper().rstrip(".")
            joined = " ".join(str(ws.cell(r, c).value or "").lower() for c in range(1, 20))
            if a == "STT" and ("hóa đơn" in joined or "chứng từ" in joined or "người" in joined):
                header_r = r
                break
        if not header_r:
            for r in range(1, 16):
                a = str(ws.cell(r, 1).value or "").strip().upper().rstrip(".")
                joined = " ".join(str(ws.cell(r, c).value or "").lower() for c in range(1, 20))
                if a == "STT" and ("số hđ" in joined or "mã sô" in joined or "mã số" in joined):
                    header_r = r
                    break
        if not header_r:
            print("bk no header", path.name, s)
            continue
        h1 = [str(ws.cell(header_r, c).value or "") for c in range(1, 22)]
        h2 = [str(ws.cell(header_r + 1, c).value or "") for c in range(1, 22)]
        merged = [(a + " " + b).strip().lower() for a, b in zip(h1, h2)]

        def col(*keys, any_of=False):
            for i, t in enumerate(merged):
                if any_of:
                    if any(k in t for k in keys):
                        return i
                elif all(k in t for k in keys):
                    return i
            return None

        i_mst = col("mã sô thuế") or col("mã số thuế") or col("mst")
        i_ten = col("tên người")
        i_so = col("số hđ") or col("so hd") or col("số hóa đơn")
        i_ngay = col("ngày hđ") or col("ngay hd") or col("ngày hóa đơn")
        i_seri = col("seri")
        i_hang = col("mặt hàng")
        i_ds = col("chưa thuế")
        i_suat = col("suất")
        i_thue = col("gtgt")
        if i_mst is None or i_so is None:
            print("bk cols missing", path.name, s, merged[:16])
            continue
        start = header_r + 1
        for r in range(header_r, min(header_r + 4, (ws.max_row or 1) + 1)):
            a = ws.cell(r, 1).value
            if isinstance(a, (int, float)) and a >= 1:
                start = r
                break
            if str(a).strip().isdigit():
                start = r
                break
        for r in range(start, (ws.max_row or 1) + 1):
            row = [ws.cell(r, c).value for c in range(1, 22)]
            if not any(row) or is_total_text(row):
                continue
            so_hd = row[i_so]
            if so_hd is None or str(so_hd).strip() == "" or "(2)" in str(so_hd):
                gt_chk = nnum(row[i_ds]) if i_ds is not None else 0
                if gt_chk == 0:
                    continue
            ngay = ndate(row[i_ngay]) if i_ngay is not None else None
            mst_goc = row[i_mst]
            mst_c = chuan_hoa_mst(mst_goc)
            gt = nnum(row[i_ds]) if i_ds is not None else 0.0
            thue = nnum(row[i_thue]) if i_thue is not None else 0.0
            suat = nnum(row[i_suat]) if i_suat is not None else 0.0
            nam = ngay.year if ngay else year_hint
            if nam and nam < 2010:
                nam = year_hint
            thang = ngay.month if ngay else None
            recs.append(
                {
                    "Chieu": chieu,
                    "Nam": nam,
                    "Quy": ((thang - 1) // 3 + 1) if thang else None,
                    "Thang": thang,
                    "Ky_hieu_HD": str(row[i_seri] or "").strip() if i_seri is not None else "",
                    "So_HD": str(so_hd or "").strip(),
                    "Ngay_HD": ngay,
                    "MST_goc": str(mst_goc or "").strip(),
                    "MST_chuan": mst_c,
                    "Ten_DoiTac": str(row[i_ten] or "").strip() if i_ten is not None else "",
                    "Mat_hang": str(row[i_hang] or "").strip() if i_hang is not None else "",
                    "GT_chua_thue": gt,
                    "Thue_suat": suat if suat <= 1 else suat / 100.0,
                    "Thue_GTGT": thue,
                    "Tong_tien": gt + thue,
                    "File_nguon": path.name,
                }
            )
    wb.close()
    return recs


def style_header(ws, row, n):
    for c in range(1, n + 1):
        cell = ws.cell(row, c)
        cell.font = FONT_H
        cell.fill = FILL_H
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN
    ws.row_dimensions[row].height = 32


def autowidth(ws, cap=48, min_w=10):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row or 1, 400), max_col=ws.max_column or 1):
        for cell in row:
            if cell.row <= 2 or cell.value is None:
                continue
            s = str(cell.value)
            if s.startswith("="):
                continue
            widths[cell.column] = max(widths.get(cell.column, min_w), min(cap, len(s) + 2))
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A4"


def write_total(ws, r, money_cols, n_cols, start=4):
    ws.cell(r, 1, "TỔNG CỘNG").font = FONT_B
    ws.cell(r, 1).alignment = Alignment(horizontal="center")
    for c in range(1, n_cols + 1):
        cell = ws.cell(r, c)
        cell.fill = FILL_T
        cell.border = TOTB
        if c in money_cols and r > start:
            col = get_column_letter(c)
            cell.value = f"=SUM({col}{start}:{col}{r-1})"
            cell.font = FONT_B
            cell.number_format = NUM
            cell.alignment = Alignment(horizontal="right")


def emit_hddt(ws, title, header, data, mst_i, danhba, f0, chieu):
    extra = ["MST_chuan", "Tình trạng hoạt động", "Cơ quan thuế quản lý", "Cảnh báo rủi ro"]
    new_h = list(header[: mst_i + 1]) + ["MST_chuan"] + list(header[mst_i + 1 :]) + extra[1:]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(new_h))
    c0 = ws.cell(1, 1, title)
    c0.font = FONT_T
    ws.cell(2, 1, f"Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Skill tong-hop-hddt-bang-ke-4-sheet | Giai đoạn 1-3").font = FONT_S
    for i, t in enumerate(new_h, 1):
        ws.cell(3, i, t)
    style_header(ws, 3, len(new_h))
    orig_money = money_idx_from_header(header)
    out_money = []
    for oi in orig_money:
        out_money.append(oi + 1 if oi <= mst_i else oi + 2)
    r = 4
    stt = 1
    ten_i = hddt_ten_idx(header, mst_i)
    gt_i = next((i for i, h in enumerate(header) if "chưa thuế" in h.lower()), 13)
    thue_i = next((i for i, h in enumerate(header) if "tiền thuế" in h.lower() or h.lower() == "tổng tiền thuế"), 14)
    ncc = {}
    for row in data:
        row = list(row) + [None] * max(0, len(header) - len(row))
        mst_c = chuan_hoa_mst(row[mst_i] if mst_i < len(row) else "")
        tt, cqt = danhba_info(danhba, mst_c)
        risk = f0_label(f0, mst_c) if chieu == "MUA" else ""
        new = list(row[: mst_i + 1]) + [mst_c] + list(row[mst_i + 1 : len(header)]) + [tt, cqt, risk]
        new[0] = stt
        stt += 1
        for c, v in enumerate(new, 1):
            cell = ws.cell(r, c, sanitize(v))
            cell.font = FONT_B if risk and c == len(new) else FONT
            cell.border = THIN
            if c in out_money:
                if isinstance(v, (int, float)):
                    cell.number_format = NUM
                cell.alignment = Alignment(horizontal="right")
            elif isinstance(v, date):
                cell.number_format = DATE_FMT
            if risk:
                ws.cell(r, len(new)).fill = FILL_R
        if chieu == "MUA" and valid_mst(mst_c):
            rec = ncc.setdefault(
                mst_digits(mst_c)[:10],
                {"mst": mst_c, "ten": str(row[ten_i] or ""), "tt": tt, "cqt": cqt, "risk": risk,
                 "hddt_n": 0, "hddt_gt": 0.0, "hddt_thue": 0.0, "bk_n": 0, "bk_gt": 0.0, "bk_thue": 0.0},
            )
            rec["hddt_n"] += 1
            rec["hddt_gt"] += nnum(row[gt_i] if gt_i < len(row) else 0)
            rec["hddt_thue"] += nnum(row[thue_i] if thue_i < len(row) else 0)
            rec["ten"] = rec["ten"] or str(row[ten_i] or "")
            rec["tt"] = rec["tt"] or tt
            rec["cqt"] = rec["cqt"] or cqt
            rec["risk"] = rec["risk"] or risk
        r += 1
    write_total(ws, r, out_money, len(new_h))
    autowidth(ws)
    return stt - 1, ncc


def emit_bk(ws, title, rows, danhba, f0, chieu, ncc):
    hdr = [
        "STT", "Nam", "Quy", "Thang", "Ky_hieu_HD", "So_HD", "Ngay_HD",
        "MST_goc", "MST_chuan", "Ten_DoiTac", "Mat_hang",
        "GT_chua_thue", "Thue_suat", "Thue_GTGT", "Tong_tien",
        "Tình trạng hoạt động", "Cơ quan thuế quản lý", "Cảnh báo rủi ro", "File_nguon",
    ]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(hdr))
    ws.cell(1, 1, title).font = FONT_T
    ws.cell(2, 1, f"Ngày chạy: {TODAY.strftime('%d/%m/%Y')} | Lọc dòng Tổng cộng | MST chuẩn BT6").font = FONT_S
    for i, t in enumerate(hdr, 1):
        ws.cell(3, i, t)
    style_header(ws, 3, len(hdr))
    money = {12, 14, 15}
    r = 4
    stt = 1
    for rec in rows:
        mst_c = rec["MST_chuan"]
        tt, cqt = danhba_info(danhba, mst_c)
        risk = f0_label(f0, mst_c) if chieu == "MUA" else ""
        vals = [
            stt, rec["Nam"], rec["Quy"], rec["Thang"], rec["Ky_hieu_HD"], rec["So_HD"], rec["Ngay_HD"],
            rec["MST_goc"], mst_c, rec["Ten_DoiTac"], rec["Mat_hang"],
            rec["GT_chua_thue"], rec["Thue_suat"], rec["Thue_GTGT"], rec["Tong_tien"],
            tt, cqt, risk, rec["File_nguon"],
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, sanitize(v))
            cell.font = FONT
            cell.border = THIN
            if c in money:
                cell.number_format = NUM
                cell.alignment = Alignment(horizontal="right")
            elif c == 13:
                cell.number_format = "0.00%"
            elif isinstance(v, date):
                cell.number_format = DATE_FMT
            if risk and c == 18:
                cell.fill = FILL_R
                cell.font = FONT_B
        if chieu == "MUA" and valid_mst(mst_c):
            key = mst_digits(mst_c)[:10]
            item = ncc.setdefault(
                key,
                {"mst": mst_c, "ten": rec["Ten_DoiTac"], "tt": tt, "cqt": cqt, "risk": risk,
                 "hddt_n": 0, "hddt_gt": 0.0, "hddt_thue": 0.0, "bk_n": 0, "bk_gt": 0.0, "bk_thue": 0.0},
            )
            item["bk_n"] += 1
            item["bk_gt"] += rec["GT_chua_thue"]
            item["bk_thue"] += rec["Thue_GTGT"]
            item["ten"] = item["ten"] or rec["Ten_DoiTac"]
            item["tt"] = item["tt"] or tt
            item["cqt"] = item["cqt"] or cqt
            item["risk"] = item["risk"] or risk
        stt += 1
        r += 1
    write_total(ws, r, money, len(hdr))
    autowidth(ws)
    return stt - 1


def src_label(s: set) -> str:
    labels = []
    if "HDDT_Mua" in s and "Bang_ke_Mua" in s:
        labels.append("Chung_Mua")
    elif "HDDT_Mua" in s:
        labels.append("HDDT_Mua")
    elif "Bang_ke_Mua" in s:
        labels.append("Bang_ke_Mua")
    if "HDDT_Ban" in s and "Bang_ke_Ban" in s:
        labels.append("Chung_Ban")
    elif "HDDT_Ban" in s:
        labels.append("HDDT_Ban")
    elif "Bang_ke_Ban" in s:
        labels.append("Bang_ke_Ban")
    return "; ".join(labels)


def emit_doi_tac(ws, partners):
    hdr = ["STT", "MST_chuan", "Ten_doi_tac", "Ghi_chu_nguon"]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    ws.cell(1, 1, f"TỔNG HỢP ĐỐI TÁC - {TEN} (MST {MST}) - {KY}").font = FONT_T
    ws.cell(2, 1, "GĐ2: Chung_Mua/Chung_Ban = có cả HĐĐT và Bảng kê; còn lại là nguồn đơn").font = FONT_S
    for i, t in enumerate(hdr, 1):
        ws.cell(3, i, t)
    style_header(ws, 3, 4)
    r = 4
    for i, rec in enumerate(sorted(partners.values(), key=lambda x: x["mst"]), 1):
        vals = [i, rec["mst"], rec["ten"], src_label(rec["src"])]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = FONT
            cell.border = THIN
        r += 1
    autowidth(ws)
    return r - 4


def emit_rui_ro(ws, ncc, f1_hits, nnt_in_ca):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
    ws.cell(1, 1, "TỔNG HỢP TRẠNG THÁI HOẠT ĐỘNG VÀ CẢNH BÁO RỦI RO NCC").font = FONT_T
    nnt_note = f"NNT {MST} trên F1: {'Có (' + str(len(f1_hits)) + ' dòng)' if f1_hits else 'Không'}. NNT trên sổ CA: {'Có' if nnt_in_ca else 'Không'}."
    ws.cell(2, 1, f"{TEN} - MST {MST} - Kỳ {KY} | {nnt_note} | NCC chỉ gắn F0 CVCBRR.").font = FONT_S

    status = defaultdict(lambda: {"n": 0, "hd_n": 0, "hd_gt": 0.0, "hd_th": 0.0, "bk_n": 0, "bk_gt": 0.0, "bk_th": 0.0})
    for rec in ncc.values():
        st = rec["tt"] or "Không rõ / Chưa tra cứu"
        d = status[st]
        d["n"] += 1
        d["hd_n"] += rec["hddt_n"]
        d["hd_gt"] += rec["hddt_gt"]
        d["hd_th"] += rec["hddt_thue"]
        d["bk_n"] += rec["bk_n"]
        d["bk_gt"] += rec["bk_gt"]
        d["bk_th"] += rec["bk_thue"]

    ws.cell(4, 1, "PHẦN I. TỔNG HỢP THEO TÌNH TRẠNG HOẠT ĐỘNG (Danh bạ NNT)").font = FONT_S
    hdr1 = [
        "STT", "Tình trạng hoạt động (ĐKKD)", "Số lượng NCC",
        "Số HĐ (HĐĐT)", "Doanh số chưa thuế (HĐĐT)", "Thuế GTGT (HĐĐT)",
        "Số HĐ (Bảng kê)", "Doanh số chưa thuế (Bảng kê)", "Thuế GTGT (Bảng kê)",
        "Chênh lệch DS (HĐĐT - BK)", "Chênh lệch Thuế (HĐĐT - BK)",
    ]
    for i, t in enumerate(hdr1, 1):
        ws.cell(5, i, t)
    style_header(ws, 5, len(hdr1))
    r = 6
    for i, (st, d) in enumerate(sorted(status.items(), key=lambda x: -x[1]["n"]), 1):
        vals = [i, st, d["n"], d["hd_n"], d["hd_gt"], d["hd_th"], d["bk_n"], d["bk_gt"], d["bk_th"],
                d["hd_gt"] - d["bk_gt"], d["hd_th"] - d["bk_th"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = FONT
            cell.border = THIN
            if c in {5, 6, 8, 9, 10, 11}:
                cell.number_format = NUM
        r += 1
    write_total(ws, r, {3, 4, 5, 6, 7, 8, 9, 10, 11}, len(hdr1), start=6)
    r += 2

    f0_n = [x for x in ncc.values() if x["risk"]]
    binh = [x for x in ncc.values() if not x["risk"]]
    ws.cell(r, 1, "PHẦN II. TỔNG HỢP THEO CẢNH BÁO RỦI RO (F0 CVCBRR - NCC mua vào)").font = FONT_S
    r += 1
    hdr2 = [
        "STT", "Phân loại rủi ro", "Số lượng NCC",
        "Số HĐ (HĐĐT)", "Doanh số chưa thuế (HĐĐT)", "Thuế GTGT (HĐĐT)",
        "Số HĐ (Bảng kê)", "Doanh số chưa thuế (Bảng kê)", "Thuế GTGT (Bảng kê)",
        "Đánh giá",
    ]
    hdr2_row = r
    for i, t in enumerate(hdr2, 1):
        ws.cell(r, i, t)
    style_header(ws, r, len(hdr2))
    r += 1
    t2_start = r
    rows2 = [
        (1, "NCC dính CVCBRR F0 (Bán hóa đơn)", f0_n, "Rủi ro cao (HD01), xác minh hàng hóa và thanh toán", FILL_R),
        (2, "NCC bình thường (Không có cảnh báo F0)", binh, "Đối chiếu theo kế hoạch thông thường", None),
    ]
    for stt, name, items, note, fill in rows2:
        hd_n = sum(x["hddt_n"] for x in items)
        hd_gt = sum(x["hddt_gt"] for x in items)
        hd_th = sum(x["hddt_thue"] for x in items)
        bk_n = sum(x["bk_n"] for x in items)
        bk_gt = sum(x["bk_gt"] for x in items)
        bk_th = sum(x["bk_thue"] for x in items)
        vals = [stt, name, len(items), hd_n, hd_gt, hd_th, bk_n, bk_gt, bk_th, note]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = FONT_B if fill else FONT
            cell.border = THIN
            if fill:
                cell.fill = fill
            if c in {5, 6, 8, 9}:
                cell.number_format = NUM
        r += 1
    write_total(ws, r, {3, 4, 5, 6, 7, 8, 9}, len(hdr2), start=t2_start)
    r += 2

    ws.cell(r, 1, "PHẦN III. DANH SÁCH CHI TIẾT NCC CÓ CẢNH BÁO F0 HOẶC ĐKKD BẤT THƯỜNG").font = FONT_S
    r += 1
    hdr3 = [
        "STT", "MST_chuan", "Tên người bán", "Tình trạng hoạt động (ĐKKD)", "Cơ quan thuế quản lý",
        "Cảnh báo rủi ro", "Nguồn dữ liệu",
        "Số HĐ (HĐĐT)", "Doanh số (HĐĐT)", "Thuế (HĐĐT)",
        "Số HĐ (Bảng kê)", "Doanh số (Bảng kê)", "Thuế (Bảng kê)",
        "Đề xuất kiểm tra",
    ]
    for i, t in enumerate(hdr3, 1):
        ws.cell(r, i, t)
    style_header(ws, r, len(hdr3))
    r += 1
    stt = 1
    for rec in sorted(ncc.values(), key=lambda x: (0 if x["risk"] else 1, -x["hddt_gt"] - x["bk_gt"])):
        tt = rec["tt"] or ""
        abnormal = bool(tt) and ("đang hoạt động" not in tt.lower())
        if not rec["risk"] and not abnormal:
            continue
        src = "Chung" if rec["hddt_n"] and rec["bk_n"] else ("HĐĐT" if rec["hddt_n"] else "Bảng kê")
        dexuat = (
            "Rà soát tính có thật của hàng hóa, hồ sơ thanh toán, loại trừ chi phí nếu HĐ bất hợp pháp"
            if rec["risk"]
            else "Kiểm tra thời điểm lập hóa đơn so với ngày cơ quan thuế thông báo ngừng hoạt động"
        )
        vals = [
            stt, rec["mst"], rec["ten"], tt, rec["cqt"], rec["risk"] or "ĐKKD bất thường", src,
            rec["hddt_n"], rec["hddt_gt"], rec["hddt_thue"], rec["bk_n"], rec["bk_gt"], rec["bk_thue"], dexuat,
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = FONT_B if rec["risk"] else FONT
            cell.border = THIN
            if rec["risk"]:
                cell.fill = FILL_R
            elif abnormal:
                cell.fill = FILL_Y
            if c in {8, 9, 10, 11, 12, 13}:
                cell.number_format = NUM
        stt += 1
        r += 1
    if stt == 1:
        ws.cell(r, 1, "Không có NCC F0 hoặc ĐKKD bất thường trong tập mua vào đã nạp.").font = FONT
        r += 1
    autowidth(ws, cap=55)
    return len(f0_n)


def main():
    print("== LOAD ==")
    danhba = load_danhba(DANHBA)
    f0 = load_f0(CVCBRR)
    f1_hits = load_f1_nnt(CVCBRR)
    ca = load_ca(CA)
    nnt_in_ca = MST in ca or MST in {mst_digits(x)[:10] for x in ca}

    print("== HDDT ==")
    h_mua, d_mua = load_hddt(HDDT_MUA, "MUA")
    h_ban, d_ban = load_hddt(HDDT_BAN, "BAN")
    mst_mua = hddt_mst_idx(h_mua, "MUA")
    mst_ban = hddt_mst_idx(h_ban, "BAN")
    ten_mua = hddt_ten_idx(h_mua, mst_mua)
    ten_ban = hddt_ten_idx(h_ban, mst_ban)

    print("== BANG KE ==")
    bk = []
    for p in sorted(BK_DIR.glob("BANG KE HOA DON *.xlsx")):
        if p.name.startswith("~$"):
            continue
        m = re.search(r"(20\d{2})", p.name)
        recs = parse_bk_file(p, int(m.group(1)) if m else None)
        print(p.name, len(recs))
        bk.extend(recs)
    bk_mua = [r for r in bk if r["Chieu"] == "MUA"]
    bk_ban = [r for r in bk if r["Chieu"] == "BAN"]
    bk_mua.sort(key=lambda x: (x["Ngay_HD"] or date.min, str(x["So_HD"])))
    bk_ban.sort(key=lambda x: (x["Ngay_HD"] or date.min, str(x["So_HD"])))
    print("BK mua", len(bk_mua), "ban", len(bk_ban))

    partners = {}

    def add_p(mst, ten, src):
        m = chuan_hoa_mst(mst)
        if not m:
            return
        rec = partners.setdefault(m, {"mst": m, "ten": ten or "", "src": set()})
        if ten and len(str(ten)) > len(rec["ten"]):
            rec["ten"] = str(ten)
        rec["src"].add(src)

    for row in d_mua:
        add_p(row[mst_mua] if mst_mua < len(row) else "", row[ten_mua] if ten_mua < len(row) else "", "HDDT_Mua")
    for row in d_ban:
        add_p(row[mst_ban] if mst_ban < len(row) else "", row[ten_ban] if ten_ban < len(row) else "", "HDDT_Ban")
    for rec in bk_mua:
        add_p(rec["MST_chuan"], rec["Ten_DoiTac"], "Bang_ke_Mua")
    for rec in bk_ban:
        add_p(rec["MST_chuan"], rec["Ten_DoiTac"], "Bang_ke_Ban")

    print("== WRITE ==")
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "HDDT_Mua_vao"
    ncc = {}
    n1, ncc = emit_hddt(
        ws1, f"TỔNG HỢP HĐĐT MUA VÀO - {TEN} (MST {MST}) - {KY}",
        h_mua, d_mua, mst_mua, danhba, f0, "MUA",
    )
    ws2 = wb.create_sheet("HDDT_Ban_ra")
    n2, _ = emit_hddt(
        ws2, f"TỔNG HỢP HĐĐT BÁN RA - {TEN} (MST {MST}) - {KY}",
        h_ban, d_ban, mst_ban, danhba, f0, "BAN",
    )
    ws3 = wb.create_sheet("Bang_ke_Mua_vao")
    n3 = emit_bk(ws3, f"BẢNG KÊ MUA VÀO (NNT NỘP) - {TEN} (MST {MST}) - {KY}", bk_mua, danhba, f0, "MUA", ncc)
    ws4 = wb.create_sheet("Bang_ke_Ban_ra")
    n4 = emit_bk(ws4, f"BẢNG KÊ BÁN RA (NNT NỘP) - {TEN} (MST {MST}) - {KY}", bk_ban, danhba, f0, "BAN", ncc)
    ws5 = wb.create_sheet("Tong_hop_doi_tac")
    n5 = emit_doi_tac(ws5, partners)
    ws6 = wb.create_sheet("Tong_hop_trang_thai_rui_ro")
    n_f0 = emit_rui_ro(ws6, ncc, f1_hits, nnt_in_ca)

    OUT_G.parent.mkdir(parents=True, exist_ok=True)
    OUT_VAULT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(lp(OUT_G))
    wb.save(lp(OUT_VAULT))
    print("saved", OUT_G, OUT_G.stat().st_size)
    print("saved", OUT_VAULT, OUT_VAULT.stat().st_size)
    print(
        "DONE",
        {
            "HDDT_Mua": n1,
            "HDDT_Ban": n2,
            "BK_Mua": n3,
            "BK_Ban": n4,
            "doi_tac": n5,
            "ncc_mua": len(ncc),
            "ncc_f0": n_f0,
            "f1_nnt": len(f1_hits),
        },
    )


if __name__ == "__main__":
    main()
