# -*- coding: utf-8 -*-
"""Boc phu luc 05-1/05-2 NLD tu XML QTT TNCN, dien IN05 Kien Thanh Tan. Ban v2."""
from __future__ import annotations

import os
import re
import shutil
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MST = "3603216935"
TEN = "CÔNG TY TNHH MỘT THÀNH VIÊN KIẾN THÀNH TÂN"
KY = "2015-2025"
YEARS = list(range(2015, 2026))
TODAY = date(2026, 11, 17)
VER = "v2"
VER_OLD = "v1"

VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\1. TD\14. KIEN THANH TAN_3603216935"
)
XML_DIR = ROOT / "1. DU LIEU THO NGANH THUE" / "1.1. To khai XML"
OUT = ROOT / "4. KET QUA PHAN TICH"
VAULT_OUT = VAULT / "ket-qua" / "KIEN_THANH_TAN_3603216935"

FONT = Font(name="Times New Roman", size=11)
FONT_B = Font(name="Times New Roman", size=11, bold=True)
FILL_GRN = PatternFill("solid", fgColor="E2EFDA")
FILL_YEL = PatternFill("solid", fgColor="FFF2CC")
FILL_RED = PatternFill("solid", fgColor="FCE4D6")
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
NUM = "#,##0"
PCT = "0.00%"
DATE_FMT = "dd/mm/yyyy"

PLUC_NGUON = {
    "PLuc_05_1_BK": "05-1",
    "PLuc_05_1_BK_QTT": "05-1",
    "PLuc_05_2_BK": "05-2",
    "PLuc_05_2_BK_QTT": "05-2",
}
PERSON_TAGS = {"BKeCTietCNhan", "ChiTietBangKe"}


def lp(p) -> str:
    s = os.path.abspath(str(p))
    if s.startswith("\\\\?\\"):
        return s
    if s.startswith("\\\\"):
        return "\\\\?\\UNC\\" + s[2:]
    return "\\\\?\\" + s


def ln(tag: str) -> str:
    return tag.split("}", 1)[-1] if tag else ""


def txt(el) -> str:
    if el is None or el.text is None:
        return ""
    return el.text.strip()


def find_first(parent, name: str):
    for ch in parent.iter():
        if ln(ch.tag) == name:
            return ch
    return None


def nnum(v):
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    sl = s.lower()
    if sl in {"false", "true", "x"}:
        return None
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def nnum0(v) -> float:
    x = nnum(v)
    return 0.0 if x is None else x


def looks_money(v) -> bool:
    x = nnum(v)
    return x is not None and abs(x) >= 1000


def is_flag_true(v) -> bool:
    s = str(v or "").strip().lower()
    return s in {"1", "true", "x", "có", "co", "yes"}


def flatten_person(el) -> dict:
    out = {}
    for ch in el.iter():
        if list(ch):
            continue
        n = ln(ch.tag)
        t = (ch.text or "").strip()
        if not n:
            continue
        if n not in out or (t and not out[n]):
            out[n] = t
    return out


def parse_person(el, nguon: str, schema: str) -> dict | None:
    m = flatten_person(el)
    ho = (m.get("ct07") or m.get("hoTen") or m.get("hoVaTen") or "").strip()
    if not ho:
        return None
    mst_cn = (m.get("ct08") or m.get("mst") or "").strip()
    cccd = (m.get("ct09") or m.get("soCCCD") or m.get("soCMND") or m.get("soDDCN") or "").strip()
    if schema == "new":
        tnct = nnum(m.get("ct12"))
        tn_mien = nnum(m.get("ct13"))
        giam_tru = nnum0(m.get("ct17")) + nnum0(m.get("ct18")) + nnum0(m.get("ct19")) + nnum0(m.get("ct20"))
        tn_tinh = nnum(m.get("ct22"))
        thue_pn = nnum(m.get("ct23"))
        thue_kt = nnum(m.get("ct24")) if nnum(m.get("ct24")) is not None else nnum(m.get("ct26"))
        uq = is_flag_true(m.get("ct10"))
        hdld_flag = m.get("ct11")
    else:
        tnct = nnum(m.get("ct11"))
        tn_mien = nnum(m.get("ct12"))
        giam_tru = nnum0(m.get("ct14")) + nnum0(m.get("ct15")) + nnum0(m.get("ct16")) + nnum0(m.get("ct17"))
        # 2016-2020: GTBT nam o ct15, BH o ct17; 2015 long nested van flatten duoc
        if not looks_money(m.get("ct14")) and looks_money(m.get("ct15")):
            giam_tru = nnum0(m.get("ct15")) + nnum0(m.get("ct16")) + nnum0(m.get("ct17"))
        tn_tinh = nnum(m.get("ct18"))
        thue_pn = nnum(m.get("ct20"))
        thue_kt = nnum(m.get("ct21")) if nnum(m.get("ct21")) is not None else nnum(m.get("ct22"))
        uq = is_flag_true(m.get("ct10")) or is_flag_true(m.get("caNhanNNUQQT"))
        hdld_flag = None
    if nguon == "05-1":
        loai_tn = "Tiền lương, tiền công"
        co_hdld = "Có"
        co_trich = ""
        cu_tru = "Không" if is_flag_true(m.get("nguoiVNSongNN_NguoiNN")) else "Có"
    else:
        loai_tn = "Không HĐLĐ / vãng lai"
        co_hdld = "Không"
        # 05-2: thue 10% thuong nam o chi tieu thue da KT
        if nnum0(tnct) > 0 and nnum0(thue_kt) == 0:
            co_trich = "Không"
        elif nnum0(thue_kt) > 0:
            co_trich = "Có"
        else:
            co_trich = ""
        cu_tru = "Không" if is_flag_true(m.get("nguoiVNSongNN_NguoiNN")) else "Có"
    return {
        "Ho_ten": ho,
        "CCCD": cccd,
        "MST_CN": mst_cn,
        "Nguon_BK": nguon,
        "Loai_thu_nhap": loai_tn,
        "Tong_TNCT": tnct,
        "Thue_da_KT": thue_kt,
        "Co_trich_10": co_trich,
        "Co_cam_ket": "",
        "Co_HDLD": co_hdld,
        "Ca_nhan_cu_tru": cu_tru,
        "TNCT_thuoc_dien_KT": tnct,
        "Tong_giam_tru": giam_tru if giam_tru else None,
        "TN_tinh_thue": tn_tinh,
        "So_nguoi_UQ": 1 if uq else 0,
        "Thue_QT_thay": thue_pn,
        "Thue_con_phai_nop": None,
        "tn_mien": tn_mien,
        "hdld_flag": hdld_flag,
    }


def detect_schema(pluc_el) -> str:
    """new = mau 953 tu 2021 (TNCT o ct12); old = TT92/ma 42 (TNCT o ct11)."""
    for ch in pluc_el.iter():
        if ln(ch.tag) not in PERSON_TAGS:
            continue
        m = flatten_person(ch)
        if looks_money(m.get("ct12")) and not looks_money(m.get("ct11")):
            return "new"
        if looks_money(m.get("ct11")):
            return "old"
    return "old"


def parse_tncn_file(path: Path) -> dict | None:
    try:
        tree = ET.parse(lp(path))
        root = tree.getroot()
    except Exception:
        return None
    ma = txt(find_first(root, "maTKhai"))
    ten_tk = txt(find_first(root, "tenTKhai"))
    ky_el = find_first(root, "KyKKhaiThue")
    ky = txt(ky_el.find(".//{*}kyKKhai")) if ky_el is not None else txt(find_first(root, "kyKKhai"))
    if not ky:
        ky = txt(find_first(root, "kyKKhai"))
    nam = 0
    if ky and "/" in ky:
        try:
            nam = int(ky.split("/")[-1])
        except ValueError:
            nam = 0
    elif ky:
        m = re.search(r"(20\d{2})", ky)
        nam = int(m.group(1)) if m else 0
    if not nam:
        m = re.search(r"(20\d{2})", path.name)
        nam = int(m.group(1)) if m else 0
    loai_tk = txt(find_first(root, "loaiTKhai")) or "C"
    so_lan = int(nnum0(txt(find_first(root, "soLan"))))
    nnt = find_first(root, "NNT")
    mst = txt(nnt.find(".//{*}mst")) if nnt is not None else txt(find_first(root, "mst"))
    ct_main = {}
    cte = find_first(root, "CTieuTKhaiChinh")
    src = cte if cte is not None else root
    for ch in src.iter():
        n = ln(ch.tag)
        if n.startswith("ct") and list(ch) == [] and (ch.text or "").strip():
            code = n[2:]
            val = nnum(ch.text)
            if val is not None and code not in ct_main:
                ct_main[code] = val
    nld = []
    has_pluc = {"05-1": False, "05-2": False}
    for e in root.iter():
        n = ln(e.tag)
        if n not in PLUC_NGUON:
            continue
        nguon = PLUC_NGUON[n]
        has_pluc[nguon] = True
        schema = detect_schema(e)
        for ch in e.iter():
            if ch is e or ln(ch.tag) not in PERSON_TAGS:
                continue
            rec = parse_person(ch, nguon, schema)
            if rec:
                rec["schema"] = schema
                rec["pluc"] = n
                nld.append(rec)
    if ma in {"42", "395", "953"} or "TNCN" in (ten_tk or "").upper() or "QTT" in (ten_tk or "").upper():
        loai = "tncn"
    else:
        loai = "khac"
    mau = "05/QTT-TNCN"
    if ma == "42":
        mau = "05/KK-TNCN"
    elif ma == "395":
        mau = "05/QTT-TNCN (TT92)"
    elif ma == "953":
        mau = "05/QTT-TNCN"
    return {
        "file": path.name,
        "path": str(path),
        "ma": ma,
        "ten": ten_tk,
        "ky": ky,
        "nam": nam,
        "loai_tkhai": loai_tk,
        "so_lan": so_lan,
        "mst": mst,
        "chi_tieu": ct_main,
        "nld": nld,
        "has_pluc": has_pluc,
        "size": os.path.getsize(lp(path)),
        "mau": mau,
        "loai": loai,
    }


def pick_latest(rows: list[dict]) -> list[dict]:
    groups = defaultdict(list)
    for x in rows:
        groups[x.get("nam")].append(x)
    out = []
    for nam, items in groups.items():
        items = sorted(
            items,
            key=lambda r: (
                0 if str(r.get("loai_tkhai")).upper() in {"B", "2"} else 1,
                -int(r.get("so_lan") or 0),
                -int(r.get("size") or 0),
            ),
        )
        pick = dict(items[0])
        pick["so_ban"] = len(items)
        pick["cac_ban"] = [i["file"] for i in items]
        out.append(pick)
    return sorted(out, key=lambda r: r.get("nam") or 0)


def tnr(cell, bold=None):
    f = cell.font
    cell.font = Font(
        name="Times New Roman",
        size=f.size if f and f.size else 11,
        bold=f.bold if bold is None and f else (bool(bold) if bold is not None else False),
        italic=False,
        color=f.color if f else None,
    )


def apply_tnr(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            tnr(cell)


def autowidth(ws, cap=60, min_w=10, max_row=4000):
    widths = {}
    for row in ws.iter_rows(max_row=min(ws.max_row or 1, max_row), max_col=ws.max_column or 1):
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            s = str(cell.value)
            if s.startswith("="):
                continue
            widths[cell.column] = max(widths.get(cell.column, min_w), min(len(s), cap) + 2)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(w, cap))


def clear_from(ws, start_row):
    if ws.max_row and ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)


def put(ws, r, c, v, bold=False, num=False, dt=False, fill=None, wrap=False, pct=False):
    cell = ws.cell(r, c, v)
    cell.font = FONT_B if bold else FONT
    cell.border = THIN
    cell.alignment = Alignment(vertical="center", wrap_text=wrap)
    if fill:
        cell.fill = fill
    if num and isinstance(v, (int, float)) and not isinstance(v, bool):
        cell.number_format = NUM
        cell.alignment = Alignment(vertical="center", horizontal="right")
    if pct:
        cell.number_format = PCT
        cell.alignment = Alignment(vertical="center", horizontal="right")
    if dt and v not in (None, ""):
        cell.number_format = DATE_FMT
        cell.alignment = Alignment(vertical="center", horizontal="center")
    return cell


def write_nld_formulas(ws, r: int, last_row: int):
    ws.cell(r, 11).value = f'=IF(OR(H{r}="",H{r}=0),"",I{r}/H{r})'
    ws.cell(r, 11).number_format = PCT
    ws.cell(r, 11).font = FONT
    ws.cell(r, 11).border = THIN
    ws.cell(r, 13).value = f'=IF(F{r}="","",IF(AND(F{r}="05-2",J{r}="Không"),"Có","Không"))'
    ws.cell(r, 13).font = FONT
    ws.cell(r, 13).border = THIN
    ws.cell(r, 15).value = f'=IF(D{r}="","",IF(COUNTIF($D$3:$D${last_row},D{r})>1,"Có","Không"))'
    ws.cell(r, 15).font = FONT
    ws.cell(r, 15).border = THIN


def qtt_tnct(ct: dict, nld: list, ma: str) -> float | None:
    # Form 953: [23] tong TNCT; form cu: [28]/[27]. Khong dung [21] (so NLD tren mau cu).
    keys = ("23", "28") if ma == "953" else ("28", "27", "23")
    for k in keys:
        if ct.get(k) not in (None, 0, 0.0):
            return ct.get(k)
    s = sum(x.get("Tong_TNCT") or 0 for x in nld)
    return s if s else None


def qtt_thue(ct: dict, ma: str) -> float | None:
    keys = ("36", "40", "29") if ma == "953" else ("36", "35", "29", "27")
    for k in keys:
        v = ct.get(k)
        if v not in (None,):
            return v
    return None


def fill_in05(ws, qtt_list: list[dict]):
    clear_from(ws, 3)
    # tinh truoc so dong de COUNTIF dung pham vi
    n_rows = 0
    for rec in qtt_list:
        n_rows += 1
        n_rows += len(rec.get("nld") or [])
        if not rec.get("nld") and rec.get("nam") == 2023:
            n_rows += 0
    last = 2 + n_rows
    r = 3
    for rec in qtt_list:
        y = rec.get("nam")
        nld = rec.get("nld") or []
        ct = rec.get("chi_tieu") or {}
        so_ld = len(nld)
        has1 = rec.get("has_pluc", {}).get("05-1")
        has2 = rec.get("has_pluc", {}).get("05-2")
        if not nld and not has1 and not has2:
            ghi = "XML QTT không có phụ lục 05-1/05-2"
        elif not nld and (has1 or has2):
            ghi = "Phụ lục có node nhưng không có dòng NLĐ có họ tên"
        else:
            n1 = sum(1 for x in nld if x.get("Nguon_BK") == "05-1")
            n2 = sum(1 for x in nld if x.get("Nguon_BK") == "05-2")
            ghi = f"Bóc từ XML: {n1} dòng 05-1" + (f", {n2} dòng 05-2" if n2 else "")
        tnct = qtt_tnct(ct, nld, rec.get("ma") or "")
        thue = qtt_thue(ct, rec.get("ma") or "")
        vals = {
            1: "QTT",
            2: y,
            3: f"Quyết toán TNCN {y}",
            4: "",
            5: MST,
            6: "",
            7: "Quyết toán năm",
            8: tnct,
            9: thue,
            10: "",
            14: ghi,
            16: rec.get("mau"),
            17: so_ld if so_ld else (None if not nld else 0),
            23: None,
            28: f"{rec.get('file')}; bản {rec.get('so_ban')}" if rec.get("so_ban", 1) > 1 else "",
            29: rec.get("file"),
        }
        if so_ld:
            vals[17] = so_ld
        for c in range(1, 30):
            put(ws, r, c, vals.get(c), num=c in {8, 9, 17, 20, 21, 22, 24, 25, 26, 27})
        write_nld_formulas(ws, r, last)
        r += 1
        nld_sorted = sorted(nld, key=lambda x: (0 if x.get("Nguon_BK") == "05-1" else 1, x.get("Ho_ten") or ""))
        for p in nld_sorted:
            for c in range(1, 30):
                put(ws, r, c, None)
            put(ws, r, 1, "NLD")
            put(ws, r, 2, y)
            put(ws, r, 3, p.get("Ho_ten"))
            put(ws, r, 4, p.get("CCCD"))
            put(ws, r, 5, p.get("MST_CN"))
            put(ws, r, 6, p.get("Nguon_BK"))
            put(ws, r, 7, p.get("Loai_thu_nhap"))
            put(ws, r, 8, p.get("Tong_TNCT"), num=True)
            put(ws, r, 9, p.get("Thue_da_KT"), num=True)
            put(ws, r, 10, p.get("Co_trich_10"))
            put(ws, r, 12, p.get("Co_cam_ket"))
            put(ws, r, 16, rec.get("mau"))
            put(ws, r, 18, p.get("Co_HDLD"))
            put(ws, r, 19, p.get("Ca_nhan_cu_tru"))
            put(ws, r, 20, p.get("TNCT_thuoc_dien_KT"), num=True)
            put(ws, r, 21, p.get("Tong_giam_tru"), num=True)
            put(ws, r, 22, p.get("TN_tinh_thue"), num=True)
            put(ws, r, 23, p.get("So_nguoi_UQ"), num=True)
            put(ws, r, 24, p.get("Thue_QT_thay"), num=True)
            put(ws, r, 29, rec.get("file"))
            write_nld_formulas(ws, r, last)
            r += 1
    autowidth(ws, cap=55)
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["N"].width = 50
    ws.column_dimensions["AC"].width = 45
    return r - 3


def bump_header(ws):
    for row in ws.iter_rows(max_row=8, max_col=20):
        for cell in row:
            if isinstance(cell, MergedCell) or not isinstance(cell.value, str):
                continue
            v = cell.value
            nv = v.replace(f"Phiên bản KQ: {VER_OLD}", f"Phiên bản KQ: {VER}")
            nv = nv.replace(f"v{VER_OLD[1:]}" if False else VER_OLD, VER) if "Phiên bản" in v or "phiên bản" in v.lower() else nv
            if "v1" in v and ("phiên bản" in v.lower() or "Phien ban" in v or "KQ:" in v):
                nv = v.replace("v1", VER)
            if nv != v:
                cell.value = nv
            tnr(cell)


def update_kq13(path_src: Path, path_dst: Path, stats: dict):
    shutil.copy2(lp(path_src), lp(path_dst))
    wb = load_workbook(lp(path_dst))
    ws = wb["KQ1_Kiem_ke"]
    bump_header(ws)
    # hang IN05: dong 14 theo mau GĐ1
    nld_n = stats["n_nld"]
    years_nld = ",".join(str(y) for y in stats["years_nld"])
    years_no = ",".join(str(y) for y in stats["years_no_nld"])
    muc = "Đủ" if nld_n else "Mỏng"
    ws.cell(14, 6).value = "XML"
    ws.cell(14, 7).value = "2015-2025"
    ws.cell(14, 8).value = f"Năm không có dòng NLĐ: {years_no}" if years_no else "Thư mục 3.3 đa nguồn trống"
    ws.cell(14, 9).value = stats["n_qtt"] + nld_n
    ws.cell(14, 9).number_format = NUM
    ws.cell(14, 10).value = 0
    ws.cell(14, 11).value = "Có"
    ws.cell(14, 12).value = muc
    ws.cell(14, 12).fill = FILL_GRN if muc == "Đủ" else FILL_YEL
    ws.cell(14, 13).value = f"{stats['n_qtt']} QTT + {nld_n} dòng NLĐ; năm có DS: {years_nld}"
    ws.cell(14, 14).value = "Đã bóc 05-1/05-2 từ XML QTT"
    ws.cell(14, 15).value = "2023 XML không có phụ lục; 3.3 đa nguồn vẫn trống"
    ws.cell(14, 17).value = "CQT 1.1 XML QTT TNCN"
    for c in range(6, 18):
        tnr(ws.cell(14, c), bold=(c == 12))
    n_du = n_mong = n_thieu = n_chua = 0
    for rr in range(10, 23):
        m = str(ws.cell(rr, 12).value or "")
        if m == "Đủ":
            n_du += 1
        elif m == "Mỏng":
            n_mong += 1
        elif m == "Thiếu":
            n_thieu += 1
        else:
            n_chua += 1
    ws["C4"] = f"{n_du} Đủ | {n_mong} Mỏng | {n_thieu} Thiếu | {n_chua} Chưa nhận"
    tnr(ws["C4"], bold=True)

    ws = wb["KQ1_Kha_nang_phan_tich"]
    bump_header(ws)
    if ws["C4"].value:
        ws["C4"] = VER
        tnr(ws["C4"])

    ws = wb["KQ1_De_nghi_bo_sung"]
    bump_header(ws)
    if ws["C3"].value:
        ws["C3"] = VER
        tnr(ws["C3"])
    # sua dong de nghi 05-1/05-2: con lai 3.3 va nam 2023
    # tim hang co IN05
    for r in range(10, (ws.max_row or 10) + 1):
        inn = str(ws.cell(r, 3).value or "")
        nd = str(ws.cell(r, 2).value or "")
        if "IN05" in inn or "05-1" in nd:
            ws.cell(r, 2).value = (
                "XML QTT TNCN năm 2023 không có phụ lục 05-1/05-2 (xin bản có bảng kê NLĐ nếu DN kê khai); "
                "tra cứu TNCN đa nguồn thư mục 3.3"
            )
            ws.cell(r, 4).value = "1.1 / 3.3"
            ws.cell(r, 5).value = "Danh sách NLĐ năm 2023; file đa nguồn CCCD"
            ws.cell(r, 8).value = (
                f"Đã bóc {nld_n} dòng NLĐ 05-1 từ XML các năm {years_nld}. "
                f"Còn thiếu: {years_no or 'không'}; 3.3 trống."
            )
            for c in range(1, 9):
                tnr(ws.cell(r, c))
                ws.cell(r, c).alignment = Alignment(vertical="center", wrap_text=c in {2, 7, 8})
            break
    apply_tnr(wb["KQ1_Kiem_ke"])
    wb.save(lp(path_dst))
    wb.close()
    return n_du, n_mong, n_thieu, n_chua


def main():
    files = [f for f in os.listdir(lp(XML_DIR)) if f.upper().startswith("TNCN") and f.lower().endswith(".xml")]
    parsed = []
    for fn in files:
        rec = parse_tncn_file(XML_DIR / fn)
        if rec and rec.get("loai") == "tncn":
            parsed.append(rec)
            print(
                "FILE",
                fn,
                "nam",
                rec["nam"],
                "ma",
                rec["ma"],
                "loaiTK",
                rec["loai_tkhai"],
                "lan",
                rec["so_lan"],
                "nld",
                len(rec["nld"]),
                "pluc",
                rec["has_pluc"],
                "ct",
                {k: rec["chi_tieu"].get(k) for k in list(rec["chi_tieu"])[:8]},
            )
            for p in rec["nld"]:
                print("   ", p["Nguon_BK"], p["Ho_ten"], "CCCD", p["CCCD"], "MST", p["MST_CN"], "TNCT", p["Tong_TNCT"], "thue", p["Thue_da_KT"])
    qtt = pick_latest(parsed)
    print("PICK", [(x["nam"], x["file"], len(x["nld"])) for x in qtt])

    stats = {
        "n_qtt": len(qtt),
        "n_nld": sum(len(x.get("nld") or []) for x in qtt),
        "years_nld": [x["nam"] for x in qtt if x.get("nld")],
        "years_no_nld": [x["nam"] for x in qtt if not x.get("nld")],
        "n05_1": sum(1 for x in qtt for p in x.get("nld") or [] if p.get("Nguon_BK") == "05-1"),
        "n05_2": sum(1 for x in qtt for p in x.get("nld") or [] if p.get("Nguon_BK") == "05-2"),
    }
    print("STATS", stats)

    name11_old = f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_{VER_OLD}.xlsx"
    name11_new = f"KQ1.1_Kiem_ke_CQT_{MST}_{KY}_{VER}.xlsx"
    name13_old = f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER_OLD}.xlsx"
    name13_new = f"KQ1.3_Danh_gia_ho_so_{MST}_{KY}_{VER}.xlsx"

    src11 = VAULT_OUT / name11_old
    if not os.path.exists(lp(src11)):
        src11 = OUT / name11_old
    dst11_cty = OUT / name11_new
    dst11_vault = VAULT_OUT / name11_new
    shutil.copy2(lp(src11), lp(dst11_cty))
    wb = load_workbook(lp(dst11_cty))
    bump_header(wb["IN05_NLD_TNCN"])
    n_write = fill_in05(wb["IN05_NLD_TNCN"], qtt)
    # cap nhat placeholder phien ban tren cac sheet khac cua KQ1.1
    for sn in wb.sheetnames:
        bump_header(wb[sn])
    apply_tnr(wb["IN05_NLD_TNCN"])
    wb.save(lp(dst11_cty))
    wb.close()
    shutil.copy2(lp(dst11_cty), lp(dst11_vault))
    print("IN05 rows written", n_write, "file", dst11_cty.name, os.path.getsize(lp(dst11_cty)))

    src13 = VAULT_OUT / name13_old
    if not os.path.exists(lp(src13)):
        src13 = OUT / name13_old
    dst13_cty = OUT / name13_new
    dst13_vault = VAULT_OUT / name13_new
    counts = update_kq13(src13, dst13_cty, stats)
    shutil.copy2(lp(dst13_cty), lp(dst13_vault))
    print("KQ1.3", dst13_cty.name, "counts", counts, os.path.getsize(lp(dst13_cty)))
    print("VAULT", dst11_vault)
    print("VAULT13", dst13_vault)
    print("DONE")


if __name__ == "__main__":
    main()
