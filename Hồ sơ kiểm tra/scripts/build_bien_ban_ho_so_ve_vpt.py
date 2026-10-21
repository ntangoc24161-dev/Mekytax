# -*- coding: utf-8 -*-
"""
Full official rebuild script for Van Phat Thanh inspection dossier.
"""
from __future__ import annotations

import os
import sys
import copy
import re
import shutil
import zipfile
from datetime import datetime
import openpyxl
import docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

TEMPLATE_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\BIEN BAN\MAUTONG HOP.docx"
TARGET_DIR = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\9. VAN PHAT THANH\0. BIEN BAN"
XLSX_PATH = os.path.join(TARGET_DIR, "Bang tinh 2021-2025.xlsx")
BRAIN_TEMPLATE_DIR = r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates"

YEAR_SLOTS = [2021, 2022, 2023, 2025]

def fmt_num(n, dash_zero: bool = False) -> str:
    if n is None:
        return "-"
    try:
        n = float(n)
    except Exception:
        return str(n)
    if abs(n) < 0.5:
        return "-" if dash_zero else "0"
    neg = n < 0
    n = abs(round(n))
    s = f"{int(n):,}".replace(",", ".")
    return f"({s})" if neg else s

def pct_tax(rate) -> str:
    try:
        return f"{round(float(rate) * 100):g}%"
    except Exception:
        return "20%"

_DV = ["", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]

def _three_digits(n: int, full: bool = False) -> str:
    tram, chuc, donvi = n // 100, (n % 100) // 10, n % 10
    parts = []
    if tram:
        parts.append(f"{_DV[tram]} trăm")
    elif full:
        parts.append("không trăm")
    if chuc > 1:
        parts.append(f"{_DV[chuc]} mươi")
        if donvi == 1:
            parts.append("mốt")
        elif donvi == 5:
            parts.append("lăm")
        elif donvi:
            parts.append(_DV[donvi])
    elif chuc == 1:
        parts.append("mười")
        if donvi == 5:
            parts.append("lăm")
        elif donvi:
            parts.append(_DV[donvi])
    elif donvi:
        if tram or full:
            parts.append("lẻ")
        parts.append(_DV[donvi])
    return " ".join(parts)

def so_bang_chu(n) -> str:
    n = int(round(abs(float(n or 0))))
    if n == 0:
        return "Không đồng"
    parts = []
    started = False
    for value, name in [(10**9, "tỷ"), (10**6, "triệu"), (10**3, "nghìn"), (1, "")]:
        chunk = n // value
        n = n % value
        if chunk:
            parts.append(_three_digits(chunk, full=started).strip() + (f" {name}" if name else ""))
            started = True
    text = " ".join(p for p in parts if p).strip()
    text = text[:1].upper() + text[1:] if text else "Không"
    return f"{text} đồng"

def load_dtchep(path: str) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["DTCHEP"]
    years: dict[int, dict] = {}
    cur = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if isinstance(a, str) and a.strip().upper().startswith("NĂM"):
            m = re.search(r"(20\d{2})", a)
            cur = int(m.group(1)) if m else None
            if cur:
                years.setdefault(cur, {})
            continue
        if cur and b:
            years[cur][str(b).strip()] = {
                "bc": ws.cell(r, 3).value or 0,
                "kt": ws.cell(r, 4).value or 0,
                "cl": ws.cell(r, 5).value or 0,
            }
    ws_p = wb["TINH PHAT + CHAM NOP (C T)"]
    penalties = {}
    total_row = None
    for r in range(1, ws_p.max_row + 1):
        y = ws_p.cell(r, 1).value
        if y in (2021, 2022, 2023, 2024, 2025):
            penalties[int(y)] = {
                "thue_tndn": ws_p.cell(r, 2).value or 0,
                "thue_gtgt": ws_p.cell(r, 3).value or 0,
                "cham_nop": ws_p.cell(r, 13).value or 0,
                "phat_20": ws_p.cell(r, 14).value or 0,
                "tong_phat": ws_p.cell(r, 15).value or 0,
                "tong": ws_p.cell(r, 17).value or 0,
                "ngay_bb": ws_p.cell(r, 4).value,
            }
        if isinstance(y, str) and "Tổng cộng" in y:
            total_row = {
                "thue_tndn": ws_p.cell(r, 2).value or 0,
                "thue_gtgt": ws_p.cell(r, 3).value or 0,
                "cham_nop": ws_p.cell(r, 13).value or 0,
                "phat_20": ws_p.cell(r, 14).value or 0,
                "tong_phat": ws_p.cell(r, 15).value or 0,
                "tong": ws_p.cell(r, 17).value or 0,
            }
    wb.close()
    return {"years": years, "penalties": penalties, "total": total_row}

def build_replacements_dict(data: dict) -> dict:
    total = data["total"] or {}
    thue_tndn = round(float(total.get("thue_tndn") or 0))
    thue_gtgt = round(float(total.get("thue_gtgt") or 0))
    phat_20 = round(float(total.get("phat_20") or 0))
    cham = round(float(total.get("cham_nop") or 0))
    tong_phat = round(float(total.get("tong_phat") or (phat_20 + cham)))
    tong = round(float(total.get("tong") or (thue_tndn + tong_phat)))

    ngay_bb = "14/10/2026"

    reps = {
        "Tên_công_ty": "CÔNG TY TNHH MỘT THÀNH VIÊN BAO BÌ GIẤY VẠN PHÁT THÀNH",
        "MST": "3601409441",
        "địa_chỉ": "Số 9/3, đường Nguyễn văn Tỏ, khu phố 2, Phường Long Hưng",
        "ĐỊA_CHỈ": "Số 9/3, đường Nguyễn văn Tỏ, khu phố 2, Phường Long Hưng, Thành phố Biên Hòa, Tỉnh Đồng Nai",
        "NNKD": "Sản xuất giấy nhăn, bìa nhăn, bao bì từ giấy và bìa",
        "GĐ": "Tô Quốc Thi",
        "KT": "Trần Thị Thoan",
        "CCCD": "096075006910",
        "ngày_cấp": "28/09/2021",
        "Ngày_cấp2": "28/09/2021",
        "Nơi_cấp": "Cục Cảnh sát Quản lý hành chính về trật tự xã hội",
        "GGCNĐKKD": "3601409441",
        "CCCDKT": "036176008691",
        "Chi_nhánh": "Cửa hàng Bao Bì Giấy Vạn Phát Thành tại số 507, Quốc Lộ 51, KP 3, Phường An Hòa (hết hiệu lực)",
        "Tài_khoản_số": "1000148510001 tại Indovina Bank - CN Đồng Nai",
        "Thời_kỳ_kiểm_tra": "Từ năm 2021 đến năm 2025",
        "Thời_hạn_kiểm_tra_": "20",
        "TĐ": "Nguyễn Thị Ánh Ngọc",
        "CVTĐ": "Chuyên viên",
        "TV": "Nguyễn Thành Đồng",
        "CVTV": "Kiểm tra viên thuế",
        "Ma_CQT": "1140968",
        "CQT": "Thuế Cơ sở 3 Thành phố Đồng Nai",
        "QĐKT": "6592/QĐ-TCS3",
        "QĐGS": "6593/QĐ-TCS3",
        "ngày_QĐ": "02/10/2026",
        "NGÀY_QĐ": "02/10/2026",
        "Ngày_công_bố": "09/10/2026",
        "Từ_ngày": "09/10/2026",
        "Đến_ngày": "13/10/2026",
        "Phân_công_trưởng_đoàn": "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng.",
        "Phân_công_thành_viên": "- Chi phí tài chính.",
        "Ngày_ký_Biên_bản_kiểm_tra": ngay_bb,
        "Ngày_báo_cáo": ngay_bb,
        "Thuế_TNDN": fmt_num(thue_tndn),
        "Thuế_GTGT": fmt_num(thue_gtgt, True),
        "Số_tiền_phạt_20": fmt_num(phat_20),
        "Số_tiền_phạt_20_BC": so_bang_chu(phat_20),
        "Số_tiền_phạt_003_TNDN": fmt_num(cham),
        "Số_tiền_phạt_003_GTGT": "-",
        "Phạt_KK_sai_VPHC": "-",
        "Tổng_phạt_kk_sai": fmt_num(phat_20),
        "Tổng_phạt": fmt_num(tong_phat),
        "Tổng_cộng_truy_thu": fmt_num(thue_tndn),
        "Tổng_cộng_truy_thu_BC": so_bang_chu(thue_tndn),
        "Tổng_cộng_số_phạt__chậm_nộp": fmt_num(cham),
        "Tổng_cộng_phạt_và_truy_thu": fmt_num(tong),
        "Tổng_cộng_phạt_và_truy_thu_BC": so_bang_chu(tong),
        "Tổng": fmt_num(tong),
        "Giảm_lỗ": "-",
        "giảm_khấu_trừ": "-",
    }
    return reps

def set_run_font(run, size_pt=None, bold=None, italic=None):
    run.font.name = "Times New Roman"
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rFonts.set(qn(a), "Times New Roman")
    if size_pt is not None:
        run.font.size = Pt(size_pt)
        sz = rPr.find(qn("w:sz"))
        if sz is None:
            sz = OxmlElement("w:sz")
            rPr.append(sz)
        sz.set(qn("w:val"), str(int(size_pt * 2)))
        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            rPr.append(szCs)
        szCs.set(qn("w:val"), str(int(size_pt * 2)))
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic

def replace_text_in_paragraph(paragraph, replacements):
    text = paragraph.text
    if not any(k in text for k in replacements):
        return
    for run in paragraph.runs:
        for k, v in replacements.items():
            pattern = f"«{k}»"
            if pattern in run.text:
                run.text = run.text.replace(pattern, v)
    if "«" in paragraph.text and "»" in paragraph.text:
        full_text = paragraph.text
        for k, v in replacements.items():
            pattern = f"«{k}»"
            if pattern in full_text:
                full_text = full_text.replace(pattern, v)
        if paragraph.runs:
            paragraph.runs[0].text = full_text
            for r in paragraph.runs[1:]:
                r.text = ""

def replace_in_doc(doc, replacements):
    for p in doc.paragraphs:
        replace_text_in_paragraph(p, replacements)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_text_in_paragraph(p, replacements)

def set_cell_text(cell, text: str):
    text = "" if text is None else str(text)
    if cell.paragraphs:
        p0 = cell.paragraphs[0]
        if p0.runs:
            p0.runs[0].text = text
            for r in p0.runs[1:]:
                r.text = ""
        else:
            p0.text = text
        for p in cell.paragraphs[1:]:
            for r in p.runs:
                r.text = ""
            if not p.runs:
                p.text = ""
    else:
        cell.text = text

def fill_five_col_row(table, row_idx: int, values: list[str]):
    row = table.rows[row_idx]
    for i, val in enumerate(values):
        if i < len(row.cells):
            set_cell_text(row.cells[i], val)

def _cost_amount(year_data: dict, key: str) -> tuple[float, float, float]:
    row = year_data.get(key, {"bc": 0, "kt": 0, "cl": 0}) or {"bc": 0, "kt": 0, "cl": 0}
    bc = float(row.get("bc") or 0)
    kt = float(row.get("kt") or 0)
    cl = row.get("cl")
    if cl is None:
        cl = kt - bc
    else:
        cl = float(cl or 0)
    return bc, kt, cl

def qldn_amount(year_data: dict) -> tuple[float, float, float]:
    return _cost_amount(year_data, "Chi phí quản lý kinh doanh")

def cpbh_amount(year_data: dict) -> tuple[float, float, float]:
    return _cost_amount(year_data, "Chi phí bán hàng")

def cptc_amount(year_data: dict) -> tuple[float, float, float]:
    return _cost_amount(year_data, "Chi phí tài chính")

def ensure_min_rows(table, min_rows: int):
    while len(table.rows) < min_rows:
        table._tbl.append(copy.deepcopy(table.rows[-1]._tr))

def tax_access(year_data: dict, year: int) -> dict:
    tn = year_data.get("Thu nhập tính thuế") or year_data.get("Thu nhập chịu thuế") or {"bc": 0, "kt": 0, "cl": 0}
    rate = year_data.get("Thuế suất thuế TNDN", {"bc": 0.2, "kt": 0.2})
    if year == 2021 and "Thuế TNDN còn phải nộp" in year_data:
        thue = year_data["Thuế TNDN còn phải nộp"]
    elif "Thuế TNDN từ hoạt động SXKD" in year_data:
        thue = year_data["Thuế TNDN từ hoạt động SXKD"]
    else:
        thue = year_data.get("Thuế TNDN còn phải nộp", {"bc": 0, "kt": 0, "cl": 0})
    ln = year_data.get("Lợi nhuận kế toán trước thuế", {"bc": 0, "kt": 0, "cl": 0})
    cptc = year_data.get("Chi phí tài chính", {"bc": 0, "kt": 0, "cl": 0})
    return {"tn": tn, "rate": rate, "thue": thue, "ln": ln, "cptc": cptc}

def replace_paragraph_text(paragraph, new_text: str):
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.text = new_text

def fill_dtchep_tables(doc: Document, data: dict):
    years = data["years"]
    # TV: CPTC (T12-T15)
    for ti, slot in zip([12, 13, 14, 15], range(4)):
        y = YEAR_SLOTS[slot]
        bc, kt, cl = cptc_amount(years[y])
        fill_five_col_row(
            doc.tables[ti],
            1,
            ["1", "Chi phí tài chính", fmt_num(bc, True), fmt_num(kt, True), fmt_num(cl, True)],
        )
    # TD: QLDN + CPBH (T18-T21)
    for ti, slot in zip([18, 19, 20, 21], range(4)):
        y = YEAR_SLOTS[slot]
        table = doc.tables[ti]
        ensure_min_rows(table, 3)
        bc_ql, kt_ql, cl_ql = qldn_amount(years[y])
        bc_bh, kt_bh, cl_bh = cpbh_amount(years[y])
        fill_five_col_row(
            table,
            1,
            ["1", "Chi phí quản lý doanh nghiệp", fmt_num(bc_ql, True), fmt_num(kt_ql, True), fmt_num(cl_ql, True)],
        )
        fill_five_col_row(
            table,
            2,
            ["2", "Chi phí bán hàng", fmt_num(bc_bh, True), fmt_num(kt_bh, True), fmt_num(cl_bh, True)],
        )
    # BBKT: 3 chi phí + TN + thuế (T24-T27)
    for ti, slot in zip([24, 25, 26, 27], range(4)):
        y = YEAR_SLOTS[slot]
        yd = years[y]
        info = tax_access(yd, y)
        bc_ql, kt_ql, cl_ql = qldn_amount(yd)
        bc_bh, kt_bh, cl_bh = cpbh_amount(yd)
        bc_tc, kt_tc, cl_tc = cptc_amount(yd)
        table = doc.tables[ti]
        need = 8 if (len(table.rows) >= 7 or y == 2025) else 7
        ensure_min_rows(table, need)
        fill_five_col_row(
            table,
            1,
            ["1", "Chi phí quản lý doanh nghiệp", fmt_num(bc_ql, True), fmt_num(kt_ql, True), fmt_num(cl_ql, True)],
        )
        fill_five_col_row(
            table,
            2,
            ["2", "Chi phí bán hàng", fmt_num(bc_bh, True), fmt_num(kt_bh, True), fmt_num(cl_bh, True)],
        )
        fill_five_col_row(
            table,
            3,
            ["3", "Chi phí tài chính", fmt_num(bc_tc, True), fmt_num(kt_tc, True), fmt_num(cl_tc, True)],
        )
        if len(table.rows) >= 8:
            fill_five_col_row(
                table,
                4,
                [
                    "4",
                    "Lợi nhuận kế toán trước thuế",
                    fmt_num(info["ln"]["bc"]),
                    fmt_num(info["ln"]["kt"]),
                    fmt_num(info["ln"]["cl"], True),
                ],
            )
            fill_five_col_row(
                table,
                5,
                [
                    "5",
                    "Thu nhập tính thuế",
                    fmt_num(info["tn"]["bc"]),
                    fmt_num(info["tn"]["kt"]),
                    fmt_num(info["tn"]["cl"], True),
                ],
            )
            fill_five_col_row(
                table,
                6,
                ["6", "Thuế suất thuế TNDN", pct_tax(info["rate"]["bc"]), pct_tax(info["rate"]["kt"]), ""],
            )
            fill_five_col_row(
                table,
                7,
                [
                    "7",
                    "Thuế TNDN từ hoạt động SXKD",
                    fmt_num(info["thue"]["bc"]),
                    fmt_num(info["thue"]["kt"]),
                    fmt_num(info["thue"]["cl"], True),
                ],
            )
        else:
            fill_five_col_row(
                table,
                4,
                [
                    "4",
                    "Thu nhập tính thuế",
                    fmt_num(info["tn"]["bc"]),
                    fmt_num(info["tn"]["kt"]),
                    fmt_num(info["tn"]["cl"], True),
                ],
            )
            fill_five_col_row(
                table,
                5,
                ["5", "Thuế suất thuế TNDN", pct_tax(info["rate"]["bc"]), pct_tax(info["rate"]["kt"]), ""],
            )
            fill_five_col_row(
                table,
                6,
                [
                    "6",
                    "Thuế TNDN từ hoạt động SXKD",
                    fmt_num(info["thue"]["bc"]),
                    fmt_num(info["thue"]["kt"]),
                    fmt_num(info["thue"]["cl"], True),
                ],
            )
    # T28 (Bảng tổng hợp truy thu)
    if len(doc.tables) > 28:
        t28 = doc.tables[28]
        for offset, slot in enumerate(range(4)):
            y = YEAR_SLOTS[slot]
            info = tax_access(years[y], y)
            row_idx = 4 + offset
            if row_idx < len(t28.rows):
                set_cell_text(t28.rows[row_idx].cells[1], str(y))
                if len(t28.rows[row_idx].cells) > 4:
                    set_cell_text(
                        t28.rows[row_idx].cells[4],
                        fmt_num(info["thue"]["cl"], True),
                    )

def cptc_detail(year: int, info: dict) -> str:
    giam = abs(round(float(info["cptc"]["cl"] or 0)))
    return f"*/ Chi phí tài chính giảm {fmt_num(giam)} đồng do Công ty hạch toán chi phí lãi vay không đúng quy định."

def _write_no_diff_year(paras, yi: int, end: int, year_label_line: str):
    replace_paragraph_text(paras[yi], year_label_line)
    text = "*/ Không chênh lệch."
    followers = list(range(yi + 1, end))
    if not followers:
        return
    first = followers[0]
    replace_paragraph_text(paras[first], text)
    for j in followers[1:]:
        replace_paragraph_text(paras[j], "")

def rewrite_year_result_block(paras, start_idx: int, years_data: dict, mode: str):
    window = []
    for i in range(start_idx, min(len(paras), start_idx + 45)):
        t = paras[i].text.strip()
        if re.search(r"(NĂM|Năm)\s*20\d{2}", t) or t.startswith("+ NĂM") or t.startswith("* NĂM"):
            window.append(i)
        if len(window) >= 4 and i > window[0] + 30:
            break
    if len(window) < 4:
        window = window[:4]
    if len(window) < 4:
        print(f"! Thiếu mốc năm (tìm được {len(window)}) từ P{start_idx}, mode={mode}")
        return

    year_idxs = window[:4]
    for slot, yi in enumerate(year_idxs):
        y = YEAR_SLOTS[slot]
        info = tax_access(years_data[y], y)
        cl = float(info["cptc"]["cl"] or 0)
        thue_cl = abs(round(float(info["thue"]["cl"] or 0)))
        tang_tn = abs(round(float(info["tn"]["cl"] or 0)))

        if mode == "td":
            end = year_idxs[slot + 1] if slot + 1 < len(year_idxs) else yi + 6
            if y == 2023:
                followers = list(range(yi + 1, end))
                replace_paragraph_text(paras[yi], "Năm 2023 Đơn vị tính: Đồng")
                if len(followers) >= 3:
                    replace_paragraph_text(paras[followers[0]], "*/ Không chênh lệch.")
                    replace_paragraph_text(paras[followers[1]], "Năm 2024 Đơn vị tính: Đồng")
                    replace_paragraph_text(paras[followers[2]], "*/ Không chênh lệch.")
                    for j in followers[3:]:
                        replace_paragraph_text(paras[j], "")
                else:
                    _write_no_diff_year(paras, yi, end, "Năm 2023 Đơn vị tính: Đồng")
            else:
                _write_no_diff_year(paras, yi, end, f"Năm {y} Đơn vị tính: Đồng")
            continue

        if mode == "tv":
            end = year_idxs[slot + 1] if slot + 1 < len(year_idxs) else yi + 6
            writable = [j for j in range(yi + 1, end) if paras[j].text.strip()]
            if abs(cl) < 1:
                if y == 2023:
                    followers = list(range(yi + 1, end))
                    replace_paragraph_text(paras[yi], "+ NĂM 2023: Đơn vị tính: Đồng")
                    if len(followers) >= 3:
                        replace_paragraph_text(paras[followers[0]], "*/ Không chênh lệch.")
                        replace_paragraph_text(paras[followers[1]], "+ NĂM 2024: Đơn vị tính: Đồng")
                        replace_paragraph_text(paras[followers[2]], "*/ Không chênh lệch.")
                        for j in followers[3:]:
                            replace_paragraph_text(paras[j], "")
                    else:
                        _write_no_diff_year(paras, yi, end, "+ NĂM 2023: Đơn vị tính: Đồng")
                else:
                    _write_no_diff_year(paras, yi, end, f"+ NĂM {y}: Đơn vị tính: Đồng")
            else:
                replace_paragraph_text(paras[yi], f"+ NĂM {y}: Đơn vị tính: Đồng")
                detail = (
                    cptc_detail(y, info)
                    + f" Thu nhập tính thuế tăng {fmt_num(tang_tn)} đồng. => Thuế TNDN truy thu: {fmt_num(thue_cl)} đồng."
                )
                if writable:
                    replace_paragraph_text(paras[writable[0]], detail)
                    for j in writable[1:]:
                        replace_paragraph_text(paras[j], "")
                else:
                    rj = yi + 1
                    while rj < len(paras) and not paras[rj].text.strip():
                        rj += 1
                    if rj < len(paras):
                        replace_paragraph_text(paras[rj], detail)
            continue

        # bbkt
        end = year_idxs[slot + 1] if slot + 1 < len(year_idxs) else yi + 8
        writable = [j for j in range(yi + 1, end) if paras[j].text.strip()]
        empties = [j for j in range(yi + 1, end) if not paras[j].text.strip()]
        if abs(cl) < 1:
            if y == 2023:
                followers = list(range(yi + 1, end))
                replace_paragraph_text(paras[yi], "* NĂM 2023 Đơn vị tính: Đồng")
                if len(followers) >= 3:
                    replace_paragraph_text(paras[followers[0]], "*/ Không chênh lệch.")
                    replace_paragraph_text(paras[followers[1]], "* NĂM 2024 Đơn vị tính: Đồng")
                    replace_paragraph_text(paras[followers[2]], "*/ Không chênh lệch.")
                    for j in followers[3:]:
                        replace_paragraph_text(paras[j], "")
                else:
                    _write_no_diff_year(paras, yi, end, "* NĂM 2023 Đơn vị tính: Đồng")
            else:
                _write_no_diff_year(paras, yi, end, f"* NĂM {y} Đơn vị tính: Đồng")
        else:
            replace_paragraph_text(paras[yi], f"* NĂM {y} Đơn vị tính: Đồng")
            lines = [
                "Giải trình số liệu chênh lệch qua kiểm tra:",
                (
                    f"Công ty báo cáo thu nhập tính thuế {fmt_num(info['tn']['bc'])} đồng, "
                    f"Đoàn kiểm tra xác định thu nhập tính thuế {fmt_num(info['tn']['kt'])} đồng, "
                    f"tăng {fmt_num(tang_tn)} đồng, nguyên nhân:"
                ),
                cptc_detail(y, info),
                f"=>Thuế TNDN truy thu qua kiểm tra: {fmt_num(thue_cl)} đồng.",
                "Giảm lỗ: không.",
            ]
            slots_idx = writable + empties
            if len(slots_idx) < 3:
                condensed = (
                    f"Giải trình số liệu chênh lệch qua kiểm tra: Công ty báo cáo thu nhập tính thuế "
                    f"{fmt_num(info['tn']['bc'])} đồng, Đoàn kiểm tra xác định thu nhập tính thuế "
                    f"{fmt_num(info['tn']['kt'])} đồng, tăng {fmt_num(tang_tn)} đồng. "
                    + cptc_detail(y, info).lstrip("*/ ").rstrip(".")
                    + f". => Thuế TNDN truy thu qua kiểm tra: {fmt_num(thue_cl)} đồng. Giảm lỗ: không."
                )
                if slots_idx:
                    replace_paragraph_text(paras[slots_idx[0]], condensed)
                    for j in slots_idx[1:]:
                        replace_paragraph_text(paras[j], "")
                else:
                    replace_paragraph_text(paras[yi + 1], condensed)
            else:
                for k, j in enumerate(slots_idx):
                    if k < len(lines):
                        replace_paragraph_text(paras[j], lines[k])
                    else:
                        replace_paragraph_text(paras[j], "")

def update_narratives_and_markers(doc: Document, data: dict):
    paras = doc.paragraphs
    years = data["years"]

    # 1) Biên bản xác nhận - thành viên
    tv_hits = [i for i, p in enumerate(paras) if p.text.strip() == "Xác nhận số liệu kiểm tra"]
    if tv_hits:
        start = None
        for h in tv_hits:
            for j in range(h, min(h + 20, len(paras))):
                if "II. Kết quả kiểm tra" in paras[j].text:
                    head = " ".join(paras[k].text for k in range(max(0, h - 15), h + 5))
                    if "Thành viên" in head and "Trưởng đoàn" not in head.split("Thành viên")[0][-40:]:
                        start = j + 1
                        break
            if start is not None:
                break
        if start is None and tv_hits:
            for j in range(tv_hits[0], min(tv_hits[0] + 25, len(paras))):
                if "II. Kết quả kiểm tra" in paras[j].text:
                    start = j + 1
                    break
        if start is not None:
            rewrite_year_result_block(paras, start, years, mode="tv")

    # 2) Biên bản xác nhận - trưởng đoàn
    td_start = None
    for h in tv_hits[1:] or tv_hits:
        head = " ".join(paras[k].text for k in range(h, min(h + 20, len(paras))))
        if "Trưởng đoàn" in head:
            for j in range(h, min(h + 30, len(paras))):
                if "II. Kết quả kiểm tra" in paras[j].text:
                    td_start = j + 1
                    break
        if td_start is not None:
            break
    if td_start is None:
        markers = (
            "Chi phí quản lý doanh nghiệp; Chi phí tài chính",
            "Chi phí quản lý doanh nghiệp.\nChi phí bán hàng",
            "Chi phí quản lý doanh nghiệp",
        )
        for i, p in enumerate(paras):
            if any(m in p.text for m in markers) and "bán hàng" in p.text.replace("\n", " "):
                for j in range(max(0, i - 2), min(i + 5, len(paras))):
                    if "II. Kết quả kiểm tra" in paras[j].text:
                        td_start = j + 1
                        break
            elif "Chi phí quản lý doanh nghiệp; Chi phí tài chính" in p.text:
                for j in range(i, min(i + 5, len(paras))):
                    if "II. Kết quả kiểm tra" in paras[j].text:
                        td_start = j + 1
                        break
            if td_start:
                break
    if td_start is not None:
        rewrite_year_result_block(paras, td_start, years, mode="td")

    # 3) Phân công text
    td_text = "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng."
    tv_text = "- Chi phí tài chính."
    for i, p in enumerate(paras):
        t = p.text.strip()
        if t in (
            "Chi phí quản lý doanh nghiệp; Chi phí tài chính.",
            "Chi phí quản lý doanh nghiệp; Chi phí tài chính",
            "Chi phí bán hàng.",
            "Chi phí bán hàng",
            "Chi phí tài chính.",
            "Chi phí tài chính",
        ):
            head = " ".join(paras[k].text for k in range(max(0, i - 8), i))
            if "Trưởng đoàn" in head and "Thành viên" not in head.split("Trưởng đoàn")[-1]:
                replace_paragraph_text(p, td_text)
            elif "Thành viên" in head:
                replace_paragraph_text(p, tv_text)
            elif "Chi phí quản lý doanh nghiệp; Chi phí tài chính" in t:
                replace_paragraph_text(p, td_text)
            elif t.startswith("Chi phí bán hàng"):
                if "Thành viên" in head:
                    replace_paragraph_text(p, tv_text)

    # 4) Biên bản kiểm tra
    bbkt = [i for i, p in enumerate(paras) if p.text.strip() == "BIÊN BẢN KIỂM TRA"]
    if bbkt:
        start = None
        for j in range(bbkt[0], min(bbkt[0] + 80, len(paras))):
            if "II. NỘI DUNG KIỂM TRA" in paras[j].text:
                start = j + 1
                break
        if start is not None:
            rewrite_year_result_block(paras, start, years, mode="bbkt")

    # Dọn các dòng mẫu rác còn sót
    for p in paras:
        t = p.text
        if "hạch toán chi phí giá vốn không tương ứng doanh thu" in t or "78.475.200" in t or "117.712.800" in t:
            replace_paragraph_text(p, "")
        if "cộng sai số học" in t or "49.414.559" in t or "970.731" in t:
            replace_paragraph_text(p, "")

def apply_nd30_to_admin_docs_only(doc: Document):
    paras = doc.paragraphs

    # --- 1. Tờ trình kiểm tra (P001 - P039) ---
    for i in range(0, min(15, len(paras))):
        p = paras[i]
        t = p.text.strip()
        if t == "TỜ TRÌNH LÃNH ĐẠO THUẾ CƠ SỞ":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=14, bold=True, italic=False)
        elif t.startswith("V/v:"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True, italic=False)

    # --- 2. Quyết định kiểm tra (P040 - P063) & Quyết định giám sát (P064 - P083) ---
    for i in range(35, min(85, len(paras))):
        p = paras[i]
        t = p.text.strip()
        if t == "QUYẾT ĐỊNH":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=14, bold=True, italic=False)
        elif t.startswith("Về việc kiểm tra") or t.startswith("Về việc giám sát"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True, italic=False)
        elif t.startswith("CÔNG TY TNHH") or t.startswith("TRƯỞNG THUẾ CƠ SỞ"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True, italic=False)
        elif t == "QUYẾT ĐỊNH:":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True, italic=False)

    # --- 3. Quyết định xử phạt VPHC (Doc 18) ---
    if len(doc.tables) > 42:
        t42 = doc.tables[42]
        c00 = t42.rows[0].cells[0]
        for p in c00.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                if "CƠ SỞ 3" in r.text or "THUẾ CƠ SỞ" in p.text:
                    set_run_font(r, size_pt=12, bold=True, italic=False)
                else:
                    set_run_font(r, size_pt=12, bold=False, italic=False)
        c01 = t42.rows[0].cells[1]
        for p in c01.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if "CỘNG HÒA" in p.text:
                for r in p.runs:
                    set_run_font(r, size_pt=12, bold=True, italic=False)
            elif "Độc lập" in p.text:
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=True, italic=False)
        r_last = t42.rows[-1]
        for p in r_last.cells[0].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                set_run_font(r, size_pt=12, bold=False, italic=False)
        for p in r_last.cells[1].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=False, italic=True)

    qd_start = -1
    for i, p in enumerate(paras):
        if p.text.strip() == "QUYẾT ĐỊNH" and i >= 540:
            qd_start = i
            break

    if qd_start != -1:
        p_title = paras[qd_start]
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.first_line_indent = Cm(0)
        p_title.paragraph_format.space_before = Pt(6)
        p_title.paragraph_format.space_after = Pt(2)
        for r in p_title.runs:
            set_run_font(r, size_pt=14, bold=True, italic=False)

        p_trich_yeu = paras[qd_start + 1]
        p_trich_yeu.text = "Về việc xử phạt vi phạm hành chính về thuế"
        p_trich_yeu.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_trich_yeu.paragraph_format.first_line_indent = Cm(0)
        p_trich_yeu.paragraph_format.space_before = Pt(2)
        p_trich_yeu.paragraph_format.space_after = Pt(6)
        for r in p_trich_yeu.runs:
            set_run_font(r, size_pt=13, bold=True, italic=False)

        for i in range(qd_start + 2, min(len(paras), qd_start + 65)):
            p = paras[i]
            t = p.text.strip()
            if not t:
                continue
            if t.startswith("TỜ TRÌNH LÃNH ĐẠO"):
                break
            if t.startswith("Căn cứ"):
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Cm(1.0)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.0
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=False, italic=True)
            elif t == "TRƯỞNG THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI":
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=True, italic=False)
            elif t.startswith("Theo đề nghị"):
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Cm(1.0)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.0
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=False, italic=False)
            elif t == "QUYẾT ĐỊNH:":
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.0
                for r in p.runs:
                    set_run_font(r, size_pt=13, bold=True, italic=False)
            elif t.startswith("Điều "):
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = Cm(1.0)
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.0
                for r in p.runs:
                    if r.text.strip().startswith("Điều"):
                        set_run_font(r, size_pt=13, bold=True, italic=False)
                    else:
                        set_run_font(r, size_pt=13, bold=None, italic=False)

    if len(doc.tables) > 43:
        t43 = doc.tables[43]
        for p in t43.rows[0].cells[0].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if "Nơi nhận" in p.text:
                for r in p.runs:
                    set_run_font(r, size_pt=12, bold=True, italic=True)
            else:
                for r in p.runs:
                    set_run_font(r, size_pt=11, bold=False, italic=False)
        for p in t43.rows[0].cells[1].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                if any(k in r.text for k in ["KT. TRƯỞNG", "PHÓ TRƯỞNG", "Phạm Đồng Xuân"]):
                    set_run_font(r, size_pt=13, bold=True, italic=False)
                else:
                    set_run_font(r, size_pt=13, bold=False, italic=False)

    # --- 4. Tờ trình xử phạt (Doc 19) ---
    for i, p in enumerate(paras):
        if "TỜ TRÌNH LÃNH ĐẠO THUẾ CƠ SỞ 3" in p.text and i >= 590:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                set_run_font(r, size_pt=14, bold=True, italic=False)
        elif "Về việc: ban hành Quyết Định xử phạt" in p.text and i >= 590:
            p.text = "Về việc: ban hành Quyết định xử phạt vi phạm trong lĩnh vực thuế."
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(6)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True, italic=False)
        elif "Nghị định số 125" in p.text and i >= 590:
            if "125/2022/NĐ-CP" in p.text:
                p.text = p.text.replace("125/2022/NĐ-CP", "125/2020/NĐ-CP")
            if "19/10/2022" in p.text:
                p.text = p.text.replace("19/10/2022", "19/10/2020")
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.0)
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.0
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=None, italic=None)
        elif "Trấn Biên, ngày 02 tháng 10 năm 2026" in p.text and i >= 620:
            p.text = p.text.replace("02 tháng 10", "14 tháng 10")
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=False, italic=True)

def build_single_page_qd_kiem_tra(master_doc, output_path):
    doc = docx.Document(TEMPLATE_PATH)
    body = doc._element.body
    sectPr = None
    for el in list(body):
        if el.tag.endswith('sectPr'):
            sectPr = copy.deepcopy(el)
            break
            
    master_body = master_doc._element.body
    all_master_elems = list(master_body)

    start = end = None
    for i, el in enumerate(all_master_elems):
        text = "".join(el.itertext())
        if start is None and el.tag.endswith("tbl") and "THUẾ THÀNH PHỐ ĐỒNG NAI" in text and i > 20:
            nxt = "".join("".join(all_master_elems[j].itertext()) for j in range(i, min(i + 5, len(all_master_elems))))
            if "Về việc kiểm tra thuế tại" in nxt:
                start = i
        if start is not None and end is None and el.tag.endswith("tbl") and "Nơi nhận" in text and "Điều 3" in text:
            end = i
            break
    if start is None or end is None:
        start, end = 27, 50
    segment = [copy.deepcopy(all_master_elems[i]) for i in range(start, end + 1)]
    
    if segment:
        el0 = segment[0]
        if el0.tag.endswith('p'):
            for br in el0.xpath('.//w:br[@w:type="page"]'):
                br.getparent().remove(br)
            texts = el0.xpath('.//w:t')
            t = "".join([node.text for node in texts if node.text]).strip()
            if not t:
                segment.pop(0)
                
    body.clear()
    for el in segment:
        body.append(el)
    if sectPr is not None:
        body.append(sectPr)
        
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2.7)
    sec.right_margin = Cm(1.8)
    
    for t in doc.tables:
        tblPr = t._element.tblPr
        tblpPr = tblPr.find(qn('w:tblpPr'))
        if tblpPr is not None:
            tblPr.remove(tblpPr)
        tblOverlap = tblPr.find(qn('w:tblOverlap'))
        if tblOverlap is not None:
            tblPr.remove(tblOverlap)
        for r_idx, row in enumerate(t.rows):
            for c_idx, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        if r.font.size and r.font.size.pt > 12.5:
                            r.font.size = Pt(12)
                            
    for p in list(doc.paragraphs):
        text = p.text.strip()
        if not text:
            if p._element.getparent() is not None:
                p._element.getparent().remove(p._element)
            continue
            
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        
        for r in p.runs:
            r.font.name = "Times New Roman"
            
        if "QUYẾT ĐỊNH" == text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.size = Pt(13)
                r.font.bold = True
        elif "Về việc kiểm tra thuế tại" in text or "CÔNG TY TNHH" in text:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.size = Pt(12)
                r.font.bold = True
        elif "TRƯỞNG THUẾ CƠ SỞ" in text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.size = Pt(12)
                r.font.bold = True
        elif text.startswith("Căn cứ") or text.startswith("Theo đề nghị"):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                r.font.size = Pt(11.5)
        elif "QUYẾT ĐỊNH:" == text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.size = Pt(12.5)
                r.font.bold = True
        elif text.startswith("Điều 1.") or text.startswith("Điều 2.") or text.startswith("Điều 3."):
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                r.font.size = Pt(12)
        else:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                if r.font.size and r.font.size.pt > 12:
                    r.font.size = Pt(12)
                    
    doc.save(output_path)
    print(f"-> Đã xuất: {os.path.basename(output_path):42s} (Chuẩn 1 trang A4)")

def build_single_page_qd_giam_sat(master_doc, output_path):
    doc = docx.Document(TEMPLATE_PATH)
    body = doc._element.body
    sectPr = None
    for el in list(body):
        if el.tag.endswith('sectPr'):
            sectPr = copy.deepcopy(el)
            break
            
    master_body = master_doc._element.body
    all_master_elems = list(master_body)

    start = end = None
    seen_kt_end = False
    for i, el in enumerate(all_master_elems):
        text = "".join(el.itertext())
        if (not seen_kt_end) and el.tag.endswith("tbl") and "Nơi nhận" in text and "Điều 3" in text and i > 20:
            seen_kt_end = True
            continue
        if seen_kt_end and start is None and el.tag.endswith("tbl") and "THUẾ THÀNH PHỐ ĐỒNG NAI" in text:
            nxt = "".join("".join(all_master_elems[j].itertext()) for j in range(i, min(i + 5, len(all_master_elems))))
            if "Về việc giám sát" in nxt:
                start = i
        if start is not None and end is None and el.tag.endswith("tbl") and "Nơi nhận" in text and "Điều 4" in text:
            end = i
            break
    if start is None or end is None:
        start, end = 52, 70
    segment = [copy.deepcopy(all_master_elems[i]) for i in range(start, end + 1)]
    
    if segment:
        el0 = segment[0]
        if el0.tag.endswith('p'):
            for br in el0.xpath('.//w:br[@w:type="page"]'):
                br.getparent().remove(br)
            texts = el0.xpath('.//w:t')
            t = "".join([node.text for node in texts if node.text]).strip()
            if not t:
                segment.pop(0)
                
    body.clear()
    for el in segment:
        body.append(el)
    if sectPr is not None:
        body.append(sectPr)
        
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2.7)
    sec.right_margin = Cm(1.8)
    
    for t in doc.tables:
        tblPr = t._element.tblPr
        tblpPr = tblPr.find(qn('w:tblpPr'))
        if tblpPr is not None:
            tblPr.remove(tblpPr)
        tblOverlap = tblPr.find(qn('w:tblOverlap'))
        if tblOverlap is not None:
            tblPr.remove(tblOverlap)
        for r_idx, row in enumerate(t.rows):
            for c_idx, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        if r.font.size and r.font.size.pt > 12.5:
                            r.font.size = Pt(12)
                            
    for p in list(doc.paragraphs):
        text = p.text.strip()
        if not text:
            if p._element.getparent() is not None:
                p._element.getparent().remove(p._element)
            continue
            
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        
        for r in p.runs:
            r.font.name = "Times New Roman"
            
        if "QUYẾT ĐỊNH" == text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.size = Pt(13)
                r.font.bold = True
        elif "Về việc giám sát" in text or "CÔNG TY TNHH" in text:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.size = Pt(12)
                r.font.bold = True
        elif "TRƯỞNG THUẾ CƠ SỞ" in text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.size = Pt(12)
                r.font.bold = True
        elif text.startswith("Căn cứ") or text.startswith("Theo đề nghị"):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                r.font.size = Pt(11.5)
        elif "QUYẾT ĐỊNH:" == text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.size = Pt(12.5)
                r.font.bold = True
        elif text.startswith("Điều 1.") or text.startswith("Điều 2.") or text.startswith("Điều 3.") or text.startswith("Điều 4."):
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                r.font.size = Pt(12)
        else:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                if r.font.size and r.font.size.pt > 12:
                    r.font.size = Pt(12)
                    
    doc.save(output_path)
    print(f"-> Đã xuất: {os.path.basename(output_path):42s} (Chuẩn 1 trang A4)")

def process_doc_04_di01(doc: Document):
    for i, p in enumerate(doc.paragraphs):
        t = p.text
        if "d) Số chi nhánh, cơ sở phụ thuộc: ." in t:
            p.text = "d) Số chi nhánh, cơ sở phụ thuộc: Cửa hàng Bao Bì Giấy Vạn Phát Thành tại số 507, Quốc Lộ 51, KP 3, Phường An Hòa (hết hiệu lực)."
        elif "Doanh nghiệp kê khai lãi từ năm 2022" in t:
            p.text = "- Doanh nghiệp kê khai kết quả hoạt động kinh doanh có lãi liên tục từ năm 2021 đến năm 2025, doanh thu thuần tăng trưởng đều qua các năm (năm 2021: 4.275.904.321 đ, năm 2022: 5.304.710.195 đ, năm 2023: 6.851.592.363 đ, năm 2024: 8.652.414.302 đ, năm 2025: 10.060.947.641 đ)."
        elif "Tỷ lệ giá vốn hàng bán trên doanh thu năm 2022" in t:
            p.text = "- Tỷ lệ giá vốn hàng bán trên doanh thu thuần duy trì ở mức cao và có xu hướng tăng dần (năm 2021: 95,21%, năm 2022: 95,57%, năm 2023: 96,43%, năm 2024: 96,57%, năm 2025: 97,84%)."
        elif "Tỷ lệ Chi phí quản lý kinh doanh trên doanh thu" in t:
            p.text = "- Tỷ lệ chi phí quản lý doanh nghiệp trên doanh thu thuần giảm mạnh qua các năm (năm 2021: 2,73%, năm 2022: 1,89%, năm 2023: 0,45%, năm 2024: 0,41%, năm 2025: 0,33%).\n- Tỷ lệ chi phí bán hàng trên doanh thu thuần (năm 2021: 0,10%, năm 2022: 0,00%, năm 2023: 0,00%, năm 2024: 0,00%, năm 2025: 0,05%).\n- Tỷ lệ chi phí tài chính trên doanh thu thuần (năm 2021: 0,73%, năm 2022: 0,64%, năm 2023: 1,22%, năm 2024: 0,77%, năm 2025: 0,03%)."
        elif "Tỷ lệ giá vốn hàng bán và chi phí quản lý kinh doanh trên doanh thu biến động không đều" in t:
            p.text = "- Chi phí quản lý doanh nghiệp, chi phí bán hàng và chi phí tài chính có sự biến động không tương ứng với tốc độ tăng trưởng doanh thu thuần của doanh nghiệp."
        elif "IV. ĐỀ XUẤT NỘI DUNG KIỂM TRA:" in t:
            if i + 1 < len(doc.paragraphs):
                doc.paragraphs[i+1].text = "- Chi phí quản lý doanh nghiệp, thời kỳ kiểm tra: Từ năm 2021 đến năm 2025.\n- Chi phí bán hàng, thời kỳ kiểm tra: Từ năm 2021 đến năm 2025.\n- Chi phí tài chính, thời kỳ kiểm tra: Từ năm 2021 đến năm 2025."
            if i + 2 < len(doc.paragraphs):
                t2 = doc.paragraphs[i + 2].text.strip()
                if "Phân_công_thành_viên" in t2 or t2 in ("Chi phí bán hàng", "Chi phí bán hàng.", "- Chi phí tài chính.", "- Chi phí tài chính"):
                    doc.paragraphs[i + 2].text = ""

def clean_doc_flow(doc: Document):
    empty_run = []
    for p in list(doc.paragraphs):
        has_pb = any(br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page' for br in p._element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br'))
        pbb = len(p._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore')) > 0
        if (not p.text.strip()) and (not has_pb) and (not pbb):
            empty_run.append(p)
        else:
            if len(empty_run) >= 3:
                for ep in empty_run[1:]:
                    el = ep._element
                    parent = el.getparent()
                    if parent is not None:
                        parent.remove(el)
            empty_run = []
    if len(empty_run) >= 3:
        for ep in empty_run[1:]:
            el = ep._element
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

def run_rebuild():
    print("=== TÁI LẬP HỒ SƠ TỪ PHÔI GỐC MAUTONG HOP.docx ===")
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    tong_path = os.path.join(TARGET_DIR, "TONG HOP.docx")
    if os.path.exists(tong_path):
        bak = os.path.join(TARGET_DIR, f"TONG HOP.bak-reapply-original-{ts}.docx")
        shutil.copy2(tong_path, bak)
        print(f"-> Đã sao lưu: {os.path.basename(bak)}")

    # 1. Đọc phôi gốc
    doc = Document(TEMPLATE_PATH)
    print(f"-> Đã mở phôi gốc: {TEMPLATE_PATH} ({len(doc.paragraphs)} đoạn văn, {len(doc.tables)} bảng)")

    # 2. Đọc số liệu DTCHEP
    data = load_dtchep(XLSX_PATH)
    reps = build_replacements_dict(data)

    # 3. Thay thế placeholders
    replace_in_doc(doc, reps)

    # 4. Xử lý Nội dung kiểm tra 3 chuyên đề & thời hạn
    for p in doc.paragraphs:
        if "«Nội_dung_kiểm_tra»" in p.text:
            p.text = "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng.\n- Chi phí tài chính."
        elif "Thời hạn kiểm tra là" in p.text and "ngày làm việc" in p.text:
            p.text = (
                "Thời hạn kiểm tra là 20 ngày làm việc "
                "(không bao gồm ngày nghỉ, ngày lễ theo quy định của pháp luật) "
                "kể từ ngày công bố Quyết định kiểm tra."
            )

    # 5. Xử lý DI-01
    process_doc_04_di01(doc)

    # 6. Điền bảng DTCHEP
    fill_dtchep_tables(doc, data)

    # 7. Cập nhật lời giải trình
    update_narratives_and_markers(doc, data)

    # 8. Chuẩn hóa NĐ 30 riêng cho nhóm văn bản hành chính (Tờ trình KT, QĐKT, QĐGS, QĐXPHC, Tờ trình XPHC)
    apply_nd30_to_admin_docs_only(doc)

    # 9. Dọn luồng trang nhẹ
    clean_doc_flow(doc)

    # 10. Lưu TONG HOP.docx
    doc.save(tong_path)
    print(f"-> Đã lưu: TONG HOP.docx ({len(doc.paragraphs)} đoạn, {len(doc.tables)} bảng)")

    # 11. Tạo 2 file Quyết định độc lập chuẩn 1 trang A4
    qd_kt = os.path.join(TARGET_DIR, "QD Kiem tra tai cty Van Phat Thanh.docx")
    qd_gs = os.path.join(TARGET_DIR, "QDgiam sat tai cty Van Phat Thanh.docx")
    build_single_page_qd_kiem_tra(doc, qd_kt)
    build_single_page_qd_giam_sat(doc, qd_gs)

    # 12. Đóng gói hoso.zip
    zip_path = os.path.join(TARGET_DIR, "hoso.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in [
            "TONG HOP.docx",
            "QD Kiem tra tai cty Van Phat Thanh.docx",
            "QDgiam sat tai cty Van Phat Thanh.docx",
        ]:
            zf.write(os.path.join(TARGET_DIR, fname), fname)
            print(f"  + ZIP: {fname}")

    print("=== HOÀN TẤT TÁI LẬP ===")

if __name__ == "__main__":
    run_rebuild()
