# -*- coding: utf-8 -*-
"""Xuat ho so di kiem tra theo ke hoach - Cong ty TNHH TMDV T&T Bien Hoa (MST 3603883505).

Nguon du lieu:
- Phoi: templates/MAUTONG HOP.docx
- So: Danh sach kiem tra quyet toan pro.xlsm (sheet 02_DANG_THUC_HIEN, hang MST 3603883505)
- Bang tinh: Bang tinh.xlsx (sheet KQKDSXKD, nam 2022-2025)

Quy cach xuat: TONG HOP.docx.
Chua co So QDKT that -> de trong theo mau `So:              /QD-TCS3` (khong dung DU_THAO-TRINH-KY).
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import build_bien_ban_ho_so_ve_vpt as base

BRAIN_DIR = SCRIPT_DIR.parent
TEMPLATE_PATH = BRAIN_DIR / "templates" / "MAUTONG HOP.docx"
TARGET_DIR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD"
    r"\12. TT_BIEN_HOA_3603883505\0. BIEN BAN"
)
BANG_TINH_PATH = TARGET_DIR / "Bang tinh.xlsx"
EXCEL_SO = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung"
    r"\Danh sach kiem tra quyet toan pro.xlsm"
)

MST = "3603883505"
COMPANY_NAME = "Công ty TNHH Thương mại Dịch vụ T&T Biên Hòa"
SO_QD_BLANK = "              /QĐ-TCS3"
SIGN_BLANK_LINES = 5
SIGNER_NAMES = {"Phạm Đồng Xuân", "Nguyễn Trường Giang", "Nguyễn Thị Ánh Ngọc", "Trịnh Thị Thanh Hương"}
CHUC_DANH_MARKERS = (
    "KT. TRƯỞNG",
    "PHÓ TRƯỞNG",
    "TRƯỞNG THUẾ CƠ SỞ",
    "TRƯỞNG ĐOÀN",
    "TỔ TRƯỞNG",
)
YEARS = (2022, 2023, 2024, 2025)


def resolve_so_qd(raw: str | None) -> str:
    s = str(raw or "").strip()
    if not s or s.upper() == "DU_THAO-TRINH-KY":
        return SO_QD_BLANK
    return s


def set_keep_with_next(paragraph: Paragraph, on: bool = True) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    tag = qn("w:keepNext")
    existing = pPr.find(tag)
    if on and existing is None:
        pPr.append(OxmlElement("w:keepNext"))
    elif not on and existing is not None:
        pPr.remove(existing)


def set_paragraph_text_preserve_spaces(paragraph: Paragraph, text: str) -> None:
    if paragraph.runs:
        paragraph.runs[0].text = text
        for r in paragraph.runs[1:]:
            r.text = ""
        run = paragraph.runs[0]
    else:
        run = paragraph.add_run(text)
    t_el = run._r.find(qn("w:t"))
    if t_el is not None:
        t_el.set(qn("xml:space"), "preserve")


def insert_paragraph_after(paragraph: Paragraph) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)


def fmt_vnd(n: float | int) -> str:
    return f"{int(round(n)):,}".replace(",", ".")


def pct(part: float, whole: float) -> str:
    if not whole:
        return "0%"
    return f"{part / whole * 100:.2f}%".replace(".", ",")


def load_company_from_so() -> dict:
    import openpyxl

    wb = openpyxl.load_workbook(EXCEL_SO, data_only=True, keep_vba=False)
    ws = None
    row = None
    for sheetname in ["02_DANG_THUC_HIEN", "01_DA_BAN_HANH"]:
        s = wb[sheetname]
        for r in range(5, s.max_row + 1):
            if str(s.cell(r, 4).value or "").strip() == MST:
                ws = s
                row = r
                break
        if row is not None:
            break

    if row is None:
        raise RuntimeError(f"Không tìm thấy MST {MST} trên sổ")

    def v(c):
        return ws.cell(row, c).value

    dia_chi = str(v(6) or "").strip()
    tk = str(v(21) or "").strip()
    tk = re.sub(r"^\+\s*", "", tk)
    tk = re.sub(r"^(TK|Tài khoản)\s*số\s*", "", tk, flags=re.IGNORECASE).strip()

    ngay_dkkd = v(19)
    if hasattr(ngay_dkkd, "strftime"):
        ngay_dkkd_s = ngay_dkkd.strftime("%d/%m/%Y")
    else:
        ngay_dkkd_s = str(ngay_dkkd or "").strip()

    thoi_han = v(33)
    thoi_han_s = str(thoi_han or "").strip()
    if thoi_han_s.endswith(".0"):
        thoi_han_s = thoi_han_s[:-2]

    return {
        "row": row,
        "ten": str(v(5) or "").strip(),
        "dia_chi": dia_chi,
        "dia_ban": str(v(7) or "").strip(),
        "ma_cqt": str(v(8) or "").strip(),
        "thoi_ky": str(v(9) or "").strip(),
        "nnkd": str(v(10) or "").strip(),
        "gd": str(v(11) or "").strip(),
        "cccd": str(v(13) or "").strip(),
        "ngay_cap_cccd": v(14),
        "noi_cap": str(v(15) or "").strip(),
        "kt": str(v(16) or "").strip() if v(16) else "",
        "cccd_kt": str(v(17) or "").strip() if v(17) else "",
        "gcn": str(v(18) or "").strip(),
        "ngay_dkkd": ngay_dkkd_s,
        "chi_nhanh": str(v(20) or "").strip(),
        "tai_khoan": tk,
        "td": str(v(22) or "").strip(),
        "cv_td": str(v(23) or "").strip(),
        "tv": str(v(24) or "").strip(),
        "cv_tv": str(v(25) or "").strip(),
        "noi_dung": str(v(26) or "").strip(),
        "phan_cong_td": str(v(27) or "").strip(),
        "phan_cong_tv": str(v(28) or "").strip(),
        "so_qdkt": str(v(29) or "").strip() if v(29) else "",
        "so_qdgs": str(v(30) or "").strip() if v(30) else "",
        "ngay_ban_hanh": v(31),
        "thoi_han": thoi_han_s,
    }


def load_kqkd_ratios() -> dict:
    import openpyxl

    wb = openpyxl.load_workbook(BANG_TINH_PATH, data_only=True)
    ws = wb["KQKDSXKD"]
    year_cols = {2022: 15, 2023: 16, 2024: 17, 2025: 18}

    def row_vals(row_idx: int) -> dict[int, float]:
        out = {}
        for y, c in year_cols.items():
            raw = ws.cell(row_idx, c).value
            try:
                out[y] = float(raw or 0)
            except Exception:
                out[y] = 0.0
        return out

    return {
        "dt": row_vals(14),
        "gv": row_vals(15),
        "bh": row_vals(20),
        "qldn": row_vals(21),
        "cptc": row_vals(18),
        "ln": row_vals(26),
        "year_cols": year_cols,
    }


def _normalize_phan_cong(text: str) -> str:
    parts = []
    for raw in (text or "").replace("\r", "").split("\n"):
        s = raw.strip().lstrip("-").strip().rstrip(".")
        if s:
            parts.append(s)
    return ("; ".join(parts) + ".") if parts else ""


def _build_noi_dung(phan_cong_td: str, phan_cong_tv: str, thoi_ky: str) -> str:
    items = []
    for block in (phan_cong_td, phan_cong_tv):
        for raw in (block or "").replace("\r", "").split("\n"):
            s = raw.strip().lstrip("-").strip().rstrip(".")
            if s:
                items.append(s)
    seen = set()
    uniq = []
    for it in items:
        key = it.lower()
        if key not in seen:
            seen.add(key)
            uniq.append(it)
    if not uniq:
        return f"- Kiểm tra chấp hành pháp luật thuế {thoi_ky}."
    return "\n".join(f"- {it} {thoi_ky}." for it in uniq)


def build_reps(info: dict) -> dict:
    so_qdkt = resolve_so_qd(info["so_qdkt"])
    so_qdgs = resolve_so_qd(info["so_qdgs"])
    ngay_qd = "..../..../2026"
    if info["ngay_ban_hanh"] and hasattr(info["ngay_ban_hanh"], "strftime"):
        ngay_qd = info["ngay_ban_hanh"].strftime("%d/%m/%Y")
    elif isinstance(info["ngay_ban_hanh"], str) and info["ngay_ban_hanh"].strip():
        ngay_qd = info["ngay_ban_hanh"].strip()

    ngay_cap = "..../..../...."
    if info["ngay_cap_cccd"] and hasattr(info["ngay_cap_cccd"], "strftime"):
        ngay_cap = info["ngay_cap_cccd"].strftime("%d/%m/%Y")

    dia_chi_full = info["dia_chi"]
    if dia_chi_full and "Đồng Nai" not in dia_chi_full:
        dia_chi_full = f"{info['dia_chi']}, Tỉnh Đồng Nai"

    cqt = "Thuế Cơ sở 3 Thành phố Đồng Nai"

    thoi_ky = info["thoi_ky"] or "Từ năm 2022 đến năm 2025"
    pc_td = _normalize_phan_cong(info["phan_cong_td"])
    pc_tv = _normalize_phan_cong(info["phan_cong_tv"])
    noi_dung = _build_noi_dung(info["phan_cong_td"], info["phan_cong_tv"], thoi_ky)

    return {
        "Tên_công_ty": COMPANY_NAME,
        "MST": MST,
        "địa_chỉ": info["dia_chi"] or "................",
        "ĐỊA_CHỈ": dia_chi_full or "................",
        "NNKD": info["nnkd"] or "Bán buôn thiết bị và linh kiện điện tử, viễn thông",
        "GĐ": info["gd"] or "Lê Tiến Trung",
        "KT": info["kt"] or ".........",
        "CCCD": info["cccd"] or "................",
        "ngày_cấp": ngay_cap,
        "Ngày_cấp2": ngay_cap,
        "Nơi_cấp": info["noi_cap"] or "Cục Cảnh sát Quản lý hành chính về trật tự xã hội",
        "GGCNĐKKD": info["gcn"] or MST,
        "CCCDKT": info["cccd_kt"] or "................",
        "Chi_nhánh": info["chi_nhanh"] or "Không",
        "Tài_khoản_số": info["tai_khoan"] or "................",
        "Thời_kỳ_kiểm_tra": thoi_ky,
        "Thời_hạn_kiểm_tra_": info["thoi_han"] or "10",
        "TĐ": info["td"] or "Nguyễn Thị Ánh Ngọc",
        "CVTĐ": info["cv_td"] or "Chuyên viên",
        "TV": info["tv"] or "Trịnh Thị Thanh Hương",
        "CVTV": info["cv_tv"] or "Kiểm tra viên thuế",
        "Ma_CQT": info["ma_cqt"] or "........",
        "CQT": cqt,
        "QĐKT": so_qdkt,
        "QĐGS": so_qdgs,
        "ngày_QĐ": ngay_qd,
        "NGÀY_QĐ": ngay_qd,
        "Ngày_công_bố": "..../..../2026",
        "Từ_ngày": "..../..../2026",
        "Đến_ngày": "..../..../2026",
        "Phân_công_trưởng_đoàn": pc_td or "- Chi phí quản lý kinh doanh.",
        "Phân_công_thành_viên": pc_tv or "- Chi phí tài chính.",
        "Nội_dung_kiểm_tra": noi_dung,
        "Ngày_ký_Biên_bản_kiểm_tra": "..../..../2026",
        "Ngày_báo_cáo": "..../..../2026",
        "Thuế_TNDN": "-",
        "Thuế_GTGT": "-",
        "Số_tiền_phạt_20": "-",
        "Số_tiền_phạt_20_BC": "-",
        "Số_tiền_phạt_003_TNDN": "-",
        "Số_tiền_phạt_003_GTGT": "-",
        "Phạt_KK_sai_VPHC": "-",
        "Tổng_phạt_kk_sai": "-",
        "Tổng_phạt": "-",
        "Tổng_cộng_truy_thu": "-",
        "Tổng_cộng_truy_thu_BC": "-",
        "Tổng_cộng_số_phạt__chậm_nộp": "-",
        "Tổng_cộng_phạt_và_truy_thu": "-",
        "Tổng_cộng_phạt_và_truy_thu_BC": "-",
        "Tổng": "-",
        "Giảm_lỗ": "-",
        "giảm_khấu_trừ": "-",
    }


def patch_noi_dung_and_thoi_han(doc: Document, noi_dung: str, thoi_han: str):
    for p in doc.paragraphs:
        if "«Nội_dung_kiểm_tra»" in p.text or p.text.strip() == "«Nội_dung_kiểm_tra»":
            base.replace_paragraph_text(p, noi_dung)
        elif "Thời hạn kiểm tra là" in p.text and "ngày làm việc" in p.text:
            han = thoi_han if thoi_han and thoi_han != "......" else "10"
            base.replace_paragraph_text(
                p,
                f"Thời hạn kiểm tra là {han} ngày làm việc "
                "(không bao gồm ngày nghỉ, ngày lễ theo quy định của pháp luật) "
                "kể từ ngày công bố Quyết định kiểm tra.",
            )


def _lai_lo_phrase(ln: dict) -> str:
    lai, lo = [], []
    for y in YEARS:
        val = ln.get(y, 0) or 0
        if val > 0:
            lai.append(str(y))
        elif val < 0:
            lo.append(str(y))
    parts = []
    if lo:
        parts.append(f"kê khai lỗ năm {', '.join(lo)}")
    if lai:
        parts.append(f"kê khai lãi năm {', '.join(lai)}")
    if not parts:
        return "Doanh nghiệp kê khai kết quả hoạt động kinh doanh giai đoạn 2022-2025"
    return "Doanh nghiệp " + "; ".join(parts)


def patch_di01_phan_tich(doc: Document, ratios: dict, chi_nhanh: str, thoi_ky: str):
    dt = ratios.get("dt") or {}
    gv = ratios.get("gv") or {}
    bh = ratios.get("bh") or {}
    qldn = ratios.get("qldn") or {}
    cptc = ratios.get("cptc") or {}
    ln = ratios.get("ln") or {}

    if all(y in dt for y in YEARS):
        dt_line = (
            f"- {_lai_lo_phrase(ln)}; "
            f"doanh thu thuần năm 2022: {fmt_vnd(dt[2022])} đ, "
            f"năm 2023: {fmt_vnd(dt[2023])} đ, "
            f"năm 2024: {fmt_vnd(dt[2024])} đ, "
            f"năm 2025: {fmt_vnd(dt[2025])} đ."
        )
        gv_line = (
            "- Tỷ lệ giá vốn hàng bán trên doanh thu thuần: "
            f"năm 2022: {pct(gv.get(2022, 0), dt[2022])}, "
            f"năm 2023: {pct(gv.get(2023, 0), dt[2023])}, "
            f"năm 2024: {pct(gv.get(2024, 0), dt[2024])}, "
            f"năm 2025: {pct(gv.get(2025, 0), dt[2025])}."
        )
        qldn_line = (
            "- Tỷ lệ chi phí quản lý kinh doanh trên doanh thu thuần: "
            f"năm 2022: {pct(qldn.get(2022, 0), dt[2022])}, "
            f"năm 2023: {pct(qldn.get(2023, 0), dt[2023])}, "
            f"năm 2024: {pct(qldn.get(2024, 0), dt[2024])}, "
            f"năm 2025: {pct(qldn.get(2025, 0), dt[2025])}."
        )
        bh_line = (
            "- Tỷ lệ chi phí bán hàng trên doanh thu thuần: "
            f"năm 2022: {pct(bh.get(2022, 0), dt[2022])}, "
            f"năm 2023: {pct(bh.get(2023, 0), dt[2023])}, "
            f"năm 2024: {pct(bh.get(2024, 0), dt[2024])}, "
            f"năm 2025: {pct(bh.get(2025, 0), dt[2025])}."
        )
        cptc_line = (
            "- Tỷ lệ chi phí tài chính trên doanh thu thuần: "
            f"năm 2022: {pct(cptc.get(2022, 0), dt[2022])}, "
            f"năm 2023: {pct(cptc.get(2023, 0), dt[2023])}, "
            f"năm 2024: {pct(cptc.get(2024, 0), dt[2024])}, "
            f"năm 2025: {pct(cptc.get(2025, 0), dt[2025])}."
        )
    else:
        dt_line = gv_line = qldn_line = bh_line = cptc_line = None

    for i, p in enumerate(doc.paragraphs):
        t = p.text
        if "d) Số chi nhánh, cơ sở phụ thuộc:" in t:
            base.replace_paragraph_text(
                p,
                f"d) Số chi nhánh, cơ sở phụ thuộc: {chi_nhanh or 'Không'}.",
            )
        elif dt_line and (
            "Doanh nghiệp kê khai lãi" in t
            or "Doanh nghiệp kê khai lỗ" in t
            or "Doanh nghiệp kê khai kết quả hoạt động kinh doanh" in t
        ):
            base.replace_paragraph_text(p, dt_line)
        elif gv_line and "Tỷ lệ giá vốn hàng bán trên doanh thu" in t:
            base.replace_paragraph_text(p, gv_line)
        elif qldn_line and (
            "Tỷ lệ Chi phí quản lý" in t
            or "Tỷ lệ chi phí quản lý" in t
        ) and "bán hàng" not in t and "tài chính" not in t:
            base.replace_paragraph_text(p, qldn_line)
        elif bh_line and "Tỷ lệ chi phí bán hàng trên doanh thu" in t:
            base.replace_paragraph_text(p, bh_line)
        elif cptc_line and "Tỷ lệ chi phí tài chính trên doanh thu" in t:
            base.replace_paragraph_text(p, cptc_line)
        elif "IV. ĐỀ XUẤT NỘI DUNG KIỂM TRA:" in t:
            dx_lines = [
                "- Chi phí quản lý kinh doanh, thời kỳ kiểm tra: Từ năm 2022 đến năm 2025.",
                "- Chi phí tài chính, thời kỳ kiểm tra: Từ năm 2022 đến năm 2025.",
            ]
            for j, line in enumerate(dx_lines):
                if i + 1 + j < len(doc.paragraphs):
                    base.replace_paragraph_text(doc.paragraphs[i + 1 + j], line)


def fix_draft_qd_numbers(doc: Document, so_qdkt: str, so_qdgs: str) -> int:
    changed = 0
    so_qdkt = resolve_so_qd(so_qdkt)
    so_qdgs = resolve_so_qd(so_qdgs)

    def fix_text(t: str, so_qd: str) -> str:
        nt = t
        nt = nt.replace("DU_THAO-TRINH-KY", so_qd)
        if so_qd == SO_QD_BLANK:
            nt = re.sub(r"(số)\s*/QĐ-TCS3", r"\1" + SO_QD_BLANK, nt, flags=re.IGNORECASE)
            nt = re.sub(r"(Số:)\s*/QĐ-TCS3", r"\1" + SO_QD_BLANK, nt)
            nt = re.sub(r"(SỐ:)\s*/QĐ-TCS3", r"\1" + SO_QD_BLANK, nt)
            nt = re.sub(r"(SỐ)\s+/QĐ-TCS3", r"\1" + SO_QD_BLANK, nt)
        nt = re.sub(r"\.{3,}/QĐ-TCS3", so_qd, nt)
        nt = re.sub(r"…+/QĐ-TCS3", so_qd, nt)
        if re.match(r"^\s*Số:\s*(DU_THAO-TRINH-KY|[.\s…]*)\s*$", t or "", flags=re.IGNORECASE):
            nt = f"Số:{so_qd}" if so_qd.startswith(" ") else f"Số: {so_qd}"
        if re.match(r"^\s*Số:\s*/QĐ-TCS3\s*$", nt or ""):
            nt = f"Số:{SO_QD_BLANK}"
        if re.match(r"^\s*Số:\s+", nt or "") and re.search(r"/QĐ-TCS3\s*$", nt or ""):
            if not re.search(r"Số:\s*\d+/QĐ-TCS3", nt or ""):
                nt = f"Số:{SO_QD_BLANK}"
        nt = nt.replace(f"..{so_qd}", so_qd).replace(f".{so_qd}", so_qd)
        return nt

    def apply_para(p, so_qd: str) -> None:
        nonlocal changed
        old = p.text
        nt = fix_text(old, so_qd)
        if nt != old:
            set_paragraph_text_preserve_spaces(p, nt)
            changed += 1

    for p in doc.paragraphs:
        so = so_qdgs if "giám sát" in p.text.lower() and "Quyết định kiểm tra" not in p.text else so_qdkt
        if "thành lập theo Quyết định số" in p.text or "theo Quyết định kiểm tra số" in p.text:
            so = so_qdkt
        apply_para(p, so)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    so = so_qdgs if "giám sát" in p.text.lower() else so_qdkt
                    if "thành lập theo Quyết định số" in p.text:
                        so = so_qdkt
                    apply_para(p, so)
    return changed


def _is_chuc_danh_line(text: str) -> bool:
    t = (text or "").strip()
    return any(m in t for m in CHUC_DANH_MARKERS)


def _is_signer_name_line(text: str) -> bool:
    t = (text or "").strip()
    if not t or _is_chuc_danh_line(t):
        return False
    if t in SIGNER_NAMES:
        return True
    if len(t) <= 40 and re.match(r"^[A-ZÀ-ỴĐ][\wÀ-ỹĐđ]*(?:[ \t]+[A-ZÀ-ỴĐ][\wÀ-ỹĐđ]*){1,4}$", t):
        return True
    return False


def ensure_signature_spacing(doc: Document, min_blank: int = SIGN_BLANK_LINES) -> int:
    fixed = 0

    def fix_paragraphs(get_paras) -> int:
        local = 0
        guard = 0
        while guard < 50:
            guard += 1
            paras = list(get_paras())
            texts = [p.text.strip() for p in paras]
            changed_here = False
            for i, t in enumerate(texts):
                if not _is_signer_name_line(t):
                    continue
                j = i - 1
                while j >= 0 and texts[j] == "":
                    j -= 1
                if j < 0 or not _is_chuc_danh_line(texts[j]):
                    continue
                block_start = j
                while block_start - 1 >= 0 and _is_chuc_danh_line(texts[block_start - 1]):
                    block_start -= 1
                blank_count = i - j - 1
                need = min_blank - blank_count
                for k in range(block_start, i):
                    set_keep_with_next(paras[k], True)
                set_keep_with_next(paras[i], False)
                if need > 0:
                    anchor = paras[i - 1] if i > 0 else paras[j]
                    for _ in range(need):
                        new_p = insert_paragraph_after(anchor)
                        set_keep_with_next(new_p, True)
                        anchor = new_p
                        local += 1
                    changed_here = True
                    break
            if not changed_here:
                break
        return local

    fixed += fix_paragraphs(lambda: doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                fixed += fix_paragraphs(lambda c=cell: c.paragraphs)
    return fixed


def ensure_times_new_roman(doc: Document) -> int:
    fixed = 0
    for p in doc.paragraphs:
        for r in p.runs:
            name = r.font.name or ""
            if "Calibri" in name or not name:
                base.set_run_font(r)
                fixed += 1
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        name = r.font.name or ""
                        if "Calibri" in name:
                            base.set_run_font(r)
                            fixed += 1
    return fixed


def remaining_placeholders(doc: Document) -> list[str]:
    texts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                texts.append(cell.text)
    return sorted(set(re.findall(r"«[^»]+»", "\n".join(texts))))


def run():
    print("=" * 72)
    print("XUẤT HỒ SƠ ĐI - T&T BIÊN HÒA (3603883505)")
    print("=" * 72)

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(TEMPLATE_PATH)
    if not EXCEL_SO.exists():
        raise FileNotFoundError(EXCEL_SO)
    if not BANG_TINH_PATH.exists():
        raise FileNotFoundError(BANG_TINH_PATH)

    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    tong_path = TARGET_DIR / "TONG HOP.docx"

    info = load_company_from_so()
    print(
        f"[so] hàng {info['row']}: {info['ten']} | {info['thoi_ky']} | "
        f"TĐ={info['td']} | TV={info['tv']}"
    )

    ratios = load_kqkd_ratios()
    reps = build_reps(info)
    base.TEMPLATE_PATH = str(TEMPLATE_PATH)

    doc = Document(str(TEMPLATE_PATH))
    print(f"[1] Mở phôi: {TEMPLATE_PATH.name} ({len(doc.paragraphs)} đoạn, {len(doc.tables)} bảng)")

    base.replace_in_doc(doc, reps)
    print("[2] Đã thay placeholder thông tin NNT/đoàn/QĐ")

    patch_noi_dung_and_thoi_han(doc, reps["Nội_dung_kiểm_tra"], reps["Thời_hạn_kiểm_tra_"])
    patch_di01_phan_tich(doc, ratios, info["chi_nhanh"], reps["Thời_kỳ_kiểm_tra"])
    print("[3] Đã cập nhật nội dung kiểm tra + DI-01 từ Bang tinh")

    base.apply_nd30_to_admin_docs_only(doc)
    qd_fixed = fix_draft_qd_numbers(doc, reps["QĐKT"], reps["QĐGS"])
    fixed = ensure_times_new_roman(doc)
    base.clean_doc_flow(doc)
    sign_fixed = ensure_signature_spacing(doc, SIGN_BLANK_LINES)
    print(
        f"[4] NĐ 30 (Tờ trình/QĐ), số QĐ {qd_fixed} chỗ, font ({fixed} run), "
        f"chữ ký +{sign_fixed} dòng trống"
    )
    left = remaining_placeholders(doc)
    if left:
        print(f"[canh bao] Còn placeholder: {left}")
    else:
        print("[ok] Không còn placeholder «...»")

    qd_fixed2 = fix_draft_qd_numbers(doc, reps["QĐKT"], reps["QĐGS"])
    sign_fixed2 = ensure_signature_spacing(doc, SIGN_BLANK_LINES)

    doc.save(str(tong_path))
    print(f"[5] Đã lưu: {tong_path}")

    for pattern in (
        "QD Kiem tra tai cty *.docx",
        "QDgiam sat tai cty *.docx",
        "QD giam sat tai cty *.docx",
        "hoso.zip",
        "TONG HOP.bak-*.docx",
    ):
        for p in TARGET_DIR.glob(pattern):
            p.unlink(missing_ok=True)
            print(f"[cleanup] removed {p.name}")

    print("=" * 72)
    print("HOÀN TẤT - chỉ giữ TONG HOP.docx (+ Bang tinh.xlsx)")
    print(f"Thư mục: {TARGET_DIR}")
    print(f" - {tong_path.name}: {tong_path.stat().st_size} bytes")
    print("=" * 72)


if __name__ == "__main__":
    run()
