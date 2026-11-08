# -*- coding: utf-8 -*-
"""Xuất hồ sơ đi kiểm tra theo kế hoạch - Công ty TNHH TM DV Tâm Hoàng Phát (MST 3603721832).

Nguồn dữ liệu:
- Phôi: templates/MAUTONG HOP.docx
- Sổ: Danh sach kiem tra quyet toan pro.xlsm (sheet 01_DA_BAN_HANH)
- Bảng tính phân tích (DI-01): Bang tinh.xlsx trong thư mục vụ

Quy cách xuất: TONG HOP.docx + 02 Quyết định độc lập + hoso.zip.
Chưa có Số QĐKT thật -> dùng mã tạm DU_THAO-TRINH-KY.
"""
from __future__ import annotations

import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import build_bien_ban_ho_so_ve_vpt as base

BRAIN_DIR = SCRIPT_DIR.parent
TEMPLATE_PATH = BRAIN_DIR / "templates" / "MAUTONG HOP.docx"
TARGET_DIR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD"
    r"\10. TAM HOANG PHAT\0. BIEN BAN"
)
BANG_TINH_PATH = TARGET_DIR / "Bang tinh.xlsx"
EXCEL_SO = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung"
    r"\Danh sach kiem tra quyet toan pro.xlsm"
)

MST = "3603721832"
# Theo sổ JAVIS PRO + quy tắc trình bày: không viết hoa toàn bộ tên công ty
COMPANY_NAME = "Công Ty TNHH Thương Mại Dịch Vụ Tâm Hoàng Phát"
COMPANY_UPPER = COMPANY_NAME  # giữ alias cũ cho chỗ gọi cũ
QD_FILE_STEM = "Tam Hoang Phat"
SO_QD_TMP = "DU_THAO-TRINH-KY"


def fmt_vnd(n: float | int) -> str:
    return f"{int(round(n)):,}".replace(",", ".")


def pct(part: float, whole: float) -> str:
    if not whole:
        return "0%"
    return f"{part / whole * 100:.2f}%".replace(".", ",")


def load_company_from_so() -> dict:
    import openpyxl

    wb = openpyxl.load_workbook(EXCEL_SO, data_only=True, keep_vba=False)
    ws = wb["01_DA_BAN_HANH"]
    row = None
    for r in range(5, ws.max_row + 1):
        if str(ws.cell(r, 4).value or "").strip() == MST:
            row = r
            break
    if row is None:
        raise RuntimeError(f"Không tìm thấy MST {MST} trên sheet 01_DA_BAN_HANH")

    def v(c):
        return ws.cell(row, c).value

    dia_chi = str(v(6) or "").strip()
    tk = str(v(21) or "").strip()
    if tk.startswith("+"):
        tk = tk.lstrip("+ ").strip()
    if tk.lower().startswith("tk số"):
        pass
    elif tk:
        # Chuẩn hóa nhẹ để khớp placeholder «Tài_khoản_số»
        tk = tk.replace("TK số", "").replace("Tài khoản số", "").strip()
        if tk and not tk.lower().startswith("số"):
            # giữ nguyên chuỗi sau khi bỏ tiền tố
            pass

    ngay_dkkd = v(19)
    if hasattr(ngay_dkkd, "strftime"):
        ngay_dkkd_s = ngay_dkkd.strftime("%d/%m/%Y")
    else:
        ngay_dkkd_s = str(ngay_dkkd or "").strip()

    return {
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
        "so_qdkt": str(v(29) or "").strip() if v(29) else "",
        "so_qdgs": str(v(30) or "").strip() if v(30) else "",
        "ngay_ban_hanh": v(31),
    }


def load_kqkd_ratios() -> dict:
    import openpyxl

    if not BANG_TINH_PATH.exists():
        return {}
    wb = openpyxl.load_workbook(BANG_TINH_PATH, data_only=True)
    ws = wb["KQKD"]
    headers = [ws.cell(2, c).value for c in range(1, 16)]
    year_cols = {}
    for c, h in enumerate(headers, start=1):
        if isinstance(h, int) and h in (2023, 2024, 2025):
            year_cols[h] = c

    def row_vals(label_part: str):
        out = {}
        for r in range(3, 20):
            name = str(ws.cell(r, 2).value or "")
            if label_part.lower() in name.lower():
                for y, c in year_cols.items():
                    out[y] = float(ws.cell(r, c).value or 0)
                return out
        return out

    return {
        "dt": row_vals("Doanh thu thuần") or row_vals("Doanh thu bán hàng"),
        "gv": row_vals("Giá vốn hàng bán"),
        "bh": row_vals("Chi phí bán hàng"),
        "qldn": row_vals("Chi phí quản lý"),
        "cptc": row_vals("Chi phí tài chính"),
        "ln": row_vals("Lợi nhuận từ hoạt động kinh doanh"),
    }


def build_reps(info: dict) -> dict:
    so_qdkt = info["so_qdkt"] or SO_QD_TMP
    so_qdgs = info["so_qdgs"] or SO_QD_TMP
    ngay_qd = "..../..../2026"
    if info["ngay_ban_hanh"] and hasattr(info["ngay_ban_hanh"], "strftime"):
        ngay_qd = info["ngay_ban_hanh"].strftime("%d/%m/%Y")
    elif isinstance(info["ngay_ban_hanh"], str) and info["ngay_ban_hanh"].strip():
        ngay_qd = info["ngay_ban_hanh"].strip()

    ngay_cap = "..../..../...."
    if info["ngay_cap_cccd"] and hasattr(info["ngay_cap_cccd"], "strftime"):
        ngay_cap = info["ngay_cap_cccd"].strftime("%d/%m/%Y")

    dia_chi_full = info["dia_chi"]
    if "Đồng Nai" not in dia_chi_full:
        dia_chi_full = f"{info['dia_chi']}, Thành phố Đồng Nai"

    # CQT ngắn gọn cho placeholder
    cqt = "Thuế Cơ sở 3 Thành phố Đồng Nai"
    if "Thuế cơ sở 3" in (info["dia_ban"] or ""):
        cqt = "Thuế Cơ sở 3 Thành phố Đồng Nai"

    tk = info["tai_khoan"]
    # Bỏ tiền tố dư nếu còn
    tk = re.sub(r"^\+\s*", "", tk)
    tk = re.sub(r"^(TK|Tài khoản)\s*số\s*", "", tk, flags=re.IGNORECASE).strip()

    return {
        "Tên_công_ty": COMPANY_NAME,
        "MST": MST,
        "địa_chỉ": info["dia_chi"],
        "ĐỊA_CHỈ": dia_chi_full,
        "NNKD": info["nnkd"] or "Lắp đặt hệ thống điện",
        "GĐ": info["gd"] or ".........",
        "KT": info["kt"] or ".........",
        "CCCD": info["cccd"] or "................",
        "ngày_cấp": ngay_cap,
        "Ngày_cấp2": ngay_cap,
        "Nơi_cấp": info["noi_cap"] or "Cục Cảnh sát Quản lý hành chính về trật tự xã hội",
        "GGCNĐKKD": info["gcn"] or MST,
        "CCCDKT": info["cccd_kt"] or "................",
        "Chi_nhánh": info["chi_nhanh"] or "Không",
        "Tài_khoản_số": tk or "................",
        "Thời_kỳ_kiểm_tra": info["thoi_ky"] or "Từ năm 2023 đến năm 2025",
        "Thời_hạn_kiểm_tra_": "20",
        "TĐ": info["td"] or "Nguyễn Thị Ánh Ngọc",
        "CVTĐ": info["cv_td"] or "Chuyên viên",
        "TV": info["tv"] or "Nguyễn Trần Thu Hà",
        "CVTV": info["cv_tv"] or "Kiểm tra viên thuế",
        "Ma_CQT": info["ma_cqt"] or "1140967",
        "CQT": cqt,
        "QĐKT": so_qdkt,
        "QĐGS": so_qdgs,
        "ngày_QĐ": ngay_qd,
        "NGÀY_QĐ": ngay_qd,
        "Ngày_công_bố": "..../..../2026",
        "Từ_ngày": "..../..../2026",
        "Đến_ngày": "..../..../2026",
        "Phân_công_trưởng_đoàn": "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng.",
        "Phân_công_thành_viên": "- Chi phí tài chính.",
        "Ngày_ký_Biên_bản_kiểm_tra": "..../..../2026",
        "Ngày_báo_cáo": "..../..../2026",
        # Hồ sơ đi: chưa có số liệu truy thu/phạt chính thức
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


def patch_noi_dung_and_thoi_han(doc: Document):
    for p in doc.paragraphs:
        if "«Nội_dung_kiểm_tra»" in p.text or p.text.strip() == "«Nội_dung_kiểm_tra»":
            base.replace_paragraph_text(
                p,
                "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng.\n- Chi phí tài chính.",
            )
        elif "Thời hạn kiểm tra là" in p.text and "ngày làm việc" in p.text:
            base.replace_paragraph_text(
                p,
                "Thời hạn kiểm tra là 20 ngày làm việc "
                "(không bao gồm ngày nghỉ, ngày lễ theo quy định của pháp luật) "
                "kể từ ngày công bố Quyết định kiểm tra.",
            )


def patch_di01_phan_tich(doc: Document, ratios: dict, chi_nhanh: str):
    """Cập nhật phần phân tích DI-01 theo số liệu Bang tinh (2023-2025)."""
    dt = ratios.get("dt") or {}
    gv = ratios.get("gv") or {}
    bh = ratios.get("bh") or {}
    qldn = ratios.get("qldn") or {}
    cptc = ratios.get("cptc") or {}

    years = [2023, 2024, 2025]
    if all(y in dt for y in years):
        dt_line = (
            "- Doanh nghiệp kê khai kết quả hoạt động kinh doanh giai đoạn 2023-2025, "
            f"doanh thu thuần năm 2023: {fmt_vnd(dt[2023])} đ, "
            f"năm 2024: {fmt_vnd(dt[2024])} đ, "
            f"năm 2025: {fmt_vnd(dt[2025])} đ."
        )
        gv_line = (
            "- Tỷ lệ giá vốn hàng bán trên doanh thu thuần: "
            f"năm 2023: {pct(gv.get(2023, 0), dt[2023])}, "
            f"năm 2024: {pct(gv.get(2024, 0), dt[2024])}, "
            f"năm 2025: {pct(gv.get(2025, 0), dt[2025])}."
        )
        qldn_line = (
            "- Tỷ lệ chi phí quản lý doanh nghiệp trên doanh thu thuần: "
            f"năm 2023: {pct(qldn.get(2023, 0), dt[2023])}, "
            f"năm 2024: {pct(qldn.get(2024, 0), dt[2024])}, "
            f"năm 2025: {pct(qldn.get(2025, 0), dt[2025])}."
        )
        bh_line = (
            "- Tỷ lệ chi phí bán hàng trên doanh thu thuần: "
            f"năm 2023: {pct(bh.get(2023, 0), dt[2023])}, "
            f"năm 2024: {pct(bh.get(2024, 0), dt[2024])}, "
            f"năm 2025: {pct(bh.get(2025, 0), dt[2025])}."
        )
        cptc_line = (
            "- Tỷ lệ chi phí tài chính trên doanh thu thuần: "
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
            or "Doanh nghiệp kê khai kết quả hoạt động kinh doanh" in t
        ):
            base.replace_paragraph_text(p, dt_line)
        elif gv_line and "Tỷ lệ giá vốn hàng bán trên doanh thu" in t:
            base.replace_paragraph_text(p, gv_line)
        elif qldn_line and (
            "Tỷ lệ Chi phí quản lý" in t or "Tỷ lệ chi phí quản lý doanh nghiệp" in t
        ) and "bán hàng" not in t and "tài chính" not in t:
            base.replace_paragraph_text(p, qldn_line)
        elif bh_line and "Tỷ lệ chi phí bán hàng trên doanh thu" in t:
            base.replace_paragraph_text(p, bh_line)
        elif cptc_line and "Tỷ lệ chi phí tài chính trên doanh thu" in t:
            base.replace_paragraph_text(p, cptc_line)
        elif "IV. ĐỀ XUẤT NỘI DUNG KIỂM TRA:" in t:
            dx_lines = [
                "- Chi phí quản lý doanh nghiệp, thời kỳ kiểm tra: Từ năm 2023 đến năm 2025.",
                "- Chi phí bán hàng, thời kỳ kiểm tra: Từ năm 2023 đến năm 2025.",
                "- Chi phí tài chính, thời kỳ kiểm tra: Từ năm 2023 đến năm 2025.",
            ]
            for j, line in enumerate(dx_lines):
                if i + 1 + j < len(doc.paragraphs):
                    base.replace_paragraph_text(doc.paragraphs[i + 1 + j], line)


def fix_draft_qd_numbers(doc: Document, so_qd: str = SO_QD_TMP) -> int:
    """Phôi QĐ dùng chuỗi chấm (.............../QĐ-TCS3) thay vì «QĐKT» - điền mã tạm/số thật."""
    changed = 0

    def fix_text(t: str) -> str:
        nt = t
        nt = re.sub(r"\.{3,}/QĐ-TCS3", so_qd, nt)
        nt = re.sub(r"…+/QĐ-TCS3", so_qd, nt)
        nt = nt.replace(f"..{so_qd}", so_qd).replace(f".{so_qd}", so_qd)
        if re.match(r"^\s*Số:\s*[\.\s]*$", nt or ""):
            nt = f"Số: {so_qd}"
        return nt

    for p in doc.paragraphs:
        nt = fix_text(p.text)
        if nt != p.text:
            base.replace_paragraph_text(p, nt)
            changed += 1
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    nt = fix_text(p.text)
                    if nt != p.text:
                        base.replace_paragraph_text(p, nt)
                        changed += 1
    return changed


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
    print("XUẤT HỒ SƠ ĐI - TÂM HOÀNG PHÁT (3603721832)")
    print("=" * 72)

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(TEMPLATE_PATH)
    if not EXCEL_SO.exists():
        raise FileNotFoundError(EXCEL_SO)

    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    tong_path = TARGET_DIR / "TONG HOP.docx"
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    if tong_path.exists():
        bak = TARGET_DIR / f"TONG HOP.bak-before-hoso-di-{ts}.docx"
        shutil.copy2(tong_path, bak)
        print(f"[backup] {bak.name}")

    info = load_company_from_so()
    print(f"[so] {info['ten']} | {info['thoi_ky']} | TĐ={info['td']} | TV={info['tv']}")
    print(f"[so] QĐKT={info['so_qdkt'] or SO_QD_TMP} (mã tạm vì chưa có số thật)" if not info["so_qdkt"] else f"[so] QĐKT={info['so_qdkt']}")

    ratios = load_kqkd_ratios()
    reps = build_reps(info)

    # Patch TEMPLATE_PATH trong module base để QD builder mở đúng phôi
    base.TEMPLATE_PATH = str(TEMPLATE_PATH)

    doc = Document(str(TEMPLATE_PATH))
    print(f"[1] Mở phôi: {TEMPLATE_PATH.name} ({len(doc.paragraphs)} đoạn, {len(doc.tables)} bảng)")

    base.replace_in_doc(doc, reps)
    print("[2] Đã thay placeholder thông tin NNT/đoàn/QĐ")

    patch_noi_dung_and_thoi_han(doc)
    patch_di01_phan_tich(doc, ratios, info["chi_nhanh"])
    print("[3] Đã cập nhật nội dung kiểm tra, thời hạn 20 ngày, DI-01")

    base.apply_nd30_to_admin_docs_only(doc)
    qd_fixed = fix_draft_qd_numbers(doc, reps["QĐKT"])
    fixed = ensure_times_new_roman(doc)
    base.clean_doc_flow(doc)
    print(f"[4] Đã chuẩn hóa NĐ 30 (Tờ trình/QĐ), điền {qd_fixed} chỗ số QĐ dạng chấm, font ({fixed} run)")

    left = remaining_placeholders(doc)
    if left:
        print(f"[canh bao] Còn placeholder: {left}")
    else:
        print("[ok] Không còn placeholder «...»")

    doc.save(str(tong_path))
    print(f"[5] Đã lưu: {tong_path}")

    qdkt = TARGET_DIR / f"QD Kiem tra tai cty {QD_FILE_STEM}.docx"
    qdgs = TARGET_DIR / f"QDgiam sat tai cty {QD_FILE_STEM}.docx"
    base.build_single_page_qd_kiem_tra(doc, str(qdkt))
    base.build_single_page_qd_giam_sat(doc, str(qdgs))

    zip_path = TARGET_DIR / "hoso.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in (tong_path, qdkt, qdgs):
            zf.write(p, arcname=p.name)
            print(f"  + ZIP: {p.name}")

    print("=" * 72)
    print("HOÀN TẤT")
    print(f"Thư mục: {TARGET_DIR}")
    for p in (tong_path, qdkt, qdgs, zip_path):
        print(f" - {p.name}: {p.stat().st_size} bytes")
    print("=" * 72)


if __name__ == "__main__":
    run()
