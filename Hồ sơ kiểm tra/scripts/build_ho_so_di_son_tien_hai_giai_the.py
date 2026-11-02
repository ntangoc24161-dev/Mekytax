# -*- coding: utf-8 -*-
"""Tạo lại TONG HOP hồ sơ đi giải thể cho Sơn Tiền Hải từ mẫu DNTN Thúy Hồng.

- Phôi nguồn: TONG HOP.docx của Thúy Hồng (mẫu kiểm tra quyết toán giải thể).
- Giữ khung/nội dung kiểm tra kiểu giải thể (GTGT, doanh thu, giá vốn, phân công như mẫu).
- Thay nhận diện NNT + thời kỳ + phân tích DI-01 theo Bang tinh Sơn Tiền Hải.
- Chưa có số QĐ: `Số:              /QĐ-TCS3` (không DU_THAO-TRINH-KY).
- Chỉ xuất TONG HOP.docx (không QĐ lẻ / zip).
"""
from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

from docx import Document

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import build_ho_so_di_son_tien_hai as sth
import build_bien_ban_ho_so_ve_vpt as base

SRC_TONG_HOP = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\8. THUY HONG"
    r"\THUY_HONG_3603437726\0. BIEN BAN\TONG HOP.docx"
)
TARGET_DIR = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\11. SON TIEN HAI\0. BIEN BAN"
)
OUT_PATH = TARGET_DIR / "TONG HOP.docx"

OLD_NAME = "Doanh Nghiệp Tư Nhân Vàng Bạc Đá Quí Thúy Hồng"
NEW_NAME = "Công ty TNHH Thương mại Sản xuất Sơn Tiền Hải"
OLD_MST = "3603437726"
NEW_MST = "3603859012"
OLD_GD = "Lê Thị Tuyết"
NEW_GD = "Phan Thế Hanh"
OLD_ADDR_CORE = "Số 01, khu phố 1, Phường Long Hưng"
NEW_ADDR_CORE = "Số 23, Tổ 1, KP 3A, Tỉnh Đồng Nai"
THOI_KY = "Từ năm 2022 đến tháng 10/2026"
THOI_KY_NGAN = "năm 2022 đến tháng 10/2026"

# Số liệu kết quả riêng của Thúy Hồng - không mang sang vụ mới.
TH_AMOUNT_PATTERNS = [
    r"313\.258\.128",
    r"1\.331\.235",
    r"299\.945\.778",
    r"15\.455\.000",
    r"2\.720\.610",
    r"4\.051\.845",
    r"13\.312\.349\.637\.700\.000",
    r"32\.428\.883\.717\.437\.000",
    r"79\.874\.099\.275",
]


def iter_paragraphs(doc: Document):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def replace_in_paragraph(p, mapping: list[tuple[str, str]]) -> bool:
    text = p.text
    if not text:
        return False
    new = text
    for old, rep in mapping:
        if old in new:
            new = new.replace(old, rep)
    if new == text:
        return False
    # Giữ khoảng trắng số QĐ nếu có.
    if "              /QĐ-TCS3" in new or new.strip().startswith("Số:"):
        sth.set_paragraph_text_preserve_spaces(p, new)
    else:
        base.replace_paragraph_text(p, new)
    return True


def build_identity_mapping(info: dict) -> list[tuple[str, str]]:
    dia_chi = info.get("dia_chi") or NEW_ADDR_CORE
    gd = info.get("gd") or NEW_GD
    pairs = [
        (OLD_NAME, NEW_NAME),
        ("Doanh nghiệp Tư Nhân Vàng Bạc Đá Quí Thúy Hồng", NEW_NAME),
        ("DOANH NGHIỆP TƯ NHÂN VÀNG BẠC ĐÁ QUÍ THÚY HỒNG", NEW_NAME.upper()),
        (OLD_MST, NEW_MST),
        (OLD_GD, gd),
        (OLD_ADDR_CORE, dia_chi),
        ("Từ năm 2017 đến năm 2024", THOI_KY),
        ("từ năm 2017 đến năm 2024", "từ năm 2022 đến tháng 10/2026"),
        ("năm 2017 – tháng 07/2024", THOI_KY_NGAN),
        ("năm 2017 - tháng 07/2024", THOI_KY_NGAN),
        ("năm 2017 – tháng 07/2024.", f"{THOI_KY_NGAN}."),
        ("2017 – tháng 07/2024", THOI_KY_NGAN),
        ("05.01.2017", info.get("ngay_dkkd") or "..../..../........"),
        ("6939/QĐ-TCS3", "              /QĐ-TCS3"),
        ("6940/QĐ-TCS3", "              /QĐ-TCS3"),
        ("Quyết định số: 6939/QĐ-TCS3", "Quyết định số:              /QĐ-TCS3"),
        ("Quyết định số 6939/QĐ-TCS3", "Quyết định số              /QĐ-TCS3"),
        ("doanh nghiệp số 6939/QĐ-TCS3", "doanh nghiệp số              /QĐ-TCS3"),
        ("số 6939/QĐ-TCS3", "số              /QĐ-TCS3"),
        ("số 6940/QĐ-TCS3", "số              /QĐ-TCS3"),
        # Ngày ban hành/công bố riêng của Thúy Hồng - để trống cho hồ sơ đi mới.
        ("ngày 19/10/2026", "ngày ..../..../2026"),
        ("ngày 23/10/2026", "ngày ..../..../2026"),
        ("ngày 20/10/2026", "ngày ..../..../2026"),
        ("đến ngày 22/10/2026", "đến ngày ..../..../2026"),
        ("từ ngày 20/10/2026", "từ ngày ..../..../2026"),
        ("19/10/2026", "..../..../2026"),
        ("23/10/2026", "..../..../2026"),
        ("20/10/2026", "..../..../2026"),
        ("22/10/2026", "..../..../2026"),
    ]
    # Dài hơn trước để tránh thay chuỗi con làm hỏng chuỗi dài.
    pairs.sort(key=lambda x: len(x[0]), reverse=True)
    return pairs


def scrub_thuy_hong_amounts(doc: Document) -> int:
    changed = 0
    for p in iter_paragraphs(doc):
        t = p.text
        if not t:
            continue
        nt = t
        for pat in TH_AMOUNT_PATTERNS:
            nt = re.sub(pat, "......", nt)
        # Câu giải trình mang dấu ấn kết quả Thúy Hồng.
        if "chưa xử lý khi giải thể" in nt and "Đoàn kiểm tra xác định" in nt:
            nt = re.sub(
                r"\*/[^\n]*chưa xử lý khi giải thể[^\n]*",
                "*/ .................................................................................................",
                nt,
            )
        if nt != t:
            base.replace_paragraph_text(p, nt)
            changed += 1
    return changed


def patch_de_xuat_and_period(doc: Document) -> int:
    """Giữ 3 nội dung kiểm tra kiểu QĐ giải thể Thúy Hồng, đổi sang kỳ Sơn Tiền Hải."""
    changed = 0
    dx = [
        f"- Thuế GTGT đầu ra, đầu vào được khấu trừ {THOI_KY_NGAN}.",
        f"- Doanh thu bán hàng và cung cấp dịch vụ từ {THOI_KY_NGAN}.",
        f"- Giá vốn hàng bán từ {THOI_KY_NGAN}.",
    ]
    paras = list(doc.paragraphs)
    for i, p in enumerate(paras):
        if "IV. ĐỀ XUẤT NỘI DUNG KIỂM TRA" in p.text:
            for j, line in enumerate(dx):
                if i + 1 + j < len(paras):
                    base.replace_paragraph_text(paras[i + 1 + j], line)
                    changed += 1
            break
    for p in paras:
        if p.text.strip() == "Thời kỳ kiểm tra: Từ năm 2022 đến tháng 10/2026.":
            continue
        if p.text.strip().startswith("Thời kỳ kiểm tra:"):
            base.replace_paragraph_text(p, f"Thời kỳ kiểm tra: {THOI_KY}.")
            changed += 1
    return changed


def patch_di01_from_bang_tinh(doc: Document, ratios: dict) -> int:
    """Cập nhật các dòng lãi/lỗ và tỷ lệ DI-01 theo Bang tinh Sơn Tiền Hải (2022-2025)."""
    if not ratios or not ratios.get("dt"):
        return 0
    dt, gv, qldn, ln = ratios["dt"], ratios["gv"], ratios["qldn"], ratios["ln"]
    years = (2022, 2023, 2024, 2025)
    if not all(y in dt for y in years):
        return 0

    lai, lo = [], []
    for y in years:
        val = ln.get(y, 0) or 0
        if val > 0:
            lai.append(str(y))
        elif val < 0:
            lo.append(str(y))
    parts = []
    if lo:
        parts.append(f"lỗ năm {', '.join(lo)}")
    if lai:
        parts.append(f"lãi năm {', '.join(lai)}")
    lai_lo = "; ".join(parts) if parts else "có phát sinh kết quả hoạt động kinh doanh"
    dt_line = (
        f"- Doanh nghiệp kê khai {lai_lo}; "
        f"doanh thu thuần năm 2022: {sth.fmt_vnd(dt[2022])} đ, "
        f"năm 2023: {sth.fmt_vnd(dt[2023])} đ, "
        f"năm 2024: {sth.fmt_vnd(dt[2024])} đ, "
        f"năm 2025: {sth.fmt_vnd(dt[2025])} đ."
    )
    gv_line = (
        "Tỷ lệ giá vốn hàng bán trên doanh thu "
        f"năm 2022 là {sth.pct(gv.get(2022, 0), dt[2022])}, "
        f"năm 2023 là {sth.pct(gv.get(2023, 0), dt[2023])}, "
        f"năm 2024 là {sth.pct(gv.get(2024, 0), dt[2024])}, "
        f"năm 2025 là {sth.pct(gv.get(2025, 0), dt[2025])}."
    )
    qldn_line = (
        "Tỷ lệ chi phí quản lý kinh doanh trên doanh thu "
        f"năm 2022 là {sth.pct(qldn.get(2022, 0), dt[2022])}, "
        f"năm 2023 là {sth.pct(qldn.get(2023, 0), dt[2023])}, "
        f"năm 2024 là {sth.pct(qldn.get(2024, 0), dt[2024])}, "
        f"năm 2025 là {sth.pct(qldn.get(2025, 0), dt[2025])}."
    )
    note_line = (
        "- Tỷ lệ giá vốn hàng bán, Chi phí quản lý kinh doanh trên doanh thu "
        "biến động không đều qua các năm, không tương ứng biến động doanh thu."
    )

    changed = 0
    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith("- Doanh nghiệp kê khai"):
            base.replace_paragraph_text(p, dt_line)
            changed += 1
        elif t.startswith("Tỷ lệ giá vốn hàng bán trên doanh thu"):
            base.replace_paragraph_text(p, gv_line)
            changed += 1
        elif t.startswith("Tỷ lệ chi phí quản lý kinh doanh trên doanh thu"):
            base.replace_paragraph_text(p, qldn_line)
            changed += 1
        elif "biến động không đều qua các năm" in t and "giá vốn" in t.lower():
            base.replace_paragraph_text(p, note_line)
            changed += 1
    return changed


def run():
    print("=" * 72)
    print("TẠO TONG HOP GIẢI THỂ - SƠN TIỀN HẢI từ mẫu THÚY HỒNG")
    print("=" * 72)
    if not SRC_TONG_HOP.exists():
        raise FileNotFoundError(SRC_TONG_HOP)
    TARGET_DIR.mkdir(parents=True, exist_ok=True)

    info = sth.load_company_from_so()
    print(f"[so] {info['ten']} | {info['thoi_ky']} | GT row={info['row']}")
    print(f"[so] TĐ={info['td']} | TV={info['tv']}")
    print(f"[so] Phân công TĐ: {info['phan_cong_td']!r}")
    print(f"[so] Phân công TV: {info['phan_cong_tv']!r}")

    ratios = sth.load_kqkd_ratios()
    print(f"[bangtinh] years={ratios.get('year_cols')}")

    # Copy phôi rồi mở (tránh sửa nhầm file Thúy Hồng).
    shutil.copy2(SRC_TONG_HOP, OUT_PATH)
    doc = Document(str(OUT_PATH))
    print(f"[1] Phôi Thúy Hồng: {len(doc.paragraphs)} đoạn, {len(doc.tables)} bảng")

    mapping = build_identity_mapping(info)
    n_id = 0
    for p in iter_paragraphs(doc):
        if replace_in_paragraph(p, mapping):
            n_id += 1
    print(f"[2] Đã thay nhận diện/thời kỳ/QĐ cũ: {n_id} đoạn")

    n_amt = scrub_thuy_hong_amounts(doc)
    print(f"[3] Đã gỡ số liệu kết quả riêng Thúy Hồng: {n_amt} đoạn")

    n_dx = patch_de_xuat_and_period(doc)
    print(f"[4] Đã giữ nội dung kiểm tra giải thể, đổi kỳ: {n_dx} đoạn")

    n_di = patch_di01_from_bang_tinh(doc, ratios)
    print(f"[5] Đã cập nhật phân tích DI-01 từ Bang tinh: {n_di} đoạn")

    # Chuẩn số trống + chữ ký.
    qd_fixed = sth.fix_draft_qd_numbers(doc, "", "")
    base.clean_doc_flow(doc)
    sign_fixed = sth.ensure_signature_spacing(doc, sth.SIGN_BLANK_LINES)
    qd_fixed2 = sth.fix_draft_qd_numbers(doc, "", "")
    print(f"[6] Số QĐ trống +{qd_fixed}+{qd_fixed2}, chữ ký +{sign_fixed}")

    doc.save(str(OUT_PATH))

    # Dọn file lẻ.
    for pattern in (
        "QD Kiem tra tai cty *.docx",
        "QDgiam sat tai cty *.docx",
        "QD giam sat tai cty *.docx",
        "hoso.zip",
        "TONG HOP.bak-*.docx",
    ):
        for p in TARGET_DIR.glob(pattern):
            p.unlink(missing_ok=True)
            print(f"[cleanup] {p.name}")

    # Kiểm nhanh.
    doc2 = Document(str(OUT_PATH))
    joined = "\n".join(p.text for p in doc2.paragraphs)
    print("[check] còn Thúy Hồng?", "Thúy Hồng" in joined)
    print("[check] còn MST cũ?", OLD_MST in joined)
    print("[check] còn 6939?", "6939" in joined)
    print("[check] có Sơn Tiền Hải?", "Sơn Tiền Hải" in joined)
    print("[check] có giải thể?", "giải thể" in joined.lower())
    print("[check] So blank sample:", "Số:              /QĐ-TCS3" in joined or any(
        "Số:              /QĐ-TCS3" in c.text for t in doc2.tables for r in t.rows for c in r.cells
    ))
    print("=" * 72)
    print(f"ĐÃ TẠO: {OUT_PATH}")
    print(f"Size: {OUT_PATH.stat().st_size} bytes")
    print("=" * 72)


if __name__ == "__main__":
    run()
