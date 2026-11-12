#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""So sánh cấu trúc và nội dung file TONG HOP.docx."""

from pathlib import Path
from docx import Document

files = {
    "Kiến Thành Tân": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/KIEN THANH TAN_3603216935/0. BIEN BAN/TONG HOP.docx",
    "Phú Lương": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/Phu Lương-GT-CTY CHUA CO KHA NANG DIEU CHINH/BIEN BAN/TONG HOP.docx",
    "Tài Dùng": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/TAI DUNG-GT/BIEN BAN/TONG HOP.docx"
}

def analyze_document(doc_path):
    """Phân tích cấu trúc document."""
    try:
        doc = Document(doc_path)

        analysis = {
            "paragraphs": len(doc.paragraphs),
            "tables": len(doc.tables),
            "sections": len(doc.sections),
        }

        # Lấy văn bản đầu tiên từ 5 đoạn đầu
        first_paras = []
        for i, para in enumerate(doc.paragraphs[:5]):
            if para.text.strip():
                first_paras.append(para.text[:80])

        analysis["first_content"] = first_paras

        # Kiểm tra bảng
        if len(doc.tables) > 0:
            first_table = doc.tables[0]
            analysis["first_table_rows"] = len(first_table.rows)
            analysis["first_table_cols"] = len(first_table.columns)

        return analysis
    except Exception as e:
        return {"error": str(e)}

print("=" * 80)
print("SO SÁNH CẤU TRÚC FILE TONG HOP.DOCX")
print("=" * 80)

data = {}
for name, path in files.items():
    if Path(path).exists():
        data[name] = analyze_document(path)
        print(f"\n{name}:")
        info = data[name]
        print(f"  - Số đoạn: {info.get('paragraphs', 'N/A')}")
        print(f"  - Số bảng: {info.get('tables', 'N/A')}")
        print(f"  - Số section: {info.get('sections', 'N/A')}")
        if "first_table_rows" in info:
            print(f"  - Bảng 1: {info['first_table_rows']} dòng × {info['first_table_cols']} cột")
        print(f"  - Nội dung đầu tiên:")
        for content in info.get('first_content', [])[:3]:
            print(f"    • {content}...")
    else:
        print(f"\n{name}: FILE KHÔNG TỒN TẠI")

# So sánh
print("\n" + "=" * 80)
print("NHẬN XÉT KHÁC NHAU")
print("=" * 80)

if "Kiến Thành Tân" in data and "Phú Lương" in data:
    kt = data["Kiến Thành Tân"]
    pl = data["Phú Lương"]

    if kt.get("paragraphs") != pl.get("paragraphs"):
        print(f"✗ Số đoạn khác: Kiến Thành Tân {kt.get('paragraphs')} vs Phú Lương {pl.get('paragraphs')}")

    if kt.get("tables") != pl.get("tables"):
        print(f"✗ Số bảng khác: Kiến Thành Tân {kt.get('tables')} vs Phú Lương {pl.get('tables')}")

print("\nGhi chú: So sánh chi tiết cần mở file bằng Word để xem layout, định dạng, placeholder.")
