#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""So sánh chi tiết 3 file TONG HOP."""

from pathlib import Path
from docx import Document
import re

files = {
    "Kiến Thành Tân (3603216935)": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/KIEN THANH TAN_3603216935/0. BIEN BAN/TONG HOP.docx",
    "Thúy Hồng (3603437726)": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/1. TD/8. THUY HONG/THUY_HONG_3603437726/0. BIEN BAN/TONG HOP.docx",
    "Sơn Tiền Hải": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/1. TD/11. SON TIEN HAI/0. BIEN BAN/TONG HOP.docx"
}

def extract_text(doc_path, max_lines=50):
    """Trích xuất nội dung từ document."""
    try:
        doc = Document(doc_path)

        info = {
            "paragraphs": len(doc.paragraphs),
            "tables": len(doc.tables),
            "sections": len(doc.sections),
            "content": [],
            "placeholders": [],
            "has_muc_iii": False,
        }

        # Trích xuất nội dung từng đoạn
        text_content = []
        for para in doc.paragraphs[:max_lines]:
            if para.text.strip():
                text_content.append(para.text.strip())
        info["content"] = text_content

        # Tìm placeholder (text trong « »)
        all_text = "\n".join([p.text for p in doc.paragraphs])
        placeholders = re.findall(r'«([^»]+)»', all_text)
        info["placeholders"] = list(set(placeholders))

        # Kiểm tra Mục III
        if "Mục III" in all_text or "NSNN" in all_text:
            info["has_muc_iii"] = True

        # Kiểm tra các section quan trọng
        sections_keywords = {
            "Quyết định kiểm tra": "QĐKT",
            "Biên bản kiểm tra": "BBKT",
            "Biên bản xác nhận": "BBXN",
            "Cam kết": "Cam kết",
            "Quyết định xử phạt": "QĐXP",
            "Hoàn thuế": "Hoàn thuế",
        }

        info["sections_found"] = {}
        for keyword, short in sections_keywords.items():
            info["sections_found"][short] = keyword in all_text

        return info
    except Exception as e:
        return {"error": str(e)}

print("=" * 100)
print("SO SÁNH 3 FILE TONG HOP GIẢI THỂ")
print("=" * 100)

data = {}
for name, path in files.items():
    if Path(path).exists():
        data[name] = extract_text(path)
        info = data[name]

        print(f"\n{name}")
        print("-" * 100)
        print(f"  Cấu trúc:")
        print(f"    - Số đoạn: {info['paragraphs']}")
        print(f"    - Số bảng: {info['tables']}")
        print(f"    - Số section: {info['sections']}")
        print(f"    - Có Mục III NSNN: {'✓ Có' if info['has_muc_iii'] else '✗ Không'}")

        print(f"\n  Các section chính:")
        for short, found in info['sections_found'].items():
            status = "✓" if found else "✗"
            print(f"    {status} {short}")

        if info['placeholders']:
            print(f"\n  Placeholder chưa điền:")
            for ph in sorted(info['placeholders'])[:10]:
                print(f"    • {ph}")

        print(f"\n  Nội dung đầu tiên (5 dòng):")
        for line in info['content'][:5]:
            print(f"    • {line[:70]}...")
    else:
        print(f"\n{name}: FILE KHÔNG TỒN TẠI")

# So sánh
print("\n" + "=" * 100)
print("PHÂN TÍCH KHÁC NHAU")
print("=" * 100)

if len(data) >= 2:
    names = list(data.keys())

    # So sánh số đoạn
    para_count = {n: data[n].get('paragraphs', 0) for n in names}
    min_para = min(para_count.values())
    max_para = max(para_count.values())
    if min_para != max_para:
        print(f"\n✗ SỐ ĐOẠN KHÁC NHAU:")
        for n, c in sorted(para_count.items(), key=lambda x: x[1]):
            diff = c - min_para
            print(f"  {n}: {c} (+{diff})")

    # So sánh số bảng
    table_count = {n: data[n].get('tables', 0) for n in names}
    min_table = min(table_count.values())
    max_table = max(table_count.values())
    if min_table != max_table:
        print(f"\n✗ SỐ BẢNG KHÁC NHAU:")
        for n, c in sorted(table_count.items(), key=lambda x: x[1]):
            diff = c - min_table
            print(f"  {n}: {c} (+{diff})")

    # So sánh placeholder
    all_placeholders = set()
    for n in names:
        all_placeholders.update(data[n].get('placeholders', []))

    if all_placeholders:
        print(f"\n✗ PLACEHOLDER CHƯA ĐIỀN (Có {len(all_placeholders)} loại):")
        for ph in sorted(all_placeholders)[:15]:
            files_with_ph = [n for n in names if ph in data[n].get('placeholders', [])]
            print(f"  • «{ph}» ← {', '.join([f.split('(')[0] for f in files_with_ph])}")

    # So sánh section
    print(f"\n✓ CẤU TRÚC SECTION (Các phần chính):")
    sections = list(data[names[0]]['sections_found'].keys())
    for section in sections:
        status = all(data[n]['sections_found'].get(section, False) for n in names)
        symbol = "✓" if status else "✗"
        print(f"  {symbol} {section}")

print("\n" + "=" * 100)
