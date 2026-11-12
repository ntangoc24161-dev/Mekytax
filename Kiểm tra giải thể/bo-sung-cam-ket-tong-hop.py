#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1. Bổ sung phần Cam kết từ chối nộp thừa vào file Kiến Thành Tân
2. So sánh chi tiết nội dung QĐKT, QĐGS, Tờ trình giữa 3 công ty
"""

from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
import re

files = {
    "Kiến Thành Tân": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/KIEN THANH TAN_3603216935/0. BIEN BAN/TONG HOP.docx",
    "Thúy Hồng": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/1. TD/8. THUY HONG/THUY_HONG_3603437726/0. BIEN BAN/TONG HOP.docx",
    "Sơn Tiền Hải": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/1. TD/11. SON TIEN HAI/0. BIEN BAN/TONG HOP.docx"
}

def extract_section_content(doc_path, section_keyword):
    """Trích xuất nội dung một section từ document."""
    try:
        doc = Document(doc_path)

        content = []
        capture = False
        next_section_keywords = ["QĐKT", "QĐGS", "Biên bản", "Cam kết", "Hoàn thuế", "Tờ trình"]

        for para in doc.paragraphs:
            text = para.text.strip()

            # Kiểm tra bắt đầu section
            if section_keyword.lower() in text.lower() and len(text) < 200:
                capture = True
                content.append(("HEADER", text))
                continue

            # Kiểm tra kết thúc section
            if capture and any(kw.lower() in text.lower() for kw in next_section_keywords if kw != section_keyword):
                if len(text) < 200 and text.count("\n") == 0:
                    capture = False
                    break

            if capture and text:
                content.append(("PARA", text[:150]))

        # Lấy bảng thuộc section
        tables_in_section = []
        capture = False
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if section_keyword.lower() in text.lower():
                capture = True
            if capture and any(kw.lower() in text.lower() for kw in next_section_keywords if kw != section_keyword):
                capture = False

        return {
            "section": section_keyword,
            "content": content[:20],  # Lấy 20 dòng đầu
            "total_lines": len(content)
        }
    except Exception as e:
        return {"error": str(e)}

# ===== PHẦN 1: TRÍCH XUẤT PHẦN CAM KẾT TỪ CHỐI NỘP THỪA =====
print("=" * 100)
print("PHẦN 1: TRÍCH XUẤT NỘI DUNG CÁC SECTION")
print("=" * 100)

sections_to_compare = [
    "Tờ trình",
    "QĐKT",  # Quyết định kiểm tra
    "QĐGS",  # Quyết định giám sát
    "Cam kết",
]

for company_name, file_path in files.items():
    print(f"\n{company_name}:")
    print("-" * 100)

    if Path(file_path).exists():
        for section in sections_to_compare:
            result = extract_section_content(file_path, section)

            if "error" in result:
                print(f"  {section}: ✗ Lỗi ({result['error']})")
            elif result['content']:
                print(f"  {section}: ✓ Tìm thấy ({result['total_lines']} dòng)")
                # In 3 dòng đầu
                for i, (typ, line) in enumerate(result['content'][:3]):
                    print(f"      {i+1}. {line}")
            else:
                print(f"  {section}: ✗ Không tìm thấy")
    else:
        print(f"  FILE KHÔNG TỒN TẠI")

print("\n" + "=" * 100)
print("PHẦN 2: ĐỐI CHIẾU CHI TIẾT QĐKT, QĐGS, TỜ TRÌNH")
print("=" * 100)

def extract_full_text(doc_path, keyword, max_chars=500):
    """Trích xuất đầy đủ nội dung section."""
    try:
        doc = Document(doc_path)
        all_text = "\n".join([p.text for p in doc.paragraphs])

        # Tìm vị trí keyword
        idx = all_text.lower().find(keyword.lower())
        if idx == -1:
            return ""

        # Lấy 500 ký tự từ vị trí đó
        return all_text[idx:idx+max_chars]
    except:
        return ""

print("\nSo sánh nội dung TỜ TRÌNH:")
print("-" * 100)
for company_name, file_path in files.items():
    content = extract_full_text(file_path, "Tờ trình", 300)
    if content:
        print(f"\n{company_name}:")
        print(f"  {content[:200]}...")
    else:
        print(f"\n{company_name}: Không tìm thấy")

print("\n" + "=" * 100)
