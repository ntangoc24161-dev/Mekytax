#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
So sánh chi tiết nội dung các văn bản (Tờ trình, QĐKT, QĐGS, v.v.)
"""

from pathlib import Path
from docx import Document
import re

files = {
    "Kiến Thành Tân": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/KIEN THANH TAN_3603216935/0. BIEN BAN/TONG HOP.docx",
    "Thúy Hồng": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/1. TD/8. THUY HONG/THUY_HONG_3603437726/0. BIEN BAN/TONG HOP.docx",
    "Sơn Tiền Hải": "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/1. TD/11. SON TIEN HAI/0. BIEN BAN/TONG HOP.docx"
}

def get_all_text(doc_path):
    """Lấy toàn bộ nội dung document."""
    try:
        doc = Document(doc_path)
        return "\n".join([p.text for p in doc.paragraphs])
    except:
        return ""

def extract_between(text, start_keyword, end_keyword, max_chars=1000):
    """Trích xuất nội dung giữa 2 từ khóa."""
    start_idx = text.lower().find(start_keyword.lower())
    if start_idx < 0:
        return None

    end_idx = text.lower().find(end_keyword.lower(), start_idx + len(start_keyword))
    if end_idx < 0:
        end_idx = start_idx + max_chars

    return text[start_idx:end_idx][:max_chars]

def extract_numbered_list(text, start_marker, max_items=5):
    """Trích xuất danh sách đánh số từ text."""
    pattern = r'^\d+\.\s+(.+)$'
    lines = text.split('\n')

    items = []
    start_found = False
    for line in lines:
        if start_marker.lower() in line.lower():
            start_found = True
            continue

        if start_found:
            match = re.match(pattern, line.strip())
            if match:
                items.append(match.group(1)[:80])
                if len(items) >= max_items:
                    break
            elif line.strip().startswith(('=', '-', '*', 'V', 'K', 'P', 'Q', 'B')) and len(line.strip()) > 50:
                break

    return items

print("=" * 100)
print("SO SÁNH CHI TIẾT NỘI DUNG CÁC VĂN BẢN (QĐKT, QĐGS, TỜ TRÌNH)")
print("=" * 100)

# Lấy toàn bộ nội dung
all_texts = {}
for name, path in files.items():
    if Path(path).exists():
        all_texts[name] = get_all_text(path)
        print(f"\n✓ {name}: {len(all_texts[name])} ký tự")
    else:
        print(f"\n✗ {name}: FILE KHÔNG TỒN TẠI")

print("\n" + "=" * 100)
print("1. SO SÁNH TỜ TRÌNH")
print("=" * 100)

for name, full_text in all_texts.items():
    content = extract_between(full_text, "TỜ TRÌNH", "Quyết định", 800)
    if content:
        print(f"\n{name}:")
        # Lấy 3 dòng đầu sau "TỜ TRÌNH"
        lines = content.split('\n')
        for i, line in enumerate(lines[1:4]):
            if line.strip():
                print(f"  {line.strip()[:80]}...")
    else:
        print(f"\n{name}: Không tìm thấy Tờ trình")

print("\n" + "=" * 100)
print("2. SO SÁNH QUYẾT ĐỊNH KIỂM TRA (QĐKT)")
print("=" * 100)

for name, full_text in all_texts.items():
    # Tìm "Quyết định" và "kiểm tra"
    qd_idx = full_text.lower().find("quyết định")
    kt_idx = full_text.lower().find("kiểm tra", qd_idx)

    if qd_idx >= 0:
        content = full_text[qd_idx:qd_idx+1000]
        print(f"\n{name}:")

        # Tìm "Nội dung công việc"
        noi_dung_idx = content.lower().find("nội dung công việc")
        if noi_dung_idx >= 0:
            noi_dung = content[noi_dung_idx:noi_dung_idx+200]
            print(f"  Nội dung: {noi_dung.split(chr(10))[0][:80]}...")

        # Tìm "Thời hạn"
        thoi_han_idx = content.lower().find("thời hạn")
        if thoi_han_idx >= 0:
            thoi_han = content[thoi_han_idx:thoi_han_idx+100]
            print(f"  Thời hạn: {thoi_han.split(chr(10))[0][:80]}...")

        # Tìm "Trưởng đoàn"
        truong_doan_idx = content.lower().find("trưởng đoàn")
        if truong_doan_idx >= 0:
            truong_doan = content[truong_doan_idx:truong_doan_idx+150]
            print(f"  Trưởng đoàn: {truong_doan.split(chr(10))[0][:80]}...")

print("\n" + "=" * 100)
print("3. SO SÁNH QUYẾT ĐỊNH GIÁM SÁT (QĐGS)")
print("=" * 100)

for name, full_text in all_texts.items():
    # Tìm "Quyết định giám sát"
    gs_idx = full_text.lower().find("giám sát")

    if gs_idx >= 0:
        content = full_text[max(0, gs_idx-100):gs_idx+800]
        print(f"\n{name}:")

        # Tìm các mục điều
        dieu_idx = content.lower().find("điều")
        if dieu_idx >= 0:
            dieu = content[dieu_idx:dieu_idx+100]
            print(f"  Điều: {dieu.split(chr(10))[0][:80]}...")

        # Lấy nội dung giám sát
        lines = content.split('\n')
        printed = 0
        for line in lines:
            if line.strip() and len(line.strip()) > 10 and printed < 2:
                print(f"  {line.strip()[:80]}...")
                printed += 1

print("\n" + "=" * 100)
print("4. KIỂM TRA PLACEHOLDER CHƯA ĐIỀN")
print("=" * 100)

placeholders = {}
for name, full_text in all_texts.items():
    ph = re.findall(r'«([^»]+)»', full_text)
    placeholders[name] = list(set(ph))
    if ph:
        print(f"\n{name}: {len(set(ph))} placeholder chưa điền")
        for p in sorted(set(ph))[:10]:
            print(f"  • «{p}»")
    else:
        print(f"\n{name}: ✓ Không có placeholder chưa điền")

print("\n" + "=" * 100)
