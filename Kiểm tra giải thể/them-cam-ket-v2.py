#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bổ sung phần Cam kết từ chối nộp thừa bằng copy-paste đơn giản.
"""

from pathlib import Path
from docx import Document

FILE_THUY_HONG = "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/1. TD/8. THUY HONG/THUY_HONG_3603437726/0. BIEN BAN/TONG HOP.docx"
FILE_KIEN_THANH_TAN = "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/KIEN THANH TAN_3603216935/0. BIEN BAN/TONG HOP.docx"

MST_KIEN_THANH = "3603216935"

print("=" * 100)
print("BỔ SUNG PHẦN CAM KẾT TỪ CHỐI NỘP THỪA (PHIÊN BẢN 2)")
print("=" * 100)

# Mở 2 file
print("\n1. Mở file...")
doc_thuy = Document(FILE_THUY_HONG)
doc_kien = Document(FILE_KIEN_THANH_TAN)
print(f"   ✓ Thúy Hồng: {len(doc_thuy.paragraphs)} đoạn, {len(doc_thuy.tables)} bảng")
print(f"   ✓ Kiến Thành Tân: {len(doc_kien.paragraphs)} đoạn, {len(doc_kien.tables)} bảng")

# Tìm phần Cam kết trong Thúy Hồng
print("\n2. Tìm phần Cam kết trong Thúy Hồng...")
cam_ket_start = -1
cam_ket_end = -1

for i, para in enumerate(doc_thuy.paragraphs):
    text = para.text.strip()
    if ("Cam kết" in text or "từ chối" in text or "thừa" in text) and len(text) > 10:
        cam_ket_start = i
        print(f"   ✓ Bắt đầu: dòng {i} - {text[:60]}...")
        break

if cam_ket_start > 0:
    # Tìm kết thúc (tìm phần tiếp theo)
    for i in range(cam_ket_start + 1, len(doc_thuy.paragraphs)):
        para = doc_thuy.paragraphs[i]
        text = para.text.strip()
        if "Phiếu chuyển" in text or "Đơn xin" in text:
            cam_ket_end = i
            print(f"   ✓ Kết thúc: dòng {i} - {text[:60]}...")
            break

    if cam_ket_end < 0:
        cam_ket_end = len(doc_thuy.paragraphs)

    # Copy phần Cam kết
    print(f"\n3. Copy {cam_ket_end - cam_ket_start} đoạn từ Thúy Hồng...")

    for i in range(cam_ket_start, cam_ket_end):
        para_thuy = doc_thuy.paragraphs[i]

        # Thêm paragraph vào Kiến Thành Tân
        new_para = doc_kien.add_paragraph(para_thuy.text)

        # Copy style
        new_para.style = para_thuy.style

        # Copy runs (formatting)
        new_para.clear()
        for run in para_thuy.runs:
            new_run = new_para.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            new_run.font.size = run.font.size

        # Thay thế placeholder MST
        if "«MST»" in new_para.text or "«MÃ SỐ THUẾ»" in new_para.text:
            new_para.text = new_para.text.replace("«MST»", MST_KIEN_THANH).replace("«MÃ SỐ THUẾ»", MST_KIEN_THANH)

    print(f"   ✓ Đã copy {cam_ket_end - cam_ket_start} đoạn")

    # Copy bảng MST (10 ô lưới)
    print(f"\n4. Copy bảng MST...")
    for table in doc_thuy.tables:
        if len(table.rows) == 1 and len(table.columns) >= 10:
            # Đây là bảng MST - copy nó
            tbl = table._element
            new_tbl = doc_kien.add_table()._element
            new_tbl.getparent().remove(new_tbl)
            doc_kien._element.body.append(tbl.__copy__())
            print(f"   ✓ Đã copy bảng MST (1 dòng × {len(table.columns)} cột)")
            break

    # Lưu file
    print(f"\n5. Lưu file Kiến Thành Tân...")
    doc_kien.save(FILE_KIEN_THANH_TAN)
    new_para_count = len(doc_kien.paragraphs)
    print(f"   ✓ Lưu thành công")
    print(f"   File mới có {new_para_count} đoạn (+{new_para_count - 770})")
    print(f"   File mới có {len(doc_kien.tables)} bảng (+{len(doc_kien.tables) - 56})")

else:
    print("   ✗ Không tìm thấy phần Cam kết trong Thúy Hồng")

print("\n" + "=" * 100)
