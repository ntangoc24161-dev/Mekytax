#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bổ sung phần Cam kết từ chối nộp thừa từ Thúy Hồng vào Kiến Thành Tân.
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches
from copy import deepcopy
import re

FILE_THUY_HONG = "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/1. TD/8. THUY HONG/THUY_HONG_3603437726/0. BIEN BAN/TONG HOP.docx"
FILE_KIEN_THANH_TAN = "G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/KIEN THANH TAN_3603216935/0. BIEN BAN/TONG HOP.docx"

MST_KIEN_THANH = "3603216935"

def extract_cam_ket_section(doc_path):
    """Trích xuất phần Cam kết từ chối nộp thừa."""
    try:
        doc = Document(doc_path)

        cam_ket_elements = []
        capture = False

        # Tìm phần Cam kết
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()

            # Bắt đầu: tìm "Cam kết" hoặc "từ chối"
            if ("Cam kết" in text or "từ chối" in text or "thừa" in text) and not capture:
                capture = True
                cam_ket_elements.append(("PARA", para))
                print(f"  Tìm thấy Cam kết ở dòng {i}: {text[:60]}...")
                continue

            # Kết thúc: tìm "Phiếu chuyển" hoặc "Đơn xin" hoặc các phần khác
            if capture and ("Phiếu chuyển" in text or "Đơn xin" in text or "Hoàn thuế" in text):
                if len(cam_ket_elements) > 5:  # Đã lấy đủ
                    capture = False
                    print(f"  Kết thúc Cam kết ở dòng {i}")
                    break

            if capture:
                cam_ket_elements.append(("PARA", para))

        # Cũng lấy các bảng phía sau
        for i, table in enumerate(doc.tables):
            # Kiểm tra xem bảng này có phải phần Cam kết không (thường là 10 ô lưới cho MST)
            if len(table.rows) == 1 and len(table.columns) >= 10:
                # Có thể là bảng MST
                cam_ket_elements.append(("TABLE", table))
                print(f"  Tìm thấy bảng MST (10 ô lưới)")
                break

        return cam_ket_elements
    except Exception as e:
        print(f"  Lỗi: {e}")
        return []

def add_cam_ket_to_file(doc_kien, cam_ket_elements, mst):
    """Thêm phần Cam kết vào cuối file."""
    try:
        # Thêm break page
        last_para = doc_kien.paragraphs[-1]
        new_para = last_para._element.addnext(deepcopy(last_para._element))
        new_para.text = ""

        # Thêm phần Cam kết
        added_count = 0
        for elem_type, elem in cam_ket_elements:
            if elem_type == "PARA":
                # Copy paragraph
                new_para = last_para._element.addnext(deepcopy(elem._element))
                # Thay thế MST nếu có placeholder
                new_para.text = elem.text.replace("«MST»", mst).replace("«MÃ SỐ THUẾ»", mst)
                added_count += 1
            elif elem_type == "TABLE":
                # Copy table
                new_table = last_para._element.addnext(deepcopy(elem._element))
                added_count += 1

        return added_count
    except Exception as e:
        print(f"  Lỗi thêm: {e}")
        return 0

# ===== THỰC HIỆN =====
print("=" * 100)
print("BỔ SUNG PHẦN CAM KẾT TỪ CHỐI NỘP THỪA")
print("=" * 100)

print("\n1. Trích xuất phần Cam kết từ Thúy Hồng...")
if Path(FILE_THUY_HONG).exists():
    cam_ket_elements = extract_cam_ket_section(FILE_THUY_HONG)
    print(f"   ✓ Trích xuất thành công ({len(cam_ket_elements)} phần tử)")
else:
    print(f"   ✗ File Thúy Hồng không tồn tại")
    cam_ket_elements = []

print("\n2. Mở file Kiến Thành Tân...")
if Path(FILE_KIEN_THANH_TAN).exists() and cam_ket_elements:
    doc_kien = Document(FILE_KIEN_THANH_TAN)
    print(f"   ✓ Mở thành công (hiện có {len(doc_kien.paragraphs)} đoạn)")

    print(f"\n3. Thêm phần Cam kết với MST {MST_KIEN_THANH}...")
    added = add_cam_ket_to_file(doc_kien, cam_ket_elements, MST_KIEN_THANH)

    if added > 0:
        print(f"   ✓ Đã thêm {added} phần tử")

        print(f"\n4. Lưu file...")
        doc_kien.save(FILE_KIEN_THANH_TAN)
        print(f"   ✓ Lưu thành công")
        print(f"   File mới có {len(doc_kien.paragraphs)} đoạn (+{len(doc_kien.paragraphs) - 770})")
    else:
        print(f"   ✗ Không thêm được")
else:
    print(f"   ✗ File không tồn tại hoặc chưa trích xuất được Cam kết")

print("\n" + "=" * 100)
