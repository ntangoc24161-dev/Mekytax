#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Xuất file TONG HOP giải thể từ phôi.
Điền số QĐ để trống theo format: "số:               /QĐ-TCS3"
"""

import sys
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from openpyxl import load_workbook

# Thông tin vụ kiểm tra
CASE_INFO = {
    "mst": "3603216935",
    "ten_dn": "Công ty TNHH Một thành viên Kiến Thành Tân",
    "truong_doan": "Nguyễn Thị Ánh Ngọc",
    "thanh_vien": "Mai Ngọc Lân",
    "noi_dung_kt": "Kiểm tra chấp hành pháp luật Thuế",
    "thoi_ky_tu": 2015,
    "thoi_ky_den": 2025,
    "so_qd": "số:               /QĐ-TCS3",
    "dia_chi": "Số 361, đường Hoàng Bá Bích, KP 4, Phường Long Bình",
    "giam_doc": "Nguyễn Thị Kim Ngân",
    "ke_toan_truong": "Nguyễn Gia Hân"
}

CASE_PATH = Path("G:/Other computers/My Laptop/1. CONG TAC KIEM TRA/1- KIEM TRA/1. KIEM TRA TAI TRU SO NNT/2026/Dang kiem tra/GIAI THE/KIEN THANH TAN_3603216935")
BANG_TINH_PATH = CASE_PATH / "0. BIEN BAN" / "Bang tinh.xlsx"
TONG_HOP_PATH = CASE_PATH / "0. BIEN BAN" / "TONG HOP.docx"

def read_bang_tinh():
    """Đọc dữ liệu từ bảng tính."""
    try:
        wb = load_workbook(str(BANG_TINH_PATH), data_only=True)

        data = {}

        # Đọc sheet KQKDSXKD (kết quả kinh doanh)
        if "KQKDSXKD" in wb.sheetnames:
            ws = wb["KQKDSXKD"]
            # Lấy doanh thu từ hàng 2 (header ở hàng 1)
            doanh_thu = {}
            for year in range(2015, 2026):
                col_index = 2015 - 2014 + (year - 2015)  # Tính cột doanh thu
                try:
                    val = ws.cell(row=2, column=col_index).value
                    doanh_thu[year] = val if val else 0
                except:
                    doanh_thu[year] = 0
            data["doanh_thu"] = doanh_thu

        # Đọc sheet TINH PHAT + CHAM NOP
        if "TINH PHAT + CHAM NOP (C T)" in wb.sheetnames:
            ws = wb["TINH PHAT + CHAM NOP (C T)"]
            # Lấy tổng phạt và chậm nộp
            try:
                tong_phat = ws.cell(row=100, column=2).value or 0  # Ví dụ
                data["tong_phat"] = tong_phat
            except:
                data["tong_phat"] = 0

        wb.close()
        return data
    except Exception as e:
        print(f"Lỗi đọc bảng tính: {e}")
        return {}

def replace_placeholder_in_doc(doc, old, new):
    """Thay thế placeholder trong document."""
    for paragraph in doc.paragraphs:
        if old in paragraph.text:
            paragraph.text = paragraph.text.replace(old, str(new))

    # Thay trong bảng
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if old in paragraph.text:
                        paragraph.text = paragraph.text.replace(old, str(new))

def fill_tong_hop():
    """Điền dữ liệu vào file TONG HOP."""
    print(f"Mở file: {TONG_HOP_PATH}")

    if not TONG_HOP_PATH.exists():
        print(f"✗ File không tồn tại: {TONG_HOP_PATH}")
        return False

    try:
        doc = Document(str(TONG_HOP_PATH))

        # Thay thế placeholder
        replacements = {
            "«MST»": CASE_INFO["mst"],
            "«TÊN DN»": CASE_INFO["ten_dn"],
            "«TRƯỞNG ĐOÀN»": CASE_INFO["truong_doan"],
            "«THÀNH VIÊN»": CASE_INFO["thanh_vien"],
            "«NỘI DUNG KT»": CASE_INFO["noi_dung_kt"],
            "«TỪ NĂM»": str(CASE_INFO["thoi_ky_tu"]),
            "«ĐẾN NĂM»": str(CASE_INFO["thoi_ky_den"]),
            "«SỐ QĐ»": CASE_INFO["so_qd"],
            "«ĐỊA CHỈ»": CASE_INFO["dia_chi"],
            "«GIÁM ĐỐC»": CASE_INFO["giam_doc"],
            "«KỌ TRƯỞNG»": CASE_INFO["ke_toan_truong"],
        }

        for old, new in replacements.items():
            replace_placeholder_in_doc(doc, old, new)

        # Lưu file
        doc.save(str(TONG_HOP_PATH))
        print(f"✓ Đã lưu file: {TONG_HOP_PATH}")
        return True

    except Exception as e:
        print(f"✗ Lỗi: {e}")
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("XUẤT FILE TÔNG HỢP GIẢI THỂ")
    print("=" * 60)
    print(f"MST: {CASE_INFO['mst']}")
    print(f"Tên DN: {CASE_INFO['ten_dn']}")
    print(f"Thời kỳ: {CASE_INFO['thoi_ky_tu']}-{CASE_INFO['thoi_ky_den']}")
    print()

    # Đọc bảng tính
    print("Đọc dữ liệu từ bảng tính...")
    data = read_bang_tinh()
    if data:
        print(f"✓ Đã lấy dữ liệu")

    # Điền vào file
    print("Điền dữ liệu vào file TONG HOP...")
    if fill_tong_hop():
        print("\n✓ HOÀN TẤT")
        sys.exit(0)
    else:
        print("\n✗ CÓ LỖI")
        sys.exit(1)
