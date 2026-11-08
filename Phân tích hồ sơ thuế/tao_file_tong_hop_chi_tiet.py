#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
from collections import defaultdict

base_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH'

# ĐỊNH DẠNG
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", name="Times New Roman", size=11)
title_font = Font(bold=True, name="Times New Roman", size=12)
warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
info_font = Font(name="Times New Roman", size=10)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
right_align = Alignment(horizontal="right", vertical="center")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Mở KQ2.1 và KQ2.2 để quét dữ liệu
kq21_path = os.path.join(base_path, 'KQ2.1_Phan_tich_CQT_3603883505_2022-2025_v4_FINAL.xlsx')
kq22_path = os.path.join(base_path, 'KQ2.2_Phan_tich_ho_so_cty_3603883505_2022-2025_v5.xlsx')

print("📖 Quét dữ liệu từ KQ2.1 và KQ2.2...")

# ============ QUÉT KQ2.1 ============
wb_kq21 = openpyxl.load_workbook(kq21_path, data_only=True)

# Quét IN07_NCC_CanhBao (NCC F0)
print("\n🔍 Quét IN07_NCC_CanhBao (NCC F0)...")
ws_in07 = wb_kq21['IN07_NCC_CanhBao']
ncc_f0_list = []

for row in range(2, min(ws_in07.max_row + 1, 200)):
    ten_ncc = ws_in07.cell(row, 2).value
    mst = ws_in07.cell(row, 3).value
    nhom_f = ws_in07.cell(row, 7).value

    if nhom_f == 'F0' and ten_ncc:
        ncc_f0_list.append({
            'ten': ten_ncc,
            'mst': mst
        })

print(f"  → Tìm thấy {len(ncc_f0_list)} NCC F0")

# Quét KQ2_NCC_mua_vao (chi tiết HĐ NCC F0 và Top-5)
print("🔍 Quét KQ2_NCC_mua_vao (chi tiết HĐ và Top-5)...")
ws_ncc_mua = wb_kq21['KQ2_NCC_mua_vao']
ncc_f0_hd = []
ncc_total_dict = defaultdict(float)

for row in range(2, min(ws_ncc_mua.max_row + 1, 800)):
    ten_ncc = ws_ncc_mua.cell(row, 2).value
    mst = ws_ncc_mua.cell(row, 3).value
    hd_so = ws_ncc_mua.cell(row, 4).value
    ngay_hd = ws_ncc_mua.cell(row, 5).value
    gia_tri = ws_ncc_mua.cell(row, 6).value
    nhom_f = ws_ncc_mua.cell(row, 13).value

    if ten_ncc and gia_tri:
        try:
            gia_tri_num = float(gia_tri) if gia_tri else 0
            ncc_total_dict[ten_ncc] += gia_tri_num

            # Nếu là F0, lưu chi tiết
            if nhom_f == 'F0':
                ncc_f0_hd.append({
                    'ten': ten_ncc,
                    'mst': mst,
                    'hd_so': hd_so,
                    'ngay': ngay_hd,
                    'gia_tri': gia_tri_num
                })
        except:
            pass

print(f"  → Tìm thấy {len(ncc_f0_hd)} hóa đơn NCC F0")

# Top-5 NCC
print("🔍 Tính Top-5 NCC...")
top5_ncc = sorted(ncc_total_dict.items(), key=lambda x: x[1], reverse=True)[:5]
total_mua = sum(ncc_total_dict.values())

print(f"  → Top-5 NCC:")
for idx, (ncc, value) in enumerate(top5_ncc, 1):
    pct = (value / total_mua * 100) if total_mua > 0 else 0
    print(f"     {idx}. {ncc}: {value:,.0f} đ ({pct:.2f}%)")

# Quét KQ2_BK_HDDT_Lech (chênh lệch HĐĐT - BK)
print("🔍 Quét KQ2_BK_HDDT_Lech (chênh lệch HĐĐT-BK)...")
if 'KQ2_BK_HDDT_Lech' in wb_kq21.sheetnames:
    ws_lech = wb_kq21['KQ2_BK_HDDT_Lech']
    lech_hddt_bk = []

    for row in range(2, min(ws_lech.max_row + 1, 100)):
        hd_so = ws_lech.cell(row, 1).value
        ngay_hddt = ws_lech.cell(row, 2).value
        ngay_bk = ws_lech.cell(row, 3).value
        noi_dung = ws_lech.cell(row, 4).value

        if hd_so and (ngay_hddt or ngay_bk):
            lech_hddt_bk.append({
                'hd_so': hd_so,
                'ngay_hddt': ngay_hddt,
                'ngay_bk': ngay_bk,
                'noi_dung': noi_dung
            })

    print(f"  → Tìm thấy {len(lech_hddt_bk)} hóa đơn lệch HĐĐT-BK")
else:
    print("  ⚠️  Không tìm thấy sheet KQ2_BK_HDDT_Lech")
    lech_hddt_bk = []

# Quét NLĐ nhiều nguồn
print("🔍 Quét KQ2_TNCN_LD2 (NLĐ nhiều nguồn)...")
if 'KQ2_TNCN_LD2' in wb_kq21.sheetnames:
    ws_nld = wb_kq21['KQ2_TNCN_LD2']
    nld_multi = []

    for row in range(2, min(ws_nld.max_row + 1, 100)):
        ten_nld = ws_nld.cell(row, 2).value
        cccd = ws_nld.cell(row, 3).value
        luong = ws_nld.cell(row, 4).value
        bq_ve = ws_nld.cell(row, 5).value
        khac = ws_nld.cell(row, 6).value

        # Nếu có 2+ nguồn
        count_source = sum(1 for x in [luong, bq_ve, khac] if x and float(x) > 0)
        if count_source >= 2 and ten_nld:
            try:
                nld_multi.append({
                    'ten': ten_nld,
                    'cccd': cccd,
                    'luong': float(luong) if luong else 0,
                    'bq_ve': float(bq_ve) if bq_ve else 0,
                    'khac': float(khac) if khac else 0
                })
            except:
                pass

    print(f"  → Tìm thấy {len(nld_multi)} NLĐ có nhiều nguồn thu nhập")
else:
    print("  ⚠️  Không tìm thấy sheet KQ2_TNCN_LD2")
    nld_multi = []

wb_kq21.close()

# ============ QUÉT KQ2.2 ============
wb_kq22 = openpyxl.load_workbook(kq22_path, data_only=True)

# Quét KQ2_511_632_Chi_tiet (HĐ bán dưới giá)
print("🔍 Quét KQ2_511_632_Chi_tiet (HĐ bán dưới giá)...")
ws_511 = wb_kq22['KQ2_511_632_Chi_tiet']
hd_duoi_gia = []

for row in range(2, min(ws_511.max_row + 1, 420)):
    hd_so = ws_511.cell(row, 1).value
    ngay_hd = ws_511.cell(row, 2).value
    nam = ws_511.cell(row, 3).value
    dt = ws_511.cell(row, 4).value
    gv = ws_511.cell(row, 5).value
    p_lai = ws_511.cell(row, 7).value

    if hd_so and p_lai:
        try:
            p_lai_num = float(p_lai) if p_lai else 0
            if p_lai_num < 0:
                hd_duoi_gia.append({
                    'hd_so': hd_so,
                    'ngay': ngay_hd,
                    'nam': nam,
                    'dt': float(dt) if dt else 0,
                    'gv': float(gv) if gv else 0,
                    'lo': p_lai_num
                })
        except:
            pass

print(f"  → Tìm thấy {len(hd_duoi_gia)} hóa đơn bán dưới giá")

# Quét KQ2_Cong_no_Tien_mat (lệch TK 131)
print("🔍 Quét KQ2_Cong_no_Tien_mat (lệch TK 131)...")
if 'KQ2_Cong_no_Tien_mat' in wb_kq22.sheetnames:
    ws_cn = wb_kq22['KQ2_Cong_no_Tien_mat']
    # Lấy dòng tổng TK 131
    for row in range(2, min(ws_cn.max_row + 1, 30)):
        chi_tieu = ws_cn.cell(row, 2).value
        if chi_tieu and '131' in str(chi_tieu):
            cn_131 = ws_cn.cell(row, 5).value
            print(f"  → TK 131 (Khách nợ): {cn_131:,.0f} đ" if cn_131 else "  → TK 131: Chưa có dữ liệu")
            break

wb_kq22.close()

print("\n✅ Quét dữ liệu xong! Giờ sẽ tạo file tổng hợp chi tiết...")
