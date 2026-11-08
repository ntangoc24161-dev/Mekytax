#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
from collections import defaultdict
from datetime import datetime

base_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH'

# ĐỊNH DẠNG
header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", name="Times New Roman", size=11)
title_font = Font(bold=True, name="Times New Roman", size=12, color="FFFFFF")
title_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
warning_font = Font(color="9C0006", bold=True, name="Times New Roman")
danger_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
danger_font = Font(color="FFFFFF", bold=True, name="Times New Roman", size=11)
info_font = Font(name="Times New Roman", size=10)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
right_align = Alignment(horizontal="right", vertical="center")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

print("📖 Quét dữ liệu từ KQ2.1 và KQ2.2...")

kq21_path = os.path.join(base_path, 'KQ2.1_Phan_tich_CQT_3603883505_2022-2025_v4_FINAL.xlsx')
kq22_path = os.path.join(base_path, 'KQ2.2_Phan_tich_ho_so_cty_3603883505_2022-2025_v5.xlsx')

wb_kq21 = openpyxl.load_workbook(kq21_path, data_only=True)
wb_kq22 = openpyxl.load_workbook(kq22_path, data_only=True)

# ============ QUÉT KQ2_NCC_mua_vao ============
print("\n🔍 Quét KQ2_NCC_mua_vao...")
ws_ncc = wb_kq21['KQ2_NCC_mua_vao']

ncc_f0_hd_list = []
ncc_total_dict = defaultdict(lambda: {'gia_tri': 0, 'mst': '', 'count': 0})

for row in range(8, min(ws_ncc.max_row + 1, 800)):
    stt = ws_ncc.cell(row, 1).value
    mst = ws_ncc.cell(row, 2).value
    ten_ncc = ws_ncc.cell(row, 3).value
    hd_so = ws_ncc.cell(row, 4).value
    ngay_hd = ws_ncc.cell(row, 5).value
    hang = ws_ncc.cell(row, 6).value
    gia_tri = ws_ncc.cell(row, 7).value

    if not ten_ncc:
        break

    try:
        gia_tri_num = float(gia_tri) if gia_tri else 0
        ncc_total_dict[ten_ncc]['gia_tri'] += gia_tri_num
        ncc_total_dict[ten_ncc]['mst'] = mst
        ncc_total_dict[ten_ncc]['count'] += 1

        # NCC F0 (NNM Group)
        if 'NNM' in str(ten_ncc):
            ncc_f0_hd_list.append({
                'stt': stt,
                'mst': mst,
                'ten': ten_ncc,
                'hd_so': hd_so,
                'ngay': ngay_hd,
                'hang': hang,
                'gia_tri': gia_tri_num
            })
    except:
        pass

print(f"  → Tìm thấy {len(ncc_f0_hd_list)} hóa đơn NCC F0 (NNM Group)")

# Top-5 NCC
total_mua = sum(x['gia_tri'] for x in ncc_total_dict.values())
top5_ncc = sorted(ncc_total_dict.items(), key=lambda x: x[1]['gia_tri'], reverse=True)[:5]

print(f"  → Top-5 NCC (Tổng mua: {total_mua:,.0f} đ):")
top5_list = []
for idx, (ten, data) in enumerate(top5_ncc, 1):
    pct = (data['gia_tri'] / total_mua * 100) if total_mua > 0 else 0
    print(f"     {idx}. {ten}: {data['gia_tri']:,.0f} đ ({pct:.2f}%)")
    top5_list.append({
        'rank': idx,
        'ten': ten,
        'mst': data['mst'],
        'gia_tri': data['gia_tri'],
        'pct': pct,
        'count': data['count']
    })

# ============ QUÉT KQ2_511_632_Chi_tiet ============
print("\n🔍 Quét KQ2_511_632_Chi_tiet (HĐ bán dưới giá)...")
ws_511 = wb_kq22['KQ2_511_632_Chi_tiet']

hd_duoi_gia_list = []

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
                dt_num = float(dt) if dt else 0
                gv_num = float(gv) if gv else 0
                hd_duoi_gia_list.append({
                    'hd_so': hd_so,
                    'ngay': ngay_hd,
                    'nam': nam,
                    'dt': dt_num,
                    'gv': gv_num,
                    'lo': p_lai_num
                })
        except:
            pass

print(f"  → Tìm thấy {len(hd_duoi_gia_list)} hóa đơn bán dưới giá")

# ============ QUÉT KQ2_BK_HDDT_Lech ============
print("🔍 Quét KQ2_BK_HDDT_Lech (chênh lệch HĐĐT-BK)...")
if 'KQ2_BK_HDDT_Lech' in wb_kq21.sheetnames:
    ws_lech = wb_kq21['KQ2_BK_HDDT_Lech']
    lech_list = []

    for row in range(2, min(ws_lech.max_row + 1, 100)):
        hd_so = ws_lech.cell(row, 1).value
        if not hd_so:
            break

        ngay_hddt = ws_lech.cell(row, 2).value
        ngay_bk = ws_lech.cell(row, 3).value
        noi_dung = ws_lech.cell(row, 4).value

        lech_list.append({
            'hd_so': hd_so,
            'ngay_hddt': ngay_hddt,
            'ngay_bk': ngay_bk,
            'noi_dung': noi_dung
        })

    print(f"  → Tìm thấy {len(lech_list)} hóa đơn lệch HĐĐT-BK")
else:
    lech_list = []
    print("  ⚠️  Không tìm thấy sheet KQ2_BK_HDDT_Lech")

wb_kq21.close()
wb_kq22.close()

print("\n✅ Quét dữ liệu xong!")
print("\n" + "=" * 100)
print("📊 TÓM TẮT DỮ LIỆU")
print("=" * 100)
print(f"\n1. NCC F0 (NNM Group): {len(ncc_f0_hd_list)} hóa đơn, tổng {sum(x['gia_tri'] for x in ncc_f0_hd_list):,.0f} đ")
print(f"2. Top-5 NCC: {sum(x['gia_tri'] for x in top5_list):,.0f} đ ({(sum(x['gia_tri'] for x in top5_list)/total_mua*100):.2f}% tổng mua)")
print(f"3. HĐ bán dưới giá: {len(hd_duoi_gia_list)} hóa đơn, tổng lỗ {sum(x['lo'] for x in hd_duoi_gia_list):,.0f} đ")
print(f"4. HĐ lệch HĐĐT-BK: {len(lech_list)} hóa đơn")

print("\n💾 Giờ sẽ tạo file Excel tổng hợp chi tiết...")

# ============ TẠO FILE TỔNG HỢP ============
wb_new = openpyxl.Workbook()
wb_new.remove(wb_new.active)

# ====== SHEET 1: NCC F0 ======
ws1 = wb_new.create_sheet("01_NCC_F0_cao_nhat")

ws1.merge_cells('A1:K1')
ws1['A1'] = "DANH SÁCH NCC F0 (CÔNG AN/CVCBRR) CẦN GIẢI TRÌNH"
ws1['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws1['A1'].fill = danger_fill

headers = ["STT", "MST", "Tên NCC", "HĐ số", "Ngày HĐ", "Năm", "Mặt hàng", "Giá trị (đ)", "Ghi chú", "Trạng thái ĐKKD", "Lý do nước ngoài mua"]
for col_num, header in enumerate(headers, start=1):
    cell = ws1.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

for idx, item in enumerate(ncc_f0_hd_list, start=4):
    row = idx
    ws1.cell(row=row, column=1).value = item['stt']
    ws1.cell(row=row, column=2).value = item['mst']
    ws1.cell(row=row, column=3).value = item['ten']
    ws1.cell(row=row, column=4).value = item['hd_so']
    ws1.cell(row=row, column=5).value = item['ngay']
    ws1.cell(row=row, column=6).value = str(item['ngay']).split('-')[0] if item['ngay'] else ''
    ws1.cell(row=row, column=7).value = item['hang']
    ws1.cell(row=row, column=8).value = item['gia_tri']
    ws1.cell(row=row, column=9).value = "Kiểm tra chứng từ gốc"
    ws1.cell(row=row, column=10).value = "[Không hoạt động]"
    ws1.cell(row=row, column=11).value = "[Xin giải trình]"

    for col in range(1, 12):
        cell = ws1.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap
        cell.border = border
        if col == 8:
            cell.alignment = right_align
            cell.number_format = '#,##0'

ws1.column_dimensions['A'].width = 5
for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
    ws1.column_dimensions[col].width = 15

# ====== SHEET 2: TOP-5 NCC ======
ws2 = wb_new.create_sheet("02_Top5_NCC")

ws2.merge_cells('A1:G1')
ws2['A1'] = "DANH SÁCH TOP-5 NCC CHIẾM 95,22% MUA VÀO"
ws2['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws2['A1'].fill = danger_fill

headers2 = ["Rank", "Tên NCC", "MST", "Tổng giá trị (đ)", "% mua vào", "Số HĐ", "Ghi chú"]
for col_num, header in enumerate(headers2, start=1):
    cell = ws2.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

for idx, item in enumerate(top5_list, start=4):
    row = idx
    ws2.cell(row=row, column=1).value = item['rank']
    ws2.cell(row=row, column=2).value = item['ten']
    ws2.cell(row=row, column=3).value = item['mst']
    ws2.cell(row=row, column=4).value = item['gia_tri']
    ws2.cell(row=row, column=5).value = item['pct']
    ws2.cell(row=row, column=6).value = item['count']
    ws2.cell(row=row, column=7).value = "Xin giải trình lý do chọn NCC"

    for col in range(1, 8):
        cell = ws2.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap
        cell.border = border
        if col == 4:
            cell.alignment = right_align
            cell.number_format = '#,##0'
        elif col == 5:
            cell.alignment = right_align
            cell.number_format = '0.00%'

ws2.column_dimensions['A'].width = 8
for col in ['B', 'C', 'D', 'E', 'F', 'G']:
    ws2.column_dimensions[col].width = 18

# ====== SHEET 3: HĐ BÁN DƯỚI GIÁ ======
ws3 = wb_new.create_sheet("03_HD_ban_duoi_gia")

ws3.merge_cells('A1:I1')
ws3['A1'] = "DANH SÁCH HÓA ĐƠN BÁN DƯỚI GIÁ VỐN"
ws3['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws3['A1'].fill = danger_fill

headers3 = ["HĐ số", "Ngày HĐ", "Năm", "Doanh thu (đ)", "Giá vốn (đ)", "Lỗ (đ)", "Biên lợi nhuận (%)", "Lý do", "Giải pháp"]
for col_num, header in enumerate(headers3, start=1):
    cell = ws3.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

for idx, item in enumerate(hd_duoi_gia_list, start=4):
    row = idx
    ws3.cell(row=row, column=1).value = item['hd_so']
    ws3.cell(row=row, column=2).value = item['ngay']
    ws3.cell(row=row, column=3).value = item['nam']
    ws3.cell(row=row, column=4).value = item['dt']
    ws3.cell(row=row, column=5).value = item['gv']
    ws3.cell(row=row, column=6).value = item['lo']

    if item['dt'] > 0:
        bien = ((item['dt'] - item['gv']) / item['dt']) * 100
        ws3.cell(row=row, column=7).value = bien
    else:
        ws3.cell(row=row, column=7).value = 0

    ws3.cell(row=row, column=8).value = "[Xin giải trình]"
    ws3.cell(row=row, column=9).value = "[Biện pháp]"

    for col in range(1, 10):
        cell = ws3.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap
        cell.border = border
        if col in [4, 5, 6]:
            cell.alignment = right_align
            cell.number_format = '#,##0'
        elif col == 7:
            cell.alignment = right_align
            cell.number_format = '0.00'

for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']:
    ws3.column_dimensions[col].width = 15

# ====== SHEET 4: LỆCH HĐĐT-BK ======
ws4 = wb_new.create_sheet("04_Lech_HDDT_BangKe")

ws4.merge_cells('A1:E1')
ws4['A1'] = f"DANH SÁCH {len(lech_list)} HÓA ĐƠN LỆCH GIỮA HĐĐT VÀ BẢNG KÊ"
ws4['A1'].font = Font(bold=True, color="FFFFFF", size=12, name="Times New Roman")
ws4['A1'].fill = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")

headers4 = ["HĐ số", "Ngày HĐĐT", "Ngày BK", "Nội dung lệch", "Lý do & Sửa chữa"]
for col_num, header in enumerate(headers4, start=1):
    cell = ws4.cell(row=3, column=col_num)
    cell.value = header
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

for idx, item in enumerate(lech_list[:50], start=4):  # Giới hạn 50 dòng
    row = idx
    ws4.cell(row=row, column=1).value = item['hd_so']
    ws4.cell(row=row, column=2).value = item['ngay_hddt']
    ws4.cell(row=row, column=3).value = item['ngay_bk']
    ws4.cell(row=row, column=4).value = item['noi_dung']
    ws4.cell(row=row, column=5).value = "[Xin giải trình]"

    for col in range(1, 6):
        cell = ws4.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap
        cell.border = border

for col in ['A', 'B', 'C', 'D', 'E']:
    ws4.column_dimensions[col].width = 18

# Lưu file
output_file = os.path.join(base_path, 'KQ3_NNT_Giai_trinh_Chi_tiet_3603883505.xlsx')
wb_new.save(output_file)

print(f"\n✅ Đã tạo file tổng hợp chi tiết!")
print(f"📁 File: {output_file}")
print(f"\n📋 Nội dung:")
print(f"  Sheet 1: 01_NCC_F0_cao_nhat - {len(ncc_f0_hd_list)} hóa đơn NCC F0")
print(f"  Sheet 2: 02_Top5_NCC - {len(top5_list)} NCC chiếm {(sum(x['gia_tri'] for x in top5_list)/total_mua*100):.2f}%")
print(f"  Sheet 3: 03_HD_ban_duoi_gia - {len(hd_duoi_gia_list)} hóa đơn bán lỗ")
print(f"  Sheet 4: 04_Lech_HDDT_BangKe - {len(lech_list)} hóa đơn lệch")
