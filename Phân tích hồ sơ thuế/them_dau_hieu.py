#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

base_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\TT_BIEN_HOA_3603883505\4. KET QUA PHAN TICH'

# Mở file tổng hợp hiện tại
file_path = os.path.join(base_path, 'KQ3_Danh_gia_Dau_hieu_Bat_thuong_3603883505.xlsx')
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# ĐỊNH DẠNG
warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
warning_font = Font(color="9C0006", bold=True, name="Times New Roman")
info_font = Font(name="Times New Roman", size=10)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Ba dấu hiệu bổ sung
new_data = [
    {
        'stt': 7,
        'loai': 'Chênh lệch HĐĐT & Bảng kê',
        'chi_tiet': 'Từ sheet KQ2_BK_HDDT_Lech: số lượng/giá trị hóa đơn không khớp giữa HĐĐT (IN01) và Bảng kê (IN02)',
        'muc_do': 'Trung bình',
        'vi_tri': 'KQ2_BK_HDDT_Lech (KQ2.1)',
        'huong_kt': '1. Xác định các HĐ lệch: theo dòng HĐ, tìm lý do (ghi nhầm số HĐ, HĐĐT sai ngày, BK sai ngày)\n2. Với mỗi HĐ lệch: xác minh chứng từ gốc\n3. Nếu HĐĐT hoặc BK sai: tác động đến tờ khai GTGT hoặc quyết toán TNDN',
        'tai_lieu': 'HĐĐT (IN01) + Bảng kê (IN02)\nSheet KQ2_BK_HDDT_Lech chi tiết\nHĐ gốc để so sánh',
        'trang_thai': 'Cần kiểm tra'
    },
    {
        'stt': 8,
        'loai': 'Lệch giữa Nợ khách (TK 131) & Bảng kê',
        'chi_tiet': 'Công nợ 131 (sổ DN) không khớp với tổng nợ từ bảng kê mua vào (IN02) hoặc từ sao kê ngân hàng',
        'muc_do': 'Trung bình',
        'vi_tri': 'KQ2_Cong_no_Tien_mat (KQ2.2)\nIN08_So_KeToan (KQ2.1)',
        'huong_kt': '1. Tính tổng công nợ từ Bảng kê (doanh thu chưa thanh toán)\n2. So sánh với số dư TK 131 ngày cuối kỳ (sổ và BCTC)\n3. Tìm các HĐ bán chưa thu hoặc thu không đủ\n4. Xác minh sao kê: tất cả dòng thanh toán từ khách có trong BK không',
        'tai_lieu': 'Sổ 131 (Khách hàng nợ) - IN08\nBảng kê doanh thu (IN02)\nSao kê ngân hàng (IN11)\nBCTC [131] Khách hàng nợ',
        'trang_thai': 'Cần kiểm tra'
    },
    {
        'stt': 9,
        'loai': 'NLĐ thu nhập từ nhiều nguồn',
        'chi_tiet': 'Nếu các NLĐ có thu nhập từ: lương (IN05), bình quân về (GDLK), khác, tạo rủi ro thuế TNCN không toàn',
        'muc_do': 'Thấp',
        'vi_tri': 'IN05_NLD_TNCN (KQ2.1)\nKQ2_TNCN_LD2 (KQ2.1)',
        'huong_kt': '1. Danh sách NLĐ thu nhập từ 2+ nguồn\n2. Tính tổng thu nhập theo thứ tự ưu tiên (lương > bình quân > khác)\n3. Kiểm tra QTT TNCN có tính đúng không (nếu có bình quân về thì được trừ 2 tháng lương)\n4. Xác minh những NLĐ có chi phí khác (tối hậu, tiền thưởng, v.v.) có kê khai đủ',
        'tai_lieu': 'Phụ lục 05-1 (Danh sách NLĐ)\nQTT TNCN từng tháng/quý\nBảng kê bình quân về (GDLK)\nHợp đồng NLĐ',
        'trang_thai': 'Cần xác nhận'
    }
]

# Thêm dữ liệu từ dòng 11 (sau 6 dấu hiệu cũ)
for idx, item in enumerate(new_data, start=11):
    row = idx

    ws.cell(row=row, column=1).value = item['stt']
    ws.cell(row=row, column=2).value = item['loai']
    ws.cell(row=row, column=3).value = item['chi_tiet']
    ws.cell(row=row, column=4).value = item['muc_do']
    ws.cell(row=row, column=5).value = item['vi_tri']
    ws.cell(row=row, column=6).value = item['huong_kt']
    ws.cell(row=row, column=7).value = item['tai_lieu']
    ws.cell(row=row, column=8).value = item['trang_thai']

    # Định dạng
    for col in range(1, 9):
        cell = ws.cell(row=row, column=col)
        cell.font = info_font
        cell.alignment = left_wrap if col > 2 else center
        cell.border = border

        if col == 4:
            if item['muc_do'] == 'Cao':
                cell.fill = warning_fill
                cell.font = warning_font
            elif item['muc_do'] == 'Trung bình':
                cell.fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")

# Lưu file
wb.save(file_path)

print("✅ Đã bổ sung 3 dấu hiệu bất thường mới!")
print("📋 Dấu hiệu #7: Chênh lệch HĐĐT & Bảng kê")
print("📋 Dấu hiệu #8: Lệch giữa Nợ khách (TK 131) & Bảng kê")
print("📋 Dấu hiệu #9: NLĐ thu nhập từ nhiều nguồn")
print(f"\n📁 File: {file_path}")
print("✓ Tổng cộng 9 dấu hiệu bất thường")
