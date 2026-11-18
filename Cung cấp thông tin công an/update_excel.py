import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    path = os.path.abspath(r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx")
    wb = excel.Workbooks.Open(path)
    ws = wb.Sheets("NHAP_HO_SO")
    
    r = 860
    data = [
        (1, '217'),
        (2, '289'),
        (3, '1806/YC-ANĐT-Đ2'),
        (4, '11/11/2026'),
        (5, 'Cơ quan An ninh điều tra - Công an thành phố Đồng Nai'),
        (6, ''),
        (7, '16/11/2026'),
        (8, 'Điều tra vụ án Mua bán trái phép HĐ, chứng từ và Trốn thuế xảy ra tại 50 CQ, DN trên địa bàn TP Đồng Nai: Cty TNHH Toàn Trí sử dụng 02 HĐ đầu vào của Cty TNHH TM XNK Nhựa Nhật Duy (MST 0316840485) tổng tiền 2.420.000 đ. Yêu cầu cung cấp: kê khai khấu trừ, hoàn thuế GTGT, chi phí TNDN; việc kê khai bổ sung điều chỉnh giảm; ưu đãi, miễn giảm thuế; tình hình xử lý VPHC về thuế/hóa đơn từ 2019 đến nay; số TK ngân hàng. Hạn 15/11/2026. ĐTV Lê Văn Hân 0969.627.979 (1032 Nguyễn Ái Quốc, P. Trảng Dài). Số đến TCS3 14490 ngày 16/11/2026.'),
        (9, '1'),
        (10, '3601035228'),
        (11, 'CÔNG TY TNHH TOÀN TRÍ'),
        (12, 'Phường Tam Hiệp'),
        (13, 'KT2'),
        (14, ''),
        (15, 'Mới'),
        (16, ''),
        (17, '16/11/2026'),
        (18, '15/11/2026'),
        (19, ''),
        (20, 'Không'),
        (21, 'Không'),
        (22, ''),
        (23, ''),
        (24, ''),
        (25, 'DB NNT 03/06/2026: TCS3, KT2, P. Tam Hiệp, Y7B KP 6. NNT đang hoạt động. Bóc tách từ CV MOI\\1806YC-ANĐT.pdf lưu folder 289.')
    ]
    for col, val in data:
        ws.Cells(r, col).Value = val
        print(f"Set col {col}: {ws.Cells(r, col).Text}")
    
    wb.Save()
    print("Saved successfully!")
    wb.Close(False)
finally:
    excel.Quit()
