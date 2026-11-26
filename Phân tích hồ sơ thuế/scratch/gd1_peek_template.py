import openpyxl
import os

BASE = r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\thu-muc-mau-ho-so-doanh-nghiep"
for name in [
    r"MAU_FILE_TONG_HOP\cu-v1.2\KQ1_MAU_Kiem_ke_danh_gia.xlsx",
    r"MAU_FILE_TONG_HOP\Bo_bieu_mau_chuan_phan_tich_ho_so_thue.xlsx",
    r"Bo_bieu_mau_chuan_phan_tich_ho_so_thue_v2.xlsx",
]:
    p = os.path.join(BASE, name)
    print("==", name, os.path.exists(p))
    if not os.path.exists(p):
        continue
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    print("sheets:", wb.sheetnames)
    for ws in wb.worksheets:
        if ws.title.lower().startswith(("04", "kq1", "kq1_", "in00", "01")) or "quy" in ws.title.lower():
            print("--", ws.title, ws.max_row, ws.max_column)
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i >= 6:
                    break
                print([c for c in row][:12])
    wb.close()
