import win32com.client
import os
import shutil
import zipfile
import re

excel_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
doc_path = r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx'

# Ensure backup directories exist
os.makedirs(r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\backups', exist_ok=True)
excel_bak = r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\backups\Danh sach kiem tra quyet toan pro.bak-20261113.xlsm'
doc_bak = r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.bak-truoc-xoa-dulieumerge-20261113.docx'

# Step 1: Backup files
shutil.copy2(excel_path, excel_bak)
shutil.copy2(doc_path, doc_bak)
print(f"Backed up Excel to: {excel_bak}")
print(f"Backed up Docx to: {doc_bak}")

field_names_58 = [
    'STT', 'Trang_thai', 'Loai_hinh', 'MST', 'Ten_NNT', 'Dia_chi', 'Ten_CQT', 'Ma_CQT',
    'Thoi_ky_KT', 'Nganh_nghe', 'Giam_doc', 'Gioi_tinh', 'CCCD_GD', 'Ngay_cap_CCCD', 'Noi_cap_CCCD',
    'Ke_toan_truong', 'CCCD_KT', 'So_DKKD', 'Ngay_cap_DKKD', 'Chi_nhanh', 'Tai_khoan_NH',
    'Truong_doan', 'Chuc_vu_TD', 'Thanh_vien', 'Chuc_vu_TV', 'Noi_dung_KT', 'Phan_cong_TD', 'Phan_cong_TV',
    'So_QDKT', 'So_QDGS', 'Ngay_QDKT', 'Ngay_cong_bo', 'Thoi_han_KT', 'Ngay_bat_dau', 'Ngay_ket_thuc',
    'Ngay_ky_BBKT', 'Ngay_bao_cao', 'So_QDXPHC', 'Ngay_QDXPHC', 'Truy_thu_TNDN', 'Truy_thu_GTGT',
    'Tong_truy_thu', 'Cham_nop_GTGT', 'Cham_nop_TNDN', 'Tong_cham_nop', 'Phat_20_phan_tram', 'Phat_thu_tuc',
    'Tong_tien_phat', 'Tong_phat_sinh', 'Giam_lo', 'Giam_khau_tru', 'Thong_bao_qua_han',
    'Tong_truy_thu_chu', 'Tong_cham_nop_chu', 'Tong_tien_phat_chu', 'Tong_phat_sinh_chu',
    'Giam_lo_chu', 'Giam_khau_tru_chu'
]

# Step 2: Modify Excel via COM
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Sheets('01_DA_BAN_HANH')
    
    # Insert new Row 1
    ws.Rows(1).Insert()
    for col_idx, f_name in enumerate(field_names_58, 1):
        ws.Cells(1, col_idx).Value = f_name
    ws.Rows(1).Hidden = True
    
    # Adjust STT formula in col A for data rows
    max_r = ws.UsedRange.Rows.Count
    for r in range(6, max_r + 1):
        if ws.Cells(r, 4).Value:
            ws.Cells(r, 1).Formula = f'=IF(D{r}<>"",ROW()-5,"")'
            
    # Delete sheet DuLieuMerge
    deleted_merge = False
    for s in wb.Sheets:
        if s.Name == 'DuLieuMerge':
            s.Delete()
            deleted_merge = True
            break
            
    print(f"DuLieuMerge sheet deleted: {deleted_merge}")
    wb.Save()
    wb.Close()
    print("Excel updated successfully!")
finally:
    excel.Quit()

# Step 3: Update Word settings.xml
tmp_zip = doc_path + '.zip'
shutil.copy2(doc_path, tmp_zip)

with zipfile.ZipFile(tmp_zip, 'r') as zin:
    files = {item.filename: zin.read(item.filename) for item in zin.infolist()}

settings_xml = files['word/settings.xml'].decode('utf-8')
settings_xml = settings_xml.replace('DuLieuMerge$', '01_DA_BAN_HANH$')
files['word/settings.xml'] = settings_xml.encode('utf-8')

with zipfile.ZipFile(doc_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for filename, content in files.items():
        zout.writestr(filename, content)

os.remove(tmp_zip)
print("MAUTONG HOP.docx updated with 01_DA_BAN_HANH$ connection!")
