import openpyxl

excel_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
wb = openpyxl.load_workbook(excel_path, keep_vba=True, data_only=False)

print("Existing defined names in workbook:")
for dn in wb.defined_names:
    print(dn.name, dn.attr_text)
