import zipfile
import re
import openpyxl

# 1. Read Word template MERGEFIELDs
with zipfile.ZipFile('templates/MAUTONG HOP.docx') as z:
    xml = z.read('word/document.xml').decode('utf-8')
    fields = set(re.findall(r'MERGEFIELD\s+"?([^"\s\\]+)', xml))
    print(f"Template has {len(fields)} unique MERGEFIELDs:")
    print(sorted(list(fields)))

# 2. Check settings.xml in Word template
with zipfile.ZipFile('templates/MAUTONG HOP.docx') as z:
    settings_xml = z.read('word/settings.xml').decode('utf-8')
    print("\nSettings mailMerge snippet:")
    if '<w:mailMerge>' in settings_xml:
        start = settings_xml.find('<w:mailMerge>')
        end = settings_xml.find('</w:mailMerge>') + len('</w:mailMerge>')
        print(settings_xml[start:end])
    else:
        print("No <w:mailMerge> in settings.xml")

# 3. Check Excel file
excel_path = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm'
wb = openpyxl.load_workbook(excel_path, keep_vba=True, data_only=False)
ws = wb['01_DA_BAN_HANH']

print(f"\n01_DA_BAN_HANH dimensions: max_row={ws.max_row}, max_column={ws.max_column}")
for r in range(1, 6):
    vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
    print(f"Row {r} (non-empty count {sum(1 for v in vals if v is not None)}): {vals[:10]}")
