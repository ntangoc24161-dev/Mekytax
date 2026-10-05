import os
import openpyxl
import docx

target = r"\\?\G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH"

# Check Tai khoan.docx
try:
    doc = docx.Document(os.path.join(target, "Tai khoan.docx"))
    print("=== Tai khoan.docx ===")
    for p in doc.paragraphs:
        if p.text.strip():
            print(" ", p.text.strip())
    for t in doc.tables:
        for r in t.rows:
            row_txt = [c.text.strip() for c in r.cells]
            print("  TABLE ROW:", " | ".join(row_txt))
except Exception as e:
    print("Error reading Tai khoan.docx:", e)

# Check Tai khoan ngan hang APT, DK.xlsx
try:
    wb = openpyxl.load_workbook(os.path.join(target, "Tai khoan ngan hang APT, DK.xlsx"), data_only=True)
    print("\n=== Tai khoan ngan hang APT, DK.xlsx ===")
    for name in wb.sheetnames:
        print(" Sheet:", name)
        ws = wb[name]
        for r in list(ws.iter_rows(values_only=True))[:10]:
            if any(r):
                print("  ", [str(x) if x is not None else '' for x in r])
except Exception as e:
    print("Error reading xlsx:", e)
