import docx
import os
import re

def search_text_in_docx(p, kw):
    if not os.path.exists(p):
        print(f"File not found: {p}")
        return
    doc = docx.Document(p)
    full = ' '.join([para.text for para in doc.paragraphs] + [c.text for t in doc.tables for r in t.rows for c in r.cells])
    matches = re.findall(rf'.{{0,40}}{kw}.{{0,40}}', full, re.IGNORECASE)
    if matches:
        print(f'=== Found "{kw}" in {p} ===')
        for m in matches[:5]:
            print('  ...', m.strip(), '...')
    else:
        print(f'No "{kw}" in {p}')

search_text_in_docx(r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAU TONG HOP HO SO DI.docx', 'giải thể')
search_text_in_docx(r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\BIEN BAN\MAUTONG HOP.docx', 'giải thể')
