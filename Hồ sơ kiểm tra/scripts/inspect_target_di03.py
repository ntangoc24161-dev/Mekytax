import docx
import os

target_folder = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\DANG KIEM TRA\KE HOACH\VAN PHAT THANH\0. BIEN BAN'
doc_di03 = docx.Document(os.path.join(target_folder, 'DI-03_DI-03.docx'))

print("=== Paragraphs in DI-03_DI-03.docx ===")
for i, p in enumerate(doc_di03.paragraphs):
    print(f"{i:02d}: {p.text}")
