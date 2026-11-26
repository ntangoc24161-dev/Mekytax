import docx

doc_thp = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\10. TAM HOANG PHAT\0. BIEN BAN\TONG HOP.docx")
print("THP Paragraphs around P570-P630 (QD XPVPHC):")
for i in range(570, 630):
    t = doc_thp.paragraphs[i].text.strip()
    if t and any(k in t.upper() for k in ["QUYẾT ĐỊNH", "ĐIỀU 1", "ĐIỀU 2", "ĐIỀU 3", "XỬ PHẠT", "TRUY THU", "CHẬM NỘP", "LÝ DO"]):
        print(f"P{i}: {t}")
