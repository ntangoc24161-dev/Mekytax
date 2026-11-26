import docx

doc_thp = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\10. TAM HOANG PHAT\0. BIEN BAN\TONG HOP.docx")
print("THP Paragraphs around P265-P360 (BBKT):")
for i in range(270, 310):
    t = doc_thp.paragraphs[i].text.strip()
    if t:
        print(f"THP P{i}: {t[:120]}")

print("\nTHP Paragraphs around P380-P420 (BCKQ):")
for i in range(380, 420):
    t = doc_thp.paragraphs[i].text.strip()
    if t:
        print(f"THP P{i}: {t[:120]}")
