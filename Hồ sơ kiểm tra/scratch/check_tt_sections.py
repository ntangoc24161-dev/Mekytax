import docx

doc = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP.docx")
print("Check BBKT text around P300-P345:")
for i in range(295, 345):
    t = doc.paragraphs[i].text.strip()
    if t:
        print(f"P{i}: {t}")

print("\nCheck BCKQ text around P400-P415:")
for i in range(398, 415):
    t = doc.paragraphs[i].text.strip()
    if t:
        print(f"P{i}: {t}")

print("\nCheck QDXP text around P595-P620:")
for i in range(595, 620):
    t = doc.paragraphs[i].text.strip()
    if t:
        print(f"P{i}: {t}")
