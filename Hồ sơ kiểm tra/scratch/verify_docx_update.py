import docx

doc = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP.docx")
print("Check BBKT text around P323-P335:")
for i in range(323, 335):
    print(f"P{i}: {doc.paragraphs[i].text.strip()}")

print("\nCheck BCKQ text around P400-P408:")
for i in range(400, 408):
    print(f"P{i}: {doc.paragraphs[i].text.strip()}")

print("\nCheck QDXP text around P595-P605:")
for i in range(595, 605):
    print(f"P{i}: {doc.paragraphs[i].text.strip()}")
