import docx

doc = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\10. TAM HOANG PHAT\0. BIEN BAN\TONG HOP.docx")
print("Number of paragraphs:", len(doc.paragraphs))
print("Number of tables:", len(doc.tables))
print("\n--- FIRST 20 PARAGRAPHS ---")
for i, p in enumerate(doc.paragraphs[:25]):
    if p.text.strip():
        print(f"P{i}: {p.text.strip()}")

print("\n--- SCAN FOR HEADINGS / SECTIONS ---")
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if any(k in txt.upper() for k in ["BIÊN BẢN KIỂM TRA", "BÁO CÁO KẾT QUẢ", "QUYẾT ĐỊNH XỬ PHẠT", "TỜ TRÌNH", "XÁC NHẬN SỐ LIỆU"]):
        print(f"P{i}: {txt}")
