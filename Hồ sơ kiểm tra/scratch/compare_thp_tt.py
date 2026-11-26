import docx

doc_thp = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\10. TAM HOANG PHAT\0. BIEN BAN\TONG HOP.docx")
print("Total paragraphs THP:", len(doc_thp.paragraphs))
print("Total tables THP:", len(doc_thp.tables))

doc_tt = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP.docx")
print("Total paragraphs TT:", len(doc_tt.paragraphs))
print("Total tables TT:", len(doc_tt.tables))

print("\n--- TT HEADINGS ---")
for i, p in enumerate(doc_tt.paragraphs):
    txt = p.text.strip()
    if any(k in txt.upper() for k in ["BIÊN BẢN KIỂM TRA", "BÁO CÁO KẾT QUẢ", "QUYẾT ĐỊNH XỬ PHẠT", "TỜ TRÌNH", "XÁC NHẬN SỐ LIỆU", "KẾ HOẠCH KIỂM TRA"]):
        print(f"P{i}: {txt}")
