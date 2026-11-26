import docx

doc_tt = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP.docx")
print("Check table contents around BBKT (P265):")
for idx, table in enumerate(doc_tt.tables):
    text_sample = " | ".join([cell.text.strip() for row in table.rows[:2] for cell in row.cells[:3]])
    if any(k in text_sample.upper() for k in ["KẾT QUẢ KIỂM TRA", "TỔNG SỐ THUẾ", "DOANH THU", "CHI PHÍ", "XỬ LÝ VỀ THUẾ"]):
        print(f"Table {idx}: {text_sample[:100]}")
