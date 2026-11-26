import docx

doc_thp = docx.Document(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\10. TAM HOANG PHAT\0. BIEN BAN\TONG HOP.docx")
print("THP BBXN around P195-P260:")
for i in range(195, 260):
    t = doc_thp.paragraphs[i].text.strip()
    if t and any(k in t.upper() for k in ["XÁC NHẬN", "NỘI DUNG", "KẾT QUẢ", "BÀ", "ÔNG", "CHI PHÍ", "TRƯỞNG ĐOÀN", "THÀNH VIÊN"]):
        print(f"P{i}: {t}")

print("\n--- Check Table 12 to 28 in THP ---")
for idx in [12, 13, 14, 15, 18, 19, 20, 21]:
    if idx < len(doc_thp.tables):
        t = doc_thp.tables[idx]
        print(f"Table {idx} rows: {len(t.rows)}, cols: {len(t.columns)}")
        for r_idx in range(min(3, len(t.rows))):
            row_txt = [c.text.strip().replace('\n', ' ') for c in t.rows[r_idx].cells]
            print(f"  R{r_idx}: {' | '.join(row_txt)}")
