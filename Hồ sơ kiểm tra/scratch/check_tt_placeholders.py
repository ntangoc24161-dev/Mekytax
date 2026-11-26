import docx
import re
import shutil

doc_path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP.docx"
doc_bak = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP_backup_before_hoso_ve.docx"
shutil.copyfile(doc_path, doc_bak)

doc = docx.Document(doc_path)

# Information:
# QD: 7516/QĐ-TCS3 ngay 20/11/2026 (hoặc ngày công bố 26/11/2026, kết thúc 27/11/2026, ký BBKT 30/11/2026)
# Số QĐKT trên sổ: 7516/QĐ-TCS3
# Số tiền:
# Truy thu TNDN 2025: 14.023.597 d
# Truy thu GTGT: 0 d
# Tong truy thu: 14.023.597 d
# Cham nop TNDN (den 30/11/2026): 1.022.320 d
# Cham nop GTGT: 0 d
# Tong cham nop: 1.022.320 d
# Phat khai sai 20%: 2.804.719 d
# Phat VPHC: 0 d
# Tong phat: 2.804.719 d
# Tong cong nghia vu: 17.850.636 d (14.023.597 + 1.022.320 + 2.804.719)
# Bang chu: Mười bảy triệu tám trăm năm mươi nghìn sáu trăm ba mươi sáu đồng.

print("Total paragraphs in TT TONG HOP:", len(doc.paragraphs))

# Search for any remaining placeholder or template text in BBKT, BCKQ, QDXP
for idx, p in enumerate(doc.paragraphs):
    txt = p.text
    if "..../..../2026" in txt or "               /QĐ-TCS3" in txt:
        print(f"P{idx}: {txt[:100]}")
