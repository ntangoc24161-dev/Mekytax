import zipfile, re

docx_path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP.docx"
with zipfile.ZipFile(docx_path, 'r') as z:
    s = z.read('word/settings.xml').decode('utf-8')
    m = re.search(r'<[^:]+:mailMerge.*?</[^:]+:mailMerge>', s, re.DOTALL)
    print(m.group(0))
