import docx
import os

doc_path = r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAU TONG HOP HO SO DI.docx'
doc = docx.Document(doc_path)

print(f"Total paragraphs: {len(doc.paragraphs)}")
print(f"Total tables: {len(doc.tables)}")

for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if any(txt.startswith(kw) for kw in ['TỜ TRÌNH', 'QUYẾT ĐỊNH', 'PHÂN TÍCH', 'THƯ TRAO ĐỔI', 'GIẤY XÁC NHẬN', 'KẾ HOẠCH GIÁM SÁT']):
        print(f"P{i:03d}: {txt[:70]}")
