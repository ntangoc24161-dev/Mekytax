import os
import json
import re
import pymupdf
from paddleocr import PaddleOCR

pdf_folder = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\HO SO MOI"
out_folder = r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo\scratch\ocr_results"
os.makedirs(out_folder, exist_ok=True)

scan_files = [
    '5645.pdf',
    'cv 9955.pdf',
    '8835-CV-05112026.pdf',
    '6210.pdf',
    '12215.pdf',
    '8746-CV-04112026.pdf',
    '6760TCS15-QLDN1 14445.pdf',
    '6136.pdf'
]

ocr = PaddleOCR(lang='vi')

summary = []

for file_name in scan_files:
    pdf_path = os.path.join(pdf_folder, file_name)
    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)
    
    # Render page 1
    page1 = doc[0]
    pix1 = page1.get_pixmap(dpi=150)
    img_path1 = os.path.join(out_folder, f"{file_name}_p1.png")
    pix1.save(img_path1)
    
    res = ocr.ocr(img_path1)
    lines = []
    if res and len(res) > 0 and 'rec_texts' in res[0]:
        lines = res[0]['rec_texts']
        
    txt_path = os.path.join(out_folder, f"{file_name}_p1_clean.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
        
    full_text = " \n ".join(lines)
    
    # Try extract so cv, ngay cv, co quan, mst
    so_cv_match = re.search(r'(?:Số|S6|So)[:\s]*([0-9A-Za-z\-/_\.]+)', full_text, re.IGNORECASE)
    ngay_cv_match = re.search(r'ngày\s+([0-9]{1,2})\s+tháng\s+([0-9]{1,2})\s+năm\s+([0-9]{4})', full_text, re.IGNORECASE)
    mst_matches = re.findall(r'\b[0-9]{10}(?:-[0-9]{3})?\b', full_text)
    
    doc.close()
    
    summary.append({
        'file': file_name,
        'pages': total_pages,
        'so_cv': so_cv_match.group(1) if so_cv_match else 'Chưa rõ',
        'ngay_cv': f"{ngay_cv_match.group(1)}/{ngay_cv_match.group(2)}/{ngay_cv_match.group(3)}" if ngay_cv_match else 'Chưa rõ',
        'mst': list(set(mst_matches)),
        'lines_sample': lines[:8]
    })

out_json = os.path.join(out_folder, "summary_8files.json")
with open(out_json, "w", encoding="utf-8") as f:
    json.dump(summary, f, ensure_ascii=False, indent=2)

print("OCR_8FILES_DONE")
