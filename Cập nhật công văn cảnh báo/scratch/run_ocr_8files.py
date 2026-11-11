import os
import sys
import json
import pymupdf
from paddleocr import PaddleOCR
from concurrent.futures import ThreadPoolExecutor

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

# Initialize OCR once
print("Initializing PaddleOCR model...")
ocr = PaddleOCR(lang='vi')

def ocr_first_page(file_name):
    try:
        pdf_path = os.path.join(pdf_folder, file_name)
        doc = pymupdf.open(pdf_path)
        page = doc[0]
        pix = page.get_pixmap(dpi=150)
        img_path = os.path.join(out_folder, f"{file_name}_p1.png")
        pix.save(img_path)
        doc.close()
        
        result = ocr.ocr(img_path)
        lines = []
        if result and result[0]:
            for line in result[0]:
                lines.append(line[1][0])
        
        res_file = os.path.join(out_folder, f"{file_name}_p1.txt")
        with open(res_file, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return file_name, True, len(lines), lines[:10]
    except Exception as e:
        return file_name, False, str(e), []

print("Running OCR sequentially for page 1 of 8 scan files...")
for f in scan_files:
    r = ocr_first_page(f)
    print(f"File: {r[0]} | Success: {r[1]} | Lines: {r[2]}")
    if r[3]:
        print("  Sample:", " | ".join(r[3][:6]))
