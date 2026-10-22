import fitz
import pytesseract
from PIL import Image
import io
import os

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
doc = fitz.open(r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI\5549.pdf')

full_text = []
for idx, page in enumerate(doc):
    pix = page.get_pixmap(dpi=150)
    img = Image.open(io.BytesIO(pix.tobytes('png')))
    txt = pytesseract.image_to_string(img, lang='eng')
    full_text.append(f"=== Trang {idx+1} ===\n" + txt)

out_path = r'D:\HSO CONG AN FILE MD\5549.md'
frontmatter = """---
type: source
source_kind: ocr
status: unprocessed
created: 2026-10-21
updated: 2026-10-21
title: "5549"
so_cv: "145/CSKT-Đ2"
ngay_cv: "2026-10-01"
co_quan_ban_hanh: "Cơ quan Cảnh sát điều tra - Công an thành phố Đồng Nai"
trich_yeu: "V/v đề nghị hủy bỏ ngăn chặn, cho phép thực hiện kê khai bổ sung thuế GTGT, thuế TNDN"
tags: [ocr, pdf, markdown]
---

# CÔNG VĂN SỐ 145/CSKT-Đ2 (FILE 5549.PDF) - CƠ QUAN CSĐT CÔNG AN THÀNH PHỐ ĐỒNG NAI

"""

with open(out_path, 'w', encoding='utf-8') as f:
    f.write(frontmatter + "\n\n".join(full_text))

print('Done writing 5549.md! Total pages:', len(full_text), 'Size:', os.path.getsize(out_path))
