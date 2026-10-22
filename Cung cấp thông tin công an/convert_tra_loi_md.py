import fitz
import pytesseract
from PIL import Image
import io
import os
import pypdf

files = [
    ('TRA LOI CV 1808_CA HAI PHONG.pdf', 'TRA LOI CV 1808_CA HAI PHONG.md', 'V/v cung cấp thông tin, hồ sơ DN theo Công văn 1808/YC-ANĐT-Đ2'),
    ('TRA LOI CV 29931_CA DONG NAI.pdf', 'TRA LOI CV 29931_CA DONG NAI.md', 'V/v cung cấp thông tin, hồ sơ DN theo Công văn 29931/YC-CSHS-CSKT')
]

cv_moi = r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\CV MOI'
out_dir = r'D:\HSO CONG AN FILE MD'

for pdf_name, md_name, trich_yeu in files:
    pdf_path = os.path.join(cv_moi, pdf_name)
    reader = pypdf.PdfReader(pdf_path)
    pages_text = []
    for i, page in enumerate(reader.pages):
        t = page.extract_text() or ''
        pages_text.append(f"## Trang {i+1}\n\n" + t)
    
    frontmatter = f"""---
type: source
source_kind: digital-pdf
status: unprocessed
created: 2026-10-21
updated: 2026-10-21
title: "{os.path.splitext(pdf_name)[0]}"
trich_yeu: "{trich_yeu}"
tags: [tra-loi, pdf, markdown]
---

# {os.path.splitext(pdf_name)[0]}

"""
    out_file = os.path.join(out_dir, md_name)
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(frontmatter + "\n\n".join(pages_text))
    print(f"Done writing {md_name} (Size: {os.path.getsize(out_file)})")
