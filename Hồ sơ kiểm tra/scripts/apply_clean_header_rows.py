import os
import shutil
import docx

orig_path = r'templates\MAUTONG HOP.docx'
bak_path = r'templates\MAUTONG HOP.bak-truoc-xoa-hang-trong-20261113.docx'

# 1. Backup original
shutil.copy2(orig_path, bak_path)
print(f"Created backup: {bak_path}")

# 2. Load doc
doc = docx.Document(orig_path)

# Tables to delete Row 1 (the redundant empty/duplicate row):
tables_to_delete_row1 = [1, 3, 5, 6, 34, 37, 42, 46]

for t_idx in tables_to_delete_row1:
    t = doc.tables[t_idx]
    row_to_del = t.rows[1]
    tr = row_to_del._tr
    tr.getparent().remove(tr)
    print(f"Deleted redundant row 1 in Table {t_idx} (now {len(t.rows)} rows)")

# Table 33: Remove redundant shapes in R1C0
t33 = doc.tables[33]
c0 = t33.rows[1].cells[0]
for p in c0.paragraphs:
    for r in list(p._p.xpath('.//w:r[.//w:pict]')):
        r.getparent().remove(r)
print(f"Cleaned redundant shapes in Table 33 R1C0")

# Save document
doc.save(orig_path)
print("Saved templates/MAUTONG HOP.docx successfully.")
