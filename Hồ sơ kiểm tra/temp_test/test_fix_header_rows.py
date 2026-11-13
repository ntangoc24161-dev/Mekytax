import docx
import shutil

src = r'templates\MAUTONG HOP.docx'
test_dst = r'temp_test\MAUTONG_HOP_test.docx'
shutil.copy2(src, test_dst)

doc = docx.Document(test_dst)

# Tables to delete Row 1:
# Table 1: QDKT (has duplicate Row 1)
# Table 3: QDGS (has empty Row 1 with lines)
# Table 5: KH Phan tich (has empty Row 1 with shape)
# Table 6: Thu trao doi (has empty Row 1 with line)
# Table 34: KH Giam sat (has empty Row 1)
# Table 37: BB Lam viec giam sat (has empty Row 1 with shape)
# Table 42: QD XPHC (has duplicate Row 1)
# Table 46: VB tham gia y kien phap ly (has duplicate Row 1)

tables_to_delete_row1 = [1, 3, 5, 6, 34, 37, 42, 46]

for t_idx in tables_to_delete_row1:
    t = doc.tables[t_idx]
    print(f"Table {t_idx} before: {len(t.rows)} rows")
    row_to_del = t.rows[1]
    tr = row_to_del._tr
    tr.getparent().remove(tr)
    print(f"Table {t_idx} after: {len(t.rows)} rows")

# Table 33: Row 1 Col 0 has extra shapes in empty paragraph
t33 = doc.tables[33]
c0 = t33.rows[1].cells[0]
print(f"Table 33 R1C0 shapes before: v:shape={c0._tc.xml.count('v:shape')}")
# Remove shapes / pict from paragraphs of R1C0
for p in c0.paragraphs:
    for r in list(p._p.xpath('.//w:r[.//w:pict]')):
        r.getparent().remove(r)
print(f"Table 33 R1C0 shapes after: v:shape={c0._tc.xml.count('v:shape')}")

doc.save(test_dst)
print("Saved test document successfully.")
