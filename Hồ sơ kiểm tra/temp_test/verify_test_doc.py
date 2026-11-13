import docx

doc = docx.Document(r'temp_test\MAUTONG_HOP_test.docx')

for t_idx, t in enumerate(doc.tables):
    first_row_text = ' | '.join(c.text.strip().replace('\n', ' ') for c in t.rows[0].cells)
    if any(k in first_row_text for k in ['CỘNG HÒA', 'THUẾ', 'TỔNG CỤC', 'ĐOÀN KIỂM TRA', 'TỔ KIỂM TRA']):
        print(f"\n=== Table {idx} (rows: {len(t.rows)}) ===" if 'idx' in locals() else f"\n=== Table {t_idx} (rows: {len(t.rows)}) ===")
        for r_i, r in enumerate(t.rows):
            for c_i, c in enumerate(r.cells):
                txt = c.text.strip().replace('\n', ' / ')
                xml = c._tc.xml
                v_line = xml.count('v:line')
                v_shape = xml.count('v:shape')
                print(f"  R{r_i}C{c_i}: '{txt}' (v:line={v_line}, v:shape={v_shape})")
