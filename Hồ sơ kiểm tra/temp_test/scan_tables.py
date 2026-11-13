import docx

doc = docx.Document(r'templates\MAUTONG HOP.docx')

for idx, t in enumerate(doc.tables):
    has_header_text = False
    for r in t.rows:
        txt = ' '.join(c.text for c in r.cells)
        if any(w in txt for w in ['CỘNG HÒA', 'THUẾ', 'TỔNG CỤC', 'ĐOÀN KIỂM TRA', 'TỔ KIỂM TRA']):
            has_header_text = True
            break
    if has_header_text:
        print(f"*** TABLE {idx}: {len(t.rows)} rows, {len(t.columns)} cols ***")
        for r_i, r in enumerate(t.rows):
            c_info = []
            for c_i, c in enumerate(r.cells):
                txt = c.text.strip().replace('\n', ' / ')
                v_line = c._tc.xml.count('v:line')
                c_info.append(f"C{c_i}: '{txt}' (lines:{v_line})")
            print(f"  Row {r_i}: " + " | ".join(c_info))
