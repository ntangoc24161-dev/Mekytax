import docx

doc = docx.Document(r'templates\MAUTONG HOP.docx')

for t_idx, t in enumerate(doc.tables):
    for r_idx, r in enumerate(t.rows):
        for c_idx, c in enumerate(r.cells):
            xml = c._tc.xml
            if 'v:line' in xml or 'v:shape' in xml:
                text = c.text.strip().replace('\n', ' / ')
                v_line = xml.count('v:line')
                v_shape = xml.count('v:shape')
                print(f"Table {t_idx} R{r_idx}C{c_idx}: text='{text}' | v:line={v_line}, v:shape={v_shape}")
