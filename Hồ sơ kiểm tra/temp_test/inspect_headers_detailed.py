import docx

doc = docx.Document(r'templates\MAUTONG HOP.docx')

# Find what heading/title follows each table
for idx, t in enumerate(doc.tables):
    # Find following paragraph
    # t._tbl is an oxml element. We can find the next sibling element in body
    p_next = t._tbl.getnext()
    next_text = ""
    while p_next is not None and not next_text:
        if p_next.tag.endswith('p'):
            next_text = p_next.text.strip()
        p_next = p_next.getnext()
    
    first_row_text = ' | '.join(c.text.strip().replace('\n', ' ') for c in t.rows[0].cells)
    if any(k in first_row_text for k in ['CỘNG HÒA', 'THUẾ', 'TỔNG CỤC', 'ĐOÀN KIỂM TRA', 'TỔ KIỂM TRA']):
        print(f"Table {idx} ({len(t.rows)} rows): Next text -> '{next_text[:60]}'")
        for r_i, r in enumerate(t.rows):
            for c_i, c in enumerate(r.cells):
                v_lines = c._tc.xml.count('v:line')
                picts = c._tc.xml.count('w:pict')
                borders = c._tc.xml.count('w:bottom')
                text = c.text.strip().replace('\n', ' ')
                print(f"  R{r_i}C{c_i}: '{text}' [v:line={v_lines}, pict={picts}, bottom_border={borders}]")
