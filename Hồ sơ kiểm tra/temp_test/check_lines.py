import docx

doc = docx.Document(r'templates\MAUTONG HOP.docx')

check_tables = [5, 8, 33, 36, 41, 42, 44]

for t_idx in check_tables:
    t = doc.tables[t_idx]
    print(f"=== CHECK TABLE {t_idx} ({len(t.rows)} rows) ===")
    for r_idx, r in enumerate(t.rows):
        for c_idx, c in enumerate(r.cells):
            print(f"  Row {r_idx} Col {c_idx}:")
            for p_idx, p in enumerate(c.paragraphs):
                p_xml = p._p.xml
                has_vline = 'v:line' in p_xml
                has_border = 'w:bottom' in p_xml or 'w:pBdr' in p_xml
                has_pict = 'w:pict' in p_xml
                print(f"    P{p_idx}: text='{p.text}' | v:line={has_vline} | pBdr={has_border} | pict={has_pict}")
                if has_vline or has_border:
                    # print snippet
                    for line in p_xml.splitlines():
                        if any(k in line for k in ['v:line', 'w:bottom', 'w:pBdr', 'stroke']):
                            print(f"      {line.strip()[:100]}")
