import docx

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

def check_p_balance(p, name):
    begins = len(p._element.xpath('.//w:fldChar[@w:fldCharType="begin"]'))
    ends = len(p._element.xpath('.//w:fldChar[@w:fldCharType="end"]'))
    seps = len(p._element.xpath('.//w:fldChar[@w:fldCharType="separate"]'))
    if begins != ends or begins != seps:
        print(f"Mismatch in {name}: begins={begins}, seps={seps}, ends={ends}")
        return 1
    return 0

mismatches = 0
for i, p in enumerate(doc.paragraphs):
    mismatches += check_p_balance(p, f"P{i}")

for ti, t in enumerate(doc.tables):
    for ri, row in enumerate(t.rows):
        for ci, cell in enumerate(row.cells):
            for pi, p in enumerate(cell.paragraphs):
                mismatches += check_p_balance(p, f"T{ti}_R{ri}_C{ci}_P{pi}")

print(f"Balance check finished. Total mismatches: {mismatches}")
