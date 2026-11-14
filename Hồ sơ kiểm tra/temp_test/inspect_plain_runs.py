import docx

doc = docx.Document(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx")

targets = [
    ("P261", doc.paragraphs[261]),
    ("P265", doc.paragraphs[265]),
    ("P313", doc.paragraphs[313]),
    ("P319", doc.paragraphs[319]),
    ("T10_R0_C0_P6", doc.tables[10].rows[0].cells[0].paragraphs[6]),
    ("T10_R0_C1_P6", doc.tables[10].rows[0].cells[1].paragraphs[6]),
    ("T16_R2_C0_P6", doc.tables[16].rows[2].cells[0].paragraphs[6]),
    ("T16_R2_C1_P6", doc.tables[16].rows[2].cells[1].paragraphs[6]),
    ("T16_R2_C2_P6", doc.tables[16].rows[2].cells[2].paragraphs[6]),
    ("T18_R1_C1_P0", doc.tables[18].rows[1].cells[1].paragraphs[0]),
    ("T19_R1_C1_P0", doc.tables[19].rows[1].cells[1].paragraphs[0]),
    ("T20_R1_C1_P0", doc.tables[20].rows[1].cells[1].paragraphs[0]),
    ("T21_R1_C1_P0", doc.tables[21].rows[1].cells[1].paragraphs[0]),
    ("T22_R2_C0_P6", doc.tables[22].rows[2].cells[0].paragraphs[6]),
    ("T22_R2_C1_P6", doc.tables[22].rows[2].cells[1].paragraphs[6]),
    ("T22_R2_C2_P6", doc.tables[22].rows[2].cells[2].paragraphs[6]),
    ("T24_R1_C1_P0", doc.tables[24].rows[1].cells[1].paragraphs[0]),
    ("T24_R2_C1_P0", doc.tables[24].rows[2].cells[1].paragraphs[0]),
    ("T25_R1_C1_P0", doc.tables[25].rows[1].cells[1].paragraphs[0]),
    ("T25_R2_C1_P0", doc.tables[25].rows[2].cells[1].paragraphs[0]),
    ("T26_R1_C1_P0", doc.tables[26].rows[1].cells[1].paragraphs[0]),
    ("T26_R2_C1_P0", doc.tables[26].rows[2].cells[1].paragraphs[0]),
    ("T27_R1_C1_P0", doc.tables[27].rows[1].cells[1].paragraphs[0]),
    ("T27_R2_C1_P0", doc.tables[27].rows[2].cells[1].paragraphs[0]),
    ("T31_R1_C0_P6", doc.tables[31].rows[1].cells[0].paragraphs[6]),
    ("T31_R1_C1_P6", doc.tables[31].rows[1].cells[1].paragraphs[6]),
]

for name, p in targets:
    print(f"=== {name} ===")
    print("Full text:", repr(p.text))
    for r in p.runs:
        print("  run:", repr(r.text))
