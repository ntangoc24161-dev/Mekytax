import zipfile
import re

with zipfile.ZipFile(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx") as z:
    xml = z.read("word/document.xml").decode("utf8")

for target in ["KT", "CCCDKT", "Tài_khoản_số"]:
    print(f"=== PARAGRAPHS FOR TARGET: {target} ===")
    matches = [m.start() for m in re.finditer(rf'MERGEFIELD\s+"?{target}"?', xml)]
    for idx, pos in enumerate(matches):
        p_start = xml.rfind("<w:p ", 0, pos)
        if p_start == -1:
            p_start = xml.rfind("<w:p>", 0, pos)
        p_end = xml.find("</w:p>", pos) + len("</w:p>")
        p_xml = xml[p_start:p_end]
        p_text = re.sub(r'<[^>]+>', '', p_xml)
        print(f"[{idx+1}] TEXT: {repr(p_text)}")
        print(f"XML snippet:\n{p_xml}\n" + "="*40)
