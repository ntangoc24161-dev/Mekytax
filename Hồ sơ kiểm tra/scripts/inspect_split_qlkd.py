import zipfile
import re

with zipfile.ZipFile(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx") as z:
    xml = z.read("word/document.xml").decode("utf8")

matches = list(re.finditer(r'kinh doanh', xml))
for idx in [14, 15, 16, 17, 18]:
    pos = matches[idx].start()
    p_start = xml.rfind("<w:p ", 0, pos)
    if p_start == -1:
        p_start = xml.rfind("<w:p>", 0, pos)
    p_end = xml.find("</w:p>", pos) + len("</w:p>")
    p_xml = xml[p_start:p_end]
    p_text = re.sub(r'<[^>]+>', '', p_xml)
    print(f"[{idx}] TEXT: {repr(p_text)}")
    print(f"XML snippet:\n{p_xml}\n" + "-"*40)
