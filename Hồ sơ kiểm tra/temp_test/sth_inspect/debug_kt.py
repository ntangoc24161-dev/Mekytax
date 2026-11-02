# -*- coding: utf-8 -*-
import re, zipfile, sys
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
xml = zipfile.ZipFile(
    Path(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\temp_test\sth_inspect\TONG HOP.docx")
).read("word/document.xml").decode("utf8")

fld = re.compile(
    r"<w:r[^>]*>"
    r"(?:<w:rPr>(?P<rpr>.*?)</w:rPr>)?"
    r'<w:fldChar w:fldCharType="begin"/>'
    r"</w:r>"
    r"<w:r[^>]*>"
    r"(?:<w:rPr>.*?</w:rPr>)?"
    r"<w:instrText[^>]*>\s*MERGEFIELD\s+\"?KT\"?\s*</w:instrText>"
    r"</w:r>"
    r"<w:r[^>]*>"
    r"(?:<w:rPr>.*?</w:rPr>)?"
    r'<w:fldChar w:fldCharType="end"/>'
    r"</w:r>",
    re.S,
)
print("matches", len(list(fld.finditer(xml))))
for i, m in enumerate(fld.finditer(xml), 1):
    g = m.group(0)
    print(f"\n#{i} len={len(g)} start={m.start()} has_tc={'</w:tc>' in g} has_p={'</w:p>' in g}")
    print(g[:200])
    print("...")
    print(g[-200:])

print("\n===== windows around instr KT =====")
for i, m in enumerate(re.finditer(r"<w:instrText[^>]*>\s*MERGEFIELD\s+\"?KT\"?\s*</w:instrText>", xml), 1):
    print(f"\nINSTR {i} @{m.start()}")
    print(xml[m.start()-400:m.end()+250])
    print("----")
