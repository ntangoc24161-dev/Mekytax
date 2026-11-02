import zipfile
import re

with zipfile.ZipFile(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx") as z:
    xml = z.read("word/document.xml").decode("utf8")

fields = re.findall(r'MERGEFIELD\s+"?([^"\s<>]+)"?', xml)
print("Unique MERGEFIELDs:", sorted(set(fields)))
for f in sorted(set(fields)):
    print(f"{f}: {fields.count(f)}")
