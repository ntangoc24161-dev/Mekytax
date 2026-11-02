import re
import zipfile

with zipfile.ZipFile(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx") as z:
    xml = z.read("word/document.xml").decode("utf8")

# Let's inspect the entire field block around MERGEFIELD "KT"
pos = xml.find('MERGEFIELD "KT"')
# Find <w:r that contains pos
r2_start = xml.rfind('<w:r ', 0, pos)
r2_end = xml.find('</w:r>', pos) + len('</w:r>')
print("Run 2 (instrText):", repr(xml[r2_start:r2_end]))

# Find preceding run (Run 1: begin)
r1_end = r2_start
r1_start = xml.rfind('<w:r ', 0, r1_end)
print("Run 1 (begin):", repr(xml[r1_start:r1_end]))

# Find succeeding run (Run 3: end)
r3_start = r2_end
r3_end = xml.find('</w:r>', r3_start) + len('</w:r>')
print("Run 3 (end):", repr(xml[r3_start:r3_end]))

full_field = xml[r1_start:r3_end]
print("Full field XML:\n", full_field)
