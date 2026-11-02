# -*- coding: utf-8 -*-
import zipfile, re, sys
sys.stdout.reconfigure(encoding="utf-8")
xml = zipfile.ZipFile(
    r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\temp_test\sth_inspect\TONG HOP.docx"
).read("word/document.xml").decode("utf8")

text = re.sub(r"</w:p>", "\n", xml)
text = re.sub(r"<w:tab[^/]*/>", "\t", text)
text = re.sub(r"<[^>]+>", "", text)
lines = text.splitlines()

print("===== lines with So / QD =====")
for i, l in enumerate(lines, 1):
    if any(k in l for k in ["Số:", "SỐ", "/QĐ-", "QĐ-TCS3", "QĐ-XPHC", "QĐ-XPHC"]):
        print(f"{i:4d}|{l.strip()[:220]}")

print("\n===== XML windows around /QĐ-TCS3 =====")
idx = 0
n = 0
while True:
    j = xml.find("/QĐ-TCS3", idx)
    if j < 0:
        break
    n += 1
    frag = xml[max(0, j-180): j+80]
    frag = re.sub(r"<[^>]+>", "", frag).replace("\n", " ")
    print(f"{n:02d} @{j}: {frag[-200:]}")
    idx = j + 1
print("total", n)

print("\n===== XML windows around Số: =====")
idx = 0
n = 0
while True:
    j = xml.find(">Số:", idx)
    if j < 0:
        j = xml.find(">Số: ", idx) 
    if j < 0:
        break
    n += 1
    frag = xml[j: j+400]
    plain = re.sub(r"<[^>]+>", "", frag)
    print(f"{n:02d} @{j}: {plain[:120]!r}")
    idx = j + 4

print("\n===== MERGEFIELD KT xml =====")
for m in re.finditer(r".{0,80}MERGEFIELD[^<]*KT.{0,200}", xml):
    print("---")
    print(m.group(0)[:300])
