import zipfile
import re

with zipfile.ZipFile(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx") as z:
    xml = z.read("word/document.xml").decode("utf8")

print("1. 'quản lý kinh doanh' raw count:", xml.count("quản lý kinh doanh"))
print("   'Quản lý kinh doanh' raw count:", xml.count("Quản lý kinh doanh"))
# check if broken across runs
pat = re.compile(r"quản\s+lý\s+kinh\s+doanh", re.I)
print("   regex match count:", len(pat.findall(xml)))

# Let's inspect each match of QLKD
for m in pat.finditer(xml):
    start = max(0, m.start() - 40)
    end = min(len(xml), m.end() + 40)
    print("   MATCH:", xml[start:end])

print("\n2. TCS3 raw occurrences:")
pat_tcs3 = re.compile(r"[^\<\>]*QĐ-TCS3[^\<\>]*")
for m in pat_tcs3.finditer(xml):
    print("   TCS3:", repr(m.group(0)))
