# -*- coding: utf-8 -*-
import zipfile, re, sys
sys.stdout.reconfigure(encoding="utf-8")
xml = zipfile.ZipFile(
    r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\temp_test\sth_inspect\TONG HOP.docx"
).read("word/document.xml").decode("utf8")

def around(needle, n=1, before=80, after=120):
    idx = 0
    found = 0
    while found < n:
        j = xml.find(needle, idx)
        if j < 0:
            print("NOT FOUND", needle[:60])
            return
        found += 1
        print(f"\n--- {needle[:40]!r} #{found} @{j} ---")
        print(xml[j-before:j+after])
        idx = j + len(needle)

# header So
around("Số:         ", n=3, before=30, after=250)
around("Số:           ", n=2, before=30, after=250)

print("\n===== SỐ QĐ bao cao =====")
j = xml.find("SỐ              /QĐ-TCS3")
print("found contiguous", j)
j2 = xml.find(">SỐ")
print("first >SỐ", j2, xml[j2:j2+300] if j2>=0 else "")

print("\n===== nhật ký /QĐ-TCS3 cells =====")
for m in re.finditer(r">\s*/QĐ-TCS3", xml):
    print(xml[m.start()-80:m.end()+80])
    print("----")

print("\n===== Phan The Hanh signature runs =====")
idx = 0
n = 0
while True:
    j = xml.find("Phan Thế Hanh", idx)
    if j < 0:
        break
    n += 1
    frag = xml[j-60:j+80]
    print(f"{n} @{j}: {frag}")
    idx = j + 1

print("\n===== quan ly kinh doanh split? =====")
phrase = "quản lý kinh doanh"
pat = "".join(re.escape(c) + r"(?:(?:</w:t>)?(?:<[^>]*>)*?(?:<w:t[^>]*>)?)?" for c in phrase)
matches = list(re.finditer(pat, xml))
print("regex matches", len(matches))
print("contiguous", xml.count(phrase))
for i, m in enumerate(matches, 1):
    g = m.group(0)
    split = "</w:t>" in g
    print(f"{i:02d} split={split} {g[:80]!r}")
