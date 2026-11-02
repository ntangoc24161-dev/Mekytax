# -*- coding: utf-8 -*-
import zipfile, re, sys
sys.stdout.reconfigure(encoding="utf-8")
xml = zipfile.ZipFile(
    r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\temp_test\sth_inspect\TONG HOP.docx"
).read("word/document.xml").decode("utf8")

frags = [
    "Số:         ",
    "Số:           ",
    "Căn cứ Quyết định kiểm tra số              /QĐ-TCS3 ngày ...../...../2026",
    "được thành lập theo Quyết định số              /QĐ-TCS3 ngày ...../...../2026.",
    "số              /QĐ-TCS3 ngày ..../..../2026",
    "SỎ              /QĐ-TCS3",
    "<w:t>              /QĐ-TCS3</w:t>",
    "<w:t>..../..../2026</w:t>",
    "……../QĐ-TCS3",
    "<w:t>…../…../202</w:t>",
    "quản lý kinh doanh",
    "-Chi phí quản lý kinh doanh",
    "- Doanh thu bán hàng và cung cấp dịch vụ",
    '<w:t xml:space="preserve">- Tài khoản số: </w:t>',
    "MERGEFIELD Tài_khoản_số",
    "Phan Thế Hanh",
    "MERGEFIELD",
    'MERGEFIELD "KT"',
    "MERGEFIELD KT",
    "Số:               /QĐ-XPHC",
    "..../..../2026-..../..../2026",
]
for f in frags:
    print(f"{xml.count(f):3d} | {f[:90]!r}")

print("\n--- instrText KT ---")
for m in re.finditer(r"<w:instrText[^>]*>[^<]*KT[^<]*</w:instrText>", xml):
    print(m.group(0)[:250])

print("\n--- around Tai khoan ---")
i = xml.find("Tài khoản số")
print(xml[i-200:i+500] if i>=0 else "NOT FOUND")
