import docx
import zipfile
import re
from pathlib import Path

SRC_XPHC = Path(r"output\3603859012_HO-SO-VE_son-tien-hai\QD XPHC Cong ty Son Tien Hai.docx")
TEST_XPHC = Path(r"temp_test\QD_XPHC_tested.docx")

with zipfile.ZipFile(SRC_XPHC) as z:
    xml = z.read("word/document.xml").decode("utf8")

QDKT = "7092/QĐ-TCS3"
NGAY = "26/10/2026"

old_ref = "số              /QĐ-TCS3 ngày ..../..../2026"
print("Old ref count in XPHC:", xml.count(old_ref))
xml = xml.replace(old_ref, f"số {QDKT} ngày {NGAY}")

with zipfile.ZipFile(SRC_XPHC) as zin, zipfile.ZipFile(TEST_XPHC, "w", zipfile.ZIP_DEFLATED) as zout:
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename == "word/document.xml":
            data = xml.encode("utf8")
        zout.writestr(it, data)

doc = docx.Document(str(TEST_XPHC))
print("Opened XPHC successfully!")
for i, p in enumerate(doc.paragraphs):
    if QDKT in p.text:
        print(f"P{i}: {p.text}")
