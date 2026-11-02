# -*- coding: utf-8 -*-
import sys
import re
import zipfile
import shutil
from pathlib import Path
import docx

SRC = Path(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx")
TEST_OUT = Path(r"temp_test\TONG_HOP_tested.docx")

QDKT = "7092/QĐ-TCS3"
QDGS = "7093/QĐ-TCS3"
NGAY = "26/10/2026"
KT = "Lê Hoài Phương"
TK = "6722949824 tại Ngân hàng TMCP Đầu tư và phát triển Việt Nam"

with zipfile.ZipFile(SRC) as z:
    xml = z.read("word/document.xml").decode("utf8")

print(f"Original XML len: {len(xml)}")

# 1. Replace 'quản lý kinh doanh' -> 'quản lý doanh nghiệp'
n_qlkd = xml.count("quản lý kinh doanh")
xml = xml.replace("quản lý kinh doanh", "quản lý doanh nghiệp")
print(f"Replaced contiguous 'quản lý kinh doanh': {n_qlkd} occurrences")

split_from = (
    '<w:t xml:space="preserve">Chi phí quản lý </w:t></w:r>'
    '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>kinh doanh</w:t></w:r>'
)
split_to = (
    '<w:t xml:space="preserve">Chi phí quản lý </w:t></w:r>'
    '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>doanh nghiệp</w:t></w:r>'
)
n_split = xml.count(split_from)
xml = xml.replace(split_from, split_to)
print(f"Replaced split 'Chi phí quản lý ' + 'kinh doanh': {n_split} occurrences")

# 2. Header QDKT: Table 1
rpr_t1 = (
    '<w:rPr><w:rFonts w:cs="Times New Roman"/><w:sz w:val="24"/>'
    '<w:szCs w:val="24"/><w:lang w:val="nl-NL"/></w:rPr>'
)
old_t1_hdr = (
    '<w:t xml:space="preserve">Số:         </w:t></w:r><w:r>'
    + rpr_t1
    + '<w:t xml:space="preserve">   </w:t></w:r>'
    + '<w:r w:rsidR="00D50A77" w:rsidRPr="009D7F91">'
    + rpr_t1
    + '<w:t xml:space="preserve">     /QĐ-TCS3</w:t>'
)
new_t1_hdr = (
    '<w:t xml:space="preserve">Số: </w:t></w:r>'
    + '<w:r w:rsidR="00D50A77" w:rsidRPr="009D7F91">'
    + rpr_t1
    + f'<w:t xml:space="preserve">{QDKT}</w:t>'
)
if old_t1_hdr in xml:
    xml = xml.replace(old_t1_hdr, new_t1_hdr, 1)
    print("Replaced Header QDKT in Table 1")

old_t1_date = "Trấn Biên, ngày       tháng       năm"
new_t1_date = f"Trấn Biên, ngày 26 tháng 10 năm 2026"
if old_t1_date in xml:
    xml = xml.replace(old_t1_date, new_t1_date, 1)
    print("Replaced Header Date in Table 1")

# Header QDGS: Table 3
old_t3_hdr = (
    '<w:t xml:space="preserve">Số:           </w:t></w:r><w:r>'
    + rpr_t1
    + '<w:t xml:space="preserve">  </w:t></w:r>'
    + '<w:r w:rsidR="008629EC" w:rsidRPr="009D7F91">'
    + rpr_t1
    + '<w:t xml:space="preserve">   /QĐ-TCS3</w:t>'
)
new_t3_hdr = (
    '<w:t xml:space="preserve">Số: </w:t></w:r>'
    + '<w:r w:rsidR="008629EC" w:rsidRPr="009D7F91">'
    + rpr_t1
    + f'<w:t xml:space="preserve">{QDGS}</w:t>'
)
if old_t3_hdr in xml:
    xml = xml.replace(old_t3_hdr, new_t3_hdr, 1)
    print("Replaced Header QDGS in Table 3")

old_t3_date = "Trấn Biên, ngày       tháng      năm"
if old_t3_date in xml:
    xml = xml.replace(old_t3_date, new_t1_date, 1)
    print("Replaced Header Date in Table 3")

# 3. References to QĐGS / QĐKT in text
xml = xml.replace(
    "Căn cứ Quyết định kiểm tra số              /QĐ-TCS3 ngày ...../...../2026",
    f"Căn cứ Quyết định kiểm tra số {QDKT} ngày {NGAY}"
)
xml = xml.replace(
    "theo Quyết định số              /QĐ-TCS3 ngày ...../...../2026",
    f"theo Quyết định số {QDKT} ngày {NGAY}"
)

old_p449 = (
    "Thực hiện Quyết định số              /QĐ-TCS3 ngày ..../..../2026 "
    "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
    "đoàn kiểm tra theo QĐ kiểm tra tại doanh nghiệp số              "
    "/QĐ-TCS3 ngày ..../..../2026"
)
new_p449 = (
    f"Thực hiện Quyết định số {QDGS} ngày {NGAY} "
    "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
    f"đoàn kiểm tra theo QĐ kiểm tra tại doanh nghiệp số {QDKT} ngày {NGAY}"
)
xml = xml.replace(old_p449, new_p449)

old_p480 = (
    "theo quyết định kiểm tra tại doanh nghiệp số              /QĐ-TCS3 "
    "ngày ..../..../2026 về việc giám sát Đoàn kiểm tra theo quyết "
    "định giám sát số              /QĐ-TCS3 ngày ..../..../2026."
)
new_p480 = (
    f"theo quyết định kiểm tra tại doanh nghiệp số {QDKT} "
    f"ngày {NGAY} về việc giám sát Đoàn kiểm tra theo quyết "
    f"định giám sát số {QDGS} ngày {NGAY}."
)
xml = xml.replace(old_p480, new_p480)

old_p501 = (
    "Thực hiện Quyết định số              /QĐ-TCS3 ngày ..../..../2026 "
    "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
    "hoạt động của Đoàn kiểm tra tại Công ty TNHH Thương mại Sản xuất Sơn Tiền Hải"
)
new_p501 = (
    f"Thực hiện Quyết định số {QDGS} ngày {NGAY} "
    "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
    "hoạt động của Đoàn kiểm tra tại Công ty TNHH Thương mại Sản xuất Sơn Tiền Hải"
)
xml = xml.replace(old_p501, new_p501)

old_gen = "số              /QĐ-TCS3 ngày ..../..../2026"
n_gen = xml.count(old_gen)
xml = xml.replace(old_gen, f"số {QDKT} ngày {NGAY}")
print(f"Replaced general references (số .../QĐ-TCS3 ngày ...): {n_gen}")

old_gen2 = "số:              /QĐ-TCS3 ngày ..../..../2026"
n_gen2 = xml.count(old_gen2)
xml = xml.replace(old_gen2, f"số: {QDKT} ngày {NGAY}")
print(f"Replaced general references with colon: {n_gen2}")

xml = xml.replace("SỐ              /QĐ-TCS3", f"SỐ {QDKT}")
xml = xml.replace("……../QĐ-TCS3 ngày …../…../202", f"{QDKT} ngày {NGAY}")
xml = xml.replace("……../QĐ-TCS3 ngày …../…../2026", f"{QDKT} ngày {NGAY}")
xml = xml.replace("……../QĐ-TCS3 ngày …../…../", f"{QDKT} ngày {NGAY}")
xml = xml.replace("……../QĐ-TCS3", QDKT)

xml = xml.replace(
    '<w:t xml:space="preserve">              /QĐ-TCS3</w:t>',
    f'<w:t xml:space="preserve">{QDKT}</w:t>'
)

# 4. Replace MERGEFIELD "KT" and MERGEFIELD KT cleanly
def replace_mergefield(doc_xml: str, field_name: str, replacement_text: str | None) -> tuple[str, int]:
    count = 0
    patterns = [f'MERGEFIELD "{field_name}"', f'MERGEFIELD {field_name}']
    for pat in patterns:
        while True:
            pos = doc_xml.find(pat)
            if pos == -1:
                break
            
            # Find Run 2 (instrText run)
            r2_start = doc_xml.rfind('<w:r', 0, pos)
            while r2_start > 0:
                prefix = doc_xml[r2_start:r2_start+5]
                if prefix.startswith('<w:r>') or prefix.startswith('<w:r '):
                    break
                r2_start = doc_xml.rfind('<w:r', 0, r2_start - 1)
            r2_end = doc_xml.find('</w:r>', pos) + len('</w:r>')
            
            # Find Run 1 (begin)
            r1_end = r2_start
            r1_start = doc_xml.rfind('<w:r', 0, r1_end)
            while r1_start > 0:
                prefix = doc_xml[r1_start:r1_start+5]
                if prefix.startswith('<w:r>') or prefix.startswith('<w:r '):
                    break
                r1_start = doc_xml.rfind('<w:r', 0, r1_start - 1)
            
            # Find Run 3 (end)
            r3_start = r2_end
            r3_end = doc_xml.find('</w:r>', r3_start) + len('</w:r>')
            
            if replacement_text is not None:
                # Extract rPr from r2
                r2_str = doc_xml[r2_start:r2_end]
                m_rpr = re.search(r'<w:rPr>.*?</w:rPr>', r2_str)
                rpr = m_rpr.group(0) if m_rpr else ""
                rep_run = f"<w:r>{rpr}<w:t>{replacement_text}</w:t></w:r>"
            else:
                rep_run = ""
            
            doc_xml = doc_xml[:r1_start] + rep_run + doc_xml[r3_end:]
            count += 1
    return doc_xml, count

xml, n_kt = replace_mergefield(xml, "KT", KT)
print(f"Replaced MERGEFIELD KT: {n_kt} occurrences (expected 6)")

xml, n_cccdkt = replace_mergefield(xml, "CCCDKT", None)
print(f"Removed MERGEFIELD CCCDKT: {n_cccdkt} occurrences (expected 1)")

# Signatures in Table 17 and Table 24:
pth = "Phan Thế Hanh"
old_sig17 = f'<w:t xml:space="preserve">        {pth}                  </w:t>'
new_sig17 = f'<w:t xml:space="preserve">        {pth}                  {KT}</w:t>'
if old_sig17 in xml:
    xml = xml.replace(old_sig17, new_sig17, 1)
    print("Updated signature Table 17")

old_sig24 = f'<w:t xml:space="preserve">{pth}             </w:t>'
new_sig24 = f'<w:t xml:space="preserve">{pth}                  {KT}</w:t>'
if old_sig24 in xml:
    xml = xml.replace(old_sig24, new_sig24, 1)
    print("Updated signature Table 24")

# 5. Bank account
old_tk = '<w:t xml:space="preserve">- Tài khoản số: </w:t>'
new_tk = f'<w:t xml:space="preserve">- Tài khoản số: {TK}.</w:t>'
if old_tk in xml:
    xml = xml.replace(old_tk, new_tk, 1)
    print("Updated Bank Account in P291")

# Remove empty Tài_khoản_số paragraph
for field_str in ['MERGEFIELD Tài_khoản_số', 'MERGEFIELD "Tài_khoản_số"']:
    while True:
        pos = xml.find(field_str)
        if pos == -1:
            break
        p_start = xml.rfind('<w:p', 0, pos)
        while p_start > 0:
            prefix = xml[p_start:p_start+5]
            if prefix.startswith('<w:p>') or prefix.startswith('<w:p '):
                break
            p_start = xml.rfind('<w:p', 0, p_start - 1)
        p_end = xml.find('</w:p>', pos) + len('</w:p>')
        xml = xml[:p_start] + xml[p_end:]
        print("Removed empty Tài_khoản_số paragraph")

# Save to TEST_OUT
with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(TEST_OUT, "w", zipfile.ZIP_DEFLATED) as zout:
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename == "word/document.xml":
            data = xml.encode("utf8")
        zout.writestr(it, data)

print(f"Saved {TEST_OUT}, size: {TEST_OUT.stat().st_size}")

doc = docx.Document(str(TEST_OUT))
print("Successfully opened with python-docx!")
print(f"Paragraphs: {len(doc.paragraphs)}, Tables: {len(doc.tables)}")

full_text = "\n".join(p.text for p in doc.paragraphs)
for t in doc.tables:
    for r in t.rows:
        for c in r.cells:
            full_text += "\n" + c.text

print("--- VERIFICATION ---")
print(f"Count '{QDKT}': {full_text.count(QDKT)}")
print(f"Count '{QDGS}': {full_text.count(QDGS)}")
print(f"Count '{KT}': {full_text.count(KT)}")
print(f"Count '6722949824': {full_text.count('6722949824')}")
print(f"Count 'quản lý kinh doanh': {full_text.lower().count('quản lý kinh doanh')}")
print(f"Count 'quản lý doanh nghiệp': {full_text.lower().count('quản lý doanh nghiệp')}")
