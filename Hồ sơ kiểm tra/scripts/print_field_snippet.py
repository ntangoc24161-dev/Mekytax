import zipfile
import re

kh_path = r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates\MAUTONG HOP.docx"

with zipfile.ZipFile(kh_path, 'r') as z:
    kh_doc = z.read('word/document.xml').decode('utf-8', errors='ignore')

# Match an entire fldChar begin to end
pattern = re.compile(r'(<w:r[^>]*>.*?w:fldCharType="begin".*?</w:r>.*?<w:r[^>]*>.*?w:fldCharType="end".*?</w:r>)', re.DOTALL)
m = pattern.search(kh_doc)
if m:
    print("Full field snippet:\n", m.group(1))
