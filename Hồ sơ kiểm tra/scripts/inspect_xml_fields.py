import zipfile
import re

with zipfile.ZipFile(r"output\3603859012_HO-SO-VE_son-tien-hai\TONG HOP.docx") as z:
    xml = z.read("word/document.xml").decode("utf8")

for target in ["KT", "CCCDKT", "Tài_khoản_số"]:
    print(f"=== TARGET: {target} ===")
    matches = [m.start() for m in re.finditer(rf'MERGEFIELD\s+"?{target}"?', xml)]
    for idx, pos in enumerate(matches):
        start = max(0, pos - 200)
        end = min(len(xml), pos + 300)
        print(f"[{idx+1}] at {pos}:\n{xml[start:end]}\n" + "-"*40)
