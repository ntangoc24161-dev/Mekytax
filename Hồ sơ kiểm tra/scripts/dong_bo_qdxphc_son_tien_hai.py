# -*- coding: utf-8 -*-
"""Dong bo file QD xu phat tach rieng cua vu Son Tien Hai voi nen phap ly
cua bo Mo Fa, giong nhu da ap cho TONG HOP.docx.

Chay: python scripts/dong_bo_qdxphc_son_tien_hai.py
"""
from __future__ import annotations

import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

VU = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
          r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\11. SON TIEN HAI\0. BIEN BAN")
QD = VU / "QD XPHC Cong ty Son Tien Hai.docx"
OUT = (Path(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra")
       / "output" / "3603859012_HO-SO-VE_son-tien-hai")

WT_SPLIT = re.compile(r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.S)
WT = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)

REPL = [
    ("1377/QĐ-CT ngày 30/06/2026", "1377/QĐ-CT ngày 30/06/2025"),
    ("13/10/2022", "13/10/2020"),
    ("Thuế cơ sổ", "Thuế cơ sở"),
    ("về việc kiểm tra quyết toán tại", "về việc kiểm tra thuế tại"),
    ("về việc kiểm tra quyết toán giải thể tại", "về việc kiểm tra thuế tại"),
    ("Biên bản kiểm tra quyết toán ký ngày", "Biên bản kiểm tra ký ngày"),
    ("được lập được lập", "được lập"),
]


def main() -> None:
    assert QD.exists(), f"khong thay {QD}"
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = VU / f"QD XPHC Cong ty Son Tien Hai.bak-{stamp}.docx"
    shutil.copy2(QD, bak)
    print(f"[bak] {bak.name}")

    zin = zipfile.ZipFile(QD)
    parts = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    xml = parts["word/document.xml"].decode("utf8")

    for old, new in REPL:
        hits = 0

        def fix(m, old=old, new=new):
            nonlocal hits
            body = m.group(2)
            if old not in body:
                return m.group(0)
            hits += body.count(old)
            return m.group(1) + body.replace(old, new) + m.group(3)

        xml = WT_SPLIT.sub(fix, xml)
        if hits:
            print(f"   {hits}x  {old[:46]!r} -> {new[:46]!r}")

    parts["word/document.xml"] = xml.encode("utf8")
    with zipfile.ZipFile(QD, "w", zipfile.ZIP_DEFLATED) as zo:
        for name, blob in parts.items():
            zo.writestr(name, blob)
    ET.fromstring(zipfile.ZipFile(QD).read("word/document.xml").decode("utf8"))
    print("[ok] XML hop le")

    con = [t for t in WT.findall(xml)
           if "quyết toán" in t or "cơ sổ" in t or "30/06/2026" in t]
    print("[check] con lai can soat:", con if con else "khong con")

    OUT.mkdir(parents=True, exist_ok=True)
    shutil.copy2(QD, OUT / QD.name)
    print(f"[out] {QD.name}")


if __name__ == "__main__":
    main()
