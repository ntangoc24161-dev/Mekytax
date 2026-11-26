# -*- coding: utf-8 -*-
"""Fix date MERGEFIELD switches so Word Mail Merge shows dd/MM/yyyy instead of Excel serials.

Usage:
  python scripts/fix_mergefield_ngay_ddmmyyyy.py [docx ...]
Default targets: templates/MAUTONG HOP.docx and templates/MAU TONG HOP HO SO DI.docx
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

DATE_FOLDED = {
    "ngay_qd", "ngay_cap", "ngay_cap2", "ngay_cong_bo", "ngay_ky_bien_ban_kiem_tra",
    "tu_ngay", "den_ngay", "ngay_bao_cao", "ngay_qdkt", "ngay_qdgs", "ngay_qdxphc",
    "ngay_cap_cccd", "ngay_cap_dkkd", "ngay_bat_dau", "ngay_ket_thuc", "ngay_ky_bbkt",
}

FIELD_RE = re.compile(
    r'^\s*MERGEFIELD\s+"?(?P<name>[^"\\\s]+)"?\s*(?P<rest>.*)$',
    re.I | re.S,
)
SWITCH_RE = re.compile(r'\\@\s*"([^"]+)"')
BROKEN_SWITCH_RE = re.compile(r'/@\s*"([^"]+)"')


def fold(s: str) -> str:
    repl = {
        "á":"a","à":"a","ả":"a","ã":"a","ạ":"a","ă":"a","ắ":"a","ằ":"a","ẳ":"a","ẵ":"a","ặ":"a",
        "â":"a","ấ":"a","ầ":"a","ẩ":"a","ẫ":"a","ậ":"a",
        "é":"e","è":"e","ẻ":"e","ẽ":"e","ẹ":"e","ê":"e","ế":"e","ề":"e","ể":"e","ễ":"e","ệ":"e",
        "í":"i","ì":"i","ỉ":"i","ĩ":"i","ị":"i",
        "ó":"o","ò":"o","ỏ":"o","õ":"o","ọ":"o","ô":"o","ố":"o","ồ":"o","ổ":"o","ỗ":"o","ộ":"o",
        "ơ":"o","ớ":"o","ờ":"o","ở":"o","ỡ":"o","ợ":"o",
        "ú":"u","ù":"u","ủ":"u","ũ":"u","ụ":"u","ư":"u","ứ":"u","ừ":"u","ử":"u","ữ":"u","ự":"u",
        "ý":"y","ỳ":"y","ỷ":"y","ỹ":"y","ỵ":"y",
        "đ":"d",
    }
    return "".join(repl.get(ch, ch) for ch in s.lower())


def is_date_field(name: str) -> bool:
    f = fold(name)
    if f in DATE_FOLDED:
        return True
    return f.startswith("ngay_") or f.startswith("tu_ngay") or f.startswith("den_ngay")


def parse_instruction(instr: str):
    m = FIELD_RE.match(instr.strip())
    if not m:
        return None, None
    return m.group("name"), (m.group("rest") or "")


def desired_switch(name: str, rest: str) -> str | None:
    if not is_date_field(name):
        return None
    if SWITCH_RE.findall(rest):
        return None
    broken = BROKEN_SWITCH_RE.findall(rest)
    if broken:
        return f'\\@ "{broken[0]}"'
    return '\\@ "dd/MM/yyyy"'


def rebuild_instruction(name: str, switch: str) -> str:
    return f' MERGEFIELD "{name}" {switch} '


def fix_document_xml(xml_bytes: bytes) -> tuple[bytes, list[str]]:
    root = etree.fromstring(xml_bytes)
    changes: list[str] = []

    candidates = [el for el in root.iter() if any(c.tag == f"{W}r" for c in el)]
    for parent in candidates:
        children = list(parent)
        i = 0
        while i < len(children):
            child = children[i]
            begin = False
            if child.tag == f"{W}r":
                for sub in child:
                    if sub.tag == f"{W}fldChar" and sub.get(f"{W}fldCharType") == "begin":
                        begin = True
                        break
            if not begin:
                i += 1
                continue

            end_idx = None
            instr_nodes = []
            sep_idx = None
            j = i + 1
            while j < len(children):
                curr = children[j]
                if curr.tag == f"{W}r":
                    for sub in curr:
                        if sub.tag == f"{W}instrText" and sep_idx is None:
                            instr_nodes.append(sub)
                        elif sub.tag == f"{W}fldChar":
                            ctype = sub.get(f"{W}fldCharType")
                            if ctype == "separate" and sep_idx is None:
                                sep_idx = j
                            elif ctype == "end":
                                end_idx = j
                                break
                    if end_idx is not None:
                        break
                j += 1

            if end_idx is None or not instr_nodes:
                i += 1
                continue

            full = "".join((n.text or "") for n in instr_nodes)
            name, rest = parse_instruction(full)
            if name is None:
                i = end_idx + 1
                continue

            switch = desired_switch(name, rest or "")
            if switch is None:
                i = end_idx + 1
                continue

            new_instr = rebuild_instruction(name, switch)
            instr_nodes[0].text = new_instr
            instr_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            for extra in instr_nodes[1:]:
                extra.text = ""
            changes.append(f"{full!r} -> {new_instr!r}")
            i = end_idx + 1

    for it in root.findall(f".//{W}instrText"):
        t = it.text or ""
        if "/@" in t and "MERGEFIELD" in t.upper():
            fixed = t.replace("/@", "\\@")
            if fixed != t:
                it.text = fixed
                changes.append(f"slash-fix {t!r} -> {fixed!r}")

    out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return out, changes


def summarize(path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    parts = re.findall(r"<w:instrText[^>]*>([^<]*)</w:instrText>", xml)
    joined = "".join(parts)
    fields = re.findall(r'MERGEFIELD\s+"?([^"\\\s]+)"?((?:\s*\\[@#/]\s*"[^"]*")*)', joined)
    by = defaultdict(list)
    for n, sw in fields:
        if is_date_field(n):
            by[n].append(sw.strip())
    print(f"=== {path.name} ===")
    for n in sorted(by, key=lambda x: x.lower()):
        print(f"  {n}: {Counter(by[n])}")
    print(f"  leftover /@ nodes: {sum(1 for p in parts if '/@' in p)}")


def patch_docx(path: Path) -> list[str]:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = path.with_suffix(path.suffix + f".bak-truoc-fix-ngay-{stamp}")
    shutil.copy2(path, bak)
    print(f"Backup: {bak}")

    with zipfile.ZipFile(path, "r") as zin:
        files = {info.filename: zin.read(info.filename) for info in zin.infolist()}
        infos = list(zin.infolist())

    new_xml, changes = fix_document_xml(files["word/document.xml"])
    files["word/document.xml"] = new_xml

    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, files[info.filename])
    tmp.replace(path)
    return changes


def main(argv: list[str]) -> int:
    brain = Path(__file__).resolve().parent.parent
    targets = [Path(a) for a in argv[1:]]
    if not targets:
        targets = [
            brain / "templates" / "MAUTONG HOP.docx",
            brain / "templates" / "MAU TONG HOP HO SO DI.docx",
        ]
    for t in targets:
        if not t.exists():
            print(f"missing: {t}")
            continue
        changes = patch_docx(t)
        print(f"{t}: {len(changes)} changes")
        summarize(t)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
