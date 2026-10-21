# -*- coding: utf-8 -*-
"""Re-apply single-item layout rules after rebuild from MAUTONG HOP.

Keep original form layouts for bien ban/ke hoach.
Re-apply only:
- blank QD issuance dates
- proper-case company name
- phan tich III/IV + remove gia von + 1.25cm bullets
- tieu ngu = quoc hieu + 1, bold, upright (no italic)
- signature gaps prefer 4 lines, shrink if to-trinh overflows 1 page
- standalone QD files stay 1 page each
"""
from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BB = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\9. VAN PHAT THANH\0. BIEN BAN"
)
LIVE = BB / "TONG HOP.docx"
QDKT = BB / "QD Kiem tra tai cty Van Phat Thanh.docx"
QDGS = BB / "QDgiam sat tai cty Van Phat Thanh.docx"
ZIPP = BB / "hoso.zip"

OLD_CO = "CÔNG TY TNHH MỘT THÀNH VIÊN BAO BÌ GIẤY VẠN PHÁT THÀNH"
NEW_CO = "Công ty TNHH Một Thành Viên Bao Bì Giấy Vạn Phát Thành"
DATE_BLANK = "Trấn Biên, ngày .... tháng .... năm 2026"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FI_125 = Cm(1.25)


def set_run_font(run, size_pt=None, bold=None, italic=None):
    run.font.name = "Times New Roman"
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rFonts.set(qn(a), "Times New Roman")
    if size_pt is not None:
        run.font.size = Pt(size_pt)
        for tag in ("w:sz", "w:szCs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rPr.append(el)
            el.set(qn("w:val"), str(int(size_pt * 2)))
    if bold is not None:
        run.bold = bold
        b = rPr.find(qn("w:b"))
        if bold:
            if b is None:
                rPr.append(OxmlElement("w:b"))
        elif b is not None:
            rPr.remove(b)
            bCs = rPr.find(qn("w:bCs"))
            if bCs is not None:
                rPr.remove(bCs)
    if italic is not None:
        run.italic = italic
        i_el = rPr.find(qn("w:i"))
        if italic:
            if i_el is None:
                rPr.append(OxmlElement("w:i"))
        else:
            if i_el is not None:
                rPr.remove(i_el)
            iCs = rPr.find(qn("w:iCs"))
            if iCs is not None:
                rPr.remove(iCs)


def replace_text_in_paragraph(p, old, new):
    full = p.text or ""
    if old not in full:
        return 0
    new_full = full.replace(old, new)
    bold = p.runs[0].bold if p.runs else None
    size = p.runs[0].font.size if p.runs else None
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(new_full)
    set_run_font(run, size_pt=(size.pt if size else 13), bold=bold)
    return 1


def replace_company_everywhere(doc):
    n = 0
    for p in doc.paragraphs:
        n += replace_text_in_paragraph(p, OLD_CO, NEW_CO)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    n += replace_text_in_paragraph(p, OLD_CO, NEW_CO)
    # center title company lines in QD
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t == NEW_CO or t.startswith(NEW_CO + ","):
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True, italic=False)
    return n


def blank_qd_header_dates(doc):
    """Blank issuance date only in QD header tables that carry Số .../QĐ-."""
    n = 0
    for tbl in doc.tables:
        joined = "\n".join(c.text for row in tbl.rows for c in row.cells)
        if not re.search(r"Số:\s*.*?/QĐ", joined) and "QĐ-TCS3" not in joined:
            continue
        # only act if this looks like a decision header (has QUYẾT or Số)
        if "Số:" not in joined and "Số :" not in joined:
            continue
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = (p.text or "").strip()
                    if not t.startswith("Trấn Biên"):
                        continue
                    if "ngày" not in t or "tháng" not in t:
                        continue
                    for r in list(p.runs):
                        r._element.getparent().remove(r._element)
                    run = p.add_run(DATE_BLANK)
                    set_run_font(run, size_pt=13, italic=True, bold=False)
                    n += 1
    return n


def clear_spacing(p, ls=1.0, sb=0, sa=3):
    pf = p.paragraph_format
    pf.line_spacing = ls
    pf.space_before = Pt(sb)
    pf.space_after = Pt(sa)


def make_bullet_para(text, size=13):
    new_p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "both")
    pPr.append(jc)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLine"), "708")  # ~1.25 cm
    pPr.append(ind)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "60")
    sp.set(qn("w:line"), "240")
    sp.set(qn("w:lineRule"), "auto")
    pPr.append(sp)
    new_p.append(pPr)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:eastAsia"), "Times New Roman")
    rPr.append(rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)
    r.append(rPr)
    te = OxmlElement("w:t")
    te.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    te.text = text
    r.append(te)
    new_p.append(r)
    return new_p


def clear_runs_set_text(p, text, size=13, bold=False, italic=False):
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run(text)
    set_run_font(run, size_pt=size, bold=bold, italic=italic)
    return run


def split_ratio_lines(raw: str):
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    if len(lines) == 1 and raw.count("- Tỷ lệ") > 1:
        parts = raw.split("- Tỷ lệ")
        lines = []
        for j, part in enumerate(parts):
            part = part.strip()
            if not part:
                continue
            if j == 0 and part.startswith("-"):
                lines.append(part)
            else:
                lines.append(("- Tỷ lệ " + part).strip())
    return lines


def fix_phan_tich(doc):
    start = end = None
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if start is None and t.startswith("PHÂN TÍCH THÔNG TIN"):
            start = i
        if start is not None and (
            t.startswith("THƯ TRAO ĐỔI") or t.startswith("Kính gửi")
        ):
            end = i
            break
    if start is None:
        print("phan tich not found")
        return
    end = end or min(len(doc.paragraphs), start + 60)
    print("phan tich range", start, end)

    # slogan size 14 in header table just before phan tich
    body = list(doc.element.body)
    last_tbl = None
    para_i = 0
    for el in body:
        tag = el.tag.split("}")[-1]
        if tag == "p":
            texts = "".join(t.text or "" for t in el.findall(".//%st" % W))
            if texts.strip().startswith("PHÂN TÍCH THÔNG TIN"):
                break
            para_i += 1
        elif tag == "tbl":
            last_tbl = el
    if last_tbl is not None:
        n_slogan = 0
        for pEl in last_tbl.findall(".//%sp" % W):
            text = "".join(t.text or "" for t in pEl.findall(".//%st" % W))
            if "Độc lập" in text or ("Độc" in text and "Tự do" in text):
                for rEl in pEl.findall("%sr" % W):
                    set_run_size_xml(rEl, 14, bold=True, italic=False)
                    n_slogan += 1
        print("phan tich slogan runs -> 14 upright:", n_slogan)

    to_delete = []
    i = start
    while i < end:
        # re-bind because deletes/inserts shift
        if i >= len(doc.paragraphs):
            break
        p = doc.paragraphs[i]
        t = (p.text or "").strip()
        replace_text_in_paragraph(p, OLD_CO, NEW_CO)

        if "Tỷ lệ giá vốn hàng bán" in t:
            to_delete.append(p._element)
            print("mark delete gia von P", i)
            i += 1
            continue

        if t.startswith("- Tỷ lệ chi phí") and (
            "\n" in (p.text or "") or (p.text or "").count("- Tỷ lệ") > 1
        ):
            lines = split_ratio_lines(p.text or "")
            keep = [
                ln
                for ln in lines
                if "giá vốn" not in ln.lower() and ln.startswith("- Tỷ lệ")
            ]
            if not keep:
                to_delete.append(p._element)
                i += 1
                continue
            clear_runs_set_text(p, keep[0], 13)
            p.paragraph_format.first_line_indent = FI_125
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            clear_spacing(p, 1.0, 0, 3)
            insert_after = p._element
            for line in keep[1:]:
                np = make_bullet_para(line)
                insert_after.addnext(np)
                insert_after = np
            print("split ratio lines", len(keep))
            end += len(keep) - 1
            i += len(keep)
            continue

        # III.1 bullets: first-line indent 1.25cm
        if t.startswith("- ") and i > start:
            # only inside III until IV
            # find if still before IV
            p.paragraph_format.first_line_indent = FI_125
            p.paragraph_format.left_indent = Cm(0)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            clear_spacing(p, 1.0, 0, 3)
            for r in p.runs:
                set_run_font(r, size_pt=13)

        if t.startswith("+/"):
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            clear_spacing(p, 1.0, 3, 3)

        if t.startswith("f)"):
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.left_indent = Cm(0)
            clear_spacing(p, 1.0, 0, 3)

        if t.startswith("III.") or t.startswith("IV."):
            clear_spacing(p, 1.0, 6, 3)
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=True)

        # IV cost lines: split and clean double dash
        if t.startswith("- Chi phí") and "\n" in (p.text or ""):
            lines = [ln.strip() for ln in p.text.splitlines() if ln.strip()]
            lines = [re.sub(r"^-\s*-\s*", "- ", ln) for ln in lines]
            clear_runs_set_text(p, lines[0], 13)
            p.paragraph_format.first_line_indent = FI_125
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            clear_spacing(p, 1.0, 0, 3)
            insert_after = p._element
            for line in lines[1:]:
                np = make_bullet_para(line)
                insert_after.addnext(np)
                insert_after = np
            print("split IV cost lines", len(lines))
            end += len(lines) - 1
            i += len(lines)
            continue

        if t.startswith("- - Chi phí"):
            clear_runs_set_text(p, re.sub(r"^-\s*-\s*", "- ", t), 13)
            p.paragraph_format.first_line_indent = FI_125
            clear_spacing(p, 1.0, 0, 3)

        # IV date: ordinary, not bold
        if t.startswith("Trấn Biên") and i > start + 20:
            clear_runs_set_text(p, t, 13, bold=False, italic=True)
            # right-align via center tab like before
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Cm(0)
            pPr = p._element.get_or_add_pPr()
            old = pPr.find(qn("w:tabs"))
            if old is not None:
                pPr.remove(old)
            tabs = OxmlElement("w:tabs")
            tab = OxmlElement("w:tab")
            tab.set(qn("w:val"), "center")
            tab.set(qn("w:pos"), "6663")
            tabs.append(tab)
            pPr.append(tabs)
            # ensure leading tab
            if not (p.text or "").startswith("\t"):
                if p.runs:
                    p.runs[0].text = "\t" + (p.runs[0].text or "").lstrip("\t")
            for r in p.runs:
                set_run_font(r, size_pt=13, bold=False, italic=True)
            print("IV date unbolded")

        i += 1

    for el in to_delete:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)
            print("deleted gia von para")


def run_size_pt(rEl):
    rPr = rEl.find("%srPr" % W)
    if rPr is None:
        return None
    sz = rPr.find("%ssz" % W)
    if sz is None:
        return None
    try:
        return int(sz.get("%sval" % W)) / 2.0
    except Exception:
        return None


def set_run_size_xml(rEl, size_pt, bold=None, italic=None):
    rPr = rEl.find("%srPr" % W)
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        rEl.insert(0, rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rFonts.set(qn(a), "Times New Roman")
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), str(int(size_pt * 2)))
    # bold
    b = rPr.find(qn("w:b"))
    if bold is True:
        if b is None:
            rPr.append(OxmlElement("w:b"))
    elif bold is False and b is not None:
        rPr.remove(b)
        bCs = rPr.find(qn("w:bCs"))
        if bCs is not None:
            rPr.remove(bCs)
    # italic OFF for slogan
    i_el = rPr.find(qn("w:i"))
    if italic is True:
        if i_el is None:
            rPr.append(OxmlElement("w:i"))
    elif italic is False:
        if i_el is not None:
            rPr.remove(i_el)
        iCs = rPr.find(qn("w:iCs"))
        if iCs is not None:
            rPr.remove(iCs)


def is_slogan_para(text):
    return "Độc lập" in text or ("Độc" in text and "Tự do" in text and "Hạnh" in text)


def is_quoc_hieu_para(text):
    return "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in text or text.strip().startswith(
        "CỘNG HÒA"
    )


def sync_tieu_ngu_in_table(tbl_el, force_slogan=None):
    quoc_sizes = []
    slogan_paras = []
    for pEl in tbl_el.findall(".//%sp" % W):
        text = "".join(t.text or "" for t in pEl.findall(".//%st" % W))
        if is_quoc_hieu_para(text):
            for rEl in pEl.findall("%sr" % W):
                sz = run_size_pt(rEl)
                if sz:
                    quoc_sizes.append(sz)
        if is_slogan_para(text):
            slogan_paras.append(pEl)
    if not slogan_paras:
        return 0, None, None
    base = max(quoc_sizes) if quoc_sizes else 13.0
    target = force_slogan if force_slogan is not None else base + 1.0
    if force_slogan is None:
        if target > 14:
            target = 14.0
        if target < 13:
            target = 13.0
    n = 0
    for pEl in slogan_paras:
        for rEl in pEl.findall("%sr" % W):
            set_run_size_xml(rEl, target, bold=True, italic=False)
            n += 1
    return n, base, target


def sync_tieu_ngu_doc(doc, label, force_phan_tich_14=True):
    total = 0
    samples = []
    # find phan tich header table index roughly by content nearby
    for ti, tbl in enumerate(doc.tables):
        joined = "\n".join(c.text for row in tbl.rows for c in row.cells)
        if "Độc lập" not in joined and "Độc" not in joined:
            continue
        if "CỘNG HÒA" not in joined:
            continue
        force = None
        n, base, target = sync_tieu_ngu_in_table(tbl._tbl, force_slogan=force)
        if not n:
            continue
        total += n
        if len(samples) < 10:
            samples.append("T%d CH=%.0f slogan->%.0f runs=%d" % (ti, base or 0, target or 0, n))
    # body paragraphs rare
    for p in doc.paragraphs:
        if is_slogan_para(p.text or ""):
            for r in p.runs:
                set_run_font(r, size_pt=14, bold=True, italic=False)
            total += 1
    print(label, "tieu ngu runs updated", total, samples)
    return total


def ensure_empty_between(doc, title_text, name_text, keep=4, search_end=None):
    ti = ni = None
    for i, p in enumerate(doc.paragraphs):
        if search_end is not None and i >= search_end:
            break
        if p.text.strip() == title_text:
            ti = i
        if ti is not None and p.text.strip() == name_text and i > ti:
            ni = i
            break
    if ti is None or ni is None:
        return "miss"
    empties = []
    el = doc.paragraphs[ti]._element.getnext()
    name_el = doc.paragraphs[ni]._element
    while el is not None and el is not name_el:
        if el.tag == qn("w:p"):
            texts = "".join(node.text or "" for node in el.iter(qn("w:t"))).strip()
            if texts == "":
                empties.append(el)
            else:
                break
        el = el.getnext()
    cur = len(empties)
    if cur > keep:
        for e in empties[keep:]:
            e.getparent().remove(e)
        return "removed %d (had %d)" % (cur - keep, cur)
    if cur < keep:
        for _ in range(keep - cur):
            empty = OxmlElement("w:p")
            pPr = OxmlElement("w:pPr")
            sp = OxmlElement("w:spacing")
            sp.set(qn("w:before"), "0")
            sp.set(qn("w:after"), "0")
            sp.set(qn("w:line"), "240")
            sp.set(qn("w:lineRule"), "auto")
            pPr.append(sp)
            empty.append(pPr)
            name_el.addprevious(empty)
        return "inserted %d (had %d)" % (keep - cur, cur)
    return "ok %d" % cur


def fix_to_trinh_sign_gaps(doc, keep=4):
    pairs = [
        ("CÁN BỘ ĐỀ XUẤT", "Nguyễn Thị Ánh Ngọc"),
        ("TỔ TRƯỞNG", "Nguyễn Trường Giang"),
        ("PHÓ TRƯỞNG THUẾ CƠ SỞ", "Phạm Đồng Xuân"),
    ]
    end = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().startswith("QUYẾT ĐỊNH") and i > 10:
            end = i
            break
    for title, name in pairs:
        print(" sign", title, ensure_empty_between(doc, title, name, keep=keep, search_end=end))


def word_page_count(path: Path) -> int | None:
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        print("win32com missing", e)
        return None
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(path), ReadOnly=True)
        doc.Repaginate()
        n = int(doc.ComputeStatistics(2))  # wdStatisticPages
        doc.Close(False)
        return n
    except Exception as e:
        print("page count error", e)
        return None
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def extract_qd_standalone(doc_src: Document, start_marker: str, end_markers: list[str], out_path: Path):
    """Copy a QD section into a fresh doc based on existing standalone skeleton.

    Safer approach: patch the existing standalone files in place rather than
    rebuilding from scratch.
    """
    raise NotImplementedError


def patch_standalone_qd(path: Path, label: str):
    d = Document(str(path))
    n_co = replace_company_everywhere(d)
    n_dt = blank_qd_header_dates(d)
    sync_tieu_ngu_doc(d, label)
    # tighten spacing for 1-page preference
    for p in d.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        pf = p.paragraph_format
        if pf.line_spacing is None or (isinstance(pf.line_spacing, float) and pf.line_spacing > 1.05):
            pf.line_spacing = 1.0
        if pf.space_after and pf.space_after.pt and pf.space_after.pt > 6:
            pf.space_after = Pt(3)
    d.save(str(path))
    print(label, "company", n_co, "blank dates", n_dt)


def verify(doc):
    print("\n=== VERIFY ===")
    upper = sum(1 for p in doc.paragraphs if OLD_CO in (p.text or ""))
    proper = sum(1 for p in doc.paragraphs if NEW_CO in (p.text or ""))
    print("company UPPER paras", upper, "PROPER paras", proper)
    gia = [i for i, p in enumerate(doc.paragraphs) if "giá vốn" in (p.text or "")]
    print("gia von paras", gia)
    # III bullets
    in_iii = False
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if t.startswith("III. PHÂN TÍCH"):
            in_iii = True
        if t.startswith("IV. ĐỀ XUẤT"):
            in_iii = False
        if in_iii and t.startswith("- "):
            fi = p.paragraph_format.first_line_indent
            print(" III bullet P%d fi=%s | %s" % (i, fi, t[:90]))
    # IV date
    for i, p in enumerate(doc.paragraphs):
        t = (p.text or "").strip()
        if t.startswith("Trấn Biên") and i > 100 and i < 140:
            print(" IV date P%d bold_runs=%s | %r" % (
                i,
                [r.bold for r in p.runs],
                t,
            ))


def main():
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = BB / ("TONG HOP.bak-reapply-don-le-%s.docx" % stamp)
    shutil.copy2(LIVE, bak)
    print("backup", bak.name)

    doc = Document(str(LIVE))
    n_co = replace_company_everywhere(doc)
    print("company replaced", n_co)
    n_dt = blank_qd_header_dates(doc)
    print("blank QD header dates", n_dt)

    # signature gaps prefer 4
    fix_to_trinh_sign_gaps(doc, keep=4)
    fix_phan_tich(doc)
    sync_tieu_ngu_doc(doc, "TONG HOP")

    # force phan tich slogan 14 again after sync (sync may have set 13+1 from local CH)
    # already done inside fix_phan_tich; re-run sync keeps +1 rule globally

    doc.save(str(LIVE))
    print("saved TONG HOP")

    # If to-trinh overflows, shrink signature gaps
    pages = word_page_count(LIVE)
    print("TONG HOP pages (whole file)", pages)

    # Check to-trinh page by exporting? Use heuristic: if Phó name pushed far,
    # shrink. Better: open Word and check page of paragraph Phạm Đồng Xuân.
    try:
        import win32com.client  # type: ignore

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        wdoc = word.Documents.Open(str(LIVE), ReadOnly=False)
        wdoc.Repaginate()
        # find "Phạm Đồng Xuân" first occurrence (to trinh)
        rng = wdoc.Content
        find = rng.Find
        find.ClearFormatting()
        find.Text = "Phạm Đồng Xuân"
        found = find.Execute()
        overflow = False
        if found:
            page_pho = int(rng.Information(3))  # wdActiveEndPageNumber
            print("first Phạm Đồng Xuân on page", page_pho)
            if page_pho > 1:
                overflow = True
        if overflow:
            print("to-trinh overflow -> shrink sign gaps to 2")
            wdoc.Close(False)
            word.Quit()
            doc2 = Document(str(LIVE))
            fix_to_trinh_sign_gaps(doc2, keep=2)
            doc2.save(str(LIVE))
            # recheck
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            wdoc = word.Documents.Open(str(LIVE), ReadOnly=True)
            wdoc.Repaginate()
            rng = wdoc.Content
            find = rng.Find
            find.Text = "Phạm Đồng Xuân"
            if find.Execute():
                print("after shrink, Phạm Đồng Xuân page", int(rng.Information(3)))
            wdoc.Close(False)
            word.Quit()
        else:
            wdoc.Close(False)
            word.Quit()
    except Exception as e:
        print("Word page check skipped/failed:", e)

    patch_standalone_qd(QDKT, "QDKT")
    patch_standalone_qd(QDGS, "QDGS")

    for path, label in ((QDKT, "QDKT"), (QDGS, "QDGS")):
        pc = word_page_count(path)
        print(label, "pages", pc)
        if pc is not None and pc > 1:
            # tighter pass
            d = Document(str(path))
            for p in d.paragraphs:
                pf = p.paragraph_format
                pf.line_spacing = 1.0
                pf.space_before = Pt(0)
                if pf.space_after is None or (pf.space_after and pf.space_after.pt > 3):
                    pf.space_after = Pt(0)
            d.save(str(path))
            print(label, "retightened, pages now", word_page_count(path))

    with ZipFile(ZIPP, "w", ZIP_DEFLATED) as z:
        for f in (LIVE, QDKT, QDGS):
            z.write(f, f.name)
    print("zip ok")

    verify(Document(str(LIVE)))


if __name__ == "__main__":
    main()
