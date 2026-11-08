# -*- coding: utf-8 -*-
"""Chuyen font Times New Roman toan van + bo sung Chi bo 13 ban hanh cho ho so nop."""
from __future__ import print_function
import os, shutil, zipfile, tempfile
from copy import copy
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

VAULT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
LAPTOP = r"H:\Other computers\My Laptop\2. CONG TAC KHAC\4. CONG TAC DANG\A PHUONG\2026\06_Tieu_chi_danh_gia\Du_thao_CB13_QuyIII_2026"
TNR = "Times New Roman"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
KH_BAN_HANH = ("(Chi bộ 13 ban hành kèm theo Quyết định số        -QĐ/CB13 "
               "ngày ..... tháng ..... năm 2026)")
BANNER = ("SHEET THEO DÕI NỘI BỘ - KHÔNG NỘP CẤP TRÊN. "
          "Chỉ dùng để tính điểm / sửa số ước tính. In và gửi sheet chính.")


def set_run_tnr(run, size=None):
    run.font.name = TNR
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), TNR)
    if size is not None:
        run.font.size = Pt(size)


def iter_all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                yield p
            for tbl in part.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p


def force_doc_defaults(doc):
    st = doc.styles["Normal"]
    st.font.name = TNR
    st.font.size = Pt(13)
    rPr = st.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), TNR)
    styles = doc.styles.element
    nsmap = {"w": W_NS}
    dd = styles.find(qn("w:docDefaults"))
    if dd is None:
        dd = OxmlElement("w:docDefaults")
        styles.insert(0, dd)
    rPrDef = dd.find(qn("w:rPrDefault"))
    if rPrDef is None:
        rPrDef = OxmlElement("w:rPrDefault")
        dd.insert(0, rPrDef)
    rPr = rPrDef.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        rPrDef.append(rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), TNR)


def insert_after(paragraph, text, italic=True, size=11, align="center"):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    np = Paragraph(new_p, paragraph._parent)
    if align == "center":
        np.alignment = 1  # WD_ALIGN_PARAGRAPH.CENTER
    np.paragraph_format.space_before = Pt(0)
    np.paragraph_format.space_after = Pt(6)
    run = np.add_run(text)
    run.italic = italic
    run.font.size = Pt(size)
    set_run_tnr(run, size)
    return np


def patch_theme_in_docx(path):
    """Doi major/minor theme font sang Times New Roman (ke ca eastAsia)."""
    tmp = path + ".tnr.tmp.docx"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/theme/theme1.xml":
                root = etree.fromstring(data)
                for el in root.xpath(".//a:latin | .//a:ea | .//a:cs",
                                     namespaces={"a": A_NS}):
                    el.set("typeface", TNR)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    os.replace(tmp, path)


def patch_docx(path):
    doc = Document(path)
    force_doc_defaults(doc)
    name = os.path.basename(path)
    added_ban_hanh = False
    replaced_star = 0

    for p in iter_all_paragraphs(doc):
        txt = (p.text or "").strip()
        if txt == "*":
            if p.runs:
                p.runs[0].text = "________"
                for extra in p.runs[1:]:
                    extra.text = ""
            else:
                r = p.add_run("________")
                r.bold = True
            replaced_star += 1
        if KH_BAN_HANH[:40] in txt:
            added_ban_hanh = True

    # Ke hoach ca nhan: them dong Chi bo 13 ban hanh neu chua co
    if name.startswith("KH_QIII-2026_") and not added_ban_hanh:
        paras = list(doc.paragraphs)
        inserted = False
        for i, p in enumerate(paras):
            if "KẾ HOẠCH CÔNG TÁC QUÝ III/2026" in (p.text or ""):
                # chen sau dong ----- neu co, khong thi ngay sau tieu de
                target = p
                if i + 1 < len(paras) and "-----" in (paras[i + 1].text or ""):
                    target = paras[i + 1]
                insert_after(target, KH_BAN_HANH, italic=True, size=11, align="center")
                inserted = True
                added_ban_hanh = True
                break
        if not inserted:
            print("  WARN: khong chen duoc dong ban hanh vao", name)

    for p in iter_all_paragraphs(doc):
        for r in p.runs:
            set_run_tnr(r)

    doc.save(path)
    patch_theme_in_docx(path)
    return {"star": replaced_star, "ban_hanh": added_ban_hanh}


def font_copy_tnr(src):
    return Font(name=TNR, size=src.size, bold=src.bold, italic=src.italic,
                underline=src.underline, color=src.color,
                vertAlign=src.vertAlign, strike=src.strike)


def set_wb_default_font(wb):
    try:
        wb._named_styles["Normal"].font.name = TNR
    except Exception:
        pass
    try:
        if wb._fonts:
            wb._fonts[0].name = TNR
    except Exception:
        pass


def patch_xlsx(path):
    wb = load_workbook(path)
    set_wb_default_font(wb)
    name = os.path.basename(path)
    rel = os.path.relpath(path, VAULT)
    notes = []

    # File 08: bo sung quoc hieu, ngay, chi bo ban hanh
    if name == "08_Thong_ke_cap_uy_vien.xlsx":
        ws = wb["Thong ke"]
        if ws["J1"].value in (None, ""):
            ws["J1"] = "ĐẢNG CỘNG SẢN VIỆT NAM"
            ws["J1"].font = Font(name=TNR, bold=True, size=11)
            ws["J1"].alignment = Alignment(horizontal="center")
            notes.append("them Quoc hieu")
        if ws["J2"].value in (None, ""):
            ws["J2"] = "Đồng Nai, ngày ..... tháng ..... năm 2026"
            ws["J2"].font = Font(name=TNR, italic=True, size=11)
            ws["J2"].alignment = Alignment(horizontal="center")
            notes.append("them ngay")
        if ws["A4"].value in (None, ""):
            ws["A4"] = "Chi bộ 13 ban hành"
            ws["A4"].font = Font(name=TNR, italic=True, size=11)
            notes.append("them 'Chi bộ 13 ban hành'")
        if ws["A2"].value == "CHI BỘ 13":
            ws["A2"].font = Font(name=TNR, bold=True, size=12, underline="single")
            notes.append("gach duoi ten co quan ban hanh")

    # File 02, 03: gach duoi CHI BO 13 neu o A3 trong
    if name == "02_Danh_sach_trich_ngang.xlsx":
        ws = wb.active
        if ws["A3"].value in (None, ""):
            ws["A3"] = "________"
            ws["A3"].font = Font(name=TNR, bold=True, size=11)
            ws["A3"].alignment = Alignment(horizontal="left")
            notes.append("gach duoi CHI BỘ 13")
        ws["A2"].font = Font(name=TNR, bold=True, size=12, underline="single")

    if name == "03_Bang_tong_hop_de_xuat.xlsx":
        ws = wb["Tong hop de xuat"]
        if ws["A3"].value in (None, ""):
            ws["A3"] = "Chi bộ 13 ban hành"
            ws["A3"].font = Font(name=TNR, italic=True, size=11)
            notes.append("them 'Chi bộ 13 ban hành'")
        ws["A2"].font = Font(name=TNR, bold=True, size=12, underline="single")

    # Ban tu danh gia: lam ro co quan ban hanh o A3
    if name.startswith("TDG_QIII-2026_") and "Khung đánh giá" in wb.sheetnames:
        ws = wb["Khung đánh giá"]
        a3 = ws["A3"].value or ""
        if "CHI BỘ 13" in a3 and "ban hành" not in a3.lower():
            # giu 2 dong Dang bo / Chi bo, khong chen them de tranh vo khung mau
            ws["A3"].font = Font(name=TNR, bold=True, size=11)
            notes.append("giu header Chi bộ 13 (khung mau)")

    # Banner sheet theo doi noi bo
    for ws in wb.worksheets:
        if ws.title in ("Tinh diem KPI", "Tinh KPI 2A-2B"):
            # chen dong canh bao vao A16 neu trong, hoac ghi de o hang 16 neu la KPI sheet
            target = None
            if ws.title == "Tinh KPI 2A-2B":
                if ws["A16"].value in (None, ""):
                    target = ws["A16"]
                elif "KHÔNG NỘP" not in str(ws["A16"].value):
                    target = ws["A16"]
            else:
                if ws["A16"].value in (None, "") or "KHÔNG NỘP" not in str(ws["A16"].value or ""):
                    target = ws["A16"]
            if target is not None and "KHÔNG NỘP" not in str(target.value or ""):
                # khong de banner de len du lieu: dung hang duoi cung + 2
                r = (ws.max_row or 1) + 2
                cell = ws.cell(row=r, column=2, value=BANNER)
                cell.font = Font(name=TNR, bold=True, italic=True, size=10, color="C00000")
                notes.append("banner sheet theo doi " + ws.title)

    # Force TNR moi o co gia tri
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                c.font = font_copy_tnr(c.font)

    set_wb_default_font(wb)
    wb.save(path)
    return notes


def copy_official_to_laptop():
    if not os.path.isdir(os.path.dirname(LAPTOP)):
        return "thieu thu muc cha laptop"
    os.makedirs(LAPTOP, exist_ok=True)
    skip = {"_scripts", "__pycache__"}
    copied = 0
    for root, dirs, files in os.walk(VAULT):
        dirs[:] = [d for d in dirs if d not in skip]
        rel = os.path.relpath(root, VAULT)
        dest_dir = LAPTOP if rel == "." else os.path.join(LAPTOP, rel)
        os.makedirs(dest_dir, exist_ok=True)
        for fn in files:
            if fn.endswith(".pyc"):
                continue
            src = os.path.join(root, fn)
            dst = os.path.join(dest_dir, fn)
            try:
                shutil.copy2(src, dst)
                copied += 1
            except Exception as e:
                print("  COPY FAIL", dst, e)
    return "copied %d files" % copied


def main():
    print("VAULT", VAULT)
    results = []
    for dirpath, dirnames, filenames in os.walk(VAULT):
        if os.path.basename(dirpath) in ("_scripts", "__pycache__"):
            dirnames[:] = []
            continue
        for fn in filenames:
            path = os.path.join(dirpath, fn)
            rel = os.path.relpath(path, VAULT)
            if fn.endswith(".docx"):
                info = patch_docx(path)
                print("DOCX", rel, info)
                results.append((rel, info))
            elif fn.endswith(".xlsx") and not fn.startswith("~$"):
                notes = patch_xlsx(path)
                print("XLSX", rel, notes)
                results.append((rel, notes))
    print("LAPTOP", copy_official_to_laptop())
    return results


if __name__ == "__main__":
    main()
