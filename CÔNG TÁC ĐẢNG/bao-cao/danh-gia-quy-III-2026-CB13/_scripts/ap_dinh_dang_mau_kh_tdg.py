# -*- coding: utf-8 -*-
"""Clone dinh dang 2 file mau KH_1.docx va TDG_1.xlsx sang file con lai o 05 va 06.
Khong doi ten, nhiem vu, KPI, cong thuc. Khong FitToPages. Khong goi table.columns.
"""
from __future__ import print_function

import io
import os
import shutil
import zipfile
from copy import deepcopy
from xml.etree import ElementTree as ET

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

VAULT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
LAPTOP = (
    r"H:\Other computers\My Laptop\2. CONG TAC KHAC\4. CONG TAC DANG"
    r"\A PHUONG\2026\06_Tieu_chi_danh_gia\Du_thao_CB13_QuyIII_2026"
)
MAU_KH = r"D:\SECOND BRAIN\javis-os\server\.staging\KH_QIII-2026_1_Le-Huu-Nghia.docx"
MAU_TDG = r"D:\SECOND BRAIN\javis-os\server\.staging\TDG_QIII-2026_1_Le-Huu-Nghia.xlsx"

DIR_KH = os.path.join(VAULT, "05_Ke_hoach_ca_nhan")
DIR_TDG = os.path.join(VAULT, "06_Ban_tu_danh_gia")

TNR = "Times New Roman"
COL_W = [441, 2857, 989, 2527, 1000, 1911, 1879, 1099, 1484, 1319]
EXTRA_W = 32
TBL_W = 15538
SIGN_W = (6814, 8724)
NOTE_PREFIX = "Số màu đỏ là ước tính của Javis"


def dual_copy(src):
    rel = os.path.relpath(src, VAULT)
    dst = os.path.join(LAPTOP, rel)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)
    return dst


def p_text(p_el):
    return "".join((t.text or "") for t in p_el.findall(".//" + qn("w:t")))


def set_tc_w(tc, width):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.insert(0, tcW)
    tcW.set(qn("w:w"), str(width))
    tcW.set(qn("w:type"), "dxa")
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
        for side in ("top", "left", "bottom", "right"):
            el = OxmlElement("w:" + side)
            el.set(qn("w:w"), "40")
            el.set(qn("w:type"), "dxa")
            tcMar.append(el)


def set_grid_span(tc, span):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    gs = tcPr.find(qn("w:gridSpan"))
    if gs is None:
        gs = OxmlElement("w:gridSpan")
        tcW = tcPr.find(qn("w:tcW"))
        if tcW is not None:
            tcW.addnext(gs)
        else:
            tcPr.insert(0, gs)
    gs.set(qn("w:val"), str(span))


def set_run_sz(r, half_points):
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r.insert(0, rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), TNR)
    val = str(half_points)
    for tag in ("sz", "szCs"):
        el = rPr.find(qn("w:" + tag))
        if el is None:
            el = OxmlElement("w:" + tag)
            rPr.append(el)
        el.set(qn("w:val"), val)


def replace_tbl_grid(tbl, widths):
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is not None:
            tblPr.addnext(grid)
        else:
            tbl.insert(0, grid)
    for child in list(grid):
        grid.remove(child)
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)


def apply_kh_file(path, header_tbl_src, sect_pr_src):
    d = Document(path)
    body = d.element.body

    old_sect = body.find(qn("w:sectPr"))
    if old_sect is not None:
        old_sect.addprevious(deepcopy(sect_pr_src))
        body.remove(old_sect)

    tbls = body.findall(qn("w:tbl"))
    if len(tbls) < 2:
        raise SystemExit("thieu bang: " + path)
    old_header = tbls[0]
    old_header.addprevious(deepcopy(header_tbl_src))
    body.remove(old_header)

    to_drop = []
    for p in body.findall(qn("w:p")):
        txt = p_text(p).strip()
        if txt == "-----" or txt.startswith("(Chi bộ 13 ban hành kèm theo Quyết định"):
            to_drop.append(p)
    for p in to_drop:
        body.remove(p)

    tbls = body.findall(qn("w:tbl"))
    plan = tbls[1]
    tblPr = plan.find(qn("w:tblPr"))
    tblW = tblPr.find(qn("w:tblW")) if tblPr is not None else None
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        if tblPr is None:
            tblPr = OxmlElement("w:tblPr")
            plan.insert(0, tblPr)
        tblPr.insert(0, tblW)
    tblW.set(qn("w:w"), str(TBL_W))
    tblW.set(qn("w:type"), "dxa")
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblPr.append(jc)
    jc.set(qn("w:val"), "center")
    replace_tbl_grid(plan, COL_W + [EXTRA_W])

    rows = plan.findall(qn("w:tr"))
    for ri, tr in enumerate(rows):
        cells = tr.findall(qn("w:tc"))
        if len(cells) >= 10:
            for ci in range(10):
                set_tc_w(cells[ci], COL_W[ci])
        elif len(cells) == 2:
            set_tc_w(cells[0], SIGN_W[0])
            set_grid_span(cells[0], 4)
            set_tc_w(cells[1], SIGN_W[1])
            set_grid_span(cells[1], 7)
        for r in tr.findall(".//" + qn("w:r")):
            set_run_sz(r, 22)
    for el in plan.findall(".//" + qn("w:sz")):
        el.set(qn("w:val"), "22")
    for el in plan.findall(".//" + qn("w:szCs")):
        el.set(qn("w:val"), "22")

    st = d.styles["Normal"]
    rPr = st.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), TNR)
    for tag in ("sz", "szCs"):
        el = rPr.find(qn("w:" + tag))
        if el is not None:
            rPr.remove(el)
    pPr = st.element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        st.element.append(pPr)
    sp = pPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        pPr.append(sp)
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "288")
    sp.set(qn("w:lineRule"), "auto")

    d.save(path)
    print("kh", os.path.basename(path), os.path.getsize(path), "drop_paras", len(to_drop))


def copy_col_widths(src_ws, dst_ws, max_col=12):
    for i in range(1, max_col + 1):
        letter = get_column_letter(i)
        sdim = src_ws.column_dimensions.get(letter)
        ddim = dst_ws.column_dimensions[letter]
        if sdim is None:
            continue
        ddim.width = sdim.width
        ddim.hidden = sdim.hidden


def copy_row_heights(src_ws, dst_ws, max_row=70):
    for i in range(1, max_row + 1):
        sdim = src_ws.row_dimensions.get(i)
        if sdim is None or sdim.height is None:
            continue
        dst_ws.row_dimensions[i].height = sdim.height


def apply_print_khung(src_ws, dst_ws):
    dst_ws.page_setup.paperSize = src_ws.page_setup.paperSize
    dst_ws.page_setup.orientation = src_ws.page_setup.orientation
    dst_ws.page_setup.fitToHeight = 0
    dst_ws.page_setup.scale = src_ws.page_setup.scale
    dst_ws.page_setup.fitToWidth = None
    dst_ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    m = src_ws.page_margins
    dst_ws.page_margins.left = m.left
    dst_ws.page_margins.right = m.right
    dst_ws.page_margins.top = m.top
    dst_ws.page_margins.bottom = m.bottom
    dst_ws.page_margins.header = m.header
    dst_ws.page_margins.footer = m.footer
    dst_ws.print_options.horizontalCentered = True
    dst_ws.print_area = src_ws.print_area
    dst_ws.sheet_view.view = "pageBreakPreview"
    if src_ws.sheet_view.zoomScale:
        dst_ws.sheet_view.zoomScale = src_ws.sheet_view.zoomScale
    dst_ws.sheet_format.defaultRowHeight = src_ws.sheet_format.defaultRowHeight
    dst_ws.sheet_format.defaultColWidth = src_ws.sheet_format.defaultColWidth
    copy_col_widths(src_ws, dst_ws, 12)
    copy_row_heights(src_ws, dst_ws, 70)


def apply_tdg_file(path, src_wb):
    wb = load_workbook(path)
    src_k = src_wb["Khung đánh giá"]
    src_k2 = src_wb["Tinh KPI 2A-2B"]
    dst_k = wb["Khung đánh giá"]
    dst_k2 = wb["Tinh KPI 2A-2B"]
    apply_print_khung(src_k, dst_k)
    apply_print_khung(src_k2, dst_k2)
    f4 = dst_k["F4"].value
    if isinstance(f4, str) and f4.startswith("Đồng Nai"):
        dst_k["F4"].value = f4.replace("Đồng Nai", "Trấn Biên", 1)
    cleared = 0
    for r in range(39, 48):
        cell = dst_k.cell(r, 9)
        val = cell.value
        if isinstance(val, str) and val.startswith(NOTE_PREFIX):
            cell.value = None
            cleared += 1
    wb.save(path)
    wb.close()
    print("tdg", os.path.basename(path), os.path.getsize(path), "cleared_notes", cleared)


NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAW_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"


def _ensure_ct_drawing(data):
    ET.register_namespace("", NS_CT)
    root = ET.fromstring(data)
    has = any(
        ov.tag.endswith("Override") and ov.get("PartName") == "/xl/drawings/drawing1.xml"
        for ov in root
    )
    if not has:
        ov = ET.SubElement(root, "{%s}Override" % NS_CT)
        ov.set("PartName", "/xl/drawings/drawing1.xml")
        ov.set("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_sheet_rels(data=None):
    ET.register_namespace("", NS_REL)
    if data:
        root = ET.fromstring(data)
    else:
        root = ET.Element("{%s}Relationships" % NS_REL)
    rid = None
    for rel in root:
        if rel.get("Type") == DRAW_TYPE:
            rid = rel.get("Id")
            break
    if rid is None:
        used = {rel.get("Id") for rel in root}
        rid = "rId2" if "rId2" not in used else "rIdDraw"
        rel = ET.SubElement(root, "{%s}Relationship" % NS_REL)
        rel.set("Id", rid)
        rel.set("Type", DRAW_TYPE)
        rel.set("Target", "../drawings/drawing1.xml")
    xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return xml, rid


def _ensure_sheet_drawing_tag(data, rid):
    import re
    text = data.decode("utf-8") if isinstance(data, bytes) else data
    rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    if "xmlns:r=" not in text[:1200]:
        text = re.sub(r"<worksheet\b", '<worksheet xmlns:r="%s"' % rns, text, count=1)
    if "<drawing " not in text:
        tag = '<drawing r:id="%s"/>' % rid
        if "</worksheet>" in text:
            text = text.replace("</worksheet>", tag + "</worksheet>")
    return text.encode("utf-8")


def inject_tdg_underline(target, sample):
    with zipfile.ZipFile(sample, "r") as zs:
        drawing = zs.read("xl/drawings/drawing1.xml")
    buf = io.BytesIO()
    with zipfile.ZipFile(target, "r") as zin:
        names = set(zin.namelist())
        files = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        infos = {item.filename: item for item in zin.infolist()}
    files["[Content_Types].xml"] = _ensure_ct_drawing(files["[Content_Types].xml"])
    rels_name = "xl/worksheets/_rels/sheet1.xml.rels"
    rels_xml, rid = _ensure_sheet_rels(files.get(rels_name))
    files[rels_name] = rels_xml
    files["xl/worksheets/sheet1.xml"] = _ensure_sheet_drawing_tag(
        files["xl/worksheets/sheet1.xml"], rid
    )
    files["xl/drawings/drawing1.xml"] = drawing
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    with open(target, "wb") as f:
        f.write(buf.getvalue())


def main():
    if not os.path.isfile(MAU_KH) or not os.path.isfile(MAU_TDG):
        raise SystemExit("thieu file mau staging")

    kh1 = os.path.join(DIR_KH, "KH_QIII-2026_1_Le-Huu-Nghia.docx")
    tdg1 = os.path.join(DIR_TDG, "TDG_QIII-2026_1_Le-Huu-Nghia.xlsx")
    shutil.copy2(MAU_KH, kh1)
    shutil.copy2(MAU_TDG, tdg1)
    dual_copy(kh1)
    dual_copy(tdg1)
    print("copied mau file 1", os.path.getsize(kh1), os.path.getsize(tdg1))

    mau_doc = Document(MAU_KH)
    header_tbl = mau_doc.tables[0]._tbl
    sect_pr = mau_doc.element.body.find(qn("w:sectPr"))
    for fn in sorted(os.listdir(DIR_KH)):
        if not fn.lower().endswith(".docx"):
            continue
        if fn.startswith("KH_QIII-2026_1_"):
            continue
        path = os.path.join(DIR_KH, fn)
        apply_kh_file(path, header_tbl, sect_pr)
        dual_copy(path)

    src_wb = load_workbook(MAU_TDG)
    for fn in sorted(os.listdir(DIR_TDG)):
        if not fn.lower().endswith(".xlsx"):
            continue
        path = os.path.join(DIR_TDG, fn)
        if fn.startswith("TDG_QIII-2026_1_"):
            continue
        apply_tdg_file(path, src_wb)
        inject_tdg_underline(path, MAU_TDG)
        dual_copy(path)
        print("injected drawing", fn, os.path.getsize(path))
    src_wb.close()


if __name__ == "__main__":
    main()
