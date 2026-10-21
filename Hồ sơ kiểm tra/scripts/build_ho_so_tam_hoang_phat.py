# -*- coding: utf-8 -*-
"""Lap ho so di ke hoach MST 3603721832 - Cty TNHH TM DV Tam Hoang Phat.

Phuong phap: dien truong truc tiep tren phoi MAUTONG HOP.docx.
Quy cach xuat: TONG HOP.docx + 2 file Quyet dinh (1 trang A4) + hoso.zip.
"""

import os
import copy
import zipfile
import shutil
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BRAIN_ROOT = "/brains/Hồ sơ kiểm tra"
TEMPLATE_PATH = os.path.join(BRAIN_ROOT, "templates", "MAUTONG HOP.docx")
TARGET_DIR = os.path.join(BRAIN_ROOT, "exports", "tam-hoang-phat-3603721832")

NOI_DUNG_KIEM_TRA = (
    "- Chi phí quản lý doanh nghiệp.\n"
    "- Chi phí bán hàng.\n"
    "- Chi phí tài chính."
)

DE_XUAT_NOI_DUNG = (
    "- Chi phí quản lý doanh nghiệp, thời kỳ kiểm tra: Từ năm 2020 đến năm 2025.\n"
    "- Chi phí bán hàng, thời kỳ kiểm tra: Từ năm 2020 đến năm 2025.\n"
    "- Chi phí tài chính, thời kỳ kiểm tra: Từ năm 2020 đến năm 2025."
)

PHAN_TICH_DT = (
    "- Doanh nghiệp đang hoạt động liên tục từ ngày 11/05/2020 đến năm 2025, "
    "ngành nghề chính lắp đặt hệ thống điện."
)
PHAN_TICH_HD = (
    "- Doanh số hóa đơn điện tử đã tổng hợp: bán ra 105.171.646.830 đồng (364 hóa đơn); "
    "mua vào 94.616.282.174 đồng (13 hóa đơn). Chênh lệch quy mô mua - bán cần đối chiếu "
    "sổ kế toán và báo cáo tài chính."
)
PHAN_TICH_CP = (
    "- Chi phí nhân công biến động mạnh: số lao động quyết toán TNCN từ 21 người (năm 2020) "
    "xuống 14 người (năm 2023) rồi tăng lên 45 người (năm 2025); tổng thu nhập chịu thuế trả "
    "cho người lao động 2020-2025 là 8.487.420.579 đồng; thuế TNCN đã khấu trừ năm 2025 là "
    "76.202.308 đồng, trong đó có khoản chi trả vãng lai.\n"
    "- Chi phí quản lý doanh nghiệp, chi phí bán hàng và chi phí tài chính cần đối chiếu "
    "BCTC với sổ kế toán chi tiết để xác định tính hợp lý, hợp lệ so với doanh thu."
)
PHAN_TICH_KQ = (
    "- Chi phí quản lý doanh nghiệp, chi phí bán hàng và chi phí tài chính có dấu hiệu "
    "biến động không tương ứng với quy mô hóa đơn và chi phí nhân công qua các năm."
)

REPLACEMENTS = {
    "Tên_công_ty": "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ TÂM HOÀNG PHÁT",
    "MST": "3603721832",
    "địa_chỉ": "30A/5, KP4, Phường Hố Nai",
    "ĐỊA_CHỈ": "30A/5, KP4, Phường Hố Nai, Thành phố Đồng Nai",
    "NNKD": "Lắp đặt hệ thống điện",
    "GĐ": "Vũ Văn Sĩ",
    "KT": "Nguyễn Thị Kim Chi",
    "CCCD": "075086016474",
    "ngày_cấp": "…..",
    "Ngày_cấp2": "11/05/2020",
    "Nơi_cấp": "Cục Cảnh sát Quản lý hành chính về trật tự xã hội",
    "GGCNĐKKD": "3603721832",
    "CCCDKT": "075188026200",
    "Chi_nhánh": "Chi nhánh Công ty TNHH TM DV Tâm Hoàng Phát (MST 3603721832-001) tại 521/59, tổ 6, khu phố 1, Phường Long Bình",
    "Tài_khoản_số": ".....",
    "Thời_kỳ_kiểm_tra": "Từ năm 2020 đến năm 2025",
    "Thời_hạn_kiểm_tra_": "20",
    "TĐ": "Nguyễn Thị Ánh Ngọc",
    "CVTĐ": "Chuyên viên",
    "TV": "Nguyễn Thành Đồng",
    "CVTV": "Kiểm tra viên thuế",
    "QĐKT": "DU_THAO-TRINH-KY",
    "QĐGS": "DU_THAO-TRINH-KY",
    "ngày_QĐ": "….. / ….. / 2026",
    "NGÀY_QĐ": "….. / ….. / 2026",
    "Phân_công_trưởng_đoàn": "Chi phí quản lý doanh nghiệp; Chi phí tài chính",
    "Phân_công_thành_viên": "Chi phí bán hàng",
    "Ma_CQT": "1140968",
    "CQT": "Thuế Cơ sở 3 Thành phố Đồng Nai",
    "Từ_ngày": "…..",
    "Đến_ngày": "…..",
    "Ngày_công_bố": "….. / ….. / 2026",
    "Ngày_báo_cáo": "…..",
    "Ngày_ký_Biên_bản_kiểm_tra": "….. / ….. / 2026",
    "Số_tiền_phạt_20": ".....",
    "Số_tiền_phạt_20_BC": ".....",
    "Phạt_KK_sai_VPHC": ".....",
    "Thuế_TNDN": ".....",
    "Thuế_GTGT": ".....",
    "Số_tiền_phạt_003_TNDN": ".....",
    "Số_tiền_phạt_003_GTGT": ".....",
    "Tổng_cộng_phạt_và_truy_thu": ".....",
    "Tổng_cộng_phạt_và_truy_thu_BC": ".....",
    "Tổng_cộng_số_phạt__chậm_nộp": ".....",
    "Tổng_cộng_truy_thu": ".....",
    "Tổng_cộng_truy_thu_BC": ".....",
    "Tổng_phạt": ".....",
    "Tổng_phạt_kk_sai": ".....",
    "Tổng": ".....",
    "Giảm_lỗ": ".....",
    "giảm_khấu_trừ": ".....",
}

# So lieu mau cua vu khac con sot trong phoi; xoa khoi ho so Tam Hoang Phat.
LEFTOVER_AMOUNTS = [
    "13.776.004.215",
    "10.643.637.220",
    "10.565.162.020",
    "8.632.787.527",
    "8.515.074.727",
    "11.260.000.000",
    "1.748.198.921",
    "288.289.977",
    "285.191.642",
    "144.281.478",
    "117.712.800",
    "78.475.200",
]


def set_run_font(paragraph, size_pt=None, bold=None):
    for r in paragraph.runs:
        r.font.name = "Times New Roman"
        r._element.rPr.rFonts.set(docx.oxml.ns.qn("w:eastAsia"), "Times New Roman")
        if size_pt is not None:
            r.font.size = Pt(size_pt)
        if bold is not None:
            r.font.bold = bold


def replace_text_in_paragraph(paragraph, replacements):
    text = paragraph.text
    if not any(("«%s»" % k) in text for k in replacements) and not any(k in text for k in leftover_keys()):
        return

    for run in paragraph.runs:
        for k, v in replacements.items():
            pattern = "«%s»" % k
            if pattern in run.text:
                run.text = run.text.replace(pattern, v)
        for amt in LEFTOVER_AMOUNTS:
            if amt in run.text:
                run.text = run.text.replace(amt, ".....")

    if "«" in paragraph.text and "»" in paragraph.text:
        full_text = paragraph.text
        for k, v in replacements.items():
            full_text = full_text.replace("«%s»" % k, v)
        if paragraph.runs:
            paragraph.runs[0].text = full_text
            for r in paragraph.runs[1:]:
                r.text = ""

    if any(amt in paragraph.text for amt in LEFTOVER_AMOUNTS):
        full_text = paragraph.text
        for amt in LEFTOVER_AMOUNTS:
            full_text = full_text.replace(amt, ".....")
        if paragraph.runs:
            paragraph.runs[0].text = full_text
            for r in paragraph.runs[1:]:
                r.text = ""


def leftover_keys():
    return LEFTOVER_AMOUNTS


def replace_in_doc(doc, replacements):
    for p in doc.paragraphs:
        replace_text_in_paragraph(p, replacements)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_text_in_paragraph(p, replacements)
    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header, section.first_page_footer):
            if part is None:
                continue
            for p in part.paragraphs:
                replace_text_in_paragraph(p, replacements)
            for t in part.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            replace_text_in_paragraph(p, replacements)


def remove_floating_table_properties(table):
    tblPr = table._element.tblPr
    if tblPr is None:
        return
    tblpPr = tblPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblpPr")
    if tblpPr is not None:
        tblPr.remove(tblpPr)
    tblOverlap = tblPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblOverlap")
    if tblOverlap is not None:
        tblPr.remove(tblOverlap)


def format_all_tables_inline(doc):
    for t in doc.tables:
        remove_floating_table_properties(t)
        txt = " ".join(c.text for row in t.rows for c in row.cells)
        if "Mẫu số:" in txt or "Mẫu:" in txt:
            t.alignment = WD_TABLE_ALIGNMENT.RIGHT
            for row in t.rows:
                for c in row.cells:
                    for p in c.paragraphs:
                        p.paragraph_format.line_spacing = 1.0
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)


def apply_tnr_size(paragraph, size_pt, spacing=1.15):
    paragraph.paragraph_format.line_spacing = spacing
    paragraph.paragraph_format.space_after = Pt(1)
    paragraph.paragraph_format.space_before = Pt(0)
    if not paragraph.runs:
        paragraph.add_run(paragraph.text)
    set_run_font(paragraph, size_pt)


def build_master_document():
    doc = docx.Document(TEMPLATE_PATH)
    replace_in_doc(doc, REPLACEMENTS)

    for p in doc.paragraphs:
        if "«Nội_dung_kiểm_tra»" in p.text or p.text.strip() == "«Nội_dung_kiểm_tra»":
            p.text = NOI_DUNG_KIEM_TRA
            apply_tnr_size(p, 12, 1.0)
            insert_newlines_as_breaks(p)
        elif "Thời hạn kiểm tra là" in p.text and "ngày làm việc" in p.text:
            p.text = (
                "Thời hạn kiểm tra là 20 ngày làm việc (không bao gồm ngày nghỉ, ngày lễ "
                "theo quy định của pháp luật) kể từ ngày công bố Quyết định kiểm tra."
            )
            apply_tnr_size(p, 12, 1.0)

    for i, p in enumerate(doc.paragraphs):
        if "d) Số chi nhánh, cơ sở phụ thuộc:" in p.text:
            p.text = (
                "d) Số chi nhánh, cơ sở phụ thuộc: Chi nhánh Công ty TNHH TM DV Tâm Hoàng Phát "
                "(MST 3603721832-001) tại 521/59, tổ 6, khu phố 1, Phường Long Bình."
            )
            apply_tnr_size(p, 13, 1.15)
        elif "Doanh nghiệp kê khai lãi từ năm 2022" in p.text:
            p.text = PHAN_TICH_DT
            apply_tnr_size(p, 13, 1.15)
        elif "Tỷ lệ giá vốn hàng bán trên doanh thu năm 2022" in p.text:
            p.text = PHAN_TICH_HD
            apply_tnr_size(p, 13, 1.15)
        elif "Tỷ lệ Chi phí quản lý kinh doanh trên doanh thu" in p.text:
            p.text = PHAN_TICH_CP
            apply_tnr_size(p, 13, 1.15)
            insert_newlines_as_breaks(p)
        elif "Tỷ lệ giá vốn hàng bán và chi phí quản lý kinh doanh trên doanh thu biến động không đều" in p.text:
            p.text = PHAN_TICH_KQ
            apply_tnr_size(p, 13, 1.15)
        elif "IV. ĐỀ XUẤT NỘI DUNG KIỂM TRA:" in p.text:
            if i + 1 < len(doc.paragraphs):
                p_next = doc.paragraphs[i + 1]
                p_next.text = DE_XUAT_NOI_DUNG
                apply_tnr_size(p_next, 13, 1.15)
                insert_newlines_as_breaks(p_next)

    format_all_tables_inline(doc)
    clean_master_document_flow(doc)
    return doc


def clean_master_document_flow(doc):
    for p in list(doc.paragraphs):
        has_pb = any(
            br.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type") == "page"
            for br in p._element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br")
        )
        pbb = len(p._element.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore")) > 0
        if not p.text.strip() and not has_pb and not pbb:
            p_el = p._element
            parent = p_el.getparent()
            if parent is not None:
                parent.remove(p_el)

    keep_names = (
        "Phạm Đồng Xuân",
        "Nguyễn Trường Giang",
        "Nguyễn Thị Ánh Ngọc",
        "Nguyễn Thành Đồng",
        "Vũ Văn Sĩ",
        "Nguyễn Thị Kim Chi",
    )
    for p in doc.paragraphs:
        p.paragraph_format.line_spacing = 1.05
        if p.paragraph_format.space_after and p.paragraph_format.space_after.pt > 2:
            p.paragraph_format.space_after = Pt(1.5)
        if p.paragraph_format.space_before and p.paragraph_format.space_before.pt > 2:
            p.paragraph_format.space_before = Pt(0)
        t = p.text.strip()
        if any(sig in t for sig in keep_names):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0

    in_thu_td = False
    for p in doc.paragraphs:
        if "THƯ TRAO ĐỔI THÔNG TIN" in p.text:
            in_thu_td = True
        elif in_thu_td and ("BIÊN BẢN" in p.text and "Công bố" in p.text):
            in_thu_td = False
        if in_thu_td:
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(0)
            for r in p.runs:
                if r.font.size and r.font.size.pt > 12:
                    r.font.size = Pt(11.5)

    in_bbkt = False
    for p in doc.paragraphs:
        if "BIÊN BẢN KIỂM TRA" in p.text:
            in_bbkt = True
        elif in_bbkt and ("Công khai và bàn giao" in p.text):
            in_bbkt = False
        if in_bbkt:
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(0)

    for t in doc.tables:
        for row in t.rows:
            trPr = row._element.trPr
            if trPr is None:
                trPr = OxmlElement("w:trPr")
                row._element.append(trPr)
            cantSplit = OxmlElement("w:cantSplit")
            trPr.append(cantSplit)
            for c in row.cells:
                for p in c.paragraphs:
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.space_before = Pt(0)


def elem_text(el):
    texts = el.findall(".//" + qn("w:t"))
    return "".join([node.text for node in texts if node.text]).strip()


def insert_newlines_as_breaks(paragraph):
    for run in paragraph.runs:
        if "\n" not in run.text:
            continue
        parts = run.text.split("\n")
        run.text = parts[0]
        for part in parts[1:]:
            br = OxmlElement("w:br")
            run._element.append(br)
            t = OxmlElement("w:t")
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = part
            run._element.append(t)


def find_segment_range(master_doc, start_marker, end_marker):
    elems = list(master_doc._element.body)
    start = None
    end = None
    for i, el in enumerate(elems):
        t = elem_text(el)
        if start is None and start_marker in t:
            start = i
            for j in range(i, -1, -1):
                tj = elem_text(elems[j])
                if "THUẾ THÀNH PHỐ ĐỒNG NAI" in tj and elems[j].tag.endswith("tbl"):
                    start = j
                    break
        if start is not None and i > start and end_marker in t:
            end = i
            for j in range(i, start, -1):
                if "Nơi nhận" in elem_text(elems[j]):
                    end = j + 1
                    break
            break
    if start is None or end is None:
        raise RuntimeError("Khong tim duoc doan: %s -> %s (start=%s end=%s)" % (start_marker, end_marker, start, end))
    return start, end


def extract_segment(master_doc, start, end):
    doc = docx.Document(TEMPLATE_PATH)
    body = doc._element.body
    sectPr = None
    for el in list(body):
        if el.tag.endswith("sectPr"):
            sectPr = copy.deepcopy(el)
            break
    all_master_elems = list(master_doc._element.body)
    segment = [copy.deepcopy(all_master_elems[i]) for i in range(start, end)]
    if segment:
        el0 = segment[0]
        if el0.tag.endswith("p"):
            for br in el0.findall(".//" + qn("w:br")):
                if br.get(qn("w:type")) == "page":
                    parent = br.getparent()
                    if parent is not None:
                        parent.remove(br)
            t = elem_text(el0)
            if not t:
                segment.pop(0)
        if segment:
            for br in segment[0].findall(".//" + qn("w:br")):
                if br.get(qn("w:type")) == "page":
                    parent = br.getparent()
                    if parent is not None:
                        parent.remove(br)
    body.clear()
    for el in segment:
        body.append(el)
    if sectPr is not None:
        body.append(sectPr)
    return doc


def format_single_page_qd(doc, title_hint):
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2.7)
    sec.right_margin = Cm(1.8)

    for t in doc.tables:
        remove_floating_table_properties(t)
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        if r.font.size and r.font.size.pt > 12.5:
                            r.font.size = Pt(12)

    for p in list(doc.paragraphs):
        text = p.text.strip()
        if not text:
            if p._element.getparent() is not None:
                p._element.getparent().remove(p._element)
            continue
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        for r in p.runs:
            r.font.name = "Times New Roman"

        if text == "QUYẾT ĐỊNH":
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.size = Pt(13)
                r.font.bold = True
        elif "Về việc" in text or "CÔNG TY TNHH" in text:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.size = Pt(12)
                r.font.bold = True
        elif "TRƯỞNG THUẾ CƠ SỞ" in text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.size = Pt(12)
                r.font.bold = True
        elif text.startswith("Căn cứ") or text.startswith("Theo đề nghị"):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                r.font.size = Pt(11.5)
        elif text == "QUYẾT ĐỊNH:":
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.size = Pt(12.5)
                r.font.bold = True
        elif text.startswith("Điều "):
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for r in p.runs:
                r.font.size = Pt(12)
        else:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                if r.font.size and r.font.size.pt > 12:
                    r.font.size = Pt(12)
    return doc


def remaining_merge_fields(doc):
    leftover = []
    for p in doc.paragraphs:
        if "«" in p.text and "»" in p.text:
            leftover.append(p.text.strip()[:120])
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if "«" in cell.text and "»" in cell.text:
                    leftover.append(cell.text.strip()[:120])
    return leftover


def main():
    print("=== LAP HO SO DI TAM HOANG PHAT MST 3603721832 ===")
    if not os.path.exists(TEMPLATE_PATH):
        raise SystemExit("THIEU PHOI: %s" % TEMPLATE_PATH)
    os.makedirs(TARGET_DIR, exist_ok=True)

    print("1. Dien du lieu tren phoi MAUTONG HOP.docx ...")
    master_doc = build_master_document()
    leftover = remaining_merge_fields(master_doc)
    if leftover:
        print("CANH BAO con truong tron:")
        for x in leftover[:20]:
            print("  -", x)

    tong_hop_path = os.path.join(TARGET_DIR, "TONG HOP.docx")
    master_doc.save(tong_hop_path)
    print("-> TONG HOP.docx (%s bytes)" % os.path.getsize(tong_hop_path))

    print("2. Xuat Quyet dinh 1 trang A4 ...")
    qd_kt_path = os.path.join(TARGET_DIR, "QD Kiem tra tai cty Tam Hoang Phat.docx")
    qd_gs_path = os.path.join(TARGET_DIR, "QDgiam sat tai cty Tam Hoang Phat.docx")
    kt_start, kt_end = find_segment_range(
        master_doc, "Về việc kiểm tra thuế tại", "Về việc giám sát hoạt động của Đoàn kiểm tra tại"
    )
    gs_start, gs_end = find_segment_range(
        master_doc, "Về việc giám sát hoạt động của Đoàn kiểm tra tại", "PHÂN TÍCH THÔNG TIN, XÁC ĐỊNH NỘI DUNG CUỘC KIỂM TRA"
    )
    print("  QDKT range", kt_start, kt_end, "QDGS range", gs_start, gs_end)
    qd_kt = extract_segment(master_doc, kt_start, kt_end)
    format_single_page_qd(qd_kt, "kiem tra")
    qd_kt.save(qd_kt_path)
    print("->", os.path.basename(qd_kt_path), os.path.getsize(qd_kt_path), "bytes")
    qd_gs = extract_segment(master_doc, gs_start, gs_end)
    format_single_page_qd(qd_gs, "giam sat")
    qd_gs.save(qd_gs_path)
    print("->", os.path.basename(qd_gs_path), os.path.getsize(qd_gs_path), "bytes")

    print("3. Dong goi hoso.zip ...")
    zip_path = os.path.join(TARGET_DIR, "hoso.zip")
    files_to_zip = [
        "TONG HOP.docx",
        "QD Kiem tra tai cty Tam Hoang Phat.docx",
        "QDgiam sat tai cty Tam Hoang Phat.docx",
    ]
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in files_to_zip:
            fpath = os.path.join(TARGET_DIR, fname)
            zf.write(fpath, fname)
            print("  +", fname)
    print("-> ZIP", zip_path, os.path.getsize(zip_path), "bytes")
    print("=== XONG ===")


if __name__ == "__main__":
    main()
