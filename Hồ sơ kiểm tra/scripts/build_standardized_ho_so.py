# -*- coding: utf-8 -*-
"""
Script chuẩn hóa toàn bộ hồ sơ kiểm tra theo đúng quy chuẩn:
1. Định dạng Quyết định kiểm tra và Quyết định giám sát chuẩn xác đúng 1 trang A4, không tràn sang trang khác.
2. Khung mẫu hồ sơ (Mẫu số: 08/QTKT-TC, Mẫu số: 12/QTKT-TC...) nằm trên cùng bên phải của văn bản, không bị nằm ở trang riêng lẻ.
3. Tự động loại bỏ toàn bộ các trang trắng không cần thiết và các trang chữ ký mồ côi trong file tổng hợp TONG HOP.docx.
4. Chỉ xuất riêng các file Quyết định độc lập, toàn bộ các phần/biểu mẫu còn lại lưu trong file tổng hợp TONG HOP.docx.
5. Đóng gói hoso.zip hoàn chỉnh.
"""

import os
import copy
import zipfile
import re
import shutil
import tempfile
import docx
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

TEMPLATE_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\BIEN BAN\MAUTONG HOP.docx"
TARGET_DIR = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\9. VAN PHAT THANH\0. BIEN BAN"
BRAIN_TEMPLATE_DIR = r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates"

REPLACEMENTS = {
    "Tên_công_ty": "CÔNG TY TNHH MỘT THÀNH VIÊN BAO BÌ GIẤY VẠN PHÁT THÀNH",
    "MST": "3601409441",
    "địa_chỉ": "Số 9/3, đường Nguyễn văn Tỏ, khu phố 2, Phường Long Hưng",
    "ĐỊA_CHỈ": "Số 9/3, đường Nguyễn văn Tỏ, khu phố 2, Phường Long Hưng, Thành phố Biên Hòa, Tỉnh Đồng Nai",
    "NNKD": "Sản xuất giấy nhăn, bìa nhăn, bao bì từ giấy và bìa",
    "GĐ": "Tô Quốc Thi",
    "KT": "Trần Thị Thoan",
    "CCCD": "096075006910",
    "ngày_cấp": "28/09/2021",
    "Ngày_cấp2": "28/09/2021",
    "Nơi_cấp": "Cục Cảnh sát Quản lý hành chính về trật tự xã hội",
    "GGCNĐKKD": "3601409441",
    "CCCDKT": "036176008691",
    "Chi_nhánh": "Cửa hàng Bao Bì Giấy Vạn Phát Thành tại số 507, Quốc Lộ 51, KP 3, Phường An Hòa (hết hiệu lực)",
    "Tài_khoản_số": "1000148510001 tại Indovina Bank - CN Đồng Nai",
    "Thời_kỳ_kiểm_tra": "Từ năm 2021 đến năm 2025",
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
    "giảm_khấu_trừ": "....."
}

def replace_text_in_paragraph(paragraph, replacements):
    text = paragraph.text
    if not any(k in text for k in replacements):
        return
    
    for run in paragraph.runs:
        for k, v in replacements.items():
            pattern = f"«{k}»"
            if pattern in run.text:
                run.text = run.text.replace(pattern, v)
    
    if "«" in paragraph.text and "»" in paragraph.text:
        full_text = paragraph.text
        for k, v in replacements.items():
            pattern = f"«{k}»"
            if pattern in full_text:
                full_text = full_text.replace(pattern, v)
        if paragraph.runs:
            paragraph.runs[0].text = full_text
            for r in paragraph.runs[1:]:
                r.text = ""

def replace_in_doc(doc, replacements):
    for p in doc.paragraphs:
        replace_text_in_paragraph(p, replacements)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_text_in_paragraph(p, replacements)

def remove_floating_table_properties(table):
    """Xóa bỏ các thuộc tính floating table (tblpPr, tblOverlap) gây lỗi nhảy trang."""
    tblPr = table._element.tblPr
    tblpPr = tblPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblpPr')
    if tblpPr is not None:
        tblPr.remove(tblpPr)
    tblOverlap = tblPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblOverlap')
    if tblOverlap is not None:
        tblPr.remove(tblOverlap)

def format_all_tables_inline(doc):
    """Đảm bảo 100% bảng trong văn bản là inline tables, không float, không nhảy trang."""
    for t in doc.tables:
        remove_floating_table_properties(t)
        txt = ' '.join(c.text for row in t.rows for c in row.cells)
        if 'Mẫu số:' in txt or 'Mẫu:' in txt:
            t.alignment = WD_TABLE_ALIGNMENT.RIGHT
            for row in t.rows:
                for c in row.cells:
                    for p in c.paragraphs:
                        p.paragraph_format.line_spacing = 1.0
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)

def build_master_document():
    doc = docx.Document(TEMPLATE_PATH)
    
    # 1. Thay thế placeholders
    replace_in_doc(doc, REPLACEMENTS)
    
    # 2. Xử lý Nội dung kiểm tra 3 chuyên đề và thời hạn kiểm tra
    for p in doc.paragraphs:
        if "«Nội_dung_kiểm_tra»" in p.text:
            p.text = "- Chi phí quản lý doanh nghiệp.\n- Chi phí bán hàng.\n- Chi phí tài chính."
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(0)
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(12)
        elif "Thời hạn kiểm tra là" in p.text and "ngày làm việc" in p.text:
            p.text = "Thời hạn kiểm tra là 20 ngày làm việc (không bao gồm ngày nghỉ, ngày lễ theo quy định của pháp luật) kể từ ngày công bố Quyết định kiểm tra."
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(0)
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(12)
                
    # 3. Xử lý DI-01 (Mẫu 08/QTKT-TC phân tích rủi ro hồ sơ NNT)
    for i, p in enumerate(doc.paragraphs):
        if "d) Số chi nhánh, cơ sở phụ thuộc: ." in p.text:
            p.text = "d) Số chi nhánh, cơ sở phụ thuộc: Cửa hàng Bao Bì Giấy Vạn Phát Thành tại số 507, Quốc Lộ 51, KP 3, Phường An Hòa (hết hiệu lực)."
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(13)
                
        elif "Doanh nghiệp kê khai lãi từ năm 2022" in p.text:
            p.text = "- Doanh nghiệp kê khai kết quả hoạt động kinh doanh có lãi liên tục từ năm 2021 đến năm 2025, doanh thu thuần tăng trưởng đều qua các năm (năm 2021: 4.275.904.321 đ, năm 2022: 5.304.710.195 đ, năm 2023: 6.851.592.363 đ, năm 2024: 8.652.414.302 đ, năm 2025: 10.060.947.641 đ)."
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(13)
                
        elif "Tỷ lệ giá vốn hàng bán trên doanh thu năm 2022" in p.text:
            p.text = "- Tỷ lệ giá vốn hàng bán trên doanh thu thuần duy trì ở mức cao và có xu hướng tăng dần (năm 2021: 95,21%, năm 2022: 95,57%, năm 2023: 96,43%, năm 2024: 96,57%, năm 2025: 97,84%)."
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(13)
                
        elif "Tỷ lệ Chi phí quản lý kinh doanh trên doanh thu" in p.text:
            p.text = "- Tỷ lệ chi phí quản lý doanh nghiệp trên doanh thu thuần giảm mạnh qua các năm (năm 2021: 2,73%, năm 2022: 1,89%, năm 2023: 0,45%, năm 2024: 0,41%, năm 2025: 0,33%).\n- Tỷ lệ chi phí bán hàng trên doanh thu thuần (năm 2021: 0,10%, năm 2022: 0,00%, năm 2023: 0,00%, năm 2024: 0,00%, năm 2025: 0,05%).\n- Tỷ lệ chi phí tài chính trên doanh thu thuần (năm 2021: 0,73%, năm 2022: 0,64%, năm 2023: 1,22%, năm 2024: 0,77%, năm 2025: 0,03%)."
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(13)
                
        elif "Tỷ lệ giá vốn hàng bán và chi phí quản lý kinh doanh trên doanh thu biến động không đều" in p.text:
            p.text = "- Chi phí quản lý doanh nghiệp, chi phí bán hàng và chi phí tài chính có sự biến động không tương ứng với tốc độ tăng trưởng doanh thu thuần của doanh nghiệp."
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                r.font.name = "Times New Roman"
                r.font.size = Pt(13)
                
        elif "IV. ĐỀ XUẤT NỘI DUNG KIỂM TRA:" in p.text:
            if i + 1 < len(doc.paragraphs):
                p_next = doc.paragraphs[i+1]
                p_next.text = "- Chi phí quản lý doanh nghiệp, thời kỳ kiểm tra: Từ năm 2021 đến năm 2025.\n- Chi phí bán hàng, thời kỳ kiểm tra: Từ năm 2021 đến năm 2025.\n- Chi phí tài chính, thời kỳ kiểm tra: Từ năm 2021 đến năm 2025."
                p_next.paragraph_format.line_spacing = 1.15
                for r in p_next.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(13)
            if i + 2 < len(doc.paragraphs) and ("Phân_công_thành_viên" in doc.paragraphs[i+2].text or "Chi phí bán hàng" in doc.paragraphs[i+2].text):
                doc.paragraphs[i+2].text = ""

    # 4. Chuyển đổi toàn bộ floating tables thành inline tables
    format_all_tables_inline(doc)
    
    # 5. Dọn dẹp khoảng trống và trang trắng trong master_doc
    clean_master_document_flow(doc)

    return doc

def clean_master_document_flow(doc):
    """
    Dọn dẹp triệt để các khoảng cách thừa, các đoạn trống cạnh ngắt trang
    để không sinh ra trang trắng hoặc đẩy chữ ký sang trang riêng.
    """
    # 1. Xóa các đoạn văn rỗng không mang ngắt trang
    for p in list(doc.paragraphs):
        has_pb = any(br.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page' for br in p._element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br'))
        pbb = len(p._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore')) > 0
        if not p.text.strip() and not has_pb and not pbb:
            p_el = p._element
            parent = p_el.getparent()
            if parent is not None:
                parent.remove(p_el)

    # 2. Tinh chỉnh khoảng cách đoạn văn tổng thể
    for p in doc.paragraphs:
        p.paragraph_format.line_spacing = 1.05
        if p.paragraph_format.space_after and p.paragraph_format.space_after.pt > 2:
            p.paragraph_format.space_after = Pt(1.5)
        if p.paragraph_format.space_before and p.paragraph_format.space_before.pt > 2:
            p.paragraph_format.space_before = Pt(0)
        t = p.text.strip()
        if any(sig in t for sig in ['Phạm Đồng Xuân', 'Nguyễn Trường Giang', 'Nguyễn Thị Ánh Ngọc', 'Tô Quốc Thi', 'Trần Thị Thoan', 'Nguyễn Thành Đồng']):
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0

    # 3. Tinh chỉnh riêng Thư trao đổi thông tin (Doc 05) để trọn vẹn 1 trang
    in_thu_td = False
    for p in doc.paragraphs:
        if 'THƯ TRAO ĐỔI THÔNG TIN' in p.text:
            in_thu_td = True
        elif in_thu_td and ('BIÊN BẢN' in p.text and 'Công bố' in p.text):
            in_thu_td = False
        if in_thu_td:
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(0)
            for r in p.runs:
                if r.font.size and r.font.size.pt > 12:
                    r.font.size = Pt(11.5)

    # 4. Tinh chỉnh Biên bản kiểm tra (Doc 10) để trọn vẹn 5 trang, không rơi chữ ký sang trang 6
    in_bbkt = False
    for p in doc.paragraphs:
        if 'BIÊN BẢN KIỂM TRA' in p.text:
            in_bbkt = True
        elif in_bbkt and ('Công khai và bàn giao' in p.text):
            in_bbkt = False
        if in_bbkt:
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(0)

    # 5. Khóa bảng không bị ngắt đôi dòng (cantSplit) và lề bảng gọn gàng
    for t in doc.tables:
        for row in t.rows:
            trPr = row._element.trPr
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                row._element.append(trPr)
            cantSplit = OxmlElement('w:cantSplit')
            trPr.append(cantSplit)
            for c in row.cells:
                for p in c.paragraphs:
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.space_before = Pt(0)

def build_single_page_qd_kiem_tra(master_doc, output_path):
    """
    Tạo file Quyết định kiểm tra độc lập và căn chỉnh chuẩn xác để nằm trọn vẹn trong ĐÚNG 1 TRANG A4.
    """
    doc = docx.Document(TEMPLATE_PATH)
    body = doc._element.body
    
    sectPr = None
    for el in list(body):
        if el.tag.endswith('sectPr'):
            sectPr = copy.deepcopy(el)
            break
            
    master_body = master_doc._element.body
    all_master_elems = list(master_body)
    
    # Range của QĐKT: từ element 39 đến 65
    segment = [copy.deepcopy(all_master_elems[i]) for i in range(39, 66)]
    
    # Xóa ngắt trang ở đầu
    if segment:
        el0 = segment[0]
        if el0.tag.endswith('p'):
            for br in el0.xpath('.//w:br[@w:type="page"]'):
                br.getparent().remove(br)
            texts = el0.xpath('.//w:t')
            t = "".join([node.text for node in texts if node.text]).strip()
            if not t:
                segment.pop(0)
                
    body.clear()
    for el in segment:
        body.append(el)
    if sectPr is not None:
        body.append(sectPr)
        
    # Căn chỉnh lề và khoảng cách để nằm trọn trong 1 trang A4
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2.7)
    sec.right_margin = Cm(1.8)
    
    # Định dạng các bảng (Header & Nơi nhận / Ký tên)
    for t in doc.tables:
        remove_floating_table_properties(t)
        for r_idx, row in enumerate(t.rows):
            for c_idx, cell in enumerate(row.cells):
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.0
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    for r in p.runs:
                        r.font.name = "Times New Roman"
                        if r.font.size and r.font.size.pt > 12.5:
                            r.font.size = Pt(12)
                            
    # Định dạng các đoạn văn
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
            
        if "QUYẾT ĐỊNH" == text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.size = Pt(13)
                r.font.bold = True
        elif "Về việc kiểm tra thuế tại" in text or "CÔNG TY TNHH" in text:
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
        elif "QUYẾT ĐỊNH:" == text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.size = Pt(12.5)
                r.font.bold = True
        elif text.startswith("Điều 1.") or text.startswith("Điều 2.") or text.startswith("Điều 3."):
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
                    
    doc.save(output_path)
    print(f"-> Đã xuất: {os.path.basename(output_path):42s} (Chuẩn 1 trang A4)")

def build_single_page_qd_giam_sat(master_doc, output_path):
    """
    Tạo file Quyết định giám sát độc lập và căn chỉnh chuẩn xác để nằm trọn vẹn trong ĐÚNG 1 TRANG A4.
    """
    doc = docx.Document(TEMPLATE_PATH)
    body = doc._element.body
    
    sectPr = None
    for el in list(body):
        if el.tag.endswith('sectPr'):
            sectPr = copy.deepcopy(el)
            break
            
    master_body = master_doc._element.body
    all_master_elems = list(master_body)
    
    # Range của QĐGS: từ element 66 đến 86
    segment = [copy.deepcopy(all_master_elems[i]) for i in range(66, 87)]
    
    # Xóa ngắt trang ở đầu
    if segment:
        el0 = segment[0]
        if el0.tag.endswith('p'):
            for br in el0.xpath('.//w:br[@w:type="page"]'):
                br.getparent().remove(br)
            texts = el0.xpath('.//w:t')
            t = "".join([node.text for node in texts if node.text]).strip()
            if not t:
                segment.pop(0)
                
    body.clear()
    for el in segment:
        body.append(el)
    if sectPr is not None:
        body.append(sectPr)
        
    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(2.7)
    sec.right_margin = Cm(1.8)
    
    for t in doc.tables:
        remove_floating_table_properties(t)
        for r_idx, row in enumerate(t.rows):
            for c_idx, cell in enumerate(row.cells):
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
            
        if "QUYẾT ĐỊNH" == text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs:
                r.font.size = Pt(13)
                r.font.bold = True
        elif "Về việc giám sát" in text or "CÔNG TY TNHH" in text:
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
        elif "QUYẾT ĐỊNH:" == text:
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.size = Pt(12.5)
                r.font.bold = True
        elif text.startswith("Điều 1.") or text.startswith("Điều 2.") or text.startswith("Điều 3.") or text.startswith("Điều 4."):
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
                    
    doc.save(output_path)
    print(f"-> Đã xuất: {os.path.basename(output_path):42s} (Chuẩn 1 trang A4)")

def main():
    print("=== BẮT ĐẦU CHUẨN HÓA TOÀN DIỆN BỘ HỒ SƠ KIỂM TRA ===")
    os.makedirs(TARGET_DIR, exist_ok=True)
    os.makedirs(BRAIN_TEMPLATE_DIR, exist_ok=True)
    
    # 1. Sao lưu phôi gốc vào templates của Brain
    brain_master_template = os.path.join(BRAIN_TEMPLATE_DIR, "MAUTONG HOP.docx")
    shutil.copy2(TEMPLATE_PATH, brain_master_template)
    print(f"-> Đã đồng bộ phôi gốc vào Brain: templates/MAUTONG HOP.docx")
    
    # 2. Tạo master document đã điền đầy đủ dữ liệu và chuẩn hóa flow
    print("\n1. Điền dữ liệu & Chuẩn hóa toàn diện phôi MAUTONG HOP.docx...")
    master_doc = build_master_document()
    
    # Lưu TONG HOP.docx (Toàn bộ 20 phần/biểu mẫu của cuộc kiểm tra)
    tong_hop_path = os.path.join(TARGET_DIR, "TONG HOP.docx")
    master_doc.save(tong_hop_path)
    print(f"-> Đã xuất file tổng hợp toàn bộ: TONG HOP.docx ({len(master_doc.paragraphs)} đoạn văn, {len(master_doc.tables)} bảng)")
    
    # 3. Xuất các file Quyết định chuẩn xác đúng 1 trang A4
    print("\n2. Xuất các file Quyết định (chuẩn 1 trang A4)...")
    qd_kt_path = os.path.join(TARGET_DIR, "QD Kiem tra tai cty Van Phat Thanh.docx")
    qd_gs_path = os.path.join(TARGET_DIR, "QDgiam sat tai cty Van Phat Thanh.docx")
    
    build_single_page_qd_kiem_tra(master_doc, qd_kt_path)
    build_single_page_qd_giam_sat(master_doc, qd_gs_path)
    
    # 4. Dọn dẹp các file hồ sơ riêng lẻ và file rác
    print("\n3. Dọn dẹp các file hồ sơ riêng lẻ và file rác...")
    files_to_remove = [
        "DI-01_DI-01.docx",
        "DI-02_DI-02.docx",
        "DI-04_DI-04.docx",
        "DI-06_DI-06.docx",
        "DI-07_DI-07.docx",
        "TONG HOP HO SO DI.docx"
    ]
    for fname in files_to_remove:
        fpath = os.path.join(TARGET_DIR, fname)
        if os.path.exists(fpath):
            try:
                os.remove(fpath)
                print(f"  - Đã dọn dẹp file hồ sơ riêng lẻ: {fname}")
            except Exception as e:
                print(f"  ! Lỗi khi dọn dẹp {fname}: {e}")
                
    for f in os.listdir(TARGET_DIR):
        if f.endswith(".bak") or f.endswith(".tmp") or f.startswith("~$"):
            try:
                os.remove(os.path.join(TARGET_DIR, f))
                print(f"  - Đã xóa file rác: {f}")
            except Exception as e:
                pass
                
    # 5. Đóng gói ZIP chỉ gồm TONG HOP.docx và 2 file Quyết định
    print("\n4. Đóng gói hoso.zip...")
    zip_path = os.path.join(TARGET_DIR, "hoso.zip")
    files_to_zip = [
        "TONG HOP.docx",
        "QD Kiem tra tai cty Van Phat Thanh.docx",
        "QDgiam sat tai cty Van Phat Thanh.docx",
    ]
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in files_to_zip:
            fpath = os.path.join(TARGET_DIR, fname)
            if os.path.exists(fpath):
                zf.write(fpath, fname)
                print(f"  + Đã thêm vào ZIP: {fname}")
                
    print(f"\n-> Gói ZIP: {zip_path} ({os.path.getsize(zip_path):,} bytes)")
    print("=== HOÀN TẤT CHUẨN HÓA ===")

if __name__ == "__main__":
    main()
