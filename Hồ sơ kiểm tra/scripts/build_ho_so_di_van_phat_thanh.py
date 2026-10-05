# -*- coding: utf-8 -*-
"""
Builder hoàn chỉnh tạo toàn bộ bộ Hồ sơ đi kiểm tra theo kế hoạch năm:
- DI-01 đến DI-07 (7 file riêng lẻ)
- TONG HOP HO SO DI.docx (File Word tổng hợp toàn bộ 7 mẫu)
- templates/MAU TONG HOP HO SO DI.docx (Template phôi tổng hợp chuẩn)
- hoso.zip (Gói nén lưu trữ)
"""

import os
import zipfile
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

TARGET_DIR = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\DANG KIEM TRA\KE HOACH\VAN PHAT THANH\0. BIEN BAN"
BRAIN_TEMPLATES_DIR = r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\templates"

DATA = {
    "Tên_công_ty": "CÔNG TY TNHH MỘT THÀNH VIÊN BAO BÌ GIẤY VẠN PHÁT THÀNH",
    "MST": "3601409441",
    "Địa_chỉ": "Số 9/3, đường Nguyễn văn Tỏ, khu phố 2, Phường Long Hưng, Thành phố Đồng Nai",
    "Kho_hàng": "Cửa hàng Bao Bì Giấy Vạn Phát Thành tại số 507, Quốc Lộ 51, KP 3, Phường An Hòa (hết hiệu lực)",
    "Ngành_nghề": "Sản xuất giấy nhăn, bìa nhăn, bao bì từ giấy và bìa",
    "GĐ": "Tô Quốc Thi",
    "CCCD": "096075006910",
    "Ngày_cấp_CCCD": "28/09/2021",
    "Nơi_cấp_CCCD": "Cục Cảnh sát Quản lý hành chính về trật tự xã hội",
    "KTT": "Trần Thị Thoan",
    "Thời_kỳ_kiểm_tra": "Từ năm 2021 đến năm 2025",
    "Thời_hạn_kiểm_tra": "20",
    "Trưởng_đoàn": "Nguyễn Thị Ánh Ngọc",
    "CV_Trưởng_đoàn": "Chuyên viên",
    "Thành_viên": "Nguyễn Thành Đồng",
    "CV_Thành_viên": "Kiểm tra viên thuế",
    "Người_giám_sát": "Nguyễn Trường Giang",
    "CV_Người_giám_sát": "Tổ trưởng Tổ Kiểm tra số 1",
    "Lãnh_đạo_ký": "Phạm Đồng Xuân",
    "CV_Lãnh_đạo": "PHÓ TRƯỞNG THUẾ CƠ SỞ",
    "Số_QĐKT": "DU_THAO-TRINH-KY",
    "Số_QĐGS": "DU_THAO-TRINH-KY",
    "Nội_dung_kiểm_tra_list": [
        "Chi phí quản lý doanh nghiệp.",
        "Chi phí bán hàng.",
        "Chi phí tài chính."
    ],
    "Phân_công_trưởng_đoàn": "Chi phí quản lý doanh nghiệp; Chi phí tài chính",
    "Phân_công_thành_viên": "Chi phí bán hàng"
}

def set_a4_margins(section):
    """Cài đặt khổ giấy A4 và canh lề chuẩn Nghị định 30/2020/NĐ-CP:
       Lề trái: 3.0 cm, Lề phải: 2.0 cm, Lề trên: 2.0 cm, Lề dưới: 2.0 cm
    """
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.0)

def set_font_run(run, font_name="Times New Roman", size_pt=13, bold=False, italic=False, underline=False, color_rgb=(0,0,0)):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}" w:eastAsia="{font_name}"/>')
    rPr.append(rFonts)

def add_p(container, text="", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=3, line_spacing=1.15, bold=False, italic=False, size_pt=13, underline=False):
    p = container.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        run = p.add_run(text)
        set_font_run(run, font_name="Times New Roman", size_pt=size_pt, bold=bold, italic=italic, underline=underline)
    return p

def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def set_cell_margins(cell, top=40, bottom=40, left=40, right=40):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def create_header_table(doc, left_lines, right_lines, left_sub="", right_sub=""):
    table = doc.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(table)
    table.columns[0].width = Cm(7.0)
    table.columns[1].width = Cm(9.0)
    
    c0 = table.cell(0, 0)
    c1 = table.cell(0, 1)
    set_cell_margins(c0, 20, 20, 0, 80)
    set_cell_margins(c1, 20, 20, 80, 0)
    
    for idx, l in enumerate(left_lines):
        bold = True if (idx >= len(left_lines)-1 or "THUẾ CƠ SỞ" in l or "TỔ KIỂM TRA" in l or "ĐOÀN KIỂM TRA" in l) else False
        add_p(c0, l, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=1, line_spacing=1.1, bold=bold, size_pt=12)
    
    for idx, l in enumerate(right_lines):
        bold = True if idx == 0 else False
        add_p(c1, l, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=1, line_spacing=1.1, bold=bold, size_pt=12)
    
    c2 = table.cell(1, 0)
    c3 = table.cell(1, 1)
    set_cell_margins(c2, 20, 20, 0, 80)
    set_cell_margins(c3, 20, 20, 80, 0)
    
    if left_sub:
        add_p(c2, left_sub, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=2, bold=False, size_pt=12)
    if right_sub:
        add_p(c3, right_sub, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=2, italic=True, size_pt=12)
        
    return table

def create_signature_table(doc, left_recipient_lines, right_signer_lines, right_name=""):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(table)
    table.columns[0].width = Cm(7.5)
    table.columns[1].width = Cm(8.5)
    
    c0 = table.cell(0, 0)
    c1 = table.cell(0, 1)
    set_cell_margins(c0, 50, 50, 0, 50)
    set_cell_margins(c1, 50, 50, 50, 0)
    
    for idx, l in enumerate(left_recipient_lines):
        if idx == 0:
            add_p(c0, l, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=1, bold=True, italic=True, size_pt=11)
        else:
            add_p(c0, l, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=1, size_pt=11)
            
    for idx, l in enumerate(right_signer_lines):
        bold = True
        size = 12 if "KT." in l else 13
        add_p(c1, l, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=2, bold=bold, size_pt=size)
        
    if right_name:
        add_p(c1, "", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0)
        add_p(c1, "", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0)
        add_p(c1, "", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0)
        add_p(c1, right_name, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=0, bold=True, size_pt=13)
        
    return table

# ==========================================
# 1. DI-02: TỜ TRÌNH
# ==========================================
def render_di02(doc, data=DATA):
    create_header_table(
        doc,
        left_lines=["THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI", "TỔ KIỂM TRA 1"],
        right_lines=["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"],
        left_sub="",
        right_sub=""
    )
    
    add_p(doc, "TỜ TRÌNH LÃNH ĐẠO THUẾ CƠ SỞ", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=3, bold=True, size_pt=14)
    add_p(doc, "V/v: Ban hành quyết định kiểm tra thuế;", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=2, italic=True, size_pt=13)
    add_p(doc, "Quyết định giám sát hoạt động của Đoàn kiểm tra", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=10, italic=True, size_pt=13)
    
    add_p(doc, "- Nội dung công việc:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=3, bold=True, size_pt=13)
    add_p(doc, f"Kiểm tra việc chấp hành pháp luật về thuế và giám sát hoạt động của Đoàn kiểm tra tại {data['Tên_công_ty']} – MST: {data['MST']}; thời kỳ kiểm tra {data['Thời_kỳ_kiểm_tra']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, size_pt=13)
    
    add_p(doc, "- Căn cứ pháp lý để giải quyết:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=3, bold=True, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 1614/QĐ-TCT ngày 13/10/2022 của Tổng cục trưởng Tổng cục Thuế về việc ban hành Quy chế giám sát hoạt động của Đoàn thanh tra, kiểm tra thuế.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 1412/QĐ-DON ngày 06/05/2026 của Trưởng Thuế Thành phố Đồng Nai về việc phê duyệt kế hoạch kiểm tra tại trụ sở người nộp thuế năm 2026 của các Thuế cơ sở thuộc Thuế tỉnh Đồng Nai;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, size_pt=13)
    add_p(doc, "Căn cứ Luật Quản lý thuế số 108/2025/QH15 ngày 10 tháng 12 năm 2025 của Quốc hội;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, size_pt=13)
    
    add_p(doc, "- Ý kiến tham gia của các đơn vị liên quan: Không.", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=6, bold=False, size_pt=13)
    
    add_p(doc, "- Đề xuất của cán bộ giải quyết:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=3, bold=True, size_pt=13)
    add_p(doc, f"Kiến nghị BLĐ Thuế Cơ sở 3 Thành phố Đồng Nai ký duyệt quyết định kiểm tra, quyết định giám sát hoạt động đoàn kiểm tra tại {data['Tên_công_ty']} – MST: {data['MST']} (có dự thảo quyết định kèm theo).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=10, size_pt=13)
    
    # Ký duyệt 3 cấp
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(tbl)
    tbl.columns[0].width = Cm(8.0)
    tbl.columns[1].width = Cm(8.0)
    
    c0 = tbl.cell(0, 0)
    c1 = tbl.cell(0, 1)
    set_cell_margins(c0, 40, 40, 0, 40)
    set_cell_margins(c1, 40, 40, 40, 0)
    
    add_p(c1, "Trấn Biên, ngày ….. tháng ….. năm 2026", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size_pt=12)
    add_p(c1, "CÁN BỘ ĐỀ XUẤT", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)
    add_p(c1, "", space_after=0)
    add_p(c1, "", space_after=0)
    add_p(c1, "", space_after=0)
    add_p(c1, data["Trưởng_đoàn"], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)
    
    add_p(doc, "Ý kiến của lãnh đạo Đội thuế:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=2, bold=True, size_pt=13)
    add_p(doc, "Kính đề nghị Lãnh đạo Tổ kiểm tra ký duyệt theo đề xuất.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, size_pt=13)
    
    tbl2 = doc.add_table(rows=1, cols=2)
    tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(tbl2)
    tbl2.columns[0].width = Cm(8.0)
    tbl2.columns[1].width = Cm(8.0)
    c2_0 = tbl2.cell(0, 0)
    c2_1 = tbl2.cell(0, 1)
    set_cell_margins(c2_0, 40, 40, 0, 40)
    set_cell_margins(c2_1, 40, 40, 40, 0)
    
    add_p(c2_1, "Trấn Biên, ngày ….. tháng ….. năm 2026", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size_pt=12)
    add_p(c2_1, "TỔ TRƯỞNG", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)
    add_p(c2_1, "", space_after=0)
    add_p(c2_1, "", space_after=0)
    add_p(c2_1, "", space_after=0)
    add_p(c2_1, data["Người_giám_sát"], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)
    
    add_p(doc, "Ý kiến phê duyệt của Lãnh đạo Thuế Cơ sở 3: ..……………………………………………………………………………………", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=10, space_after=6, size_pt=13)
    
    tbl3 = doc.add_table(rows=1, cols=2)
    tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(tbl3)
    tbl3.columns[0].width = Cm(8.0)
    tbl3.columns[1].width = Cm(8.0)
    c3_0 = tbl3.cell(0, 0)
    c3_1 = tbl3.cell(0, 1)
    set_cell_margins(c3_0, 40, 40, 0, 40)
    set_cell_margins(c3_1, 40, 40, 40, 0)
    
    add_p(c3_1, "Trấn Biên, ngày ….. tháng ….. năm 2026", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size_pt=12)
    add_p(c3_1, "KT. TRƯỞNG THUẾ CƠ SỞ", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=12)
    add_p(c3_1, "PHÓ TRƯỞNG THUẾ CƠ SỞ", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)
    add_p(c3_1, "", space_after=0)
    add_p(c3_1, "", space_after=0)
    add_p(c3_1, "", space_after=0)
    add_p(c3_1, data["Lãnh_đạo_ký"], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)

# ==========================================
# 2. DI-03: QUYẾT ĐỊNH KIỂM TRA THUẾ
# ==========================================
def render_di03(doc, data=DATA):
    create_header_table(
        doc,
        left_lines=["THUẾ THÀNH PHỐ ĐỒNG NAI", "THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI"],
        right_lines=["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"],
        left_sub=f"Số: {data['Số_QĐKT']}",
        right_sub="Trấn Biên, ngày       tháng       năm 2026"
    )
    
    add_p(doc, "QUYẾT ĐỊNH", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=3, bold=True, size_pt=14)
    add_p(doc, "Về việc kiểm tra thuế tại", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=2, bold=True, size_pt=13)
    add_p(doc, data["Tên_công_ty"], align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4, bold=True, size_pt=13)
    add_p(doc, "TRƯỞNG THUẾ CƠ SỞ  3 THÀNH PHỐ ĐỒNG NAI", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=8, bold=True, size_pt=13)
    
    add_p(doc, "Căn cứ Luật Quản lý thuế số 108/2025/QH15 ngày 10 tháng 12 năm 2025 và các văn bản hướng dẫn thi hành;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ các Luật thuế và các văn bản hướng dẫn thi hành;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 1377/QĐ-CT ngày 30/06/2026 của Cục trưởng Cục Thuế quy định chức năng, nhiệm vụ, quyền hạn của Thuế cơ sổ thuộc Thuế tỉnh, thành phố;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 3737/QĐ-CT ngày 31/12/2025 của Cục trưởng Cục Thuế quy định mô hình tổ chức quản lý thuế của Thuế cơ sở;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 1412/QĐ-DON ngày 06/05/2026 của Trưởng Thuế Thành phố Đồng Nai về việc phê duyệt kế hoạch kiểm tra tại trụ sở người nộp thuế năm 2026 của các Thuế cơ sở thuộc Thuế tỉnh Đồng Nai;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 2618/QĐ-TCS3 ngày 20/05/2026 của Trưởng Thuế Cơ sở 3 Thành phố Đồng Nai về việc giao quyền ra Quyết định Kiểm tra thuế tại trụ sở của người nộp thuế;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Theo đề nghị của Tổ trưởng Tổ Kiểm tra số 1.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=8, italic=True, size_pt=13)
    
    add_p(doc, "QUYẾT ĐỊNH:", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=6, bold=True, size_pt=14)
    
    p_d1 = add_p(doc, f"Điều 1. Kiểm tra việc chấp hành pháp luật thuế tại {data['Tên_công_ty']}, mã số thuế: {data['MST']} về các nội dung sau:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, bold=False, size_pt=13)
    p_d1.runs[0].font.bold = False
    
    # 3 nội dung kiểm tra chuẩn xác
    for nd in data["Nội_dung_kiểm_tra_list"]:
        add_p(doc, f"- {nd}", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
        
    add_p(doc, f"Thời kỳ kiểm tra: {data['Thời_kỳ_kiểm_tra']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, f"Thời hạn kiểm tra là {data['Thời_hạn_kiểm_tra']} ngày làm việc, kể từ ngày công bố Quyết định kiểm tra.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=6, size_pt=13)
    
    add_p(doc, "Điều 2. Thành lập Đoàn kiểm tra để thực hiện nhiệm vụ ghi tại Điều 1 gồm có:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, f"- Bà: {data['Trưởng_đoàn']}\t- Chức vụ: {data['CV_Trưởng_đoàn']}\t- Trưởng đoàn;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"- Ông: {data['Thành_viên']}\t- Chức vụ: {data['CV_Thành_viên']}\t- Thành viên.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "Đoàn kiểm tra có nhiệm vụ, quyền hạn quy định tại Điều 22 của Luật Quản lý thuế số 108/2025/QH15.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=6, size_pt=13)
    
    add_p(doc, "Điều 3. Quyết định này có hiệu lực kể từ ngày ký. Công ty được kiểm tra có tên tại Điều 1 và Đoàn kiểm tra chịu trách nhiệm thi hành Quyết định này./.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=10, size_pt=13)
    
    create_signature_table(
        doc,
        left_recipient_lines=["Nơi nhận:", "- Như Điều 3;", "- Lưu: VT; Tổ KTR1 (04b)."],
        right_signer_lines=["KT. TRƯỞNG THUẾ CƠ SỞ", "PHÓ TRƯỞNG THUẾ CƠ SỞ"],
        right_name=data["Lãnh_đạo_ký"]
    )

# ==========================================
# 3. DI-05: QUYẾT ĐỊNH GIÁM SÁT
# ==========================================
def render_di05(doc, data=DATA):
    create_header_table(
        doc,
        left_lines=["THUẾ THÀNH PHỐ ĐỒNG NAI", "THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI"],
        right_lines=["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"],
        left_sub=f"Số: {data['Số_QĐGS']}",
        right_sub="Trấn Biên, ngày      tháng     năm 2026"
    )
    
    add_p(doc, "QUYẾT ĐỊNH", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=3, bold=True, size_pt=14)
    add_p(doc, "Về việc giám sát hoạt động của Đoàn kiểm tra tại", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=2, bold=True, size_pt=13)
    add_p(doc, data["Tên_công_ty"], align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4, bold=True, size_pt=13)
    add_p(doc, "TRƯỞNG THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=8, bold=True, size_pt=13)
    
    add_p(doc, "Căn cứ Luật Quản lý thuế số 108/2025/QH15 ngày 10 tháng 12 năm 2025 và các văn bản hướng dẫn thi hành;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ Luật Thanh tra ngày 14 tháng 11 năm 2022 và các văn bản hướng dẫn thi hành;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 1614/QĐ-TCT ngày 13/10/2022 của Tổng cục trưởng Tổng cục Thuế về việc ban hành Quy chế giám sát hoạt động của Đoàn thanh tra, kiểm tra thuế;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 3737/QĐ-CT ngày 31/12/2025 của Cục trưởng Cục Thuế quy định mô hình tổ chức quản lý thuế của Thuế cơ sở;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Căn cứ Quyết định số 1412/QĐ-DON ngày 06/05/2026 của Trưởng Thuế Thành phố Đồng Nai về việc phê duyệt kế hoạch kiểm tra tại trụ sở người nộp thuế năm 2026 của các Thuế cơ sở thuộc Thuế tỉnh Đồng Nai;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, f"Căn cứ Quyết định kiểm tra số {data['Số_QĐKT']} ngày ...../...../2026 của Thuế Cơ sở 3 Thành phố Đồng Nai về việc kiểm tra thuế tại {data['Tên_công_ty']};", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, italic=True, size_pt=13)
    add_p(doc, "Theo đề nghị của Tổ trưởng Tổ Kiểm tra số 1.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=8, italic=True, size_pt=13)
    
    add_p(doc, "QUYẾT ĐỊNH:", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4, space_after=6, bold=True, size_pt=14)
    
    add_p(doc, f"Điều 1. Thực hiện giám sát đối với hoạt động kiểm tra của Đoàn kiểm tra được thành lập theo Quyết định số {data['Số_QĐKT']} ngày ...../...../2026 của Thuế Cơ sở 3 Thành phố Đồng Nai.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, "Thời hạn giám sát được tính kể từ ngày công bố quyết định kiểm tra đến thời điểm kết thúc việc kiểm tra tại trụ sở người nộp thuế.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=4, size_pt=13)
    
    add_p(doc, f"Điều 2. Giao Ông {data['Người_giám_sát']} – Chức vụ: {data['CV_Người_giám_sát']} giám sát hoạt động của Đoàn kiểm tra để thực hiện nhiệm vụ tại Điều 1.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=4, size_pt=13)
    add_p(doc, f"Điều 3. Ông {data['Người_giám_sát']} có nhiệm vụ giám sát hoạt động của Đoàn kiểm tra theo quy định tại Quy chế giám sát hoạt động của Đoàn thanh tra, kiểm tra thuế ban hành kèm theo Quyết định số 1614/QĐ-TCT ngày 13/10/2022 của Tổng cục trưởng Tổng cục Thuế.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=4, size_pt=13)
    add_p(doc, f"Điều 4. Ông {data['Người_giám_sát']}, Đoàn Kiểm tra, {data['Tên_công_ty']} và các tổ chức, cá nhân có liên quan chịu trách nhiệm thi hành quyết định này./.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=10, size_pt=13)
    
    create_signature_table(
        doc,
        left_recipient_lines=["Nơi nhận:", "- Như Điều 4;", "- Lưu: VT; Tổ KTR1 (03b)."],
        right_signer_lines=["KT. TRƯỞNG THUẾ CƠ SỞ", "PHÓ TRƯỞNG THUẾ CƠ SỞ"],
        right_name=data["Lãnh_đạo_ký"]
    )

# ==========================================
# 4. DI-01: BẢN PHÂN TÍCH THÔNG TIN (Mẫu 08/QTKT-TC)
# ==========================================
def render_di01(doc, data=DATA):
    create_header_table(
        doc,
        left_lines=["THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI", "ĐOÀN KIỂM TRA"],
        right_lines=["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"],
        left_sub="",
        right_sub=""
    )
    
    add_p(doc, "PHÂN TÍCH THÔNG TIN, XÁC ĐỊNH NỘI DUNG CUỘC KIỂM TRA", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=2, bold=True, size_pt=14)
    add_p(doc, "(Mẫu số 08/QTKT-TC ban hành kèm theo Quyết định số 1161/QĐ-CT ngày 14/08/2026 của Cục Thuế)", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=10, italic=True, size_pt=12)
    
    # I. THÔNG TIN CHUNG
    add_p(doc, "I. THÔNG TIN CHUNG:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, "1. Đoàn kiểm tra (Nhóm phân tích):", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=2, bold=True, size_pt=13)
    add_p(doc, f"a) Trưởng đoàn: {data['Trưởng_đoàn']}. Chức vụ: {data['CV_Trưởng_đoàn']}. Bộ phận kiểm tra: Tổ Kiểm tra số 1, Thuế Cơ sở 3 Thành phố Đồng Nai.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"b) Thành viên: {data['Thành_viên']}. Chức vụ: {data['CV_Thành_viên']}. Bộ phận kiểm tra: Tổ Kiểm tra số 1, Thuế Cơ sở 3 Thành phố Đồng Nai.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, size_pt=13)
    
    add_p(doc, "2. Đối tượng được kiểm tra:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=2, bold=True, size_pt=13)
    add_p(doc, f"a) Tên người nộp thuế: {data['Tên_công_ty']};", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"b) Mã số thuế: {data['MST']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"c) Địa chỉ: {data['Địa_chỉ']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"d) Số chi nhánh, cơ sở phụ thuộc: {data['Kho_hàng']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "đ) Niên độ kế toán: Từ ngày 01 tháng 01 đến ngày 31 tháng 12.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"e) Ngành nghề kinh doanh: {data['Ngành_nghề']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"g) Đại diện pháp luật: Giám đốc {data['GĐ']} (CCCD số: {data['CCCD']} do {data['Nơi_cấp_CCCD']} cấp ngày {data['Ngày_cấp_CCCD']}); Kế toán trưởng: {data['KTT']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    
    # II. TÌNH HÌNH CHẤP HÀNH
    add_p(doc, "II. TÌNH HÌNH CHẤP HÀNH PHÁP LUẬT THUẾ:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, "1. Tình hình chấp hành và các vi phạm trong các năm gần nhất:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=2, bold=True, size_pt=13)
    add_p(doc, "a) Tình hình nộp các tờ khai thuế: Nộp đầy đủ theo quy định.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "b) Tình hình nộp báo cáo quyết toán thuế: Nộp đầy đủ theo quy định.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "c) Quyết định xử lý vi phạm hành chính: Không có.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "d) Vi phạm về chế độ hóa đơn: Không có.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "e) Nộp thuế quá hạn: Không có.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "f) Lỗ liên tục từ 2 năm trở lên: Không", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "g) Có hoàn thuế GTGT: Không", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"2. Tình hình kiểm tra, thanh tra của cơ quan thuế: Chưa kiểm tra tại trụ sở NNT giai đoạn {data['Thời_kỳ_kiểm_tra']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, "3. Thông tin thu thập được ngoài cơ quan thuế: Chưa phát hiện thông tin bất thường.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=4, size_pt=13)
    
    # III. PHÂN TÍCH ĐÁNH GIÁ THÔNG TIN
    add_p(doc, "III. PHÂN TÍCH, ĐÁNH GIÁ THÔNG TIN:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, "1. Đối với các trường hợp theo kế hoạch, chuyên đề (trừ hoàn thuế):", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=2, bold=True, size_pt=13)
    add_p(doc, "a) Phân tích các chỉ tiêu rủi ro trên báo cáo tài chính:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=2, italic=True, size_pt=13)
    add_p(doc, "+/ Báo cáo kết quả hoạt động kinh doanh:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=2, bold=True, size_pt=13)
    add_p(doc, "- Doanh nghiệp kê khai hoạt động kinh doanh có lãi liên tục các năm từ năm 2021 đến năm 2025: Lợi nhuận trước thuế năm 2021 là 51.517.353 đồng; năm 2022 là 100.785.336 đồng; năm 2023 là 130.367.408 đồng; năm 2024 là 194.365.311 đồng; năm 2025 là 252.235.769 đồng.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Doanh thu thuần tăng trưởng liên tục qua các năm: năm 2021 là 4.275.904.321 đồng; năm 2022 là 5.304.710.195 đồng; năm 2023 là 6.851.592.363 đồng; năm 2024 là 8.652.414.302 đồng; năm 2025 là 10.060.947.641 đồng.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Tỷ lệ giá vốn hàng bán trên doanh thu thuần năm 2021 là 95,21% (giá vốn: 4.071.273.089 đ); năm 2022 là 95,57% (giá vốn: 5.069.923.139 đ); năm 2023 là 96,43% (giá vốn: 6.606.725.467 đ); năm 2024 là 96,57% (giá vốn: 8.355.881.944 đ); năm 2025 là 97,84% (giá vốn: 9.843.387.408 đ).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Tỷ lệ Chi phí quản lý kinh doanh (quản lý doanh nghiệp) trên doanh thu thuần năm 2021 là 2,73% (116.737.526 đ); năm 2022 là 1,89% (100.035.053 đ); năm 2023 là 0,45% (31.097.237 đ); năm 2024 là 0,41% (35.440.399 đ); năm 2025 là 0,33% (32.946.558 đ).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Tỷ lệ Chi phí bán hàng trên doanh thu thuần năm 2021 là 0,10% (4.092.970 đ); các năm 2022, 2023, 2024 không phát sinh (0%); năm 2025 là 0,05% (4.800.000 đ).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Tỷ lệ Chi phí tài chính trên doanh thu thuần năm 2021 là 0,73% (31.181.930 đ); năm 2022 là 0,64% (33.975.531 đ); năm 2023 là 1,22% (83.787.601 đ); năm 2024 là 0,77% (67.053.036 đ); năm 2025 là 0,03% (2.748.835 đ).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Đánh giá: Tỷ lệ giá vốn hàng bán duy trì ở mức rất cao và có xu hướng tăng dần (từ 95,21% đến 97,84%); tỷ lệ chi phí quản lý doanh nghiệp giảm mạnh từ năm 2023 đến năm 2025 (chỉ còn dưới 0,5% doanh thu); chi phí tài chính và chi phí bán hàng biến động không tương ứng với tốc độ tăng trưởng doanh thu.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "+/ Bảng cân đối kế toán: Các chỉ tiêu tài sản, nguồn vốn biến động tương ứng với quy mô sản xuất kinh doanh.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, "b) Phân tích các chỉ tiêu rủi ro trên tờ khai thuế: Kê khai cơ bản khớp đúng theo số liệu trên Báo cáo tài chính và sổ sách kế toán.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, "2. Đối với các trường hợp hoàn thuế: thực hiện phân tích thông tin theo nội dung đề nghị hoàn thuế (kiểm tra trước hoàn thuế) hoặc quyết định hoàn thuế (kiểm tra sau hoàn thuế): Không phát sinh.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, "3. Các trường hợp khác: Không phát sinh.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=4, size_pt=13)
    
    # IV. ĐỀ XUẤT NỘI DUNG KIỂM TRA
    add_p(doc, "IV. ĐỀ XUẤT NỘI DUNG KIỂM TRA:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=3, bold=True, size_pt=13)
    add_p(doc, f"- Chi phí quản lý doanh nghiệp, thời kỳ kiểm tra: {data['Thời_kỳ_kiểm_tra']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"- Chi phí bán hàng, thời kỳ kiểm tra: {data['Thời_kỳ_kiểm_tra']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"- Chi phí tài chính, thời kỳ kiểm tra: {data['Thời_kỳ_kiểm_tra']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=10, size_pt=13)
    
    # Chữ ký
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(tbl)
    tbl.columns[0].width = Cm(8.0)
    tbl.columns[1].width = Cm(8.0)
    c0 = tbl.cell(0, 0)
    c1 = tbl.cell(0, 1)
    set_cell_margins(c0, 40, 40, 0, 40)
    set_cell_margins(c1, 40, 40, 40, 0)
    
    add_p(c1, "Trấn Biên, ngày ….. tháng …. năm 2026", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size_pt=12)
    add_p(c1, "NGƯỜI PHÂN TÍCH / TRƯỞNG ĐOÀN KIỂM TRA", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=12)
    add_p(c1, "", space_after=0)
    add_p(c1, "", space_after=0)
    add_p(c1, "", space_after=0)
    add_p(c1, data["Trưởng_đoàn"], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)

# ==========================================
# 5. DI-07: THƯ TRAO ĐỔI THÔNG TIN
# ==========================================
def render_di07(doc, data=DATA):
    create_header_table(
        doc,
        left_lines=["THUẾ THÀNH PHỐ ĐỒNG NAI", "THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI"],
        right_lines=["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"],
        left_sub="",
        right_sub="Trấn Biên, ngày      tháng     năm 2026"
    )
    
    add_p(doc, "THƯ TRAO ĐỔI THÔNG TIN", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=4, bold=True, size_pt=14)
    add_p(doc, f"Kính gửi: {data['Tên_công_ty']}", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=10, bold=True, size_pt=13)
    
    add_p(doc, "Thuế Cơ sở 3 Thành phố Đồng Nai xin gửi lời chào trân trọng và cảm ơn sự hợp tác của Quý Doanh nghiệp, Người nộp thuế trong việc thực hiện nghĩa vụ thuế đối với Nhà nước.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    add_p(doc, "Với mục tiêu xây dựng cơ quan thuế Văn hóa – Hiện đại, đội ngũ cán bộ công chức Thuế chuyên nghiệp – liêm chính, nâng cao chất lượng phục vụ và tạo mọi điều kiện thuận lợi cho Người nộp thuế;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    add_p(doc, "Thuế Cơ sở 3 Thành phố Đồng Nai cam kết luôn lắng nghe, nắm bắt các khó khăn, vướng mắc trong quá trình thực hiện chính sách thuế để kịp thời tháo gỡ, hỗ trợ Người nộp thuế.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    add_p(doc, "Chúng tôi mong muốn Người nộp thuế luôn thực hiện tốt chính sách pháp luật thuế, đặc biệt là phối hợp chặt chẽ, tạo điều kiện thuận lợi để Đoàn kiểm tra thuế hoàn thành nhiệm vụ theo quy định pháp luật.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    add_p(doc, "Thuế Cơ sở 3 Thành phố Đồng Nai đề nghị Người nộp thuế: Trong quá trình thực hiện các thủ tục hành chính thuế cũng như quá trình làm việc với Đoàn kiểm tra, nếu phát hiện công chức thuế có hành vi sách nhiễu, phiền hà, tiêu cực hoặc vi phạm Quy chế làm việc, đề nghị phản ánh ngay cho cơ quan thuế.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    add_p(doc, "Mọi thông tin phản ánh của Người nộp thuế xin gửi về hòm thư điện tử và số điện thoại của Lãnh đạo bộ phận thanh tra, kiểm tra thuế; Lãnh đạo bộ phận kiểm tra nội bộ và Tổ trưởng tổ giám sát hoặc người giám sát là người được người ra quyết định thanh tra, kiểm tra thuế giao thực hiện nhiệm vụ giám sát, Cụ thể:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    
    add_p(doc, "- Phó Trưởng Thuế cơ sở: Ông Phạm Đồng Xuân. Số điện thoại: 0913.327.749. Email: pdxuanbho.don@gdt.gov.vn", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, "- Tổ trưởng Tổ Kiểm tra thuế 1 (giám sát đoàn kiểm tra): Ông Nguyễn Trường Giang. Số điện thoại: 0918.423.955. Email: ntgiang.don@gdt.gov.vn", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=6, size_pt=13)
    
    add_p(doc, "Thuế Cơ sở 3 Thành phố Đồng Nai trân trọng cảm ơn sự hợp tác của Quý Doanh nghiệp, Người nộp thuế!", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=10, size_pt=13)
    
    create_signature_table(
        doc,
        left_recipient_lines=["Nơi nhận:", "Như trên;", "- Lưu: KTR1 (02b)."],
        right_signer_lines=["KT. TRƯỞNG THUẾ CƠ SỞ", "PHÓ TRƯỞNG THUẾ CƠ SỞ"],
        right_name=data["Lãnh_đạo_ký"]
    )

# ==========================================
# 6. DI-04: GIẤY XÁC NHẬN GIAO NHẬN QĐ
# ==========================================
def render_di04(doc, data=DATA):
    create_header_table(
        doc,
        left_lines=[],
        right_lines=["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"],
        left_sub="",
        right_sub="Trấn Biên, ngày         tháng     năm 2026"
    )
    
    add_p(doc, "GIẤY XÁC NHẬN", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=3, bold=True, size_pt=14)
    add_p(doc, f"Giao Quyết Định số {data['Số_QĐKT']} ngày …../…../2026", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=10, bold=True, size_pt=13)
    
    add_p(doc, "1. Chúng tôi đại diện cơ quan quản lý thuế, gồm:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, f"- Họ và tên: {data['Trưởng_đoàn']};", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"- Chức vụ: {data['CV_Trưởng_đoàn']};", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Đơn vị công tác: Tổ Kiểm tra số 1, Thuế Cơ sở 3 Thành phố Đồng Nai.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    
    add_p(doc, "2. Chúng tôi đại diện bên nhận văn bản/quyết định gồm:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, f"- Họ và tên: {data['GĐ']};", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"- Số CCCD: {data['CCCD']} cấp ngày: {data['Ngày_cấp_CCCD']} tại: {data['Nơi_cấp_CCCD']};", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Giấy ủy quyền số …………….. ngày …………….. (nếu có);", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, f"- Chức vụ: Giám đốc; Đại diện cho {data['Tên_công_ty']}.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    
    add_p(doc, f"Đã nhận Quyết Định số: {data['Số_QĐKT']} ngày …../…../2026 do người đại diện cơ quan thuế nói trên giao trực tiếp tại: văn phòng Tổ Kiểm tra số 1, Thuế Cơ sở 3 Thành phố Đồng Nai vào lúc: 9 giờ 00 ngày …… tháng …… năm 2026.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=4, space_after=4, size_pt=13)
    add_p(doc, "Ý kiến của người nhận quyết định hoặc lý do không nhận (nếu có): ………… …………………………………………………………………………………………………………………………………………………………………………", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=4, size_pt=13)
    add_p(doc, "Giấy xác nhận được lập thành 02 bản giống nhau, mỗi bên giao và nhận giữ 01 bản.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=10, size_pt=13)
    
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(tbl)
    tbl.columns[0].width = Cm(8.0)
    tbl.columns[1].width = Cm(8.0)
    c0 = tbl.cell(0, 0)
    c1 = tbl.cell(0, 1)
    set_cell_margins(c0, 40, 40, 0, 40)
    set_cell_margins(c1, 40, 40, 40, 0)
    
    add_p(c0, "Người được phân công giao văn bản", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=12)
    add_p(c0, "(Ký, ghi rõ họ tên)", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size_pt=11)
    add_p(c0, "", space_after=0)
    add_p(c0, "", space_after=0)
    add_p(c0, "", space_after=0)
    add_p(c0, data["Trưởng_đoàn"], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)
    
    add_p(c1, "Người nhận văn bản", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=12)
    add_p(c1, "(Ký, ghi rõ họ tên)", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size_pt=11)
    add_p(c1, "", space_after=0)
    add_p(c1, "", space_after=0)
    add_p(c1, "", space_after=0)
    add_p(c1, data["GĐ"], align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size_pt=13)

# ==========================================
# 7. DI-06: KẾ HOẠCH GIÁM SÁT
# ==========================================
def render_di06(doc, data=DATA):
    create_header_table(
        doc,
        left_lines=["THUẾ THÀNH PHỐ ĐỒNG NAI", "THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI"],
        right_lines=["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "Độc lập - Tự do - Hạnh phúc"],
        left_sub="",
        right_sub="Trấn Biên, ngày 28  tháng 07  năm 2026"
    )
    
    add_p(doc, "KẾ HOẠCH GIÁM SÁT HOẠT ĐỘNG", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=2, bold=True, size_pt=14)
    add_p(doc, "CỦA ĐOÀN KIỂM TRA THUẾ", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8, bold=True, size_pt=14)
    
    add_p(doc, f"Thực hiện Quyết định số {data['Số_QĐGS']} ngày ...../...../2026 của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát đoàn kiểm tra theo QĐ kiểm tra tại doanh nghiệp số {data['Số_QĐKT']} ngày ...../...../2026 do ông {data['Người_giám_sát']} làm giám sát đoàn kiểm tra, lập kế hoạch tiến hành giám sát như sau:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=6, size_pt=13)
    
    add_p(doc, "I. Mục đích, yêu cầu", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, "Căn cứ theo Quyết định số 1614/QĐ-TCT ngày 13/10/2022 của Tổng cục trưởng Tổng cục Thuế về việc ban hành Quy chế giám sát hoạt động của Đoàn thanh tra, kiểm tra thuế.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    
    add_p(doc, "II. Nội dung giám sát", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, "1. Giám sát việc chấp hành pháp luật của người được giám sát bao gồm các nội dung sau:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "a. Việc thực hiện các quy định của pháp luật về trình tự, thủ tục, thời hạn tiến hành một cuộc kiểm tra thuế theo quy trình kiểm tra thuế;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "b. Việc thực hiện các quyền trong hoạt động kiểm tra thuế đối với người được giám sát, bao gồm: căn cứ, thẩm quyền, trình tự, thủ tục thực hiện các quyền;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "c. Việc thực hiện các quy định của pháp luật đối với người được giám sát, bao gồm: quy định về những hành vi bị nghiêm cấm, nghĩa vụ của Trưởng đoàn kiểm tra, thành viên Đoàn kiểm tra thuế; các quy định về việc tiếp nhận, chuyển giao hồ sơ vụ việc có dấu hiệu tội phạm để truy cứu trách nhiệm hình sự.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, size_pt=13)
    
    add_p(doc, "2. Giám sát việc thực hiện nhiệm vụ và triển khai hoạt động kiểm tra thuế của Đoàn kiểm tra bao gồm các nội dung sau:", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=2, size_pt=13)
    add_p(doc, "a. Tiến độ và kết quả đã đạt được so với yêu cầu theo quyết định kiểm tra thuế, kế hoạch tiến hành kiểm tra;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "b. Khó khăn, vướng mắc phát sinh trong hoạt động kiểm tra thuế và tác động đối với việc hoàn thành kế hoạch tiến hành kiểm tra;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "c. Việc thực hiện nhiệm vụ, triển khai hoạt động kiểm tra thuế của người được giám sát;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "d. Việc xử lý ý kiến khác nhau giữa Trưởng đoàn và thành viên Đoàn kiểm tra về những vấn đề liên quan đến nội dung kiểm tra thuế;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "e. Các nội dung tại Khoản 6, điều 7 Quy chế giám sát ban hành kèm theo QĐ 1614;", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "f. Các nội dung khác có liên quan đến việc thực hiện nhiệm vụ, triển khai hoạt động kiểm tra của Đoàn kiểm tra thuế.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=3, size_pt=13)
    add_p(doc, "3. Các nội dung giám sát hoạt động của Đoàn kiểm tra thuế được thực hiện theo mẫu 05-TTr ban hành kèm theo Quy chế giám sát ban hành kèm theo QĐ 1614.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    
    add_p(doc, "III. Hình thức giám sát", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, "- Theo điều 14 của QĐ 1614 (tùy trường hợp).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=4, size_pt=13)
    
    add_p(doc, "IV. Tổ chức thực hiện", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=4, space_after=2, bold=True, size_pt=13)
    add_p(doc, "- Tiến độ thực hiện:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=2, bold=False, size_pt=13)
    add_p(doc, "+ Chậm nhất 05 ngày làm việc sau khi quyết định giám sát hoạt động của Đoàn kiểm tra thuế được ký, người được phân công giám sát thực hiện nhiệm vụ giám sát của mình theo quy chế giám sát của QĐ 1614.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "+ Yêu cầu Đoàn kiểm tra thực hiện theo khoản 1 Điều 7 của QĐ 1614.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Chế độ thông tin báo cáo: Người được giám sát thực hiện tại khoản 2 Điều 16 của QĐ 1614 (nếu có).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Điều kiện vật chất đảm bảo hoạt động giám sát.", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=2, size_pt=13)
    add_p(doc, "- Những vấn đề khác (nếu có).", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=10, size_pt=13)
    
    create_signature_table(
        doc,
        left_recipient_lines=["Phê duyệt của Lãnh đạo Thuế cơ sở", "", "", "", "", data["Lãnh_đạo_ký"]],
        right_signer_lines=["Tổ trưởng Tổ giám sát hoặc công chức", "được giao nhiệm vụ giám sát", "", "", "", data["Người_giám_sát"]],
        right_name=""
    )

def build_all():
    os.makedirs(TARGET_DIR, exist_ok=True)
    os.makedirs(BRAIN_TEMPLATES_DIR, exist_ok=True)
    
    # 1. Tạo 7 file riêng lẻ
    file_map = [
        ("DI-02_DI-02.docx", render_di02),
        ("DI-03_DI-03.docx", render_di03),
        ("DI-05_DI-05.docx", render_di05),
        ("DI-01_DI-01.docx", render_di01),
        ("DI-07_DI-07.docx", render_di07),
        ("DI-04_DI-04.docx", render_di04),
        ("DI-06_DI-06.docx", render_di06),
    ]
    
    for fname, render_func in file_map:
        doc = docx.Document()
        set_a4_margins(doc.sections[0])
        render_func(doc, DATA)
        fpath = os.path.join(TARGET_DIR, fname)
        doc.save(fpath)
        print(f"Generated single file: {fname}")
        
    # 2. Tạo file TỔNG HỢP HỒ SƠ ĐI (TONG HOP HO SO DI.docx)
    doc_th = docx.Document()
    set_a4_margins(doc_th.sections[0])
    
    # Thứ tự trình ký: DI-02, DI-03, DI-05, DI-01, DI-07, DI-04, DI-06
    for idx, (fname, render_func) in enumerate(file_map):
        if idx > 0:
            doc_th.add_page_break()
        render_func(doc_th, DATA)
        
    th_path = os.path.join(TARGET_DIR, "TONG HOP HO SO DI.docx")
    doc_th.save(th_path)
    print(f"Generated combined file: {th_path}")
    
    # 3. Tạo file template phôi chuẩn trong brain templates
    template_path = os.path.join(BRAIN_TEMPLATES_DIR, "MAU TONG HOP HO SO DI.docx")
    doc_th.save(template_path)
    print(f"Saved template: {template_path}")
    
    # 4. Đóng gói hoso.zip
    zip_path = os.path.join(TARGET_DIR, "hoso.zip")
    files_to_zip = [fname for fname, _ in file_map] + ["TONG HOP HO SO DI.docx"]
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fn in files_to_zip:
            fp = os.path.join(TARGET_DIR, fn)
            if os.path.exists(fp):
                zf.write(fp, arcname=fn)
    print(f"Created zip archive: {zip_path} with {len(files_to_zip)} files")

if __name__ == "__main__":
    build_all()
