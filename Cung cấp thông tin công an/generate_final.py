import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl
import os
import shutil

# --- Helper functions for formatting ---

def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m_name, m_val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m_name)
        node.set(qn('w:w'), str(m_val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_width(cell, width_dxa):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(width_dxa))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def set_table_indent(table, indent_dxa):
    tblPr = table._tbl.tblPr
    tblInd = OxmlElement('w:tblInd')
    tblInd.set(qn('w:w'), str(indent_dxa))
    tblInd.set(qn('w:type'), 'dxa')
    tblPr.append(tblInd)

# --- Load Data ---

with open(r'sources\4020.md', 'r', encoding='utf-8') as f:
    content = f.read()

rows_4020 = []
for line in content.splitlines():
    if '|' in line and not ':-:' in line and not 'Tên doanh nghiệp' in line and not '**TỔNG**' in line:
        parts = [p.strip() for p in line.split('|')[1:-1]]
        if len(parts) >= 8:
            rows_4020.append({
                'stt': parts[0],
                'name': parts[1],
                'mst': parts[2],
                'so_hd': parts[3],
                'ds_chua_thue': parts[4],
                'thue_vat': parts[5],
                'ds_co_thue': parts[6],
                'dia_chi': parts[7]
            })

wb_nnt = openpyxl.load_workbook(r'sources\DB NNT 03.06.2026.xlsx', data_only=True)
ws_nnt = wb_nnt['NNT']
headers = [cell.value for cell in ws_nnt[3]]

xlsx_nnt = {}
for r in range(4, ws_nnt.max_row + 1):
    mst = str(ws_nnt.cell(r, 4).value or '').strip()
    if mst.endswith('.0'):
        mst = mst[:-2]
    mst_clean = mst.replace('-', '').strip()
    if mst_clean:
        row_dict = {headers[c-1]: ws_nnt.cell(r, c).value for c in range(1, len(headers)+1)}
        xlsx_nnt[mst_clean] = row_dict
        if mst_clean.startswith('0') and len(mst_clean) == 10:
            xlsx_nnt[mst_clean[1:]] = row_dict
        elif len(mst_clean) == 9:
            xlsx_nnt['0' + mst_clean] = row_dict

print(f"Loaded DB NNT rows: {len(xlsx_nnt)}")

matched_tcs3 = []
not_matched_tcs3 = []

for item in rows_4020:
    mst = item['mst'].strip().replace('-', '')
    nnt_info = xlsx_nnt.get(mst)
    if not nnt_info:
        if mst.startswith('0') and len(mst) == 10:
            nnt_info = xlsx_nnt.get(mst[1:])
        elif len(mst) == 9:
            nnt_info = xlsx_nnt.get('0' + mst)
    
    is_bh = 'Biên Hoà' in item['dia_chi'] or 'Biên Hòa' in item['dia_chi'] or (nnt_info is not None and ('Biên Hòa' in str(nnt_info.get('Tên Quận Huyện') or '') or 'Biên Hoà' in str(nnt_info.get('Tên Quận Huyện') or '')))
    
    entry = {
        'stt_4020': item['stt'],
        'name_4020': item['name'],
        'mst': item['mst'],
        'so_hd': item['so_hd'],
        'ds_chua_thue': item['ds_chua_thue'],
        'thue_vat': item['thue_vat'],
        'ds_co_thue': item['ds_co_thue'],
        'dia_chi_4020': item['dia_chi'],
        'nnt_info': nnt_info
    }
    if is_bh or nnt_info is not None:
        matched_tcs3.append(entry)
    else:
        not_matched_tcs3.append(entry)

print(f"Matched TCS3 = {len(matched_tcs3)}, Not matched TCS3 = {len(not_matched_tcs3)}")

# --- 1. BUILD DOCUMENT: TRA LOI CV 5734 ---

doc = docx.Document()

# Page Setup: A4, Margins: Left 2.5cm, Right 2.0cm, Top 2.0cm, Bottom 2.0cm
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.0)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)

# Style Normal
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(14)
style.font.color.rgb = RGBColor(0, 0, 0)

# BẢNG HEADER 3 HÀNG (tblInd = -850 dxa / -1.5cm, Cột 1 = 9.1cm / 5159 dxa, Cột 2 = 10.2cm / 5783 dxa)
table_hdr = doc.add_table(rows=3, cols=2)
table_hdr.alignment = WD_TABLE_ALIGNMENT.LEFT
remove_table_borders(table_hdr)
set_table_indent(table_hdr, -850)

for row in table_hdr.rows:
    set_cell_width(row.cells[0], 5159)
    set_cell_width(row.cells[1], 5783)
    set_cell_margins(row.cells[0], top=0, bottom=0, left=0, right=0)
    set_cell_margins(row.cells[1], top=0, bottom=0, left=0, right=0)

# Hàng 1 - Cột trái
cell_00 = table_hdr.cell(0, 0)
p = cell_00.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("THUẾ THÀNH PHỐ ĐỒNG NAI")
r.font.name = "Times New Roman"
r.font.size = Pt(13)

p = cell_00.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI")
r.font.name = "Times New Roman"
r.font.size = Pt(13)
r.bold = True

p = cell_00.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("────────────") # 12 ký tự kẻ
r.font.name = "Times New Roman"
r.font.size = Pt(10)

# Hàng 1 - Cột phải
cell_01 = table_hdr.cell(0, 1)
p = cell_01.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
r.font.name = "Times New Roman"
r.font.size = Pt(13)
r.bold = True

p = cell_01.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("Độc lập - Tự do - Hạnh phúc")
r.font.name = "Times New Roman"
r.font.size = Pt(14)
r.bold = True

p = cell_01.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("────────────────────────") # 24 ký tự kẻ
r.font.name = "Times New Roman"
r.font.size = Pt(10)

# Hàng 2 - Số & Ngày tháng
cell_10 = table_hdr.cell(1, 0)
p = cell_10.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.left_indent = Cm(2.0)
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Số:          ")
r.font.name = "Times New Roman"
r.font.size = Pt(13)

cell_11 = table_hdr.cell(1, 1)
p = cell_11.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("Trấn Biên, ngày       tháng       năm 2026")
r.font.name = "Times New Roman"
r.font.size = Pt(13)
r.italic = True

# Hàng 3 - Trích yếu
cell_20 = table_hdr.cell(2, 0)
p = cell_20.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("V/v cung cấp thông tin, tài liệu theo Công văn số 5734/CSKT-Đ2")
r.font.name = "Times New Roman"
r.font.size = Pt(12.5)

cell_21 = table_hdr.cell(2, 1)
cell_21.paragraphs[0].text = ""

# DÒNG KÍNH GỬI
p_kg = doc.add_paragraph()
p_kg.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_kg.paragraph_format.space_before = Pt(18)
p_kg.paragraph_format.space_after = Pt(14)
p_kg.paragraph_format.line_spacing = 1.15
r_kg = p_kg.add_run("Kính gửi: Thuế thành phố Đồng Nai")
r_kg.font.name = "Times New Roman"
r_kg.font.size = Pt(14)

# ĐOẠN CĂN CỨ MỞ ĐẦU (2 đoạn)
p_cc1 = doc.add_paragraph()
p_cc1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_cc1.paragraph_format.first_line_indent = Cm(1.25)
p_cc1.paragraph_format.space_before = Pt(3)
p_cc1.paragraph_format.space_after = Pt(0)
p_cc1.paragraph_format.line_spacing = 1.15
r_cc1 = p_cc1.add_run("Thuế cơ sở 3 thành phố Đồng Nai nhận được Công văn số 16451/DON-KTr1 ngày 25/08/2026 của Thuế thành phố Đồng Nai về việc chuyển thông tin liên quan đến vi phạm pháp luật về thuế, hóa đơn theo yêu cầu của Cơ quan Cảnh sát điều tra - Công an thành phố Đồng Nai tại Công văn số 5734/CSKT-Đ2 ngày 19/07/2026.")
r_cc1.font.name = "Times New Roman"
r_cc1.font.size = Pt(14)

p_cc2 = doc.add_paragraph()
p_cc2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_cc2.paragraph_format.first_line_indent = Cm(1.25)
p_cc2.paragraph_format.space_before = Pt(3)
p_cc2.paragraph_format.space_after = Pt(0)
p_cc2.paragraph_format.line_spacing = 1.15
r_cc2 = p_cc2.add_run("Căn cứ dữ liệu quản lý thuế, Thuế cơ sở 3 thành phố Đồng Nai báo cáo kết quả rà soát, xử lý đối với các doanh nghiệp thuộc địa bàn quản lý như sau:")
r_cc2.font.name = "Times New Roman"
r_cc2.font.size = Pt(14)

# THÂN BÀI - CHI TIẾT 26 DOANH NGHIỆP THUỘC TCS3
p_sec1 = doc.add_paragraph()
p_sec1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_sec1.paragraph_format.first_line_indent = Cm(1.25)
p_sec1.paragraph_format.space_before = Pt(6)
p_sec1.paragraph_format.space_after = Pt(0)
p_sec1.paragraph_format.line_spacing = 1.15
r_sec1 = p_sec1.add_run(f"I. Đối với các doanh nghiệp thuộc địa bàn Thuế cơ sở 3 quản lý ({len(matched_tcs3)} doanh nghiệp):")
r_sec1.bold = True

for idx, item in enumerate(matched_tcs3, 1):
    nnt = item['nnt_info'] or {}
    
    # Tiêu đề DN
    p_dn = doc.add_paragraph()
    p_dn.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_dn.paragraph_format.first_line_indent = Cm(1.25)
    p_dn.paragraph_format.space_before = Pt(6)
    p_dn.paragraph_format.space_after = Pt(0)
    p_dn.paragraph_format.line_spacing = 1.15
    r_dn = p_dn.add_run(f"{idx}. {item['name_4020'].upper()}, MST: {item['mst']}.")
    r_dn.bold = True
    
    # Địa điểm ĐKKD
    p_dc = doc.add_paragraph()
    p_dc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_dc.paragraph_format.first_line_indent = Cm(1.25)
    p_dc.paragraph_format.space_before = Pt(3)
    p_dc.paragraph_format.space_after = Pt(0)
    p_dc.paragraph_format.line_spacing = 1.15
    
    dc_parts = []
    if nnt.get('Số nhà/Đường phố'):
        dc_parts.append(str(nnt.get('Số nhà/Đường phố')).strip())
    if nnt.get('Tên Phường Xã'):
        dc_parts.append(str(nnt.get('Tên Phường Xã')).strip())
    if nnt.get('Tên Quận Huyện'):
        dc_parts.append(str(nnt.get('Tên Quận Huyện')).strip())
    dc_str = ', '.join(dc_parts) if dc_parts else item['dia_chi_4020']
    
    p_dc.add_run(f"+ Địa điểm đăng ký kinh doanh: {dc_str};")
    
    # Tình trạng hoạt động
    p_tt = doc.add_paragraph()
    p_tt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_tt.paragraph_format.first_line_indent = Cm(1.25)
    p_tt.paragraph_format.space_before = Pt(3)
    p_tt.paragraph_format.space_after = Pt(0)
    p_tt.paragraph_format.line_spacing = 1.15
    st_code = str(nnt.get('Trạng thái ĐKT tổ chức') or nnt.get('Trạng thái NNT') or '').strip()
    if st_code == '00':
        tt_text = "+ Tình trạng hoạt động: Người nộp thuế đang hoạt động (trạng thái 00);"
    else:
        ld = nnt.get('Lý do đóng MST tổ chức')
        if ld:
            tt_text = f"+ Tình trạng hoạt động: Người nộp thuế ngừng hoạt động ({ld});"
        elif st_code:
            tt_text = f"+ Tình trạng hoạt động: Người nộp thuế ngừng hoạt động (trạng thái {st_code});"
        else:
            tt_text = "+ Tình trạng hoạt động: Người nộp thuế ngừng hoạt động;"
    p_tt.add_run(tt_text)
    
    # Thông tin hóa đơn theo CV 5734
    p_hd = doc.add_paragraph()
    p_hd.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_hd.paragraph_format.first_line_indent = Cm(1.25)
    p_hd.paragraph_format.space_before = Pt(3)
    p_hd.paragraph_format.space_after = Pt(0)
    p_hd.paragraph_format.line_spacing = 1.15
    p_hd.add_run(f"+ Số liệu hóa đơn liên quan theo Công văn số 5734/CSKT-Đ2: {item['so_hd']} hóa đơn; Doanh số mua chưa có thuế: {item['ds_chua_thue']} đồng; Thuế GTGT: {item['thue_vat']} đồng; Tổng doanh số: {item['ds_co_thue']} đồng;")
    
    # Tình hình rà soát xử lý vi phạm thuế, hóa đơn
    p_xl = doc.add_paragraph()
    p_xl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_xl.paragraph_format.first_line_indent = Cm(1.25)
    p_xl.paragraph_format.space_before = Pt(3)
    p_xl.paragraph_format.space_after = Pt(0)
    p_xl.paragraph_format.line_spacing = 1.15
    
    if st_code == '00':
        p_xl.add_run("+ Tình hình rà soát, xử lý vi phạm về thuế, hóa đơn: Tổ Kiểm tra số 1 đã ban hành Phiếu chuyển chuyển Tổ Quản lý doanh nghiệp rà soát, đối chiếu hồ sơ khai thuế, yêu cầu doanh nghiệp giải trình, loại trừ chi phí và khấu trừ thuế GTGT đầu vào đối với các hóa đơn nêu trên; lập biên bản và ban hành quyết định xử phạt vi phạm hành chính theo quy định. ")
        r_red = p_xl.add_run("[CHƯA XÁC ĐỊNH]")
        r_red.font.color.rgb = RGBColor(255, 0, 0)
        r_red.bold = True
        r_red_tail = p_xl.add_run(" (kết quả xử phạt chi tiết đang trong thời hạn giải trình và thực hiện của Tổ Quản lý doanh nghiệp).")
        r_red_tail.font.color.rgb = RGBColor(255, 0, 0)
    else:
        p_xl.add_run("+ Tình hình rà soát, xử lý vi phạm về thuế, hóa đơn: Người nộp thuế đã ngừng hoạt động tại địa chỉ đăng ký kinh doanh; Cơ quan Thuế đã ban hành thông báo và khóa mã số thuế trên hệ thống quản lý. Cơ quan Thuế tiếp tục theo dõi, phối hợp xử lý theo quy định khi người nộp thuế thực hiện thủ tục khôi phục mã số thuế hoặc phát sinh thông tin mới.")

# PHẦN II: Các DN ngoài địa bàn
p_sec2 = doc.add_paragraph()
p_sec2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_sec2.paragraph_format.first_line_indent = Cm(1.25)
p_sec2.paragraph_format.space_before = Pt(6)
p_sec2.paragraph_format.space_after = Pt(0)
p_sec2.paragraph_format.line_spacing = 1.15
r_sec2 = p_sec2.add_run(f"II. Đối với các doanh nghiệp ngoài địa bàn quản lý của Thuế cơ sở 3 ({len(not_matched_tcs3)} doanh nghiệp):")
r_sec2.bold = True

p_sec2_body = doc.add_paragraph()
p_sec2_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_sec2_body.paragraph_format.first_line_indent = Cm(1.25)
p_sec2_body.paragraph_format.space_before = Pt(3)
p_sec2_body.paragraph_format.space_after = Pt(0)
p_sec2_body.paragraph_format.line_spacing = 1.15
p_sec2_body.add_run(f"Qua rà soát danh sách kèm theo Công văn số 5734/CSKT-Đ2, có {len(not_matched_tcs3)} doanh nghiệp có địa chỉ trụ sở tại các huyện Long Thành, Nhơn Trạch, Trảng Bom, Định Quán và thành phố Long Khánh (không thuộc địa bàn quản lý của Thuế cơ sở 3 thành phố Đồng Nai). Do đó, Thuế cơ sở 3 không có dữ liệu quản lý thuế để thực hiện rà soát, xử lý đối với {len(not_matched_tcs3)} đơn vị này.")

p_ket = doc.add_paragraph()
p_ket.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_ket.paragraph_format.first_line_indent = Cm(1.25)
p_ket.paragraph_format.space_before = Pt(3)
p_ket.paragraph_format.space_after = Pt(12)
p_ket.paragraph_format.line_spacing = 1.15
p_ket.add_run("Thuế cơ sở 3 thành phố Đồng Nai báo cáo Thuế thành phố Đồng Nai để tổng hợp theo quy định./.")

# BẢNG CHỮ KÝ CUỐI TRANG
table_sign = doc.add_table(rows=1, cols=2)
table_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
remove_table_borders(table_sign)
set_cell_width(table_sign.cell(0, 0), 4252) # 7.5cm
set_cell_width(table_sign.cell(0, 1), 5102) # 9.0cm
set_cell_margins(table_sign.cell(0, 0), top=0, bottom=0, left=0, right=0)
set_cell_margins(table_sign.cell(0, 1), top=0, bottom=0, left=0, right=0)

# Cột Nơi nhận
cell_s0 = table_sign.cell(0, 0)
p_nn_title = cell_s0.paragraphs[0]
p_nn_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_nn_title.paragraph_format.space_before = Pt(0)
p_nn_title.paragraph_format.space_after = Pt(2)
p_nn_title.paragraph_format.line_spacing = 1.0
r_nn = p_nn_title.add_run("Nơi nhận:")
r_nn.font.name = "Times New Roman"
r_nn.font.size = Pt(12)
r_nn.bold = True
r_nn.italic = True

for nn_item in ["- Như trên;", "- BLĐ thuế cơ sở 3 (để b/c);", "- Lưu: VT, KT1 (Ngọc)."]:
    p_nn = cell_s0.add_paragraph()
    p_nn.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_nn.paragraph_format.space_before = Pt(0)
    p_nn.paragraph_format.space_after = Pt(1)
    p_nn.paragraph_format.line_spacing = 1.0
    r = p_nn.add_run(nn_item)
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)

# Cột Chữ ký
cell_s1 = table_sign.cell(0, 1)
p_sig_title1 = cell_s1.paragraphs[0]
p_sig_title1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sig_title1.paragraph_format.space_before = Pt(0)
p_sig_title1.paragraph_format.space_after = Pt(0)
p_sig_title1.paragraph_format.line_spacing = 1.0
r = p_sig_title1.add_run("KT. TRƯỞNG THUẾ CƠ SỞ")
r.font.name = "Times New Roman"
r.font.size = Pt(13)
r.bold = True

p_sig_title2 = cell_s1.add_paragraph()
p_sig_title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sig_title2.paragraph_format.space_before = Pt(0)
p_sig_title2.paragraph_format.space_after = Pt(0)
p_sig_title2.paragraph_format.line_spacing = 1.0
r = p_sig_title2.add_run("PHÓ TRƯỞNG THUẾ CƠ SỞ")
r.font.name = "Times New Roman"
r.font.size = Pt(13)
r.bold = True

# 5 dòng trống
for _ in range(5):
    p_sp = cell_s1.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(0)
    p_sp.paragraph_format.line_spacing = 1.0

p_sig_name = cell_s1.add_paragraph()
p_sig_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sig_name.paragraph_format.space_before = Pt(0)
p_sig_name.paragraph_format.space_after = Pt(0)
p_sig_name.paragraph_format.line_spacing = 1.0
r = p_sig_name.add_run("Phạm Đồng Xuân")
r.font.name = "Times New Roman"
r.font.size = Pt(14)
r.bold = True

# Save TRA LOI docx (both Drive and Vault)
drive_dir = r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\240. CV 5734-CA DONG NAI"
vault_dir = r"D:\SECOND BRAIN\javis-os\brains\Cung cấp thông tin công an"

cv_name1 = "TRA LOI CV 5734_THUE TP DONG NAI_du thao.docx"
cv_name2 = "TRA LOI CV 5734_CA DONG NAI_du thao.docx"

cv_drive_path1 = os.path.join(drive_dir, cv_name1)
cv_drive_path2 = os.path.join(drive_dir, cv_name2)
cv_vault_path1 = os.path.join(vault_dir, cv_name1)

doc.save(cv_drive_path1)
doc.save(cv_drive_path2)
doc.save(cv_vault_path1)
print(f"CV Saved Drive 1: {cv_drive_path1} (exists: {os.path.exists(cv_drive_path1)})")
print(f"CV Saved Drive 2: {cv_drive_path2} (exists: {os.path.exists(cv_drive_path2)})")
print(f"CV Saved Vault: {cv_vault_path1} (exists: {os.path.exists(cv_vault_path1)})")


# --- 2. BUILD DOCUMENT: to trinh 5734.docx ---

doc_tt = docx.Document()
sec_tt = doc_tt.sections[0]
sec_tt.page_width = Cm(21.0)
sec_tt.page_height = Cm(29.7)
sec_tt.left_margin = Cm(3.0)
sec_tt.right_margin = Cm(2.0)
sec_tt.top_margin = Cm(2.5)
sec_tt.bottom_margin = Cm(1.0)

style_tt = doc_tt.styles['Normal']
style_tt.font.name = 'Times New Roman'
style_tt.font.size = Pt(13)
style_tt.font.color.rgb = RGBColor(0, 0, 0)

# Header Tờ trình (8.5cm & 10.1cm)
table_tt_hdr = doc_tt.add_table(rows=1, cols=2)
table_tt_hdr.alignment = WD_TABLE_ALIGNMENT.LEFT
remove_table_borders(table_tt_hdr)
set_table_indent(table_tt_hdr, -850)
set_cell_width(table_tt_hdr.cell(0, 0), 4819) # 8.5cm
set_cell_width(table_tt_hdr.cell(0, 1), 5726) # 10.1cm
set_cell_margins(table_tt_hdr.cell(0, 0), top=0, bottom=0, left=0, right=0)
set_cell_margins(table_tt_hdr.cell(0, 1), top=0, bottom=0, left=0, right=0)

# Cột trái Header TT
cell_tt00 = table_tt_hdr.cell(0, 0)
p = cell_tt00.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI")
r.font.name = "Times New Roman"
r.font.size = Pt(12)

p = cell_tt00.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("TỔ KIỂM TRA SỐ 1")
r.font.name = "Times New Roman"
r.font.size = Pt(13)
r.bold = True

p = cell_tt00.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("───────")
r.font.name = "Times New Roman"
r.font.size = Pt(10)

# Cột phải Header TT
cell_tt01 = table_tt_hdr.cell(0, 1)
p = cell_tt01.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
r.font.name = "Times New Roman"
r.font.size = Pt(12.5)
r.bold = True

p = cell_tt01.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("Độc lập - Tự do - Hạnh phúc")
r.font.name = "Times New Roman"
r.font.size = Pt(13)
r.bold = True

p = cell_tt01.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
p.paragraph_format.line_spacing = 1.0
r = p.add_run("───────────────────") # 23 ký tự
r.font.name = "Times New Roman"
r.font.size = Pt(10)

# Tiêu đề TỜ TRÌNH LÃNH ĐẠO
p_title = doc_tt.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.space_before = Pt(14)
p_title.paragraph_format.space_after = Pt(6)
r = p_title.add_run("TỜ TRÌNH LÃNH ĐẠO")
r.font.name = "Times New Roman"
r.font.size = Pt(14)
r.bold = True

# Trích yếu Về việc:
p_vv = doc_tt.add_paragraph()
p_vv.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_vv.paragraph_format.space_before = Pt(0)
p_vv.paragraph_format.space_after = Pt(10)
r = p_vv.add_run("Về việc: báo cáo kết quả rà soát theo Công văn số 16451/DON-KTr1")
r.font.name = "Times New Roman"
r.font.size = Pt(13)
r.bold = True

# Nội dung Tờ trình
p_nd = doc_tt.add_paragraph()
p_nd.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_nd.paragraph_format.first_line_indent = Cm(1.25)
p_nd.paragraph_format.space_before = Pt(3)
p_nd.paragraph_format.space_after = Pt(0)
p_nd.paragraph_format.line_spacing = 1.15
r = p_nd.add_run("- Nội dung công việc: ")
r.bold = True
r2 = p_nd.add_run("Căn cứ Công văn số 16451/DON-KTr1 ngày 25/08/2026 của Thuế thành phố Đồng Nai về việc chuyển thông tin liên quan đến vi phạm pháp luật về thuế, hóa đơn theo yêu cầu của Cơ quan Cảnh sát điều tra - Công an thành phố Đồng Nai tại Công văn số 5734/CSKT-Đ2 ngày 19/07/2026 đối với các doanh nghiệp liên quan đến vụ án mua bán trái phép hóa đơn.")

p_yk = doc_tt.add_paragraph()
p_yk.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_yk.paragraph_format.first_line_indent = Cm(1.25)
p_yk.paragraph_format.space_before = Pt(3)
p_yk.paragraph_format.space_after = Pt(0)
p_yk.paragraph_format.line_spacing = 1.15
r = p_yk.add_run("- Ý kiến tham gia của các đơn vị liên quan: ")
r.bold = True
r2 = p_yk.add_run("Tổ Kiểm tra số 1 đã phát hành Phiếu chuyển chuyển Tổ Quản lý doanh nghiệp 1 và Tổ Quản lý doanh nghiệp 2 để tiến hành rà soát, xử lý vi phạm đối với các doanh nghiệp thuộc địa bàn quản lý.")

p_dx = doc_tt.add_paragraph()
p_dx.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_dx.paragraph_format.first_line_indent = Cm(1.25)
p_dx.paragraph_format.space_before = Pt(3)
p_dx.paragraph_format.space_after = Pt(8)
p_dx.paragraph_format.line_spacing = 1.15
r = p_dx.add_run("- Đề xuất của cán bộ giải quyết: ")
r.bold = True
r2 = p_dx.add_run("Kính trình Ban Lãnh đạo Thuế cơ sở 3 thành phố Đồng Nai xem xét, phê duyệt và ký ban hành dự thảo công văn báo cáo kết quả nêu trên gửi Thuế thành phố Đồng Nai.")

# 3 Khối Ký duyệt (Cán bộ đề xuất, Tổ trưởng, Lãnh đạo phê duyệt)
# Khối 1: Cán bộ đề xuất
tbl_sig1 = doc_tt.add_table(rows=1, cols=2)
remove_table_borders(tbl_sig1)
set_cell_width(tbl_sig1.cell(0, 0), 4535) # 8.0cm
set_cell_width(tbl_sig1.cell(0, 1), 4535) # 8.0cm
c1 = tbl_sig1.cell(0, 1)
p = c1.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Trấn Biên, ngày     tháng    năm 2026")
r.italic = True
r.font.size = Pt(12)
p = c1.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("CÁN BỘ ĐỀ XUẤT")
r.bold = True
r.font.size = Pt(13)
for _ in range(3):
    p = c1.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
p = c1.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Nguyễn Thị Ánh Ngọc")
r.bold = True
r.font.size = Pt(13)

# Khối 2: Ý kiến Tổ trưởng & Ký
p_yk_tt = doc_tt.add_paragraph()
p_yk_tt.paragraph_format.first_line_indent = Cm(1.25)
p_yk_tt.paragraph_format.space_before = Pt(4)
p_yk_tt.paragraph_format.space_after = Pt(2)
p_yk_tt.add_run("Ý kiến của lãnh đạo Tổ KTR1: ............................................................................")

tbl_sig2 = doc_tt.add_table(rows=1, cols=2)
remove_table_borders(tbl_sig2)
set_cell_width(tbl_sig2.cell(0, 0), 4535)
set_cell_width(tbl_sig2.cell(0, 1), 4535)
c2 = tbl_sig2.cell(0, 1)
p = c2.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Trấn Biên, ngày     tháng    năm 2026")
r.italic = True
r.font.size = Pt(12)
p = c2.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("TỔ TRƯỞNG")
r.bold = True
r.font.size = Pt(13)
for _ in range(3):
    p = c2.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
p = c2.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(4)
r = p.add_run("Nguyễn Trường Giang")
r.bold = True
r.font.size = Pt(13)

# Khối 3: Ý kiến Lãnh đạo & Phê duyệt
p_yk_ld = doc_tt.add_paragraph()
p_yk_ld.paragraph_format.first_line_indent = Cm(1.25)
p_yk_ld.paragraph_format.space_before = Pt(4)
p_yk_ld.paragraph_format.space_after = Pt(2)
p_yk_ld.add_run("Ý kiến phê duyệt của Lãnh đạo TCS3: ................................................................")

tbl_sig3 = doc_tt.add_table(rows=1, cols=2)
remove_table_borders(tbl_sig3)
set_cell_width(tbl_sig3.cell(0, 0), 4535)
set_cell_width(tbl_sig3.cell(0, 1), 4535)
c3 = tbl_sig3.cell(0, 1)
p = c3.paragraphs[0]
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Trấn Biên, ngày     tháng    năm 2026")
r.italic = True
r.font.size = Pt(12)
p = c3.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("KT. TRƯỞNG THUẾ CƠ SỞ")
r.bold = True
r.font.size = Pt(12)
p = c3.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("PHÓ TRƯỞNG THUẾ CƠ SỞ")
r.bold = True
r.font.size = Pt(12)
for _ in range(3):
    p = c3.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
p = c3.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Phạm Đồng Xuân")
r.bold = True
r.font.size = Pt(13)

tt_name = "to trinh 5734.docx"
tt_drive_path = os.path.join(drive_dir, tt_name)
tt_vault_path = os.path.join(vault_dir, tt_name)

doc_tt.save(tt_drive_path)
doc_tt.save(tt_vault_path)
print(f"TT Saved Drive: {tt_drive_path} (exists: {os.path.exists(tt_drive_path)})")
print(f"TT Saved Vault: {tt_vault_path} (exists: {os.path.exists(tt_vault_path)})")

print("ALL FILES GENERATED AND SAVED SUCCESSFULLY!")
