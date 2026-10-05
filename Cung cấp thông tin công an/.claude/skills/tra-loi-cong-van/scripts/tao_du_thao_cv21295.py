import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import os

target_path = r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\254. CV 21295-CA HAI PHONG-KT1+KT2+KT3\TRA LOI CV 21295_CA HAI PHONG_du thao.docx"

doc = docx.Document()

# Page setup: A4, standard margins
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.0)
section.bottom_margin = Cm(2.0)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.0)

# Set Normal style font
style_normal = doc.styles['Normal']
style_normal.font.name = 'Times New Roman'
style_normal.font.size = Pt(14)
style_normal.font.color.rgb = RGBColor(0, 0, 0)

# Helper for cell margins
def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# Helper for table borders none
def set_table_borders_none(table):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(r'''
        <w:tblBorders %s>
            <w:top w:val="none"/>
            <w:left w:val="none"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
            <w:insideH w:val="none"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''' % nsdecls('w'))
    tblPr.append(tblBorders)

# Helper for fixed table layout and negative indent
def set_header_table_properties(table, left_col_dxa, right_col_dxa, tbl_ind_dxa):
    tblPr = table._tbl.tblPr
    total_w = left_col_dxa + right_col_dxa
    
    # Remove existing tblW if any
    for el in tblPr.findall(qn('w:tblW')):
        tblPr.remove(el)
    
    tblW = parse_xml(r'<w:tblW %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), total_w))
    tblPr.append(tblW)
    
    # Indent (tblInd)
    tblInd = parse_xml(r'<w:tblInd %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), tbl_ind_dxa))
    tblPr.append(tblInd)
    
    # Layout fixed
    tblLayout = parse_xml(r'<w:tblLayout %s w:type="fixed"/>' % nsdecls('w'))
    tblPr.append(tblLayout)
    
    # tblGrid
    tblGrid = table._tbl.find(qn('w:tblGrid'))
    if tblGrid is not None:
        table._tbl.remove(tblGrid)
    new_grid = parse_xml(r'''
        <w:tblGrid %s>
            <w:gridCol w:w="%d"/>
            <w:gridCol w:w="%d"/>
        </w:tblGrid>
    ''' % (nsdecls('w'), left_col_dxa, right_col_dxa))
    # insert tblGrid right after tblPr
    table._tbl.insert(table._tbl.index(tblPr) + 1, new_grid)

# --- 1. HEADER TABLE (3 rows, 2 cols) ---
tbl_hdr = doc.add_table(rows=3, cols=2)
set_table_borders_none(tbl_hdr)

LEFT_DXA = 5159   # 9.1 cm
RIGHT_DXA = 5783  # 10.2 cm
TBL_IND = -850    # -1.5 cm

set_header_table_properties(tbl_hdr, LEFT_DXA, RIGHT_DXA, TBL_IND)

for row in tbl_hdr.rows:
    for idx, cell in enumerate(row.cells):
        set_cell_margins(cell, 0, 0, 0, 0)
        tcPr = cell._tc.get_or_add_tcPr()
        for el in tcPr.findall(qn('w:tcW')):
            tcPr.remove(el)
        tcW = parse_xml(r'<w:tcW %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), LEFT_DXA if idx == 0 else RIGHT_DXA))
        tcPr.append(tcW)

# Row 1, Cell (0, 0): Co quan ban hanh
c00 = tbl_hdr.cell(0, 0)
p00 = c00.paragraphs[0]
p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
p00.paragraph_format.space_before = Pt(0)
p00.paragraph_format.space_after = Pt(0)
p00.paragraph_format.line_spacing = 1.0

# Dong 1: THUẾ THÀNH PHỐ ĐỒNG NAI (13, chu dung)
r1 = p00.add_run("THUẾ THÀNH PHỐ ĐỒNG NAI\n")
r1.font.name = "Times New Roman"
r1.font.size = Pt(13)
r1.font.bold = False

# Dong 2: THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI (13, in dam)
r2 = p00.add_run("THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI\n")
r2.font.name = "Times New Roman"
r2.font.size = Pt(13)
r2.font.bold = True

# Dong 3: Duong ke ngang 1/3
r3 = p00.add_run("─" * 12)
r3.font.name = "Times New Roman"
r3.font.size = Pt(10)
r3.font.bold = False

# Row 1, Cell (0, 1): Quoc hieu & Tieu ngu
c01 = tbl_hdr.cell(0, 1)
p01 = c01.paragraphs[0]
p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
p01.paragraph_format.space_before = Pt(0)
p01.paragraph_format.space_after = Pt(0)
p01.paragraph_format.line_spacing = 1.0

# Dong 1: CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM (13, in dam)
rq1 = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
rq1.font.name = "Times New Roman"
rq1.font.size = Pt(13)
rq1.font.bold = True

# Dong 2: Độc lập - Tự do - Hạnh phúc (14, in dam)
rq2 = p01.add_run("Độc lập - Tự do - Hạnh phúc\n")
rq2.font.name = "Times New Roman"
rq2.font.size = Pt(14)
rq2.font.bold = True

# Dong 3: Net gach ngang lien duoi vua het chu (24 ky tu)
rq3 = p01.add_run("─" * 24)
rq3.font.name = "Times New Roman"
rq3.font.size = Pt(10)
rq3.font.bold = False

# Row 2, Cell (1, 0): So cong van
c10 = tbl_hdr.cell(1, 0)
p10 = c10.paragraphs[0]
p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
p10.paragraph_format.space_before = Pt(4)
p10.paragraph_format.space_after = Pt(2)
rs = p10.add_run("Số:          ")
rs.font.name = "Times New Roman"
rs.font.size = Pt(13)
rs.font.bold = False

# Row 2, Cell (1, 1): Ngay thang
c11 = tbl_hdr.cell(1, 1)
p11 = c11.paragraphs[0]
p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
p11.paragraph_format.space_before = Pt(4)
p11.paragraph_format.space_after = Pt(2)
rd = p11.add_run("Đồng Nai, ngày       tháng       năm 2026")
rd.font.name = "Times New Roman"
rd.font.size = Pt(13)
rd.font.italic = True

# Row 3, Cell (2, 0): Trich yeu (chu dung, khong in nghieng, khong in dam)
c20 = tbl_hdr.cell(2, 0)
p20 = c20.paragraphs[0]
p20.alignment = WD_ALIGN_PARAGRAPH.CENTER
p20.paragraph_format.space_before = Pt(2)
p20.paragraph_format.space_after = Pt(2)
p20.paragraph_format.line_spacing = 1.15
rty = p20.add_run("V/v cung cấp thông tin, hồ sơ DN theo Công văn 21295/VPCQCSĐT(Đ1).")
rty.font.name = "Times New Roman"
rty.font.size = Pt(13)
rty.font.bold = False
rty.font.italic = False

# Row 3, Cell (2, 1): O trong
c21 = tbl_hdr.cell(2, 1)
p21 = c21.paragraphs[0]
p21.paragraph_format.space_before = Pt(0)
p21.paragraph_format.space_after = Pt(0)

# --- 2. DONG KINH GUI ---
pkg = doc.add_paragraph()
pkg.alignment = WD_ALIGN_PARAGRAPH.CENTER
pkg.paragraph_format.space_before = Pt(18)
pkg.paragraph_format.space_after = Pt(12)
pkg.paragraph_format.line_spacing = 1.15
rkg = pkg.add_run("Kính gửi: Cơ quan Cảnh sát điều tra - Công an thành phố Hải Phòng")
rkg.font.name = "Times New Roman"
rkg.font.size = Pt(14)
rkg.font.bold = False

# Helper for body paragraphs
def add_body_p(text, bold=False, space_before=Pt(3), space_after=Pt(0), first_line_ind=Cm(1.25), align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = space_before
    p.paragraph_format.space_after = space_after
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.first_line_indent = first_line_ind
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(14)
    run.font.bold = bold
    return p

# --- 3. THAN BAI ---
add_body_p("Căn cứ Công văn số 21295/VPCQCSĐT(Đ1) ngày 21/09/2026 của Cơ quan Cảnh sát điều tra - Công an thành phố Hải Phòng (dấu đến Thuế cơ sở 3 TP Đồng Nai số 14020 ngày 24/09/2026) về việc đề nghị cung cấp thông tin, tài liệu liên quan đến vụ án Bùi Thanh An cùng đồng phạm mua bán trái phép hóa đơn theo Quyết định khởi tố vụ án hình sự số 466/QĐ-VPCQCSĐT ngày 09/04/2026;")
add_body_p("Căn cứ cơ sở dữ liệu quản lý thuế đối với các doanh nghiệp do Thuế cơ sở 3 thành phố Đồng Nai quản lý;")
add_body_p("Thuế cơ sở 3 thành phố Đồng Nai phúc đáp và cung cấp các thông tin, tài liệu đối với các doanh nghiệp thuộc phạm vi quản lý của Tổ Kiểm tra thuế số 1 như sau:")

# Cong ty 1: NYSE VN
add_body_p("1. CÔNG TY TNHH DTPT NYSE VN, MST: 3604056586, NGÀY BẮT ĐẦU HOẠT ĐỘNG: 10/10/2025.", bold=True, space_before=Pt(3))
add_body_p("Thông tin cung cấp:")
add_body_p("+ Địa điểm đăng ký kinh doanh: Số 43/23, đường Cơ Giới, tổ 29, khu phố Thái Hòa, Phường Hố Nai, TP Đồng Nai;")
add_body_p("+ Giấy chứng nhận ĐKKD: Số 3604056586 do Sở Kế hoạch và Đầu tư tỉnh Đồng Nai cấp lần đầu ngày 10/10/2025;")
add_body_p("+ Tình trạng hoạt động: NNT không hoạt động tại địa chỉ đã đăng ký theo Thông báo số 4936/TB-TCS3 ngày 10/03/2026 của Thuế cơ sở 3 thành phố Đồng Nai (ngày đóng trạng thái trên hệ thống: 11/03/2026);")
add_body_p("+ Người đại diện theo pháp luật / Giám đốc: Danh Thanh Tiền (sinh ngày 01/01/1990; SĐT: 0901.298.735; Email: danhtien3579@gmail.com);")
add_body_p("+ Cán bộ quản lý thuế: Thuộc Tổ Kiểm tra thuế số 1 quản lý địa bàn Phường Hố Nai (chưa phân công cán bộ quản lý trực tiếp do doanh nghiệp đã ngừng hoạt động tại địa chỉ đăng ký ngay sau khi thành lập);")
add_body_p("+ Số tài khoản ngân hàng: Doanh nghiệp không đăng ký thông tin tài khoản ngân hàng với Cơ quan Thuế;")
add_body_p("+ Tình hình thực hiện nghĩa vụ thuế: NNT không còn nợ thuế (tính đến ngày 25/09/2026);")
add_body_p("+ Tình hình thanh tra, kiểm tra, xử lý vi phạm: Cơ quan Thuế chưa tiến hành thanh tra, kiểm tra và chưa ban hành quyết định xử phạt vi phạm hành chính đối với doanh nghiệp.")
add_body_p("Hồ sơ, tài liệu cung cấp bao gồm:")
add_body_p("+ Bản sao Thông tin đăng ký thuế, đăng ký doanh nghiệp cấp lần đầu ngày 10/10/2025;")
add_body_p("+ Thông báo số 4936/TB-TCS3 ngày 10/03/2026 về việc NNT không hoạt động tại địa chỉ đã đăng ký;")
add_body_p("+ Tờ khai thuế GTGT từ quý IV/2025 đến quý II/2026;")
add_body_p("+ Tờ khai quyết toán thuế TNDN, quyết toán thuế TNCN, Báo cáo tài chính năm 2025 (đơn vị chưa nộp trên hệ thống do ngừng hoạt động);")
add_body_p("+ Bảng kê chi tiết hóa đơn điện tử mua vào - bán ra từ tháng 10/2025 đến nay (kèm file mềm Excel xuất từ hệ thống HĐĐT gửi về địa chỉ email vpcqcsdt.catphp@gmail.com theo yêu cầu).")

# Cong ty 2: YD FASHION VIET
add_body_p("2. CÔNG TY TNHH MTV YD FASHION VIET, MST: 3604042992, NGÀY BẮT ĐẦU HOẠT ĐỘNG: 08/08/2025.", bold=True, space_before=Pt(3))
add_body_p("Thông tin cung cấp:")
add_body_p("+ Địa điểm đăng ký kinh doanh: Hẻm 1023/71 đường Bùi Văn Hòa, tổ 16, khu phố 7, Phường Long Bình, TP Đồng Nai;")
add_body_p("+ Giấy chứng nhận ĐKKD: Số 3604042992 do Sở Kế hoạch và Đầu tư tỉnh Đồng Nai cấp lần đầu ngày 08/08/2025;")
add_body_p("+ Tình trạng hoạt động: NNT không hoạt động tại địa chỉ đã đăng ký theo Thông báo số 9273/TB-TCS3 ngày 29/04/2026 của Thuế cơ sở 3 thành phố Đồng Nai (ngày đóng trạng thái trên hệ thống: 29/04/2026);")
add_body_p("+ Người đại diện theo pháp luật / Giám đốc: Phan Văn Vinh (sinh ngày 25/04/1999; SĐT: 0389763073; Email: thoitrangyendung99@gmail.com);")
add_body_p("+ Cán bộ quản lý thuế: Cán bộ Nguyễn Văn Thông - Tổ Kiểm tra thuế số 1, Thuế cơ sở 3 thành phố Đồng Nai;")
add_body_p("+ Số tài khoản ngân hàng: Doanh nghiệp không đăng ký thông tin tài khoản ngân hàng với Cơ quan Thuế;")
add_body_p("+ Tình hình thực hiện nghĩa vụ thuế: NNT không còn nợ thuế (tính đến ngày 25/09/2026);")
add_body_p("+ Tình hình thanh tra, kiểm tra, xử lý vi phạm: Cơ quan Thuế chưa tiến hành thanh tra, kiểm tra và chưa ban hành quyết định xử phạt vi phạm hành chính đối với doanh nghiệp.")
add_body_p("Hồ sơ, tài liệu cung cấp bao gồm:")
add_body_p("+ Bản sao Giấy chứng nhận đăng ký doanh nghiệp cấp lần đầu ngày 08/08/2025;")
add_body_p("+ Thông báo số 9273/TB-TCS3 ngày 29/04/2026 về việc NNT không hoạt động tại địa chỉ đã đăng ký;")
add_body_p("+ Thông báo chấp nhận đăng ký sử dụng hóa đơn điện tử theo Nghị định số 123/2020/NĐ-CP;")
add_body_p("+ Tờ khai thuế GTGT từ quý III/2025 đến quý II/2026;")
add_body_p("+ Tờ khai quyết toán thuế TNDN, quyết toán thuế TNCN, Báo cáo tài chính năm 2025;")
add_body_p("+ Bảng kê chi tiết hóa đơn điện tử mua vào - bán ra từ tháng 08/2025 đến nay (kèm file mềm Excel xuất từ hệ thống HĐĐT gửi về địa chỉ email vpcqcsdt.catphp@gmail.com theo yêu cầu).")

# Dieu phoi 2 cty con lai
add_body_p("Đối với 02 doanh nghiệp còn lại tại Công văn số 21295/VPCQCSĐT(Đ1): Công ty TNHH ĐT VERDATECH (MST: 3604083491) thuộc địa bàn quản lý của Tổ Kiểm tra thuế số 2 và Công ty TNHH Đầu tư Thương mại Dịch vụ Phát triển Lâm Tri (MST: 3604080331) thuộc địa bàn quản lý của Tổ Kiểm tra thuế số 3. Các Tổ kiểm tra đang thực hiện rà soát, thu thập hồ sơ và lập văn bản cung cấp riêng cho Cơ quan Cảnh sát điều tra theo quy định.")
add_body_p("Thuế cơ sở 3 thành phố Đồng Nai phúc đáp để Cơ quan Cảnh sát điều tra - Công an thành phố Hải Phòng được biết và phối hợp công tác./.")

# --- 4. SIGNATURE TABLE (1 row, 2 cols) ---
tbl_sig = doc.add_table(rows=1, cols=2)
set_table_borders_none(tbl_sig)

SIG_LEFT_DXA = 4252   # 7.5 cm
SIG_RIGHT_DXA = 5103  # 9.0 cm
SIG_TOTAL = SIG_LEFT_DXA + SIG_RIGHT_DXA

tblPr_sig = tbl_sig._tbl.tblPr
for el in tblPr_sig.findall(qn('w:tblW')):
    tblPr_sig.remove(el)
tblW_sig = parse_xml(r'<w:tblW %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), SIG_TOTAL))
tblPr_sig.append(tblW_sig)
tblLayout_sig = parse_xml(r'<w:tblLayout %s w:type="fixed"/>' % nsdecls('w'))
tblPr_sig.append(tblLayout_sig)

tblGrid_sig = tbl_sig._tbl.find(qn('w:tblGrid'))
if tblGrid_sig is not None:
    tbl_sig._tbl.remove(tblGrid_sig)
new_grid_sig = parse_xml(r'''
    <w:tblGrid %s>
        <w:gridCol w:w="%d"/>
        <w:gridCol w:w="%d"/>
    </w:tblGrid>
''' % (nsdecls('w'), SIG_LEFT_DXA, SIG_RIGHT_DXA))
tbl_sig._tbl.insert(tbl_sig._tbl.index(tblPr_sig) + 1, new_grid_sig)

# Left cell: Noi nhan
c_s0 = tbl_sig.cell(0, 0)
tcPr_s0 = c_s0._tc.get_or_add_tcPr()
for el in tcPr_s0.findall(qn('w:tcW')):
    tcPr_s0.remove(el)
tcPr_s0.append(parse_xml(r'<w:tcW %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), SIG_LEFT_DXA)))
tcPr_s0.append(parse_xml(r'<w:vAlign %s w:val="top"/>' % nsdecls('w')))

p_s0 = c_s0.paragraphs[0]
p_s0.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_s0.paragraph_format.space_before = Pt(6)
p_s0.paragraph_format.space_after = Pt(0)
p_s0.paragraph_format.line_spacing = 1.0

r_nn_title = p_s0.add_run("Nơi nhận:\n")
r_nn_title.font.name = "Times New Roman"
r_nn_title.font.size = Pt(12)
r_nn_title.font.bold = True
r_nn_title.font.italic = True

for line in ["- Như trên;\n", "- BLĐ thuế cơ sở 3 (để b/c);\n", "- Lưu: VT, KT1 (Tr)."]:
    r_nn_item = p_s0.add_run(line)
    r_nn_item.font.name = "Times New Roman"
    r_nn_item.font.size = Pt(11)
    r_nn_item.font.bold = False
    r_nn_item.font.italic = False

# Right cell: Chuc danh & Chu ky
c_s1 = tbl_sig.cell(0, 1)
tcPr_s1 = c_s1._tc.get_or_add_tcPr()
for el in tcPr_s1.findall(qn('w:tcW')):
    tcPr_s1.remove(el)
tcPr_s1.append(parse_xml(r'<w:tcW %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), SIG_RIGHT_DXA)))
tcPr_s1.append(parse_xml(r'<w:vAlign %s w:val="top"/>' % nsdecls('w')))

p_s1 = c_s1.paragraphs[0]
p_s1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_s1.paragraph_format.space_before = Pt(6)
p_s1.paragraph_format.space_after = Pt(0)
p_s1.paragraph_format.line_spacing = 1.0

r_cd1 = p_s1.add_run("KT. TRƯỞNG THUẾ CƠ SỞ\n")
r_cd1.font.name = "Times New Roman"
r_cd1.font.size = Pt(13)
r_cd1.font.bold = True

r_cd2 = p_s1.add_run("PHÓ TRƯỞNG THUẾ CƠ SỞ\n\n\n\n\n")
r_cd2.font.name = "Times New Roman"
r_cd2.font.size = Pt(13)
r_cd2.font.bold = True

r_sign = p_s1.add_run("Phạm Đồng Xuân")
r_sign.font.name = "Times New Roman"
r_sign.font.size = Pt(14)
r_sign.font.bold = True

# Save document
doc.save(target_path)
print("SUCCESS: Saved document to", target_path)
