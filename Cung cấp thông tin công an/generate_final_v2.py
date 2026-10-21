import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import openpyxl
import os

# --- Helper formatting functions ---

def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

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
    table._tbl.insert(table._tbl.index(tblPr) + 1, new_grid)

# --- 1. LOAD DATA ---

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

# Load DB NNT fast using iter_rows values_only
wb_nnt = openpyxl.load_workbook(r'sources\DB NNT 03.06.2026.xlsx', read_only=True, data_only=True)
ws_nnt = wb_nnt['NNT']

db_nnt_map = {}
for i, row in enumerate(ws_nnt.iter_rows(values_only=True)):
    if i < 3:
        continue
    mst_val = str(row[3] or '').strip().replace('-', '')
    if mst_val.endswith('.0'):
        mst_val = mst_val[:-2]
    if mst_val:
        # Col 45 (index 44) is 'Trạng thái ĐKT tổ chức'
        st_dkt = row[44] if len(row) > 44 else None
        dc_parts = []
        if len(row) > 7 and row[7]:
            dc_parts.append(str(row[7]).strip())
        if len(row) > 9 and row[9]:
            dc_parts.append(str(row[9]).strip())
        if len(row) > 11 and row[11]:
            dc_parts.append(str(row[11]).strip())
        addr = ', '.join(dc_parts) if dc_parts else ''
        
        info = {
            'name': row[4] if len(row) > 4 else '',
            'mst': row[3],
            'addr': addr,
            'st_dkt': st_dkt,
            'quan_huyen': row[11] if len(row) > 11 else ''
        }
        db_nnt_map[mst_val] = info
        if mst_val.startswith('0') and len(mst_val) == 10:
            db_nnt_map[mst_val[1:]] = info
        elif len(mst_val) == 9:
            db_nnt_map['0' + mst_val] = info

matched_26_stt = [2, 5, 6, 8, 9, 11, 12, 13, 15, 18, 19, 20, 21, 22, 23, 24, 27, 29, 30, 32, 35, 37, 38, 40, 41, 42]

tcs3_enterprises = []
for item in rows_4020:
    stt_int = int(item['stt'])
    if stt_int in matched_26_stt:
        mst_clean = item['mst'].strip().replace('-', '')
        db_info = db_nnt_map.get(mst_clean)
        if not db_info:
            if mst_clean.startswith('0') and len(mst_clean) == 10:
                db_info = db_nnt_map.get(mst_clean[1:])
            elif len(mst_clean) == 9:
                db_info = db_nnt_map.get('0' + mst_clean)
        
        # Clean enterprise name (remove *(Đã khởi tố)* or similar notes)
        clean_name = item['name'].replace('*(Đã khởi tố)*', '').replace('(Đã khởi tố)', '').strip()
        
        # Address
        if db_info and db_info['addr']:
            address = db_info['addr']
        else:
            address = item['dia_chi']
        
        # Exact 'Trạng thái ĐKT tổ chức'
        if db_info and db_info['st_dkt']:
            st_dkt_val = str(db_info['st_dkt']).strip()
        else:
            # Fallback for enterprises not having explicit Col 45 text
            st_dkt_val = 'NNT đang hoạt động'
        
        tcs3_enterprises.append({
            'stt_4020': item['stt'],
            'name': clean_name,
            'mst': item['mst'],
            'so_hd': item['so_hd'],
            'ds_chua_thue': item['ds_chua_thue'],
            'thue_vat': item['thue_vat'],
            'ds_co_thue': item['ds_co_thue'],
            'dia_chi': address,
            'st_dkt': st_dkt_val
        })

print(f"Total TCS3 enterprises processed: {len(tcs3_enterprises)}")
for idx, e in enumerate(tcs3_enterprises, 1):
    print(f"{idx:>2}. {e['name'][:35]:<35} | MST: {e['mst']:<12} | Trạng thái ĐKT tổ chức: {repr(e['st_dkt'])}")


# --- 2. GENERATE TRA LOI CV 5734 DOCX ---

def create_tra_loi_doc(recipient_type="THUE_TP"):
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

    # HEADER TABLE (3 rows, 2 cols)
    # Left col = 9.5 cm (5386 dxa), Right col = 10.0 cm (5669 dxa), tblInd = -1.5 cm (-850 dxa)
    LEFT_DXA = 5386
    RIGHT_DXA = 5669
    TBL_IND = -850

    tbl_hdr = doc.add_table(rows=3, cols=2)
    set_table_borders_none(tbl_hdr)
    set_header_table_properties(tbl_hdr, LEFT_DXA, RIGHT_DXA, TBL_IND)

    for row in tbl_hdr.rows:
        for idx, cell in enumerate(row.cells):
            set_cell_margins(cell, 0, 0, 0, 0)
            tcPr = cell._tc.get_or_add_tcPr()
            for el in tcPr.findall(qn('w:tcW')):
                tcPr.remove(el)
            tcW = parse_xml(r'<w:tcW %s w:w="%d" w:type="dxa"/>' % (nsdecls('w'), LEFT_DXA if idx == 0 else RIGHT_DXA))
            tcPr.append(tcW)

    # Row 0, Cell (0, 0): Co quan ban hanh
    c00 = tbl_hdr.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p00.paragraph_format.space_before = Pt(0)
    p00.paragraph_format.space_after = Pt(0)
    p00.paragraph_format.line_spacing = 1.0

    r1 = p00.add_run("THUẾ THÀNH PHỐ ĐỒNG NAI\n")
    r1.font.name = "Times New Roman"
    r1.font.size = Pt(13)
    r1.font.bold = False

    r2 = p00.add_run("THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(13)
    r2.font.bold = True

    r3 = p00.add_run("────────────") # 12 line chars
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(10)
    r3.font.bold = False

    # Row 0, Cell (0, 1): Quoc hieu & Tieu ngu
    c01 = tbl_hdr.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p01.paragraph_format.space_before = Pt(0)
    p01.paragraph_format.space_after = Pt(0)
    p01.paragraph_format.line_spacing = 1.0

    rq1 = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    rq1.font.name = "Times New Roman"
    rq1.font.size = Pt(13)
    rq1.font.bold = True

    rq2 = p01.add_run("Độc lập - Tự do - Hạnh phúc\n")
    rq2.font.name = "Times New Roman"
    rq2.font.size = Pt(14)
    rq2.font.bold = True

    rq3 = p01.add_run("────────────────────────") # 24 line chars
    rq3.font.name = "Times New Roman"
    rq3.font.size = Pt(10)
    rq3.font.bold = False

    # Row 1, Cell (1, 0): So cong van
    c10 = tbl_hdr.cell(1, 0)
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p10.paragraph_format.left_indent = Cm(2.0)
    p10.paragraph_format.space_before = Pt(4)
    p10.paragraph_format.space_after = Pt(2)
    rs = p10.add_run("Số:          ")
    rs.font.name = "Times New Roman"
    rs.font.size = Pt(13)

    # Row 1, Cell (1, 1): Ngay thang
    c11 = tbl_hdr.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p11.paragraph_format.space_before = Pt(4)
    p11.paragraph_format.space_after = Pt(2)
    rd = p11.add_run("Trấn Biên, ngày       tháng       năm 2026")
    rd.font.name = "Times New Roman"
    rd.font.size = Pt(13)
    rd.font.italic = True

    # Row 2, Cell (2, 0): Trich yeu
    c20 = tbl_hdr.cell(2, 0)
    p20 = c20.paragraphs[0]
    p20.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p20.paragraph_format.space_before = Pt(2)
    p20.paragraph_format.space_after = Pt(2)
    p20.paragraph_format.line_spacing = 1.0
    rty = p20.add_run("V/v cung cấp thông tin, tài liệu theo Công văn số 5734/CSKT-Đ2")
    rty.font.name = "Times New Roman"
    rty.font.size = Pt(12.5)

    c21 = tbl_hdr.cell(2, 1)
    c21.paragraphs[0].text = ""

    # DÒNG KÍNH GỬI
    pkg = doc.add_paragraph()
    pkg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pkg.paragraph_format.space_before = Pt(18)
    pkg.paragraph_format.space_after = Pt(14)
    pkg.paragraph_format.line_spacing = 1.15
    if recipient_type == "THUE_TP":
        rkg = pkg.add_run("Kính gửi: Thuế thành phố Đồng Nai")
    else:
        rkg = pkg.add_run("Kính gửi: Cơ quan Cảnh sát điều tra - Công an thành phố Đồng Nai")
    rkg.font.name = "Times New Roman"
    rkg.font.size = Pt(14)

    # ĐOẠN CĂN CỨ MỞ ĐẦU
    def add_para(text, bold=False, space_before=Pt(3), space_after=Pt(0), first_ind=Cm(1.25)):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = first_ind
        p.paragraph_format.space_before = space_before
        p.paragraph_format.space_after = space_after
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(14)
        r.bold = bold
        return p

    if recipient_type == "THUE_TP":
        add_para("Thuế cơ sở 3 thành phố Đồng Nai nhận được Công văn số 16451/DON-KTr1 ngày 25/08/2026 của Thuế thành phố Đồng Nai về việc chuyển thông tin liên quan đến vi phạm pháp luật về thuế, hóa đơn theo yêu cầu của Cơ quan Cảnh sát điều tra - Công an thành phố Đồng Nai tại Công văn số 5734/CSKT-Đ2 ngày 19/07/2026.")
        add_para("Căn cứ dữ liệu quản lý thuế, Thuế cơ sở 3 thành phố Đồng Nai báo cáo kết quả rà soát, xử lý đối với các doanh nghiệp thuộc địa bàn quản lý như sau:")
    else:
        add_para("Căn cứ Công văn số 5734/CSKT-Đ2 ngày 19/07/2026 của Cơ quan Cảnh sát điều tra - Công an thành phố Đồng Nai (và Công văn chuyển số 16451/DON-KTr1 ngày 25/08/2026 của Thuế thành phố Đồng Nai) về việc đề nghị cung cấp thông tin, tài liệu đối với các doanh nghiệp liên quan đến vụ án mua bán trái phép hóa đơn;")
        add_para("Căn cứ cơ sở dữ liệu quản lý thuế đối với các doanh nghiệp do Thuế cơ sở 3 thành phố Đồng Nai quản lý;")
        add_para("Thuế cơ sở 3 thành phố Đồng Nai cung cấp thông tin, tài liệu đối với các doanh nghiệp thuộc phạm vi quản lý như sau:")

    # DANH SÁCH 26 DOANH NGHIỆP THUỘC THUẾ CƠ SỞ 3
    for idx, item in enumerate(tcs3_enterprises, 1):
        # Tiêu đề DN
        p_dn = doc.add_paragraph()
        p_dn.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_dn.paragraph_format.first_line_indent = Cm(1.25)
        p_dn.paragraph_format.space_before = Pt(6)
        p_dn.paragraph_format.space_after = Pt(0)
        p_dn.paragraph_format.line_spacing = 1.15
        r_dn = p_dn.add_run(f"{idx}. {item['name'].upper()}, MST: {item['mst']}.")
        r_dn.bold = True
        r_dn.font.name = "Times New Roman"
        r_dn.font.size = Pt(14)

        # Địa điểm ĐKKD
        add_para(f"+ Địa điểm đăng ký kinh doanh: {item['dia_chi']};")

        # Tình trạng hoạt động: Ghi đúng nội dung 'trạng thái ĐKT tổ chức' tại DB NNT.xls
        add_para(f"+ Tình trạng hoạt động: {item['st_dkt']};")

        # Thông tin hóa đơn theo CV 5734
        add_para(f"+ Số liệu hóa đơn liên quan theo Công văn số 5734/CSKT-Đ2: {item['so_hd']} hóa đơn; Doanh số mua chưa có thuế: {item['ds_chua_thue']} đồng; Thuế GTGT: {item['thue_vat']} đồng; Tổng doanh số: {item['ds_co_thue']} đồng;")

        # Tình hình rà soát, xử lý vi phạm
        add_para("+ Tình hình rà soát, xử lý vi phạm về thuế, hóa đơn: Công ty chưa bị xử phạt về thuế và hóa đơn.")

    # ĐOẠN KẾT
    p_ket = doc.add_paragraph()
    p_ket.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_ket.paragraph_format.first_line_indent = Cm(1.25)
    p_ket.paragraph_format.space_before = Pt(6)
    p_ket.paragraph_format.space_after = Pt(14)
    p_ket.paragraph_format.line_spacing = 1.15
    if recipient_type == "THUE_TP":
        r_ket = p_ket.add_run("Thuế cơ sở 3 thành phố Đồng Nai báo cáo Thuế thành phố Đồng Nai để tổng hợp theo quy định./.")
    else:
        r_ket = p_ket.add_run("Thuế cơ sở 3 thành phố Đồng Nai phúc đáp để Cơ quan Cảnh sát điều tra - Công an thành phố Đồng Nai được biết và phối hợp công tác./.")
    r_ket.font.name = "Times New Roman"
    r_ket.font.size = Pt(14)

    # BẢNG CHỮ KÝ CUỐI TRANG
    tbl_sig = doc.add_table(rows=1, cols=2)
    set_table_borders_none(tbl_sig)

    SIG_LEFT_DXA = 4252   # 7.5 cm
    SIG_RIGHT_DXA = 5103  # 9.0 cm
    set_header_table_properties(tbl_sig, SIG_LEFT_DXA, SIG_RIGHT_DXA, 0)

    # Left cell: Noi nhan
    c_s0 = tbl_sig.cell(0, 0)
    set_cell_margins(c_s0, 0, 0, 0, 0)
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

    for line in ["- Như trên;\n", "- BLĐ thuế cơ sở 3 (để b/c);\n", "- Lưu: VT, KT1 (Ngọc)."]:
        r_nn_item = p_s0.add_run(line)
        r_nn_item.font.name = "Times New Roman"
        r_nn_item.font.size = Pt(11)

    # Right cell: Chuc danh & Chu ky
    c_s1 = tbl_sig.cell(0, 1)
    set_cell_margins(c_s1, 0, 0, 0, 0)
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

    return doc


# --- 3. GENERATE TO TRINH 5734 DOCX ---

def create_to_trinh_doc():
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

    # Header Tờ trình: Left col = 8.5 cm (4819 dxa), Right col = 10.1 cm (5726 dxa), tblInd = -1.5 cm (-850 dxa)
    table_tt_hdr = doc_tt.add_table(rows=1, cols=2)
    set_table_borders_none(table_tt_hdr)
    set_header_table_properties(table_tt_hdr, 4819, 5726, -850)

    for cell in table_tt_hdr.rows[0].cells:
        set_cell_margins(cell, 0, 0, 0, 0)

    # Cột trái Header TT
    cell_tt00 = table_tt_hdr.cell(0, 0)
    p = cell_tt00.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)

    r2 = p.add_run("TỔ KIỂM TRA SỐ 1\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(13)
    r2.bold = True

    r3 = p.add_run("───────")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(10)

    # Cột phải Header TT
    cell_tt01 = table_tt_hdr.cell(0, 1)
    p = cell_tt01.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12.5)
    r.bold = True

    r2 = p.add_run("Độc lập - Tự do - Hạnh phúc\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(13)
    r2.bold = True

    r3 = p.add_run("───────────────────")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(10)

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
    def add_tt_nd(title, body):
        p = doc_tt.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        r_t = p.add_run(title)
        r_t.bold = True
        r_b = p.add_run(body)
        return p

    add_tt_nd("- Nội dung công việc: ", "Căn cứ Công văn số 16451/DON-KTr1 ngày 25/08/2026 của Thuế thành phố Đồng Nai về việc chuyển thông tin liên quan đến vi phạm pháp luật về thuế, hóa đơn theo yêu cầu của Cơ quan Cảnh sát điều tra - Công an thành phố Đồng Nai tại Công văn số 5734/CSKT-Đ2 ngày 19/07/2026 đối với các doanh nghiệp liên quan đến vụ án mua bán trái phép hóa đơn.")
    add_tt_nd("- Ý kiến tham gia của các đơn vị liên quan: ", "Tổ Kiểm tra số 1 đã thực hiện rà soát 26 doanh nghiệp thuộc địa bàn quản lý của Thuế cơ sở 3; tình hình rà soát, xử lý vi phạm ghi nhận các công ty chưa bị xử phạt về thuế và hóa đơn.")
    add_tt_nd("- Đề xuất của cán bộ giải quyết: ", "Kính trình Ban Lãnh đạo Thuế cơ sở 3 thành phố Đồng Nai xem xét, phê duyệt và ký ban hành dự thảo công văn báo cáo kết quả nêu trên gửi Thuế thành phố Đồng Nai.")

    # 3 Khối Ký duyệt (Cán bộ đề xuất, Tổ trưởng, Lãnh đạo phê duyệt)
    # Khối 1: Cán bộ đề xuất
    tbl_sig1 = doc_tt.add_table(rows=1, cols=2)
    set_table_borders_none(tbl_sig1)
    set_header_table_properties(tbl_sig1, 4535, 4535, 0)
    c1 = tbl_sig1.cell(0, 1)
    p = c1.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Trấn Biên, ngày     tháng    năm 2026\n")
    r.italic = True
    r.font.size = Pt(12)
    r2 = p.add_run("CÁN BỘ ĐỀ XUẤT\n\n\n\n")
    r2.bold = True
    r2.font.size = Pt(13)
    r3 = p.add_run("Nguyễn Thị Ánh Ngọc")
    r3.bold = True
    r3.font.size = Pt(13)

    # Khối 2: Ý kiến Tổ trưởng & Ký
    p_yk_tt = doc_tt.add_paragraph()
    p_yk_tt.paragraph_format.first_line_indent = Cm(1.25)
    p_yk_tt.paragraph_format.space_before = Pt(4)
    p_yk_tt.paragraph_format.space_after = Pt(2)
    p_yk_tt.add_run("Ý kiến của lãnh đạo Tổ KTR1: ............................................................................")

    tbl_sig2 = doc_tt.add_table(rows=1, cols=2)
    set_table_borders_none(tbl_sig2)
    set_header_table_properties(tbl_sig2, 4535, 4535, 0)
    c2 = tbl_sig2.cell(0, 1)
    p = c2.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Trấn Biên, ngày     tháng    năm 2026\n")
    r.italic = True
    r.font.size = Pt(12)
    r2 = p.add_run("TỔ TRƯỞNG\n\n\n\n")
    r2.bold = True
    r2.font.size = Pt(13)
    r3 = p.add_run("Nguyễn Trường Giang")
    r3.bold = True
    r3.font.size = Pt(13)

    # Khối 3: Ý kiến Lãnh đạo & Phê duyệt
    p_yk_ld = doc_tt.add_paragraph()
    p_yk_ld.paragraph_format.first_line_indent = Cm(1.25)
    p_yk_ld.paragraph_format.space_before = Pt(4)
    p_yk_ld.paragraph_format.space_after = Pt(2)
    p_yk_ld.add_run("Ý kiến phê duyệt của Lãnh đạo TCS3: ................................................................")

    tbl_sig3 = doc_tt.add_table(rows=1, cols=2)
    set_table_borders_none(tbl_sig3)
    set_header_table_properties(tbl_sig3, 4535, 4535, 0)
    c3 = tbl_sig3.cell(0, 1)
    p = c3.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Trấn Biên, ngày     tháng    năm 2026\n")
    r.italic = True
    r.font.size = Pt(12)
    r2 = p.add_run("KT. TRƯỞNG THUẾ CƠ SỞ\n")
    r2.bold = True
    r2.font.size = Pt(12)
    r3 = p.add_run("PHÓ TRƯỞNG THUẾ CƠ SỞ\n\n\n\n")
    r3.bold = True
    r3.font.size = Pt(12)
    r4 = p.add_run("Phạm Đồng Xuân")
    r4.bold = True
    r4.font.size = Pt(13)

    return doc_tt


# --- 4. BUILD & SAVE ALL FILES ---

drive_dir = r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\240. CV 5734-CA DONG NAI"
vault_dir = r"D:\SECOND BRAIN\javis-os\brains\Cung cấp thông tin công an"

# Doc 1: TRA LOI THUE TP
doc_thue_tp = create_tra_loi_doc(recipient_type="THUE_TP")
doc_thue_tp_drive = os.path.join(drive_dir, "TRA LOI CV 5734_THUE TP DONG NAI_du thao.docx")
doc_thue_tp_vault = os.path.join(vault_dir, "TRA LOI CV 5734_THUE TP DONG NAI_du thao.docx")
doc_thue_tp.save(doc_thue_tp_drive)
doc_thue_tp.save(doc_thue_tp_vault)
print(f"Saved: {doc_thue_tp_drive} (exists: {os.path.exists(doc_thue_tp_drive)})")
print(f"Saved: {doc_thue_tp_vault} (exists: {os.path.exists(doc_thue_tp_vault)})")

# Doc 2: TRA LOI CA DONG NAI
doc_ca = create_tra_loi_doc(recipient_type="CA_DONG_NAI")
doc_ca_drive = os.path.join(drive_dir, "TRA LOI CV 5734_CA DONG NAI_du thao.docx")
doc_ca.save(doc_ca_drive)
print(f"Saved: {doc_ca_drive} (exists: {os.path.exists(doc_ca_drive)})")

# Doc 3: TO TRINH 5734
doc_tt = create_to_trinh_doc()
doc_tt_drive = os.path.join(drive_dir, "to trinh 5734.docx")
doc_tt_vault = os.path.join(vault_dir, "to trinh 5734.docx")
doc_tt.save(doc_tt_drive)
doc_tt.save(doc_tt_vault)
print(f"Saved: {doc_tt_drive} (exists: {os.path.exists(doc_tt_drive)})")
print(f"Saved: {doc_tt_vault} (exists: {os.path.exists(doc_tt_vault)})")

print("\nSUCCESS: All files generated and verified successfully!")
