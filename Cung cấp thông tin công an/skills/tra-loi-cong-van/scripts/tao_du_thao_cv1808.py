# -*- coding: utf-8 -*-
"""Dự thảo trả lời Công văn 1808/YC-ANĐT-Đ2 (cụm 1821+1808) - CA Hải Phòng."""
import os

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

RED = RGBColor(255, 0, 0)
BLACK = RGBColor(0, 0, 0)
FONT = "Times New Roman"

VAULT = "/brains/Cung cấp thông tin công an"
DRIVE_CANDIDATES = [
    r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\242. CV 1821-CA HAI PHONG-KT1+KT2",
    r"I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\206. CV 1821+1808-CA HAI PHONG-KT1",
]


def set_run_font(run, size=14, bold=False, italic=False, color=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color if color is not None else BLACK
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT)


def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in (("w:top", top), ("w:bottom", bottom), ("w:left", left), ("w:right", right)):
        node = OxmlElement(m)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders_none(table):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        """
        <w:tblBorders %s>
            <w:top w:val="none"/>
            <w:left w:val="none"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
            <w:insideH w:val="none"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    """
        % nsdecls("w")
    )
    tblPr.append(tblBorders)


def set_header_table_properties(table, left_col_dxa, right_col_dxa, tbl_ind_dxa):
    tblPr = table._tbl.tblPr
    total_w = left_col_dxa + right_col_dxa
    for el in tblPr.findall(qn("w:tblW")):
        tblPr.remove(el)
    tblPr.append(parse_xml(r'<w:tblW %s w:w="%d" w:type="dxa"/>' % (nsdecls("w"), total_w)))
    tblPr.append(parse_xml(r'<w:tblInd %s w:w="%d" w:type="dxa"/>' % (nsdecls("w"), tbl_ind_dxa)))
    tblPr.append(parse_xml(r'<w:tblLayout %s w:type="fixed"/>' % nsdecls("w")))
    tblGrid = table._tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        table._tbl.remove(tblGrid)
    new_grid = parse_xml(
        """
        <w:tblGrid %s>
            <w:gridCol w:w="%d"/>
            <w:gridCol w:w="%d"/>
        </w:tblGrid>
    """
        % (nsdecls("w"), left_col_dxa, right_col_dxa)
    )
    table._tbl.insert(table._tbl.index(tblPr) + 1, new_grid)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            set_cell_margins(cell, 0, 0, 0, 0)
            tcPr = cell._tc.get_or_add_tcPr()
            for el in tcPr.findall(qn("w:tcW")):
                tcPr.remove(el)
            w = left_col_dxa if idx == 0 else right_col_dxa
            tcPr.append(parse_xml(r'<w:tcW %s w:w="%d" w:type="dxa"/>' % (nsdecls("w"), w)))


def add_runs(p, parts, size=14):
    """parts: list of (text, bold, italic, color)."""
    for text, bold, italic, color in parts:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, italic=italic, color=color)
    return p


def add_body(doc, parts, bold=False, size=14, first_ind=True, align="justify", before=3, after=0):
    p = doc.add_paragraph()
    if align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "left":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.first_line_indent = Cm(1.25) if first_ind else Cm(0)
    if isinstance(parts, str):
        parts = [(parts, bold, False, BLACK)]
    add_runs(p, parts, size=size)
    return p


def set_cell_width(cell, dxa):
    tcPr = cell._tc.get_or_add_tcPr()
    for el in tcPr.findall(qn("w:tcW")):
        tcPr.remove(el)
    tcPr.append(parse_xml(r'<w:tcW %s w:w="%d" w:type="dxa"/>' % (nsdecls("w"), dxa)))


def valign_top(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(r'<w:vAlign %s w:val="top"/>' % nsdecls("w")))


def create_cong_van():
    doc = docx.Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(14)
    style.font.color.rgb = BLACK
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), FONT)

    LEFT_DXA = 5159
    RIGHT_DXA = 5783
    TBL_IND = -850

    tbl = doc.add_table(rows=3, cols=2)
    set_table_borders_none(tbl)
    set_header_table_properties(tbl, LEFT_DXA, RIGHT_DXA, TBL_IND)

    c00 = tbl.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p00.paragraph_format.space_before = Pt(0)
    p00.paragraph_format.space_after = Pt(0)
    p00.paragraph_format.line_spacing = 1.0
    r = p00.add_run("THUẾ THÀNH PHỐ ĐỒNG NAI\n")
    set_run_font(r, size=13, bold=False)
    r = p00.add_run("THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI\n")
    set_run_font(r, size=13, bold=True)
    r = p00.add_run("─" * 12)
    set_run_font(r, size=10)

    c01 = tbl.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p01.paragraph_format.space_before = Pt(0)
    p01.paragraph_format.space_after = Pt(0)
    p01.paragraph_format.line_spacing = 1.0
    r = p01.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    set_run_font(r, size=13, bold=True)
    r = p01.add_run("Độc lập - Tự do - Hạnh phúc\n")
    set_run_font(r, size=14, bold=True)
    r = p01.add_run("─" * 24)
    set_run_font(r, size=10)

    c10 = tbl.cell(1, 0)
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p10.paragraph_format.left_indent = Cm(2)
    p10.paragraph_format.space_before = Pt(4)
    p10.paragraph_format.space_after = Pt(2)
    r = p10.add_run("Số:          ")
    set_run_font(r, size=13)

    c11 = tbl.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p11.paragraph_format.space_before = Pt(4)
    p11.paragraph_format.space_after = Pt(2)
    r = p11.add_run("Đồng Nai, ngày       tháng       năm 2026")
    set_run_font(r, size=13, italic=True)

    TY_DXA = 3402
    c20 = tbl.cell(2, 0)
    set_cell_width(c20, TY_DXA)
    p20 = c20.paragraphs[0]
    p20.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p20.paragraph_format.space_before = Pt(2)
    p20.paragraph_format.space_after = Pt(2)
    p20.paragraph_format.line_spacing = 1.15
    r = p20.add_run("V/v cung cấp thông tin, hồ sơ DN theo Công văn 1808/YC-ANĐT-Đ2.")
    set_run_font(r, size=13)

    pkg = doc.add_paragraph()
    pkg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pkg.paragraph_format.space_before = Pt(18)
    pkg.paragraph_format.space_after = Pt(12)
    pkg.paragraph_format.line_spacing = 1.15
    r = pkg.add_run("Kính gửi: Cơ quan An ninh điều tra - Công an thành phố Hải Phòng")
    set_run_font(r, size=14)

    add_body(
        doc,
        "Thuế cơ sở 3 thành phố Đồng Nai nhận được Công văn số 1808/YC-ANĐT-Đ2 ngày 09/04/2026 của Cơ quan An ninh điều tra - Công an thành phố Hải Phòng (cùng Công văn số 1821/YC-ANĐT-Đ2 cùng ngày; dấu đến Thuế cơ sở 3 thành phố Đồng Nai số 12242 ngày 07/09/2026) về việc đề nghị cung cấp thông tin, hồ sơ doanh nghiệp liên quan vụ án Ngô Văn Tới mua bán trái phép hóa đơn theo Quyết định số 32/QĐ-ANĐT-Đ2.",
    )
    add_body(doc, "Căn cứ dữ liệu quản lý thuế, Thuế cơ sở 3 thành phố Đồng Nai trả lời như sau:")

    # 1. Truong Anh Nguyen - DB NNT live
    add_body(
        doc,
        "1. CÔNG TY TNHH THƯƠNG MẠI TRƯỜNG ANH NGUYÊN, MST: 3603989685, NGÀY BẮT ĐẦU HOẠT ĐỘNG: 04/11/2024.",
        bold=True,
    )
    add_body(doc, "Thông tin cung cấp:")
    add_body(
        doc,
        "+ Địa điểm đăng ký kinh doanh: Số 31/8, khu phố 2A, tổ 23, Phường Long Bình, thành phố Đồng Nai;",
    )
    add_body(
        doc,
        "+ Giấy chứng nhận ĐKKD: Số 3603989685 do Sở Kế hoạch và Đầu tư tỉnh Đồng Nai cấp lần đầu ngày 04/11/2024;",
    )
    add_body(
        doc,
        "+ Tình trạng hoạt động: NNT không hoạt động tại địa chỉ đã đăng ký theo Thông báo số 1851/TB-CCTKV ngày 14/04/2025 (ngày đóng trạng thái trên hệ thống: 15/04/2025);",
    )
    add_body(
        doc,
        "+ Người đại diện theo pháp luật / Giám đốc: Lê Văn Trường Nhỏ (sinh ngày 01/01/1972; SĐT: 0778582675);",
    )
    add_body(
        doc,
        "+ Cán bộ quản lý thuế: Cán bộ Đoàn Thanh Tùng - Tổ Kiểm tra thuế số 1, Thuế cơ sở 3 thành phố Đồng Nai;",
    )
    add_body(doc, "+ Số tài khoản ngân hàng: NNT không đăng ký với Cơ quan Thuế;")
    add_body(
        doc,
        [
            ("+ Tình hình thực hiện nghĩa vụ thuế: ", False, False, BLACK),
            (
                "Chưa có kết xuất TMS hoặc sổ theo dõi nợ của đơn vị tại thời điểm lập dự thảo.",
                False,
                False,
                RED,
            ),
        ],
    )
    add_body(
        doc,
        "+ Tình hình thanh tra, kiểm tra, xử lý vi phạm về thuế: Công ty chưa bị xử phạt về thuế và hóa đơn.",
    )
    add_body(doc, "Hồ sơ, tài liệu cung cấp bao gồm:")
    add_body(doc, "+ Bản sao Giấy chứng nhận đăng ký doanh nghiệp cấp lần đầu ngày 04/11/2024;")
    add_body(
        doc,
        "+ Thông báo số 1851/TB-CCTKV ngày 14/04/2025 về việc NNT không hoạt động tại địa chỉ đã đăng ký;",
    )
    add_body(doc, "+ Tờ khai thuế GTGT từ quý 04/2024 đến quý 02/2025;")
    add_body(
        doc,
        "+ Bảng kê chi tiết hóa đơn điện tử mua vào - bán ra từ quý 04/2024 đến quý 02/2025.",
    )

    # 2. Shapemoon - khong co tren DB NNT 03.06.2026, fallback masothue 06/10/2026
    add_body(
        doc,
        "2. CÔNG TY TNHH XÂY DỰNG SHAPEMOON, MST: 3604014466, NGÀY BẮT ĐẦU HOẠT ĐỘNG: 18/04/2025.",
        bold=True,
    )
    add_body(doc, "Thông tin cung cấp:")
    add_body(
        doc,
        "+ Địa điểm đăng ký kinh doanh: Số 350, Tổ 14, Ấp 2, Xã Tân An, thành phố Đồng Nai;",
    )
    add_body(
        doc,
        [
            ("+ Giấy chứng nhận ĐKKD: Số 3604014466 cấp lần đầu ngày 18/04/2025", False, False, BLACK),
            (" (chưa đối chiếu bản GCN trên CSDL NNT);", False, False, RED),
        ],
    )
    add_body(
        doc,
        [
            (
                "+ Tình trạng hoạt động: NNT không hoạt động tại địa chỉ đã đăng ký",
                False,
                False,
                BLACK,
            ),
            (
                " (chưa có số/ngày Thông báo đóng MST trên CSDL nội bộ DB NNT.xls);",
                False,
                False,
                RED,
            ),
        ],
    )
    add_body(
        doc,
        [
            (
                "+ Người đại diện theo pháp luật / Giám đốc: Nguyễn Viết Huy Hoàng (SĐT: 0935235734)",
                False,
                False,
                BLACK,
            ),
            (" (nguồn masothue.com, chưa có trên DB NNT.xls);", False, False, RED),
        ],
    )
    add_body(
        doc,
        "+ Cán bộ quản lý thuế: Thuộc Tổ Kiểm tra thuế số 2 quản lý địa bàn Xã Tân An;",
    )
    add_body(
        doc,
        [
            ("+ Số tài khoản ngân hàng: ", False, False, BLACK),
            (
                "Chưa có trong CSDL NNT (doanh nghiệp không có trên DB NNT.xls).",
                False,
                False,
                RED,
            ),
        ],
    )
    add_body(
        doc,
        [
            ("+ Tình hình thực hiện nghĩa vụ thuế: ", False, False, BLACK),
            (
                "Chưa có kết xuất TMS hoặc sổ theo dõi nợ của đơn vị tại thời điểm lập dự thảo.",
                False,
                False,
                RED,
            ),
        ],
    )
    add_body(
        doc,
        "+ Tình hình thanh tra, kiểm tra, xử lý vi phạm về thuế: Công ty chưa bị xử phạt về thuế và hóa đơn.",
    )
    add_body(doc, "Hồ sơ, tài liệu cung cấp bao gồm:")
    add_body(doc, "+ Bản sao Giấy chứng nhận đăng ký doanh nghiệp cấp lần đầu ngày 18/04/2025;")
    add_body(
        doc,
        [
            (
                "+ Thông báo NNT không hoạt động tại địa chỉ đã đăng ký: ",
                False,
                False,
                BLACK,
            ),
            ("chưa có số/ngày Thông báo trên CSDL nội bộ;", False, False, RED),
        ],
    )
    add_body(doc, "+ Tờ khai thuế GTGT từ quý 02/2025 đến quý 03/2026;")
    add_body(
        doc,
        "+ Bảng kê chi tiết hóa đơn điện tử mua vào - bán ra từ quý 02/2025 đến quý 03/2026.",
    )

    add_body(
        doc,
        "Thuế cơ sở 3 thành phố Đồng Nai phúc đáp để Cơ quan An ninh điều tra - Công an thành phố Hải Phòng được biết và phối hợp công tác./.",
    )

    SIG_LEFT = 4252
    SIG_RIGHT = 5103
    tbl_sig = doc.add_table(rows=1, cols=2)
    set_table_borders_none(tbl_sig)
    set_header_table_properties(tbl_sig, SIG_LEFT, SIG_RIGHT, 0)

    c_s0 = tbl_sig.cell(0, 0)
    valign_top(c_s0)
    p_s0 = c_s0.paragraphs[0]
    p_s0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_s0.paragraph_format.space_before = Pt(6)
    p_s0.paragraph_format.space_after = Pt(0)
    p_s0.paragraph_format.line_spacing = 1.0
    r = p_s0.add_run("Nơi nhận:\n")
    set_run_font(r, size=12, bold=True, italic=True)
    for line in ("- Như trên;\n", "- BLĐ thuế cơ sở 3 (để b/c);\n", "- Lưu: VT, KT1 (Tr)."):
        r = p_s0.add_run(line)
        set_run_font(r, size=11)

    c_s1 = tbl_sig.cell(0, 1)
    valign_top(c_s1)
    p_s1 = c_s1.paragraphs[0]
    p_s1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_s1.paragraph_format.space_before = Pt(6)
    p_s1.paragraph_format.space_after = Pt(0)
    p_s1.paragraph_format.line_spacing = 1.0
    r = p_s1.add_run("KT. TRƯỞNG THUẾ CƠ SỞ\n")
    set_run_font(r, size=13, bold=True)
    r = p_s1.add_run("PHÓ TRƯỞNG THUẾ CƠ SỞ\n\n\n\n\n")
    set_run_font(r, size=13, bold=True)
    r = p_s1.add_run("Phạm Đồng Xuân")
    set_run_font(r, size=14, bold=True)
    return doc


def _sig_block(cell, lines):
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    for text, size, bold, italic in lines:
        r = p.add_run(text)
        set_run_font(r, size=size, bold=bold, italic=italic)


def create_to_trinh():
    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(1.0)

    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(13)
    style.font.color.rgb = BLACK

    tbl = doc.add_table(rows=1, cols=2)
    set_table_borders_none(tbl)
    set_header_table_properties(tbl, 4819, 5726, -850)

    c00 = tbl.cell(0, 0)
    p = c00.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI\n")
    set_run_font(r, size=12)
    r = p.add_run("TỔ KIỂM TRA SỐ 1\n")
    set_run_font(r, size=13, bold=True)
    r = p.add_run("─" * 12)
    set_run_font(r, size=10)

    c01 = tbl.cell(0, 1)
    p = c01.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
    set_run_font(r, size=13, bold=True)
    r = p.add_run("Độc lập - Tự do - Hạnh phúc\n")
    set_run_font(r, size=14, bold=True)
    r = p.add_run("─" * 23)
    set_run_font(r, size=10)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(6)
    r = p_title.add_run("TỜ TRÌNH LÃNH ĐẠO")
    set_run_font(r, size=14, bold=True)

    p_vv = doc.add_paragraph()
    p_vv.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_vv.paragraph_format.space_before = Pt(0)
    p_vv.paragraph_format.space_after = Pt(10)
    r = p_vv.add_run("Về việc: cung cấp thông tin, hồ sơ DN theo Công văn số 1808/YC-ANĐT-Đ2")
    set_run_font(r, size=13, bold=False)

    def nd(title, body):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(title)
        set_run_font(r, size=13, bold=True)
        r = p.add_run(body)
        set_run_font(r, size=13, bold=False)

    nd(
        "- Nội dung công việc: ",
        "Căn cứ Công văn số 1808/YC-ANĐT-Đ2 ngày 09/04/2026 của Cơ quan An ninh điều tra - Công an thành phố Hải Phòng (cùng Công văn số 1821/YC-ANĐT-Đ2 cùng ngày; dấu đến số 12242 ngày 07/09/2026) về việc đề nghị cung cấp thông tin, hồ sơ doanh nghiệp.",
    )
    nd(
        "- Ý kiến tham gia của các đơn vị liên quan: ",
        "Tổ Kiểm tra số 1 đã rà soát dữ liệu quản lý thuế và dự thảo công văn trả lời theo yêu cầu.",
    )
    nd(
        "- Đề xuất của cán bộ giải quyết: ",
        "Kính trình Ban Lãnh đạo Thuế cơ sở 3 thành phố Đồng Nai xem xét, phê duyệt và ký ban hành dự thảo công văn trả lời Cơ quan An ninh điều tra - Công an thành phố Hải Phòng.",
    )

    # Bang 3 hang: CB de xuat / Y kien To KTR1 / To truong
    tbl3 = doc.add_table(rows=3, cols=2)
    set_table_borders_none(tbl3)
    set_header_table_properties(tbl3, 4535, 4535, 0)

    _sig_block(
        tbl3.cell(0, 1),
        [
            ("Trấn Biên, ngày     tháng    năm 2026\n", 12, False, True),
            ("CÁN BỘ ĐỀ XUẤT\n\n\n\n", 13, True, False),
            ("Nguyễn Thị Ánh Ngọc", 13, True, False),
        ],
    )

    merged = tbl3.cell(1, 0).merge(tbl3.cell(1, 1))
    p = merged.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Cm(1.25)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(
        "Ý kiến của lãnh đạo Tổ KTR1: ............................................................................"
    )
    set_run_font(r, size=13)

    _sig_block(
        tbl3.cell(2, 1),
        [
            ("Trấn Biên, ngày     tháng    năm 2026\n", 12, False, True),
            ("TỔ TRƯỞNG\n\n\n\n", 13, True, False),
            ("Nguyễn Trường Giang", 13, True, False),
        ],
    )

    p_yk = doc.add_paragraph()
    p_yk.paragraph_format.first_line_indent = Cm(1.25)
    p_yk.paragraph_format.space_before = Pt(4)
    p_yk.paragraph_format.space_after = Pt(2)
    p_yk.paragraph_format.line_spacing = 1.15
    r = p_yk.add_run(
        "Ý kiến phê duyệt của Lãnh đạo TCS3: ................................................................"
    )
    set_run_font(r, size=13)

    tbl_ld = doc.add_table(rows=1, cols=2)
    set_table_borders_none(tbl_ld)
    set_header_table_properties(tbl_ld, 4535, 4535, 0)
    _sig_block(
        tbl_ld.cell(0, 1),
        [
            ("Trấn Biên, ngày     tháng    năm 2026\n", 12, False, True),
            ("KT. TRƯỞNG THUẾ CƠ SỞ\n", 12, True, False),
            ("PHÓ TRƯỞNG THUẾ CƠ SỞ\n\n\n\n", 12, True, False),
            ("Phạm Đồng Xuân", 13, True, False),
        ],
    )
    return doc


def save_all():
    cv_name = "TRA LOI CV 1808_CA HAI PHONG_du thao.docx"
    tt_name = "to trinh 1808.docx"
    doc_cv = create_cong_van()
    doc_tt = create_to_trinh()

    saved = []
    vault_cv = os.path.join(VAULT, cv_name)
    vault_tt = os.path.join(VAULT, tt_name)
    doc_cv.save(vault_cv)
    doc_tt.save(vault_tt)
    saved.append(vault_cv)
    saved.append(vault_tt)

    for d in DRIVE_CANDIDATES:
        if os.path.isdir(d):
            p1 = os.path.join(d, cv_name)
            p2 = os.path.join(d, tt_name)
            doc_cv.save(p1)
            doc_tt.save(p2)
            saved.append(p1)
            saved.append(p2)
            break
    return saved


if __name__ == "__main__":
    paths = save_all()
    for p in paths:
        print("SAVED", p, "exists", os.path.exists(p), "size", os.path.getsize(p) if os.path.exists(p) else 0)
