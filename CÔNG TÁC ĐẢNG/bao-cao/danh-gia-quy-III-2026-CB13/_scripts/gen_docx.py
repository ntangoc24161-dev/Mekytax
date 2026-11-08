# -*- coding: utf-8 -*-
"""Sinh cac van ban .docx cho ho so danh gia Quy III/2026 - Chi bo 13."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from cb13_data import *
from cb13_noidung import NHIEM_VU, ANH_XA, CHUNG, CHO_DIEN
import cb13_uoc_tinh as UT

DO_RGB = RGBColor(0xC0, 0x00, 0x00)
DO_TUP = (0xC0, 0x00, 0x00)
TNR = "Times New Roman"


def set_run_tnr(run):
    run.font.name = TNR
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), TNR)


def diem(n):
    """Tong diem uoc tinh cua 1 dong chi."""
    return 30 + UT.diem_nhom_b(n["stt"], TIEU_CHI_B)

OUT = sys.argv[1]


def new_doc(landscape=False):
    d = Document()
    st = d.styles["Normal"]
    st.font.name = TNR
    st.font.size = Pt(14)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), TNR)
    st.element.rPr.rFonts.set(qn("w:ascii"), TNR)
    st.element.rPr.rFonts.set(qn("w:hAnsi"), TNR)
    st.element.rPr.rFonts.set(qn("w:cs"), TNR)
    pf = st.paragraph_format
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.line_spacing = 1.2
    s = d.sections[0]
    if landscape:
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = Cm(29.7), Cm(21)
    s.top_margin, s.bottom_margin = Cm(2.0), Cm(2.0)
    s.left_margin, s.right_margin = Cm(3.0), Cm(1.5)
    return d


def p(doc, text="", bold=False, italic=False, size=13, align="left",
      before=0, after=0, indent=None, caps=False, color=None):
    par = doc.add_paragraph()
    par.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                     "right": WD_ALIGN_PARAGRAPH.RIGHT, "just": WD_ALIGN_PARAGRAPH.JUSTIFY}[align]
    par.paragraph_format.space_before = Pt(before)
    par.paragraph_format.space_after = Pt(after)
    if indent is not None:
        par.paragraph_format.first_line_indent = Cm(indent)
    r = par.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    set_run_tnr(r)
    if caps:
        r.font.all_caps = True
    if color:
        r.font.color.rgb = RGBColor(*color)
    return par


def header_dang(doc, so_vb="Số:", loai=""):
    """Header 3 dong mau 04. Khong dung table.columns."""
    from docx.oxml import OxmlElement
    t = doc.add_table(rows=3, cols=2)
    tbl = t._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl.insert(0, grid)
    for child in list(grid):
        grid.remove(child)
    for w in (5353, 4819):
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(w))
        grid.append(col)
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), "10172")
    tblW.set(qn("w:type"), "dxa")

    def cp(cell, text, bold=False, size=13, italic=False):
        par = cell.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_after = Pt(0)
        r = par.add_run(text)
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        set_run_tnr(r)

    r0, r1, r2 = t.rows[0], t.rows[1], t.rows[2]
    cp(r0.cells[0], DANG_BO, bold=True, size=12)
    cp(r0.cells[1], "ĐẢNG CỘNG SẢN VIỆT NAM", bold=True, size=13)
    cp(r1.cells[0], CHI_BO, bold=True, size=13)
    cp(r1.cells[1], "")
    cp(r2.cells[0], so_vb or "Số:", size=13)
    cp(r2.cells[1], "Trấn Biên, ngày ..... tháng ..... năm 2026", italic=True, size=13)
    for row in t.rows:
        tcs = row._tr.findall(qn("w:tc"))
        for tc, w in zip(tcs, (5353, 4819)):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.append(tcW)
            tcW.set(qn("w:w"), str(w))
            tcW.set(qn("w:type"), "dxa")
            va = tcPr.find(qn("w:vAlign"))
            if va is None:
                va = OxmlElement("w:vAlign")
                tcPr.append(va)
            va.set(qn("w:val"), "center")
    return t


def set_tbl_borders(tbl):
    from docx.oxml import OxmlElement
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "6")
        e.set(qn("w:space"), "0"); e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)


def table(doc, headers, rows, widths=None, size=11, hdr_size=11, red_cols=()):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_tbl_borders(t)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.paragraphs[0].add_run(h)
        r.bold = True; r.font.size = Pt(hdr_size); set_run_tnr(r)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            par = cells[i].paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = par.add_run(str(v))
            r.font.size = Pt(size); set_run_tnr(r)
            if i in red_cols:
                r.font.color.rgb = DO_RGB
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if widths:
        for ri in range(len(t.rows)):
            for i, w in enumerate(widths):
                t.rows[ri].cells[i].width = Cm(w)
    return t


def ky_ten(doc, chuc="T/M CHI BỘ 13", danh="BÍ THƯ", ten="Lê Hữu Nghĩa", noi_nhan=None):
    doc.add_paragraph()
    t = doc.add_table(rows=1, cols=2)
    c0, c1 = t.rows[0].cells
    c0.width, c1.width = Cm(8.0), Cm(8.0)
    par = c0.paragraphs[0]
    r = par.add_run("Nơi nhận:"); r.bold = True; r.italic = True; r.font.size = Pt(11)
    set_run_tnr(r)
    for nn in (noi_nhan or []):
        pp = c0.add_paragraph(); pp.paragraph_format.space_after = Pt(0)
        rr = pp.add_run("- " + nn + ";"); rr.font.size = Pt(11); set_run_tnr(rr)

    def cc(text, bold=False, italic=False, size=13, first=False):
        pp = c1.paragraphs[0] if first else c1.add_paragraph()
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.space_after = Pt(0)
        rr = pp.add_run(text); rr.bold = bold; rr.italic = italic
        rr.font.size = Pt(size); set_run_tnr(rr)
    cc(chuc, bold=True, first=True)
    cc(danh, bold=True)
    cc("(Ký, ghi rõ họ tên và đóng dấu)", italic=True, size=11)
    cc(""); cc(""); cc("")
    cc(ten, bold=True)


CAN_CU = [
 "Căn cứ Quy định số 84-QĐ/TU ngày 24/11/2025 của Ban Thường vụ Thành ủy Đồng Nai về đánh giá, xếp loại chất lượng định kỳ đối với cán bộ lãnh đạo, quản lý;",
 "Căn cứ Hướng dẫn số 07-HD/TU ngày 19/6/2026 của Ban Thường vụ Thành ủy Đồng Nai về đánh giá, xếp loại chất lượng định kỳ đối với cán bộ lãnh đạo, quản lý;",
 "Căn cứ Công văn số 2951-CV/BTCTU ngày 12/9/2026 của Ban Tổ chức Thành ủy;",
 "Căn cứ Công văn số 4988-CV/ĐUUBND ngày 09/10/2026 của Đảng ủy Ủy ban nhân dân thành phố Đồng Nai về đánh giá, xếp loại chất lượng định kỳ quý III/2026 đối với cán bộ lãnh đạo, quản lý;",
 "Căn cứ Quy chế làm việc của Chi ủy Chi bộ 13 nhiệm kỳ 2025-2030 (ban hành kèm theo Quyết định số 05-QĐ/CB ngày 10/2/2026) và phân công nhiệm vụ đối với các đồng chí chi ủy viên;",
]

NOI_NHAN = ["Ban Thường vụ Đảng ủy Thuế thành phố Đồng Nai", "Văn phòng Đảng ủy Thuế thành phố",
            "Ủy ban Kiểm tra Đảng ủy Thuế thành phố", "Lưu Chi bộ 13"]


# ===================== 01. TO TRINH =====================
def to_trinh():
    d = new_doc()
    header_dang(d, so_vb="Số:        -TTr/CB13")
    p(d, "", after=6)
    p(d, "TỜ TRÌNH", bold=True, size=14, align="center")
    p(d, "về việc đề xuất mức xếp loại chất lượng đối với cán bộ", bold=True, align="center")
    p(d, "thuộc diện Ban Thường vụ Đảng ủy Thuế thành phố quản lý quý III/2026", bold=True, align="center")
    p(d, "-----", align="center", after=10)
    p(d, "Kính gửi: Ban Thường vụ Đảng ủy Thuế thành phố Đồng Nai.", bold=True, align="center", after=10)
    for c in CAN_CU:
        p(d, c, align="just", indent=1.0, after=4)
    p(d, "Chi bộ 13 báo cáo và đề xuất mức xếp loại chất lượng quý III/2026 đối với các đồng chí "
         "trong cấp ủy Chi bộ 13 thuộc diện Ban Thường vụ Đảng ủy Thuế thành phố quản lý như sau:",
      align="just", indent=1.0, after=8)

    p(d, "1. Khái quát tình hình tổ chức thực hiện", bold=True, after=4)
    p(d, "Chi bộ 13 trực thuộc Đảng bộ Thuế thành phố Đồng Nai, tính đến ngày 02/11/2026 có 109 đảng viên "
         "đang công tác, được phân thành 03 tổ đảng (Tổ đảng 1: 43 đồng chí; Tổ đảng 2: 30 đồng chí; "
         "Tổ đảng 3: 36 đồng chí). Cấp ủy Chi bộ gồm 07 đồng chí: 01 Bí thư, 01 Phó Bí thư và 05 Chi ủy viên.",
      align="just", indent=1.0, after=4)
    p(d, "Thực hiện chỉ đạo của cấp trên, Chi bộ đã ban hành quyết định phân công nhiệm vụ và danh mục "
         "sản phẩm, công việc quý III/2026 đối với từng đồng chí chi ủy viên; tổ chức cho từng cá nhân "
         "xây dựng kế hoạch công tác quý và bản tự đánh giá, xếp loại theo khung chấm điểm 30 điểm tiêu chí "
         "chung và 70 điểm kết quả thực hiện nhiệm vụ; chi ủy, chi bộ đã họp nhận xét, chấm điểm và đề xuất "
         "mức xếp loại đối với từng đồng chí.", align="just", indent=1.0, after=8)

    p(d, "2. Kết quả thực hiện nhiệm vụ và đề xuất mức xếp loại", bold=True, after=4)
    rows = []
    for n in NHAN_SU:
        rows.append([n["stt"], n["ten_hoa_thuong"], n["cv_dang"], n["linh_vuc_ngan"],
                     UT.so(diem(n)), XEP_LOAI])
    table(d, ["TT", "Họ và tên", "Chức vụ Đảng", "Lĩnh vực được phân công phụ trách",
              "Tổng điểm (ước tính)", "Mức xếp loại đề xuất"], rows,
          widths=[0.9, 3.0, 3.8, 4.2, 2.1, 3.0], size=10, hdr_size=10, red_cols=(4,))
    p(d, "", after=4)
    p(d, "Trong quý III/2026, 07/07 đồng chí trong cấp ủy Chi bộ 13 hoàn thành 100% nhiệm vụ, chỉ tiêu "
         "được giao theo kế hoạch công tác quý và danh mục sản phẩm, công việc đã được cá thể hóa; "
         "không có đồng chí nào bị xử lý kỷ luật, không có biểu hiện suy thoái về tư tưởng chính trị, "
         "đạo đức, lối sống. Trên cơ sở kết quả chấm điểm theo khung 30 + 70, Chi bộ đề xuất xếp loại "
         "cả 07 đồng chí ở mức " + XEP_LOAI + ".",
      align="just", indent=1.0, after=4)
    dmin = min(diem(n) for n in NHAN_SU)
    dmax = max(diem(n) for n in NHAN_SU)
    p(d, "Tổng điểm của 07 đồng chí trong khoảng %s đến %s điểm trên thang 100 điểm "
         "(30 điểm nhóm tiêu chí chung và %s đến %s điểm nhóm kết quả thực hiện nhiệm vụ). "
         "Mức đề xuất là hoàn thành tốt nhiệm vụ: hoàn thành 100%% nhiệm vụ được giao, "
         "không đặt tiêu chí vượt mức. Các số liệu điểm nêu trên là SỐ ƯỚC TÍNH, "
         "Chi bộ thay bằng kết quả chấm điểm thực tế tại bản tự đánh giá của từng đồng chí "
         "trước khi trình ký."
         % (UT.so(dmin), UT.so(dmax), UT.so(dmin - 30), UT.so(dmax - 30)),
      align="just", indent=1.0, after=8, color=DO_TUP, italic=True)

    p(d, "3. Đề nghị", bold=True, after=4)
    p(d, "Chi bộ 13 kính đề nghị Ban Thường vụ Đảng ủy Thuế thành phố Đồng Nai xem xét, quyết định "
         "mức xếp loại chất lượng quý III/2026 đối với 07 đồng chí trong cấp ủy Chi bộ 13 nêu trên.",
      align="just", indent=1.0, after=4)
    p(d, "Hồ sơ kèm theo: (1) Danh sách trích ngang; (2) Biểu tổng hợp kết quả thực hiện nhiệm vụ và "
         "đề xuất xếp loại; (3) Quyết định ban hành danh mục sản phẩm, công việc quý III/2026 và danh mục "
         "của từng đồng chí; (4) Kế hoạch công tác quý III/2026 của từng cá nhân; (5) Bản tự đánh giá, "
         "xếp loại của từng cá nhân; (6) Báo cáo kết quả đánh giá, xếp loại quý III/2026; "
         "(7) Biểu thống kê kết quả đánh giá, xếp loại cấp ủy viên.", align="just", indent=1.0, after=4)
    p(d, "Chi bộ 13 trân trọng báo cáo./.", align="just", indent=1.0)
    ky_ten(d, noi_nhan=NOI_NHAN)
    d.save(os.path.join(OUT, "01_To_trinh_de_xuat_xep_loai.docx"))


# ===================== 04. QUYET DINH + DANH MUC =====================
def quyet_dinh():
    d = new_doc()
    header_dang(d, so_vb="Số:        -QĐ/CB13")
    p(d, "", after=6)
    p(d, "QUYẾT ĐỊNH", bold=True, size=14, align="center")
    p(d, "về việc ban hành danh mục sản phẩm, công việc và phân công nhiệm vụ", bold=True, align="center")
    p(d, "đối với các đồng chí trong cấp ủy Chi bộ 13 quý III/2026", bold=True, align="center")
    p(d, "-----", align="center", after=10)
    p(d, "CHI ỦY CHI BỘ 13", bold=True, align="center", after=6)
    for c in CAN_CU[:4]:
        p(d, "- " + c, align="just", indent=1.0, after=4)
    p(d, "- Căn cứ Nghị quyết số 05-NQ/TU ngày 01/5/2026 của Ban Chấp hành Đảng bộ thành phố Đồng Nai "
         "về phát động đợt thi đua đặc biệt 500 ngày đêm;", align="just", indent=1.0, after=4)
    p(d, "- Xét đề nghị của đồng chí Phó Bí thư Chi bộ phụ trách công tác tuyên giáo,",
      align="just", indent=1.0, after=8)
    p(d, "QUYẾT ĐỊNH:", bold=True, align="center", after=8)

    p(d, "Điều 1. Ban hành kèm theo Quyết định này Danh mục sản phẩm, công việc quý III/2026 đã được "
         "cá thể hóa đối với 07 đồng chí trong cấp ủy Chi bộ 13 (Phụ lục kèm theo).",
      align="just", indent=1.0, after=6)
    p(d, "Điều 2. Phân công lĩnh vực phụ trách đối với các đồng chí trong cấp ủy Chi bộ 13 quý III/2026 "
         "như sau:", align="just", indent=1.0, after=6)
    for n in NHAN_SU:
        p(d, "%d. Đồng chí %s - %s: %s" % (n["stt"], n["ten_hoa_thuong"], n["cv_dang"], n["linh_vuc"]),
          align="just", indent=1.0, after=5)
    p(d, "Điều 3. Các đồng chí chi ủy viên căn cứ danh mục sản phẩm, công việc và lĩnh vực được phân công "
         "để xây dựng kế hoạch công tác quý III/2026 của cá nhân, tổ chức thực hiện và tự đánh giá, "
         "xếp loại theo quy định. Đồng chí Phó Bí thư Chi bộ và đồng chí phụ trách công tác đảng vụ "
         "có trách nhiệm theo dõi, đôn đốc, tổng hợp kết quả báo cáo chi ủy.", align="just", indent=1.0, after=6)
    p(d, "Điều 4. Quyết định này có hiệu lực kể từ ngày ký. Chi ủy viên Chi bộ 13 và các tổ đảng "
         "chịu trách nhiệm thi hành Quyết định này./.", align="just", indent=1.0)
    ky_ten(d, chuc="T/M CHI ỦY CHI BỘ 13", noi_nhan=NOI_NHAN)

    # ---- Phu luc danh muc san pham ----
    d.add_page_break()
    p(d, "PHỤ LỤC", bold=True, size=14, align="center")
    p(d, "DANH MỤC SẢN PHẨM, CÔNG VIỆC QUÝ III/2026 ĐÃ CÁ THỂ HÓA", bold=True, align="center")
    p(d, "ĐỐI VỚI CÁC ĐỒNG CHÍ TRONG CẤP ỦY CHI BỘ 13", bold=True, align="center")
    p(d, "(Kèm theo Quyết định số        -QĐ/CB13 ngày ..... tháng ..... năm 2026 của Chi ủy Chi bộ 13)",
      italic=True, size=11, align="center", after=6)
    p(d, "Hệ số quy đổi áp dụng theo Phụ lục 2A: 1,0 - công việc chuẩn (100 điểm quy đổi); "
         "1,2 - báo cáo, nhiệm vụ định kỳ phức tạp hơn; 1,5 - tờ trình, dự thảo văn bản trình cấp ủy; "
         "2,0 - nhiệm vụ trọng tâm, quy định, bộ tiêu chí.", italic=True, size=10, align="just", after=8)

    for n in NHAN_SU:
        p(d, "%d. Đồng chí %s - %s" % (n["stt"], n["ten_hoa_thuong"], n["cv_dang"]),
          bold=True, after=4)
        rows = []
        tong = 0.0
        for i, nv in enumerate(NHIEM_VU[n["ten"]], 1):
            noi_dung, loai, chi_tieu, tien_do, san_pham, phoi_hop, cap_trinh, thoi_han, sl, hs = nv
            diem = round(sl * hs * 100)
            tong += diem
            rows.append([i, noi_dung, san_pham, "%d" % sl,
                         ("%.1f" % hs).replace(".", ","), "%d" % diem, thoi_han])
        rows.append(["", "TỔNG ĐIỂM QUY ĐỔI KẾ HOẠCH", "", "", "", "%d" % tong, ""])
        table(d, ["TT", "Nhiệm vụ, công việc được giao", "Sản phẩm cụ thể", "Số lượng",
                  "Hệ số", "Điểm quy đổi KH", "Thời hạn"], rows,
              widths=[0.9, 5.0, 4.0, 1.2, 1.0, 1.6, 2.3], size=9, hdr_size=9)
        p(d, "", after=8)
    d.save(os.path.join(OUT, "04_Quyet_dinh_phan_cong_va_danh_muc_san_pham.docx"))


# ===================== 05. KE HOACH CA NHAN =====================
def ke_hoach_ca_nhan():
    sub = os.path.join(OUT, "05_Ke_hoach_ca_nhan")
    os.makedirs(sub, exist_ok=True)
    import unicodedata
    for n in NHAN_SU:
        d = new_doc(landscape=True)
        header_dang(d)
        p(d, "", after=6)
        p(d, "KẾ HOẠCH CÔNG TÁC QUÝ III/2026", bold=True, size=14, align="center")
        p(d, "-----", align="center", after=2)
        p(d, "(Chi bộ 13 ban hành kèm theo Quyết định số        -QĐ/CB13 "
             "ngày ..... tháng ..... năm 2026)", italic=True, size=11, align="center", after=8)
        p(d, "Họ và tên: %s                     Ngày sinh: %s" % (n["ten_hoa_thuong"], n["ns"]), after=3)
        p(d, "Chức vụ Đảng: %s" % n["cv_dang"], after=3)
        p(d, "Chức vụ chính quyền: %s" % (n["cv_cq"] or "Không"), after=3)
        p(d, "Chức vụ đoàn thể: %s" % (n["cv_dt"] or "Không"), after=3)
        p(d, "Đơn vị công tác: Chi bộ 13 - Đảng bộ Thuế thành phố Đồng Nai", after=3)
        p(d, "Lĩnh vực được phân công phụ trách: %s" % n["linh_vuc"], align="just", after=8)
        rows = []
        for i, nv in enumerate(NHIEM_VU[n["ten"]], 1):
            noi_dung, loai, chi_tieu, tien_do, san_pham, phoi_hop, cap_trinh, thoi_han, sl, hs = nv
            rows.append([i, noi_dung, loai, chi_tieu, tien_do, san_pham,
                         n["ten_hoa_thuong"], phoi_hop, cap_trinh, thoi_han])
        table(d, ["TT", "Nội dung nhiệm vụ", "Phân loại nhiệm vụ", "Chỉ tiêu, yêu cầu thực hiện",
                  "Tiến độ thực hiện", "Sản phẩm công việc cụ thể", "Cá nhân chủ trì",
                  "Đơn vị/cá nhân phối hợp", "Cấp trình/Thẩm quyền", "Thời hạn hoàn thành"],
              rows, widths=[0.8, 5.2, 1.8, 4.6, 1.8, 4.2, 2.0, 2.6, 2.4, 2.0], size=9, hdr_size=9)
        p(d, "", after=6)
        d.add_paragraph()
        t = d.add_table(rows=1, cols=2)
        c0, c1 = t.rows[0].cells

        def cc(cell, text, bold=False, italic=False, size=12, first=False):
            pp = cell.paragraphs[0] if first else cell.add_paragraph()
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.paragraph_format.space_after = Pt(0)
            rr = pp.add_run(text); rr.bold = bold; rr.italic = italic
            rr.font.size = Pt(size); set_run_tnr(rr)
        cc(c0, "PHÊ DUYỆT CỦA CẤP CÓ THẨM QUYỀN", bold=True, first=True)
        cc(c0, n["xac_nhan"], bold=True)
        cc(c0, "(Ký, ghi rõ họ tên)", italic=True, size=11)
        cc(c0, ""); cc(c0, ""); cc(c0, "")
        cc(c0, "Lê Hữu Nghĩa" if n["stt"] != 1 else "", bold=True)
        cc(c1, "CÁ NHÂN THỰC HIỆN", bold=True, first=True)
        cc(c1, "(Ký, ghi rõ họ tên)", italic=True, size=11)
        cc(c1, ""); cc(c1, ""); cc(c1, ""); cc(c1, "")
        cc(c1, n["ten_hoa_thuong"], bold=True)
        slug = unicodedata.normalize("NFD", n["ten_hoa_thuong"])
        slug = "".join(ch for ch in slug if unicodedata.category(ch) != "Mn")
        slug = slug.replace("Đ", "D").replace("đ", "d").replace(" ", "-")
        d.save(os.path.join(sub, "KH_QIII-2026_%d_%s.docx" % (n["stt"], slug)))


# ===================== 07. BAO CAO KET QUA =====================
def bao_cao():
    d = new_doc()
    header_dang(d, so_vb="Số:        -BC/CB13")
    p(d, "", after=6)
    p(d, "BÁO CÁO", bold=True, size=14, align="center")
    p(d, "kết quả đánh giá, xếp loại chất lượng định kỳ quý III/2026", bold=True, align="center")
    p(d, "đối với các đồng chí trong cấp ủy Chi bộ 13", bold=True, align="center")
    p(d, "-----", align="center", after=10)

    p(d, "I. TÌNH HÌNH, KẾT QUẢ THỰC HIỆN", bold=True, after=5)
    p(d, "1. Công tác lãnh đạo, chỉ đạo", bold=True, after=4)
    p(d, "Chi ủy Chi bộ 13 đã tổ chức quán triệt đến toàn thể chi ủy viên và đảng viên các văn bản: "
         "Quy định số 84-QĐ/TU ngày 24/11/2025 của Ban Thường vụ Thành ủy; Hướng dẫn số 07-HD/TU ngày "
         "19/6/2026 của Ban Thường vụ Thành ủy; Công văn số 2951-CV/BTCTU ngày 12/9/2026 của Ban Tổ chức "
         "Thành ủy và Công văn số 4988-CV/ĐUUBND ngày 09/10/2026 về đánh giá, xếp loại chất lượng định kỳ "
         "quý III/2026.", align="just", indent=1.0, after=4)
    p(d, "Chi ủy đã ban hành quyết định phân công nhiệm vụ và danh mục sản phẩm, công việc quý III/2026 "
         "đã cá thể hóa đối với 07 đồng chí chi ủy viên; chỉ đạo từng đồng chí xây dựng kế hoạch công tác "
         "quý và bản tự đánh giá theo khung chấm điểm 30 + 70; phương pháp tính điểm KPI thực hiện theo "
         "Phụ lục 2A và 2B.", align="just", indent=1.0, after=4)
    p(d, "Đồng chí Chi ủy viên phụ trách công tác kiểm tra, giám sát được phân công theo dõi, giám sát "
         "quá trình đánh giá, xếp loại, bảo đảm thực chất, đúng quy trình, không nể nang, hình thức.",
      align="just", indent=1.0, after=6)

    p(d, "2. Kết quả thực hiện", bold=True, after=4)
    p(d, "2.1. Về tổ chức thực hiện", bold=True, italic=True, after=3)
    p(d, "100% cá nhân thuộc diện đánh giá đã hoàn thành kế hoạch công tác quý III/2026 và bản tự đánh giá, "
         "xếp loại theo mẫu. Chi bộ tổ chức hội nghị đánh giá, xếp loại theo hình thức họp trực tiếp; "
         "trình tự: cá nhân trình bày bản tự đánh giá - tập thể chi bộ tham gia ý kiến - chi ủy nhận xét, "
         "chấm điểm - biểu quyết đề xuất mức xếp loại.", align="just", indent=1.0, after=4)
    p(d, "Tập thể chi ủy và đồng chí Bí thư Chi bộ đã phát huy vai trò trong việc nhận xét, đánh giá; "
         "kết quả được tổng hợp, thẩm định và đề xuất trình Ban Thường vụ Đảng ủy Thuế thành phố "
         "xem xét, quyết định.", align="just", indent=1.0, after=5)

    p(d, "2.2. Kết quả đánh giá, xếp loại", bold=True, italic=True, after=3)
    p(d, "* Kết quả đề xuất đánh giá, xếp loại đối với các đồng chí trong cấp ủy Chi bộ 13:",
      italic=True, indent=1.0, after=3)
    rows = [[n["stt"], n["ten_hoa_thuong"], n["cv_dang"], UT.so(diem(n)), XEP_LOAI]
            for n in NHAN_SU]
    table(d, ["TT", "Họ và tên", "Chức vụ Đảng", "Tổng điểm (ước tính)",
              "Mức xếp loại đề xuất"], rows,
          widths=[0.9, 3.2, 5.4, 2.1, 3.4], size=10, hdr_size=10, red_cols=(3,))
    p(d, "", after=4)
    p(d, "Tổng hợp: 07/07 đồng chí được đề xuất " + XEP_LOAI + " (đạt 100%); không có đồng chí "
         "hoàn thành xuất sắc nhiệm vụ, hoàn thành nhiệm vụ hoặc không hoàn thành nhiệm vụ.",
      align="just", indent=1.0, after=4)
    ds = [diem(n) for n in NHAN_SU]
    p(d, "Điểm bình quân của cấp ủy là %s điểm, thấp nhất %s điểm, cao nhất %s điểm. "
         "Đây là SỐ ƯỚC TÍNH theo dạng hình số liệu của bộ hồ sơ mẫu, Chi bộ thay bằng kết quả "
         "chấm điểm thực tế trước khi trình ký."
         % (UT.so(round(sum(ds) / len(ds), 2)), UT.so(min(ds)), UT.so(max(ds))),
      align="just", indent=1.0, after=4, color=DO_TUP, italic=True)
    p(d, "* Kết quả đánh giá, xếp loại cấp ủy viên thuộc diện Chi bộ quản lý: Chi bộ 13 không có "
         "đảng bộ bộ phận trực thuộc; cấp ủy viên của Chi bộ đều thuộc diện Ban Thường vụ Đảng ủy "
         "Thuế thành phố quản lý nên không phát sinh nội dung này.", italic=True, align="just",
      indent=1.0, after=4)
    p(d, "Trong quý III/2026 không có trường hợp không hoàn thành nhiệm vụ.", align="just",
      indent=1.0, after=5)

    p(d, "2.3. Nhận xét chung về chất lượng đội ngũ cán bộ lãnh đạo, quản lý quý III/2026",
      bold=True, italic=True, after=3)
    p(d, "Các đồng chí trong cấp ủy Chi bộ 13 có bản lĩnh chính trị vững vàng, phẩm chất đạo đức, "
         "lối sống trong sáng, chấp hành nghiêm Điều lệ Đảng, pháp luật của Nhà nước và quy định của "
         "cơ quan; có tinh thần trách nhiệm, chủ động trong lĩnh vực được phân công; cơ bản đáp ứng "
         "yêu cầu nhiệm vụ của Chi bộ có quy mô 109 đảng viên.", align="just", indent=1.0, after=6)

    p(d, "II. ĐÁNH GIÁ CHUNG", bold=True, after=5)
    p(d, "1. Ưu điểm và nguyên nhân", bold=True, after=3)
    for s in ["Việc phân công nhiệm vụ đã được cá thể hóa đến từng đồng chí chi ủy viên, gắn với sản phẩm "
              "cụ thể và hệ số quy đổi, tạo cơ sở định lượng cho việc chấm điểm.",
              "Quy trình đánh giá bảo đảm dân chủ, công khai; kết quả được tập thể chi bộ thảo luận và "
              "biểu quyết.",
              "Hồ sơ đánh giá được lập đầy đủ theo thành phần yêu cầu, có minh chứng kèm theo."]:
        p(d, "- " + s, align="just", indent=1.0, after=3)
    p(d, "Nguyên nhân: chi ủy quan tâm chỉ đạo, phân công rõ người, rõ việc, rõ sản phẩm; đội ngũ "
         "chi ủy viên có trình độ chuyên môn và lý luận chính trị đáp ứng yêu cầu.",
      align="just", indent=1.0, after=5)
    p(d, "2. Hạn chế, khuyết điểm và nguyên nhân", bold=True, after=3)
    for s in ["Việc xây dựng kế hoạch công tác quý và danh mục sản phẩm của một số đồng chí còn chậm "
              "so với thời điểm đầu quý.",
              "Việc lượng hóa chỉ tiêu theo 04 trụ (số lượng, chất lượng, tiến độ, năng lực lãnh đạo) "
              "còn có nội dung định tính, phụ thuộc nhận xét của tập thể.",
              "Việc lưu giữ minh chứng đối với một số nhiệm vụ thường xuyên chưa đầy đủ, hệ thống."]:
        p(d, "- " + s, align="just", indent=1.0, after=3)
    p(d, "Nguyên nhân: đây là quý đầu áp dụng khung chấm điểm theo Quy định 84-QĐ/TU và Hướng dẫn "
         "07-HD/TU đối với chi bộ; khối lượng công việc chuyên môn trong quý lớn.",
      align="just", indent=1.0, after=5)
    p(d, "3. Khó khăn, vướng mắc", bold=True, after=3)
    for s in ["Chi bộ có 109 đảng viên, quy mô lớn, trong khi chi ủy chỉ có 07 đồng chí, đều kiêm nhiệm "
              "công tác chuyên môn nên khó bố trí thời gian cho công tác Đảng.",
              "Một số tiêu chí trong khung 70 điểm được thiết kế cho cấp ủy có đơn vị trực thuộc, "
              "khi áp dụng cho chi bộ không có đảng bộ bộ phận thì khó xác định sản phẩm tương ứng.",
              "Thời điểm nhận văn bản chỉ đạo sát với thời hạn gửi hồ sơ, gây áp lực về tiến độ."]:
        p(d, "- " + s, align="just", indent=1.0, after=3)
    p(d, "", after=4)
    p(d, "III. ĐỀ XUẤT, KIẾN NGHỊ", bold=True, after=4)
    for s in ["Đề nghị Ban Thường vụ Đảng ủy Thuế thành phố ban hành văn bản hướng dẫn thống nhất cách "
              "xác định sản phẩm và hệ số quy đổi đối với các tiêu chí mang tính định tính, để các chi bộ "
              "áp dụng đồng nhất.",
              "Đề nghị Đảng ủy Thuế thành phố chỉ đạo việc giao danh mục sản phẩm, công việc từ đầu quý "
              "để cá nhân có cơ sở xây dựng kế hoạch ngay từ tháng đầu tiên của quý.",
              "Đề nghị tổ chức tập huấn phương pháp tính điểm KPI theo Phụ lục 2A và 2B cho chi ủy "
              "các chi bộ trực thuộc."]:
        p(d, "- " + s, align="just", indent=1.0, after=3)
    p(d, "", after=3)
    p(d, "Chi bộ 13 báo cáo Ban Thường vụ Đảng ủy Thuế thành phố Đồng Nai./.", align="just", indent=1.0)
    ky_ten(d, noi_nhan=NOI_NHAN)
    d.save(os.path.join(OUT, "07_Bao_cao_ket_qua_QIII_2026.docx"))


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    to_trinh(); print("01 to trinh OK")
    quyet_dinh(); print("04 quyet dinh OK")
    ke_hoach_ca_nhan(); print("05 ke hoach ca nhan OK")
    bao_cao(); print("07 bao cao OK")
