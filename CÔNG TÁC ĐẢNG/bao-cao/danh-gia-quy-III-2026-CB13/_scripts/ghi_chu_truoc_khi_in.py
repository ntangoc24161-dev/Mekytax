# -*- coding: utf-8 -*-
"""Ghi chu Word: file nao in / khong in, de theo doi truoc khi nop."""
import os, shutil
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

VAULT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
LAPTOP = r"H:\Other computers\My Laptop\2. CONG TAC KHAC\4. CONG TAC DANG\A PHUONG\2026\06_Tieu_chi_danh_gia\Du_thao_CB13_QuyIII_2026"
TNR = "Times New Roman"
DO = RGBColor(0xC0, 0x00, 0x00)
FN = "00_GHI_CHU_TRUOC_KHI_IN.docx"


def set_run_tnr(run):
    run.font.name = TNR
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), TNR)


def p(doc, text="", bold=False, italic=False, size=13, align="left",
      before=0, after=4, color=None):
    par = doc.add_paragraph()
    par.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "just": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }[align]
    par.paragraph_format.space_before = Pt(before)
    par.paragraph_format.space_after = Pt(after)
    par.paragraph_format.line_spacing = 1.15
    r = par.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    set_run_tnr(r)
    if color:
        r.font.color.rgb = color
    return par


def set_borders(tbl):
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "6")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)


def table(doc, headers, rows, widths, hdr_fill=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_borders(t)
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.paragraphs[0].add_run(h)
        r.bold = True
        r.font.size = Pt(11)
        set_run_tnr(r)
        if hdr_fill:
            shd = OxmlElement("w:shd")
            shd.set(qn("w:fill"), hdr_fill)
            shd.set(qn("w:val"), "clear")
            c._tc.get_or_add_tcPr().append(shd)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            par = cells[i].paragraphs[0]
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
            r = par.add_run(str(v))
            r.font.size = Pt(11)
            set_run_tnr(r)
    for ri in range(len(t.rows)):
        for i, w in enumerate(widths):
            t.rows[ri].cells[i].width = Cm(w)
    return t


def build():
    d = Document()
    st = d.styles["Normal"]
    st.font.name = TNR
    st.font.size = Pt(13)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), TNR)
    st.element.rPr.rFonts.set(qn("w:ascii"), TNR)
    st.element.rPr.rFonts.set(qn("w:hAnsi"), TNR)
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin, s.bottom_margin = Cm(2.0), Cm(2.0)
    s.left_margin, s.right_margin = Cm(2.5), Cm(2.0)

    p(d, "GHI CHÚ THEO DÕI TRƯỚC KHI IN", bold=True, size=16, align="center", after=2)
    p(d, "Hồ sơ đánh giá, xếp loại chất lượng quý III/2026 - Chi bộ 13",
      bold=True, size=13, align="center", after=2)
    p(d, "File này là ghi chú nội bộ. KHÔNG in, KHÔNG nộp Đảng ủy Thuế thành phố.",
      italic=True, size=12, align="center", color=DO, after=8)
    p(d, "Ngày lập: 03/11/2026. Nơi nộp chính thức: Ban Thường vụ Đảng ủy Thuế thành phố Đồng Nai. "
         "Mức đề xuất: Hoàn thành tốt nhiệm vụ cho 07/07 đồng chí cấp ủy.",
      align="just", after=8)

    p(d, "1. KHÔNG IN, KHÔNG NỘP", bold=True, size=14, before=6, after=6)
    p(d, "Những file và sheet dưới đây Javis tự tạo để theo dõi, tính điểm, sinh hồ sơ. "
         "Bỏ ra khỏi bộ in.", align="just")
    table(d,
          ["STT", "File / sheet", "Vai trò", "Cách xử lý khi in"],
          [
              ["1", "00_GHI_CHU_TRUOC_KHI_IN.docx (file này)",
               "Ghi chú theo dõi trước khi in", "Không in"],
              ["2", "00_CHECKLIST_HO_SO.md",
               "Checklist nội bộ, giả định, rủi ro", "Không in"],
              ["3", "Thư mục _scripts/ (kể cả __pycache__)",
               "Script sinh Word/Excel", "Không in, không copy sang bộ nộp"],
              ["4", "Sheet Tinh diem KPI trong 03_Bang_tong_hop_de_xuat.xlsx",
               "Máy tính điểm sống 7 đồng chí",
               "Không in sheet này; chỉ in sheet Tong hop de xuat"],
              ["5", "Sheet Tinh KPI 2A-2B trong 7 file TDG_*.xlsx",
               "Máy tính 4 trụ A/B/C/D",
               "Không in sheet này; chỉ in sheet Khung đánh giá"],
          ],
          [1.2, 5.5, 4.5, 5.0],
          hdr_fill="F4B183")
    p(d, "", after=2)
    p(d, "Hai sheet tính điểm vẫn để trong file Excel để sửa số thì điểm tự chạy. "
         "Sheet có dòng đỏ: SHEET THEO DÕI NỘI BỘ - KHÔNG NỘP CẤP TRÊN.",
      italic=True, size=12, after=10)

    p(d, "2. IN VÀ NỘP - 8 SẢN PHẨM", bold=True, size=14, before=4, after=6)
    p(d, "Đúng danh mục Công văn 4988, nơi nhận đã đổi sang Đảng ủy Thuế thành phố Đồng Nai. "
         "Font Times New Roman toàn văn. Cơ quan ban hành: Chi bộ 13.",
      align="just")
    table(d,
          ["Mã", "Sản phẩm", "File in", "Lưu ý khi in"],
          [
              ["A1", "Tờ trình đề xuất xếp loại",
               "01_To_trinh_de_xuat_xep_loai.docx",
               "Chi bộ 13 ban hành. Điểm cột đỏ là ước tính."],
              ["A2", "Danh sách trích ngang",
               "02_Danh_sach_trich_ngang.xlsx",
               "In 1 sheet Trich ngang. Landscape."],
              ["A3", "Biểu tổng hợp đề xuất",
               "03_Bang_tong_hop_de_xuat.xlsx",
               "CHỈ in sheet Tong hop de xuat. Không in sheet Tinh diem KPI."],
              ["A4", "Quyết định phân công + danh mục sản phẩm",
               "04_Quyet_dinh_phan_cong_va_danh_muc_san_pham.docx",
               "Chi ủy Chi bộ 13 ban hành. In cả phụ lục."],
              ["A5", "Kế hoạch công tác quý cá nhân (7 file)",
               "05_Ke_hoach_ca_nhan/KH_QIII-2026_1..7_*.docx",
               "Mỗi đồng chí 1 file. Có dòng Chi bộ 13 ban hành kèm theo QĐ."],
              ["A6", "Bản tự đánh giá 30 + 70 (7 file)",
               "06_Ban_tu_danh_gia/TDG_QIII-2026_1..7_*.xlsx",
               "CHỈ in sheet Khung đánh giá. Không in sheet Tinh KPI 2A-2B."],
              ["B1", "Báo cáo kết quả theo đề cương",
               "07_Bao_cao_ket_qua_QIII_2026.docx",
               "Chi bộ 13 ban hành. Điểm bình quân in đỏ là ước tính."],
              ["B2", "Biểu thống kê cấp ủy viên",
               "08_Thong_ke_cap_uy_vien.xlsx",
               "In 1 sheet Thong ke. Landscape."],
          ],
          [1.2, 4.4, 5.6, 5.0],
          hdr_fill="C6EFCE")
    p(d, "", after=2)
    p(d, "Cộng: 3 file Word đơn + 7 kế hoạch + 1 quyết định + 3 Excel đơn (trích ngang, tổng hợp, thống kê) "
         "+ 7 bản tự đánh giá = 21 file hồ sơ. In ra khoảng 18 bản giấy "
         "(7 kế hoạch + 7 tự đánh giá + 4 văn bản/biểu của Chi bộ).",
      align="just", after=10)

    p(d, "3. THỨ TỰ IN GỢI Ý", bold=True, size=14, before=4, after=6)
    for i, s in enumerate([
        "Tờ trình 01 -> Danh sách trích ngang 02 -> Biểu tổng hợp 03 (sheet Tong hop de xuat).",
        "Quyết định 04 (cả phụ lục danh mục sản phẩm).",
        "7 kế hoạch cá nhân, đúng STT 1 Lê Hữu Nghĩa đến 7 Nguyễn Thị Lệ An.",
        "7 bản tự đánh giá, chỉ sheet Khung đánh giá, cùng thứ tự.",
        "Báo cáo 07 -> Biểu thống kê 08.",
    ], 1):
        p(d, "%d. %s" % (i, s), after=3)

    p(d, "4. CHỮ VÀ SỐ MÀU ĐỎ", bold=True, size=14, before=8, after=6)
    p(d, "Mọi chữ và số màu đỏ là ước tính của Javis, căn cứ dạng hình file mẫu Danh gia PCVP.xlsx. "
         "Số đen là dữ liệu thật từ danh sách đảng viên đã chuẩn hóa hoặc là quyết định của Chi bộ "
         "(danh sách 7 đồng chí, mức HTTNV, phân công).",
      align="just")
    p(d, "Điểm ước tính từng đồng chí (thang 100): Nghĩa 96,01; Luận 96,88; Trung 94,86; "
         "Phương 94,00; Lan 94,86; Xuân 95,72; Lệ An 97,74. Bình quân 95,72. "
         "Anh sửa 8 ô nhập ở sheet Tinh KPI 2A-2B thì điểm ở Khung đánh giá tự chạy.",
      align="just", after=10)

    p(d, "5. VIỆC PHẢI CHỐT TRƯỚC KHI TRÌNH KÝ", bold=True, size=14, before=4, after=6)
    for i, s in enumerate([
        "Số, ngày văn bản cho tờ trình, quyết định, báo cáo.",
        "Minh chứng quý III (số và ngày văn bản) thay chỗ chữ đỏ / [ĐIỀN KẾT QUẢ THỰC TẾ].",
        "Phân công căn quy chế 05-QĐ/CB ngày 10/2/2026 và file phân công chi ủy: "
        "Nghĩa phụ trách chung; Luận tuyên giáo và điều hành thay Bí thư; Trung kiểm tra giám sát; "
        "Phương đảng vụ - hồ sơ đảng viên; Lan tài chính Đảng, đảng phí, biên bản Chi ủy; "
        "Xuân dân vận và Đoàn Thanh niên; Lệ An công tác Đảng, quần chúng ưu tú.",
        "HTTNV cho 7/7: hoàn thành 100% nhiệm vụ, không đặt tiêu chí vượt mức. "
        "Không dùng mức HTXSNV nên không cần chứng minh 30% vượt mức.",
        "Hạn CV 4988 (14/10 và 16/10/2026) đã qua so với ngày lập 03/11/2026. "
        "Kế hoạch quý III đang lập sau khi quý kết thúc.",
    ], 1):
        p(d, "%d. %s" % (i, s), align="just", after=3)

    p(d, "6. NƠI LƯU FILE", bold=True, size=14, before=8, after=6)
    p(d, "Trên máy: H:\\Other computers\\My Laptop\\2. CONG TAC KHAC\\4. CONG TAC DANG\\"
         "A PHUONG\\2026\\06_Tieu_chi_danh_gia\\Du_thao_CB13_QuyIII_2026\\",
      size=11, after=3)
    p(d, "Trong vault: bao-cao/danh-gia-quy-III-2026-CB13/", size=11, after=8)
    p(d, "./.", align="right", after=0)

    out = os.path.join(VAULT, FN)
    d.save(out)
    os.makedirs(LAPTOP, exist_ok=True)
    shutil.copy2(out, os.path.join(LAPTOP, FN))
    print("WROTE", out)
    print("COPIED", os.path.join(LAPTOP, FN))
    print("SIZE", os.path.getsize(out))


if __name__ == "__main__":
    build()
