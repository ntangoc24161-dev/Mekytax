# -*- coding: utf-8 -*-
"""Sinh cac bieu Excel cho ho so danh gia Quy III/2026 - Chi bo 13."""
import os, sys, shutil, unicodedata
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from cb13_data import *
from cb13_noidung import NHIEM_VU, ANH_XA, CHUNG, CHO_DIEN
import cb13_uoc_tinh as UT

OUT = sys.argv[1]
TNR = "Times New Roman"


def set_tnr(wb, bo_sheet_rong=True):
    """Dat font Times New Roman cho MOI o co du lieu + style Normal cua workbook.

    Giu nguyen bold/italic/size/mau da gan truoc do, chi doi ten font.
    """
    try:
        wb._named_styles["Normal"].font.name = TNR
    except Exception:
        pass
    for ws in list(wb.worksheets):
        if bo_sheet_rong and ws.max_row == 1 and ws.max_column == 1 \
                and ws.cell(1, 1).value is None and len(wb.worksheets) > 1:
            wb.remove(ws)
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.value is None and c.font.name in (None, TNR):
                    continue
                f = c.font
                c.font = Font(name=TNR, size=f.size, bold=f.bold, italic=f.italic,
                              underline=f.underline, color=f.color,
                              vertAlign=f.vertAlign, strike=f.strike)
    return wb
TPL = sys.argv[2]  # file khung ban tu danh gia cap uy

THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
HDR_FILL = PatternFill("solid", fgColor="D9E1F2")


def slugify(s):
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.replace("Đ", "D").replace("đ", "d").replace(" ", "-")


def style_sheet(ws, r0, r1, c1, cn, widths):
    for r in range(r0, r1 + 1):
        for c in range(c1, cn + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BOX
            cell.alignment = CTR if c == c1 else WRAP
    for i, w in enumerate(widths, start=c1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ===================== 02. DANH SACH TRICH NGANG =====================
def trich_ngang():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trich ngang"
    ws["A1"] = DANG_BO; ws["A1"].font = Font(name=TNR, bold=True, size=11)
    ws["A2"] = CHI_BO; ws["A2"].font = Font(name=TNR, bold=True, size=12, underline="single")
    ws["A3"] = "________"
    ws["A3"].font = Font(name=TNR, bold=True, size=11)
    ws["F1"] = "ĐẢNG CỘNG SẢN VIỆT NAM"; ws["F1"].font = Font(name=TNR, bold=True, size=11)
    ws["F2"] = "Đồng Nai, ngày ..... tháng ..... năm 2026"; ws["F2"].font = Font(name=TNR, italic=True, size=11)
    ws["A4"] = "DANH SÁCH TRÍCH NGANG"; ws["A4"].font = Font(name=TNR, bold=True, size=14)
    ws["A5"] = ("Đề xuất mức xếp loại chất lượng quý III/2026 đối với cán bộ thuộc diện "
                "Ban Thường vụ Đảng ủy Thuế thành phố quản lý")
    ws["A5"].font = Font(bold=True, size=11)
    ws["A6"] = "(Kèm theo Tờ trình số        -TTr/CB13 ngày ..... tháng ..... năm 2026 của Chi bộ 13)"
    ws["A6"].font = Font(italic=True, size=10)

    hdr = ["TT", "Họ và tên", "Ngày sinh", "Giới tính", "Số định danh cá nhân",
           "Ngày vào Đảng", "Ngày chính thức", "GD phổ thông", "Trình độ chuyên môn",
           "Lý luận chính trị", "Chức vụ Đảng", "Chức vụ chính quyền", "Chức vụ đoàn thể",
           "Đơn vị công tác", "Lĩnh vực phụ trách", "Mức xếp loại đề xuất QIII/2026"]
    r = 8
    for i, h in enumerate(hdr, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(bold=True, size=10); c.alignment = CTR; c.fill = HDR_FILL
    for n in NHAN_SU:
        r += 1
        vals = [n["stt"], n["ten"], n["ns"], n["gt"], n["cccd"], n["vd"], n["ct"],
                n["gdpt"], n["cm"], n["llct"], n["cv_dang"], n["cv_cq"] or "Không",
                n["cv_dt"] or "Không", "Chi bộ 13 - Đảng bộ Thuế thành phố Đồng Nai",
                n["linh_vuc_ngan"], XEP_LOAI]
        for i, v in enumerate(vals, 1):
            cc = ws.cell(row=r, column=i, value=v)
            cc.font = Font(size=10)
    style_sheet(ws, 8, r, 1, 16,
                [5, 22, 11, 8, 16, 12, 12, 10, 14, 12, 34, 20, 12, 30, 26, 26])
    ws.row_dimensions[8].height = 46
    for rr in range(9, r + 1):
        ws.row_dimensions[rr].height = 40
    r += 2
    ws.cell(row=r, column=2, value="NGƯỜI LẬP").font = Font(bold=True, size=11)
    ws.cell(row=r, column=11, value="T/M CHI BỘ 13").font = Font(bold=True, size=11)
    ws.cell(row=r + 1, column=11, value="BÍ THƯ").font = Font(bold=True, size=11)
    ws.cell(row=r + 1, column=2, value="(Ký, ghi rõ họ tên)").font = Font(italic=True, size=10)
    ws.cell(row=r + 2, column=11, value="(Ký, ghi rõ họ tên và đóng dấu)").font = Font(italic=True, size=10)
    ws.cell(row=r + 5, column=2, value="Nguyễn Đăng Phương").font = Font(bold=True, size=11)
    ws.cell(row=r + 5, column=11, value="Lê Hữu Nghĩa").font = Font(bold=True, size=11)
    ws.cell(row=r + 7, column=1,
            value="Ghi chú: 04 mức xếp loại - Hoàn thành xuất sắc nhiệm vụ / Hoàn thành tốt nhiệm vụ / "
                  "Hoàn thành nhiệm vụ / Không hoàn thành nhiệm vụ.").font = Font(italic=True, size=10)
    ws.page_setup.orientation = "landscape"
    ws.freeze_panes = "A9"
    set_tnr(wb)
    wb.save(os.path.join(OUT, "02_Danh_sach_trich_ngang.xlsx"))


NL = chr(10)
TD_SHEET = "Tinh diem KPI"


def tinh_diem_sheet(wb):
    """Sheet tinh diem song cho bieu tong hop: sua KPI thi tong diem tu doi."""
    k = wb.create_sheet(TD_SHEET)
    k["A1"] = "BẢNG TÍNH ĐIỂM ƯỚC TÍNH QUÝ III/2026 - 7 ĐỒNG CHÍ CHI ỦY CHI BỘ 13"
    k["A1"].font = Font(bold=True, size=12)
    k["A2"] = ("Sửa ô KPI (%) màu đỏ của từng tiêu chí thì Điểm nhóm B, Tổng điểm và cột Ghi chú "
               "ở sheet Tong hop de xuat tự cập nhật. KPI từng tiêu chí lấy từ sheet "
               "'Tinh KPI 2A-2B' trong bản tự đánh giá của từng đồng chí.")
    k["A2"].font = Font(italic=True, size=10)
    hdr = ["TT", "Họ và tên"] + ["KPI TC%d (%%)" % t for t, _, _ in TIEU_CHI_B] + \
          ["Điểm nhóm A (30)", "Điểm nhóm B (70)", "Tổng điểm (100)"]
    for i, h in enumerate(hdr, 1):
        c = k.cell(row=3, column=i, value=h)
        c.font = Font(bold=True, size=9); c.alignment = CTR; c.fill = HDR_FILL
    k.cell(row=4, column=2, value="Điểm tối đa từng tiêu chí").font = Font(bold=True, size=9)
    for i, (tt, diem, _) in enumerate(TIEU_CHI_B):
        c = k.cell(row=4, column=3 + i, value=diem)
        c.font = Font(bold=True, size=9); c.alignment = CTR
    r = 4
    for n in NHAN_SU:
        r += 1
        profs = UT.phan_bo(n["stt"])
        k.cell(row=r, column=1, value=n["stt"])
        k.cell(row=r, column=2, value=n["ten_hoa_thuong"]).font = Font(size=10)
        for i, pr in enumerate(profs):
            c = k.cell(row=r, column=3 + i, value=UT.kpi(pr))
            c.font = Font(size=10, color=UT.DO); c.alignment = CTR
        k.cell(row=r, column=12, value=30).font = Font(size=10)
        cb = k.cell(row=r, column=13,
                    value="=ROUND(SUMPRODUCT(C%d:K%d,$C$4:$K$4)/100,2)" % (r, r))
        cb.font = Font(size=10)
        ct = k.cell(row=r, column=14, value="=ROUND(L%d+M%d,2)" % (r, r))
        ct.font = Font(bold=True, size=10)
    r += 1
    k.cell(row=r, column=2, value="Trung bình 7 đồng chí").font = Font(bold=True, size=9)
    k.cell(row=r, column=13, value="=ROUND(AVERAGE(M5:M11),2)").font = Font(bold=True, size=9)
    k.cell(row=r, column=14, value="=ROUND(AVERAGE(N5:N11),2)").font = Font(bold=True, size=9)
    style_sheet(k, 3, r, 1, 14,
                [5, 24] + [11] * 9 + [14, 14, 14])
    k.row_dimensions[3].height = 40
    k.cell(row=r + 2, column=2,
           value="Mức đề xuất HTTNV: hoàn thành 100% nhiệm vụ, chất lượng tốt, không đặt tiêu chí "
                 "vượt mức. Ước tính dùng mức hoàn thành tốt (KPI 92,25%) và hoàn thành (KPI 86,50%). "
                 "Chi bộ thay bằng minh chứng thực tế trước khi trình ký."
           ).font = Font(italic=True, size=10, color=UT.DO)
    k.cell(row=r + 4, column=2,
           value="SHEET THEO DÕI NỘI BỘ - KHÔNG NỘP CẤP TRÊN. Chỉ dùng để tính điểm / sửa số ước tính."
           ).font = Font(bold=True, italic=True, size=10, color="C00000")


# ===================== 03. BIEU TONG HOP =====================
def tong_hop():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tong hop de xuat"
    ws["A1"] = DANG_BO; ws["A1"].font = Font(name=TNR, bold=True, size=11)
    ws["A2"] = CHI_BO; ws["A2"].font = Font(name=TNR, bold=True, size=12, underline="single")
    ws["A3"] = "Chi bộ 13 ban hành"
    ws["A3"].font = Font(name=TNR, italic=True, size=11)
    ws["E1"] = "ĐẢNG CỘNG SẢN VIỆT NAM"; ws["E1"].font = Font(name=TNR, bold=True, size=11)
    ws["E2"] = "Đồng Nai, ngày ..... tháng ..... năm 2026"; ws["E2"].font = Font(name=TNR, italic=True, size=11)
    ws["A4"] = "BIỂU TỔNG HỢP"; ws["A4"].font = Font(name=TNR, bold=True, size=14)
    ws["A5"] = ("KẾT QUẢ THỰC HIỆN NHIỆM VỤ, CHỈ TIÊU ĐƯỢC GIAO VÀ ĐỀ XUẤT MỨC ĐỘ ĐÁNH GIÁ, XẾP LOẠI "
                "CÁN BỘ THUỘC DIỆN BAN THƯỜNG VỤ ĐẢNG ỦY THUẾ THÀNH PHỐ QUẢN LÝ QUÝ III/2026")
    ws["A5"].font = Font(bold=True, size=11)
    hdr = ["TT", "Họ và tên", "Chức vụ, đơn vị (chức danh chủ yếu)",
           "Những nhiệm vụ, chỉ tiêu chủ yếu được giao quý III/2026",
           "Tóm tắt kết quả chủ yếu trong thực hiện chức trách, nhiệm vụ quý III/2026",
           "Đánh giá, xếp loại cán bộ", "Ghi chú"]
    r = 7
    for i, h in enumerate(hdr, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(bold=True, size=10); c.alignment = CTR; c.fill = HDR_FILL
    for n in NHAN_SU:
        r += 1
        nvs = NHIEM_VU[n["ten"]]
        profs = UT.phan_bo(n["stt"])
        amap = ANH_XA[n["ten"]]
        nv_text = NL.join("%d. %s (chỉ tiêu: %s)" % (i, x[0], x[2]) for i, x in enumerate(nvs, 1))
        kq_lines = []
        for i, x in enumerate(nvs, 1):
            tc = next((t for t, j in amap.items() if j == i - 1), None)
            pr = profs[tc - 1] if tc else UT.TOT
            kq_lines.append("%d. %s - Mức %s: số lượng %s%%, chất lượng %s%%, "
                            "tiến độ %s%%, năng lực lãnh đạo %s%%; KPI %s%%."
                            % (i, x[4], UT.NHAN[pr], UT.so(pr[0] * 100), UT.so(pr[1] * 100),
                               UT.so(pr[2] * 100), UT.so(pr[3] * 100), UT.so(UT.kpi(pr))))
        kq_text = NL.join(kq_lines)
        kr = r - 7 + 4  # dong tuong ung tai sheet tinh diem
        ghi_chu = ('="Tổng điểm: "&TEXT(\'%s\'!N%d,"0.00")&"/100 '
                   '(nhóm A 30 + nhóm B "&TEXT(\'%s\'!M%d,"0.00")&")."'
                   % (TD_SHEET, kr, TD_SHEET, kr))
        vals = [n["stt"], n["ten"], "%s / %s" % (n["cv_dang"], n["cv_cq"] or "Không"),
                nv_text, kq_text, XEP_LOAI, ghi_chu]
        for i, v in enumerate(vals, 1):
            cc = ws.cell(row=r, column=i, value=v)
            cc.font = Font(size=10, color=UT.DO) if i in (5, 7) else Font(size=10)
        ws.row_dimensions[r].height = 190
    style_sheet(ws, 7, r, 1, 7, [5, 22, 26, 58, 58, 24, 14])
    ws.row_dimensions[7].height = 56
    r += 2
    ws.cell(row=r, column=5, value="T/M CẤP ỦY CHI BỘ 13").font = Font(bold=True, size=11)
    ws.cell(row=r + 1, column=5, value="BÍ THƯ").font = Font(bold=True, size=11)
    ws.cell(row=r + 2, column=5,
            value="(Xác lập thời điểm, ký, ghi rõ họ tên và đóng dấu)").font = Font(italic=True, size=10)
    ws.cell(row=r + 5, column=5, value="Lê Hữu Nghĩa").font = Font(bold=True, size=11)
    ws.cell(row=r + 7, column=1,
            value="Lưu ý: toàn bộ chữ và số MÀU ĐỎ ở cột 'Tóm tắt kết quả' và cột 'Ghi chú' là ước tính "
                  "của Javis theo dạng hình số liệu file mẫu Danh gia PCVP.xlsx. Chi bộ sửa các ô KPI "
                  "màu đỏ tại sheet '%s' thì cột Ghi chú tự cập nhật tổng điểm; đồng thời thay bằng "
                  "số liệu, số và ngày văn bản minh chứng thực tế quý III/2026 trước khi trình ký."
                  % TD_SHEET
            ).font = Font(italic=True, size=10, color=UT.DO)
    tinh_diem_sheet(wb)
    ws.page_setup.orientation = "landscape"
    set_tnr(wb)
    wb.save(os.path.join(OUT, "03_Bang_tong_hop_de_xuat.xlsx"))


# ===================== 08. THONG KE CAP UY VIEN =====================
def thong_ke():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Thong ke"
    ws["A1"] = DANG_BO; ws["A1"].font = Font(name=TNR, bold=True, size=11)
    ws["A2"] = CHI_BO; ws["A2"].font = Font(name=TNR, bold=True, size=12, underline="single")
    ws["J1"] = "ĐẢNG CỘNG SẢN VIỆT NAM"; ws["J1"].font = Font(name=TNR, bold=True, size=11)
    ws["J1"].alignment = Alignment(horizontal="center")
    ws["J2"] = "Đồng Nai, ngày ..... tháng ..... năm 2026"
    ws["J2"].font = Font(name=TNR, italic=True, size=11)
    ws["J2"].alignment = Alignment(horizontal="center")
    ws["A3"] = ("THỐNG KÊ KẾT QUẢ ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG QUÝ III/2026 ĐỐI VỚI CẤP ỦY VIÊN "
                "TRONG CẤP ỦY TRỰC THUỘC CƠ SỞ")
    ws["A3"].font = Font(name=TNR, bold=True, size=12)
    ws["A4"] = "Chi bộ 13 ban hành"
    ws["A4"].font = Font(name=TNR, italic=True, size=11)
    hdr = ["TT", "Chức danh", "Có mặt cuối quý III/2026", "Chưa đánh giá, xếp loại",
           "Đã đánh giá, xếp loại", "Hoàn thành tốt nhiệm vụ", "Trong đó: Hoàn thành xuất sắc nhiệm vụ",
           "Hoàn thành nhiệm vụ", "Không hoàn thành nhiệm vụ", "Ghi chú"]
    r = 5
    for i, h in enumerate(hdr, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.font = Font(bold=True, size=10); c.alignment = CTR; c.fill = HDR_FILL
    data = [
        ("1", "Bí thư Chi bộ", 1, 0, 1, 1, 0, 0, 0, "Lê Hữu Nghĩa"),
        ("2", "Phó Bí thư Chi bộ", 1, 0, 1, 1, 0, 0, 0, "Trần Ngọc Luận"),
        ("3", "Chi ủy viên", 5, 0, 5, 5, 0, 0, 0,
         "Nguyễn Khắc Trung, Nguyễn Đăng Phương, Lê Thị Lan, Phạm Đồng Xuân, Nguyễn Thị Lệ An"),
        ("4", "Ủy viên UBKT, Chủ nhiệm, Phó Chủ nhiệm UBKT đảng ủy trực thuộc cơ sở", 0, 0, 0, 0, 0, 0, 0,
         "Chi bộ 13 không có UBKT; công tác kiểm tra, giám sát do 01 chi ủy viên phụ trách"),
        ("5", "Lãnh đạo cấp phòng và tương đương (nếu có)", 0, 0, 0, 0, 0, 0, 0,
         "Không thuộc đối tượng thống kê tại biểu này"),
    ]
    for row in data:
        r += 1
        for i, v in enumerate(row, 1):
            cc = ws.cell(row=r, column=i, value=v)
            cc.font = Font(size=10)
        ws.row_dimensions[r].height = 34
    r += 1
    ws.cell(row=r, column=2, value="TỔNG CỘNG").font = Font(bold=True, size=10)
    for col in range(3, 10):
        L = get_column_letter(col)
        ws.cell(row=r, column=col, value="=SUM(%s6:%s%d)" % (L, L, r - 1)).font = Font(bold=True, size=10)
    style_sheet(ws, 5, r, 1, 10, [5, 38, 14, 14, 14, 14, 18, 14, 14, 52])
    ws.row_dimensions[5].height = 56
    r += 2
    ws.cell(row=r, column=1,
            value="Ghi chú: cột 'Hoàn thành tốt nhiệm vụ' ghi số đồng chí xếp loại ở mức Hoàn thành tốt "
                  "nhiệm vụ; cột 'Trong đó: Hoàn thành xuất sắc nhiệm vụ' ghi số đồng chí xếp loại "
                  "Hoàn thành xuất sắc nhiệm vụ. Chi bộ đối chiếu lại cách ghi theo hướng dẫn của "
                  "Đảng ủy Thuế thành phố trước khi gửi.").font = Font(italic=True, size=10, color="C00000")
    r += 3
    ws.cell(row=r, column=6, value="T/M CHI BỘ 13").font = Font(bold=True, size=11)
    ws.cell(row=r + 1, column=6, value="BÍ THƯ").font = Font(bold=True, size=11)
    ws.cell(row=r + 4, column=6, value="Lê Hữu Nghĩa").font = Font(bold=True, size=11)
    ws.page_setup.orientation = "landscape"
    set_tnr(wb)
    wb.save(os.path.join(OUT, "08_Thong_ke_cap_uy_vien.xlsx"))


# ===================== 06. BAN TU DANH GIA (7 file) =====================
ROWS_A = list(range(17, 26)) + list(range(27, 31)) + list(range(32, 36))
ROWS_B = list(range(39, 48))
KPI_SHEET = "Tinh KPI 2A-2B"


def tu_danh_gia():
    sub = os.path.join(OUT, "06_Ban_tu_danh_gia")
    os.makedirs(sub, exist_ok=True)
    from cb13_noidung import NHIEM_VU as NV
    for n in NHAN_SU:
        wb = openpyxl.load_workbook(TPL)
        ws = wb["Khung đánh giá"]
        ws["A3"] = DANG_BO + "\nCHI BỘ 13"
        ws["F4"] = "Đồng Nai, ngày ..... tháng ..... năm 2026"
        ws["A6"] = "Quý III, Năm 2026"
        ws["B7"] = "Họ và tên: %s                    Ngày sinh: %s" % (n["ten_hoa_thuong"], n["ns"])
        ws["B8"] = "Chức vụ Đảng: %s" % n["cv_dang"]
        ws["B9"] = "Chức vụ chính quyền: %s" % (n["cv_cq"] or "Không")
        ws["B10"] = "Chức vụ đoàn thể: %s" % (n["cv_dt"] or "Không")
        ws["B11"] = "Đơn vị công tác: Chi bộ 13 - Đảng bộ Thuế thành phố Đồng Nai"

        # --- Nhom A: danh dau dam bao, diem dat = diem toi da ---
        for r in ROWS_A:
            ws.cell(row=r, column=5, value="x").alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=8, value="=G%d" % r)
        ws["H16"] = "=SUM(H17:H25)"
        ws["H26"] = "=SUM(H27:H30)"
        ws["H31"] = "=SUM(H32:H35)"
        ws["H36"] = "=H16+H26+H31"
        for r in (16, 26, 31, 36):
            ws.cell(row=r, column=8).font = Font(bold=True, size=11)

        # --- Nhom B ---
        nvs = NV[n["ten"]]
        amap = ANH_XA[n["ten"]]
        profs = UT.phan_bo(n["stt"])
        nv_theo_tc = []
        for idx, r in enumerate(ROWS_B):
            tc = idx + 1
            j = amap.get(tc)
            nv = None
            if j is not None:
                nv = nvs[j]
                muc_tieu = "%s\nChỉ tiêu: %s\nSản phẩm: %s\nThời hạn: %s" % (nv[0], nv[2], nv[4], nv[7])
            else:
                muc_tieu = CHUNG.get(tc, "")
            nv_theo_tc.append(nv)
            ws.cell(row=r, column=4, value=muc_tieu).alignment = WRAP
            ce = ws.cell(row=r, column=5, value=UT.ket_qua_text(nv, profs[idx]))
            ce.alignment = WRAP
            ce.font = Font(size=10, color=UT.DO)
            ws.cell(row=r, column=6,
                    value="=IF('%s'!K%d=\"\",\"\",'%s'!K%d)" % (KPI_SHEET, idx + 4, KPI_SHEET, idx + 4))
            ws.cell(row=r, column=8, value="=IFERROR(F%d/100*G%d,0)" % (r, r))
            ci = ws.cell(row=r, column=9,
                         value="Số màu đỏ là ước tính của Javis. Sửa 8 ô nhập tại sheet '%s' "
                               "dòng %d thì cột Điểm KPI (%%) và Điểm đạt tự cập nhật."
                               % (KPI_SHEET, idx + 4))
            ci.alignment = WRAP
            ci.font = Font(size=9, italic=True, color=UT.DO)
            ws.row_dimensions[r].height = 120
        ws["H48"] = "=SUM(H39:H47)"
        ws["H49"] = "=H36+H48"
        ws["H48"].font = Font(bold=True, size=11)
        ws["H49"].font = Font(bold=True, size=12)
        ws.column_dimensions["D"].width = 52
        ws.column_dimensions["E"].width = 46
        ws.column_dimensions["I"].width = 30

        ws["B50"] = "II. Tự đề xuất xếp loại mức chất lượng: %s" % XEP_LOAI
        ws["B50"].font = Font(bold=True, size=12)
        ws["F53"] = "(Ký, ghi rõ họ tên)\n\n\n\n%s" % n["ten_hoa_thuong"]
        ws["F53"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        ws.row_dimensions[53].height = 76
        ws["D58"] = ("%s\n(Xác lập thời điểm, ký, ghi rõ họ tên và đóng dấu)\n\n\n\n%s"
                     % (n["xac_nhan"], "Lê Hữu Nghĩa" if n["stt"] != 1 else ""))
        ws["D58"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        ws.row_dimensions[58].height = 92
        ws.cell(row=65, column=1,
                value="LƯU Ý KHI HOÀN THIỆN: (1) Mọi số và chữ MÀU ĐỎ là ước tính của Javis theo dạng "
                      "hình số liệu file mẫu Danh gia PCVP.xlsx%s - thay bằng số liệu và số, ngày văn bản "
                      "minh chứng thực tế quý III/2026. (2) Cột 'Điểm KPI (%%)' tự tính "
                      "từ sheet '%s', phải nhập đủ 8 ô số lượng/chất lượng/tiến độ/lãnh đạo (kế hoạch và "
                      "thực tế) thì mới ra điểm. (3) Mức tự đề xuất đang ghi %s theo chỉ đạo; "
                      "hoàn thành 100%% nhiệm vụ, không đặt tiêu chí vượt mức."
                      % (CHO_DIEN, KPI_SHEET, XEP_LOAI)).font = Font(italic=True, size=10, color="C00000")

        # --- Sheet tinh KPI 2A-2B ---
        k = wb.create_sheet(KPI_SHEET)
        k["A1"] = "BẢNG TÍNH ĐIỂM KPI THEO PHỤ LỤC 2A VÀ 2B - %s - QUÝ III/2026" % n["ten"]
        k["A1"].font = Font(bold=True, size=12)
        k["A2"] = ("KPI (%) = trung bình của 4 tỷ lệ: A số lượng, B chất lượng, C tiến độ, "
                   "D năng lực lãnh đạo, chỉ đạo, điều hành. Mỗi trụ = điểm quy đổi thực tế / "
                   "điểm quy đổi kế hoạch. Nhập đủ 8 ô mới ra kết quả.")
        k["A2"].font = Font(italic=True, size=10)
        khdr = ["TT", "Tiêu chí (Nhóm B)", "A - Số lượng (KH)", "A - Số lượng (TT)",
                "B - Chất lượng (KH)", "B - Chất lượng (TT)", "C - Tiến độ (KH)", "C - Tiến độ (TT)",
                "D - Lãnh đạo (KH)", "D - Lãnh đạo (TT)", "KPI (%)", "Điểm tối đa", "Điểm đạt"]
        for i, h in enumerate(khdr, 1):
            c = k.cell(row=3, column=i, value=h)
            c.font = Font(bold=True, size=10); c.alignment = CTR; c.fill = HDR_FILL
        for idx, (tt, diem, ten_tc) in enumerate(TIEU_CHI_B):
            r = idx + 4
            k.cell(row=r, column=1, value=tt)
            k.cell(row=r, column=2, value=ten_tc).alignment = WRAP
            for off, val in enumerate(UT.o_nhap(nv_theo_tc[idx], profs[idx])):
                cn = k.cell(row=r, column=3 + off, value=val)
                cn.font = Font(size=10, color=UT.DO)
                cn.alignment = CTR
            k.cell(row=r, column=11,
                   value="=IF(COUNT(C%d:J%d)<8,\"\",ROUND(AVERAGE(D%d/C%d,F%d/E%d,H%d/G%d,J%d/I%d)*100,2))"
                         % (r, r, r, r, r, r, r, r, r, r))
            k.cell(row=r, column=12, value=diem)
            k.cell(row=r, column=13, value="=IFERROR(K%d/100*L%d,0)" % (r, r))
            k.row_dimensions[r].height = 54
        rtot = 4 + len(TIEU_CHI_B)
        k.cell(row=rtot, column=2, value="TỔNG NHÓM B (70 điểm)").font = Font(bold=True, size=10)
        k.cell(row=rtot, column=12, value="=SUM(L4:L12)").font = Font(bold=True, size=10)
        k.cell(row=rtot, column=13, value="=SUM(M4:M12)").font = Font(bold=True, size=10)
        style_sheet(k, 3, rtot, 1, 13, [5, 44, 12, 12, 12, 12, 12, 12, 12, 12, 11, 11, 11])
        k.row_dimensions[3].height = 46
        k.cell(row=rtot + 1, column=2,
               value="CÁC SỐ MÀU ĐỎ (cột C đến J) LÀ ƯỚC TÍNH CỦA JAVIS, căn cứ dạng hình số liệu của "
                     "file mẫu Danh gia PCVP.xlsx. Chi bộ thay bằng số thực tế quý III/2026; cột "
                     "KPI (%), Điểm đạt và tổng điểm ở sheet Khung đánh giá tự tính lại."
               ).font = Font(bold=True, italic=True, size=10, color=UT.DO)
        k.cell(row=rtot + 2, column=2,
               value="Ghi chú: điểm quy đổi kế hoạch của từng trụ lấy theo danh mục sản phẩm, công việc "
                     "đã cá thể hóa (số lượng x hệ số x 100). Điểm quy đổi thực tế lấy theo khối lượng, "
                     "chất lượng, tiến độ thực tế đã hoàn thành trong quý."
               ).font = Font(italic=True, size=10)
        k.cell(row=rtot + 4, column=2,
               value="SHEET THEO DÕI NỘI BỘ - KHÔNG NỘP CẤP TRÊN. Chỉ dùng để tính điểm / sửa số ước tính. "
                     "In và gửi sheet Khung đánh giá."
               ).font = Font(bold=True, italic=True, size=10, color="C00000")

        fn = "TDG_QIII-2026_%d_%s.xlsx" % (n["stt"], slugify(n["ten_hoa_thuong"]))
        set_tnr(wb)
        wb.save(os.path.join(sub, fn))


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    trich_ngang(); print("02 trich ngang OK")
    tong_hop(); print("03 tong hop OK")
    thong_ke(); print("08 thong ke OK")
    tu_danh_gia(); print("06 ban tu danh gia OK")
