# -*- coding: utf-8 -*-
"""Chi sua NOI DUNG ho so Quy III/2026 CB13: phan cong chinh thuc + HTTNV.
Khong doi font, border, width, mau, cong thuc (tru khi thay value o nhap KPI).
"""
import os, sys, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from copy import copy, deepcopy
from docx import Document
from openpyxl import load_workbook
from openpyxl.styles import Font
from cb13_data import *
from cb13_noidung import NHIEM_VU, ANH_XA, CHUNG, CHO_DIEN
import cb13_uoc_tinh as UT

VAULT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
LAPTOP = (r"H:\Other computers\My Laptop\2. CONG TAC KHAC\4. CONG TAC DANG"
          r"\A PHUONG\2026\06_Tieu_chi_danh_gia\Du_thao_CB13_QuyIII_2026")
NL = chr(10)
KPI_SHEET = "Tinh KPI 2A-2B"
TD_SHEET = "Tinh diem KPI"
ROWS_B = list(range(39, 48))
LOG = []


def log(msg):
    LOG.append(msg)
    print(msg)


def set_para_text(p, text):
    """Ghi de chu, giu format run dau."""
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def set_cell_text(cell, text):
    paras = cell.paragraphs
    if not paras:
        return
    set_para_text(paras[0], text if text is not None else "")
    for p in paras[1:]:
        set_para_text(p, "")


def replace_in_para(p, pairs):
    t = p.text
    if not t:
        return False
    n = t
    for a, b in pairs:
        n = n.replace(a, b)
    if n != t:
        set_para_text(p, n)
        return True
    return False


def walk_docx(doc, pairs):
    n = 0
    for p in doc.paragraphs:
        if replace_in_para(p, pairs):
            n += 1
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if replace_in_para(p, pairs):
                        n += 1
    return n


def insert_tr_before(table, tmpl_row, before_row):
    """Chen 1 hang copy format, khong doi style bang."""
    new_tr = deepcopy(tmpl_row._tr)
    before_row._tr.addprevious(new_tr)
    return table.rows[before_row._index - 1] if hasattr(before_row, "_index") else None


def insert_row_before_last(table):
    """Copy hang du lieu cuoi (truoc TONG) roi chen dung truoc hang TONG."""
    if len(table.rows) < 3:
        return table.rows[-1]
    tmpl = table.rows[-2]
    new_tr = deepcopy(tmpl._tr)
    table.rows[-1]._tr.addprevious(new_tr)
    return table.rows[-2]


def append_copied_row(table):
    """Copy hang cuoi roi noi them 1 hang (bang ke hoach ca nhan)."""
    tmpl = table.rows[-1]
    new_tr = deepcopy(tmpl._tr)
    table._tbl.append(new_tr)
    return table.rows[-1]


def diem(n):
    return 30 + UT.diem_nhom_b(n["stt"], TIEU_CHI_B)


def by_ten():
    return {n["ten"]: n for n in NHAN_SU}


def by_hoa():
    return {n["ten_hoa_thuong"]: n for n in NHAN_SU}


# ---------- WORD ----------
def patch_01(path):
    d = Document(path)
    dmin = min(diem(n) for n in NHAN_SU)
    dmax = max(diem(n) for n in NHAN_SU)
    pairs = [
        ("Căn cứ Quy chế làm việc của Chi bộ 13 và kết quả phân công nhiệm vụ đối với các đồng chí chi ủy viên quý III/2026;",
         "Căn cứ Quy chế làm việc của Chi ủy Chi bộ 13 nhiệm kỳ 2025-2030 (ban hành kèm theo Quyết định số 05-QĐ/CB ngày 10/2/2026) và phân công nhiệm vụ đối với các đồng chí chi ủy viên;"),
        ("Trên cơ sở kết quả chấm điểm theo khung 30 + 70 và tỷ lệ nhiệm vụ hoàn thành vượt mức, Chi bộ đề xuất xếp loại cả 07 đồng chí ở mức Hoàn thành xuất sắc nhiệm vụ.",
         "Trên cơ sở kết quả chấm điểm theo khung 30 + 70, Chi bộ đề xuất xếp loại cả 07 đồng chí ở mức Hoàn thành tốt nhiệm vụ."),
    ]
    walk_docx(d, pairs)
    # doan diem uoc tinh (do, nghieng) - thay toan van
    for p in d.paragraphs:
        if "mỗi đồng chí có 03 trong 09 tiêu chí" in p.text or "đáp ứng điều kiện từ 30%" in p.text:
            set_para_text(
                p,
                "Tổng điểm của 07 đồng chí trong khoảng %s đến %s điểm trên thang 100 điểm "
                "(30 điểm nhóm tiêu chí chung và %s đến %s điểm nhóm kết quả thực hiện nhiệm vụ). "
                "Mức đề xuất là hoàn thành tốt nhiệm vụ: hoàn thành 100%% nhiệm vụ được giao, "
                "không đặt tiêu chí vượt mức."
                % (UT.so(dmin), UT.so(dmax), UT.so(dmin - 30), UT.so(dmax - 30)),
            )
    # bang 7 dong chi: TT, ten, cv, linh vuc ngan, diem, xep loai
    people = by_hoa()
    for t in d.tables:
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if "Họ và tên" in hdr and "Mức xếp loại" in hdr:
            for row in t.rows[1:]:
                ten = row.cells[1].text.strip()
                n = people.get(ten)
                if not n:
                    continue
                set_cell_text(row.cells[2], n["cv_dang"])
                set_cell_text(row.cells[3], n["linh_vuc_ngan"])
                set_cell_text(row.cells[4], UT.so(diem(n)))
                set_cell_text(row.cells[5], XEP_LOAI)
    d.save(path)
    log("patched 01 " + os.path.basename(path))


def patch_04(path):
    d = Document(path)
    people_hoa = by_hoa()
    walk_docx(d, [
        ("Xét đề nghị của đồng chí Phó Bí thư Chi bộ phụ trách công tác tổ chức - đảng viên,",
         "Xét đề nghị của đồng chí Phó Bí thư Chi bộ phụ trách công tác tuyên giáo,"),
        ("Đồng chí Phó Bí thư Chi bộ và đồng chí Trợ lý Bí thư có trách nhiệm",
         "Đồng chí Phó Bí thư Chi bộ và đồng chí phụ trách công tác đảng vụ có trách nhiệm"),
    ])
    # Dieu 2: doan "N. Dong chi TEN - CV: linh vuc"
    for p in d.paragraphs:
        t = p.text.strip()
        for n in NHAN_SU:
            prefix = "%d. Đồng chí %s - " % (n["stt"], n["ten_hoa_thuong"])
            if t.startswith(prefix) or t.startswith("%d. Đồng chí %s" % (n["stt"], n["ten_hoa_thuong"])):
                set_para_text(
                    p,
                    "%d. Đồng chí %s - %s: %s"
                    % (n["stt"], n["ten_hoa_thuong"], n["cv_dang"], n["linh_vuc"]),
                )
            heading = "%d. Đồng chí %s - %s" % (n["stt"], n["ten_hoa_thuong"], n["cv_dang"])
            # tieu de phu luc (khong co dau :)
            if t.startswith("%d. Đồng chí %s" % (n["stt"], n["ten_hoa_thuong"])) and ":" not in t:
                set_para_text(p, heading)
    # bang phu luc: 7 bang nhiem vu (bo bang noi nhan)
    nv_tables = []
    for t in d.tables:
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if "Nhiệm vụ" in hdr and "Hệ số" in hdr:
            nv_tables.append(t)
    if len(nv_tables) != 7:
        log("WARN 04: so bang nhiem vu = %d (ky vong 7)" % len(nv_tables))
    for n, t in zip(NHAN_SU, nv_tables):
        nvs = NHIEM_VU[n["ten"]]
        body = t.rows[1:]
        # hang cuoi la TONG
        while body and "TỔNG" in body[-1].cells[1].text and (len(body) - 1) < len(nvs):
            insert_row_before_last(t)
            body = t.rows[1:]
        data_rows = body[:-1] if body and "TỔNG" in body[-1].cells[1].text else body
        for i, nv in enumerate(nvs):
            if i >= len(data_rows):
                log("WARN 04 %s thieu hang nv %d" % (n["ten"], i + 1))
                break
            noi_dung, loai, chi_tieu, tien_do, san_pham, phoi_hop, cap_trinh, thoi_han, sl, hs = nv
            diem_qd = round(sl * hs * 100)
            row = data_rows[i]
            set_cell_text(row.cells[0], str(i + 1))
            set_cell_text(row.cells[1], noi_dung)
            set_cell_text(row.cells[2], san_pham)
            set_cell_text(row.cells[3], "%d" % sl)
            set_cell_text(row.cells[4], ("%.1f" % hs).replace(".", ","))
            set_cell_text(row.cells[5], "%d" % diem_qd)
            set_cell_text(row.cells[6], thoi_han)
        tong = sum(round(nv[8] * nv[9] * 100) for nv in nvs)
        if body and "TỔNG" in body[-1].cells[1].text:
            set_cell_text(body[-1].cells[5], "%d" % tong)
    d.save(path)
    log("patched 04 " + os.path.basename(path))


def patch_kh(path, n):
    d = Document(path)
    for p in d.paragraphs:
        if p.text.startswith("Chức vụ Đảng:"):
            set_para_text(p, "Chức vụ Đảng: %s" % n["cv_dang"])
        elif p.text.startswith("Lĩnh vực được phân công phụ trách:"):
            set_para_text(p, "Lĩnh vực được phân công phụ trách: %s" % n["linh_vuc"])
        elif p.text.strip() in ("Ghi chú:", "Ghi chú"):
            set_para_text(p, "")
        elif p.text.strip().startswith("- Nhiệm vụ trọng tâm:") or p.text.strip().startswith("- Nhiệm vụ thường xuyên:") or p.text.strip().startswith("- Nhiệm vụ đột xuất"):
            set_para_text(p, "")
    nvs = NHIEM_VU[n["ten"]]
    for t in d.tables:
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if "Nội dung nhiệm vụ" in hdr or "Nội dung nhiệm vụ" in t.rows[0].cells[1].text:
            while (len(t.rows) - 1) < len(nvs):
                append_copied_row(t)
            for i, nv in enumerate(nvs):
                if i + 1 >= len(t.rows):
                    log("WARN KH %s thieu hang %d" % (n["ten"], i + 1))
                    break
                row = t.rows[i + 1]
                noi_dung, loai, chi_tieu, tien_do, san_pham, phoi_hop, cap_trinh, thoi_han, sl, hs = nv
                vals = [str(i + 1), noi_dung, loai, chi_tieu, tien_do, san_pham,
                        n["ten_hoa_thuong"], phoi_hop, cap_trinh, thoi_han]
                for c, v in enumerate(vals):
                    if c < len(row.cells):
                        set_cell_text(row.cells[c], v)
    d.save(path)
    log("patched KH " + os.path.basename(path))


def patch_07(path):
    d = Document(path)
    walk_docx(d, [
        ("không có đồng chí hoàn thành tốt nhiệm vụ, hoàn thành nhiệm vụ hoặc không hoàn thành nhiệm vụ.",
         "không có đồng chí hoàn thành xuất sắc nhiệm vụ, hoàn thành nhiệm vụ hoặc không hoàn thành nhiệm vụ."),
        ("07/07 đồng chí được đề xuất Hoàn thành xuất sắc nhiệm vụ",
         "07/07 đồng chí được đề xuất Hoàn thành tốt nhiệm vụ"),
    ])
    people = by_hoa()
    ds = [diem(n) for n in NHAN_SU]
    for p in d.paragraphs:
        if "Điểm bình quân của cấp ủy" in p.text:
            set_para_text(
                p,
                "Điểm bình quân của cấp ủy là %s điểm, thấp nhất %s điểm, cao nhất %s điểm."
                % (UT.so(round(sum(ds) / len(ds), 2)), UT.so(min(ds)), UT.so(max(ds))),
            )
    for t in d.tables:
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if "Họ và tên" in hdr and "Mức xếp loại" in hdr:
            for row in t.rows[1:]:
                ten = row.cells[1].text.strip()
                n = people.get(ten)
                if not n:
                    continue
                set_cell_text(row.cells[2], n["cv_dang"])
                set_cell_text(row.cells[3], UT.so(diem(n)))
                set_cell_text(row.cells[4], XEP_LOAI)
    d.save(path)
    log("patched 07 " + os.path.basename(path))


def patch_ghi_chu(path):
    d = Document(path)
    walk_docx(d, [
        ("Mức đề xuất: Hoàn thành xuất sắc nhiệm vụ cho 07/07 đồng chí cấp ủy.",
         "Mức đề xuất: Hoàn thành tốt nhiệm vụ cho 07/07 đồng chí cấp ủy."),
        ("(danh sách 7 đồng chí, mức HTXSNV, phân công).",
         "(danh sách 7 đồng chí, mức HTTNV, phân công)."),
        ("Rà 3 lĩnh vực Javis tự đề xuất: Luận (tổ chức - đảng viên), Xuân (tuyên giáo, đoàn thể), "
         "Lệ An (dân vận, thi đua). User đã chốt Nghĩa, Phương, Lan, Trung.",
         "Phân công đã căn quy chế 05-QĐ/CB ngày 10/2/2026 và file phân công chi ủy: "
         "Nghĩa phụ trách chung; Luận tuyên giáo và điều hành thay Bí thư; Trung kiểm tra giám sát; "
         "Phương đảng vụ - hồ sơ đảng viên; Lan tài chính Đảng, đảng phí, biên bản Chi ủy; "
         "Xuân dân vận và Đoàn Thanh niên; Lệ An công tác Đảng, quần chúng ưu tú."),
        ("HTXSNV cho 7/7 phải chứng minh được 100% nhiệm vụ hoàn thành và từ 30% vượt mức. "
         "CV 4988 cho phép cấp trên hạ mức nếu hồ sơ không chứng minh được.",
         "HTTNV cho 7/7: hoàn thành 100% nhiệm vụ, không đặt tiêu chí vượt mức. "
         "Không dùng mức HTXSNV nên không cần chứng minh 30% vượt mức."),
    ])
    # doan diem cu
    for p in d.paragraphs:
        if "Điểm ước tính từng đồng chí" in p.text:
            bits = "; ".join(
                "%s %s" % (n["ten_hoa_thuong"].split()[-1], UT.so(diem(n)))
                for n in NHAN_SU
            )
            avg = UT.so(round(sum(diem(n) for n in NHAN_SU) / 7.0, 2))
            set_para_text(
                p,
                "Điểm ước tính từng đồng chí (thang 100): %s. Bình quân %s. "
                "Anh sửa 8 ô nhập ở sheet Tinh KPI 2A-2B thì điểm ở Khung đánh giá tự chạy."
                % (bits, avg),
            )
    d.save(path)
    log("patched ghi chu " + os.path.basename(path))


# ---------- EXCEL ----------
def keep_font(cell, value, color=None):
    f = cell.font
    cell.value = value
    kw = dict(name=f.name, size=f.size, bold=f.bold, italic=f.italic,
              underline=f.underline, vertAlign=f.vertAlign, strike=f.strike)
    if color:
        kw["color"] = color
    elif f.color is not None:
        kw["color"] = f.color
    cell.font = Font(**kw)


def patch_02(path):
    wb = load_workbook(path)
    ws = wb.active
    for i, n in enumerate(NHAN_SU):
        r = 9 + i
        keep_font(ws.cell(r, 11), n["cv_dang"])
        keep_font(ws.cell(r, 15), n["linh_vuc_ngan"])
        keep_font(ws.cell(r, 16), XEP_LOAI)
    wb.save(path)
    log("patched 02 " + os.path.basename(path))


def patch_03(path):
    wb = load_workbook(path)
    ws = wb["Tong hop de xuat"]
    for i, n in enumerate(NHAN_SU):
        r = 8 + i
        nvs = NHIEM_VU[n["ten"]]
        profs = UT.phan_bo(n["stt"])
        amap = ANH_XA[n["ten"]]
        nv_text = NL.join("%d. %s (chỉ tiêu: %s)" % (j, x[0], x[2]) for j, x in enumerate(nvs, 1))
        kq_lines = []
        for j, x in enumerate(nvs, 1):
            tc = next((t for t, k in amap.items() if k == j - 1), None)
            pr = profs[tc - 1] if tc else UT.TOT
            kq_lines.append("%d. %s - Mức %s: số lượng %s%%, chất lượng %s%%, "
                            "tiến độ %s%%, năng lực lãnh đạo %s%%; KPI %s%%."
                            % (j, x[4], UT.NHAN[pr], UT.so(pr[0] * 100), UT.so(pr[1] * 100),
                               UT.so(pr[2] * 100), UT.so(pr[3] * 100), UT.so(UT.kpi(pr))))
        keep_font(ws.cell(r, 3), "%s / %s" % (n["cv_dang"], n["cv_cq"] or "Không"))
        keep_font(ws.cell(r, 4), nv_text)
        keep_font(ws.cell(r, 5), NL.join(kq_lines), color=UT.DO)
        keep_font(ws.cell(r, 6), XEP_LOAI)
        kr = r - 7 + 4
        ghi_chu = ('="Tổng điểm: "&TEXT(\'%s\'!N%d,"0.00")&"/100 '
                   '(nhóm A 30 + nhóm B "&TEXT(\'%s\'!M%d,"0.00")&")."'
                   % (TD_SHEET, kr, TD_SHEET, kr))
        keep_font(ws.cell(r, 7), ghi_chu, color=UT.DO)
    for row in ws.iter_rows(min_row=1, max_row=30, min_col=1, max_col=1):
        c = row[0]
        if c.value and isinstance(c.value, str) and ("Lưu ý:" in c.value or "Javis" in c.value):
            keep_font(c, "")
    if TD_SHEET in wb.sheetnames:
        k = wb[TD_SHEET]
        for i, n in enumerate(NHAN_SU):
            r = 5 + i
            profs = UT.phan_bo(n["stt"])
            for j, pr in enumerate(profs):
                keep_font(k.cell(r, 3 + j), UT.kpi(pr), color=UT.DO)
        # ghi chu dieu kien
        for row in k.iter_rows(min_row=1, max_row=20, min_col=1, max_col=2):
            for c in row:
                if not c.value or not isinstance(c.value, str):
                    continue
                if "Sửa ô KPI" in c.value or "thì Điểm nhóm B" in c.value or "SHEET THEO DÕI" in c.value:
                    keep_font(c, "")
                elif "HTXSNV" in c.value or "Mức đề xuất HTTNV" in c.value:
                    keep_font(c, "Mức đề xuất: hoàn thành tốt nhiệm vụ.")
    wb.save(path)
    log("patched 03 " + os.path.basename(path))


def patch_08(path):
    wb = load_workbook(path)
    ws = wb.active
    # hang 6 BT, 7 PBT, 8 CUV: cot F=HTTNV, G=HTXSNV
    keep_font(ws.cell(6, 6), 1)
    keep_font(ws.cell(6, 7), 0)
    keep_font(ws.cell(7, 6), 1)
    keep_font(ws.cell(7, 7), 0)
    keep_font(ws.cell(8, 6), 5)
    keep_font(ws.cell(8, 7), 0)
    wb.save(path)
    log("patched 08 " + os.path.basename(path))


def patch_tdg(path, n):
    wb = load_workbook(path)
    ws = wb["Khung đánh giá"]
    keep_font(ws["B8"], "Chức vụ Đảng: %s" % n["cv_dang"])
    nvs = NHIEM_VU[n["ten"]]
    amap = ANH_XA[n["ten"]]
    profs = UT.phan_bo(n["stt"])
    nv_theo_tc = []
    for idx, r in enumerate(ROWS_B):
        tc = idx + 1
        j = amap.get(tc)
        nv = nvs[j] if j is not None else None
        nv_theo_tc.append(nv)
        if nv is not None:
            muc_tieu = "%s\nChỉ tiêu: %s\nSản phẩm: %s\nThời hạn: %s" % (nv[0], nv[2], nv[4], nv[7])
        else:
            muc_tieu = CHUNG.get(tc, "")
        keep_font(ws.cell(r, 4), muc_tieu)
        keep_font(ws.cell(r, 5), UT.ket_qua_text(nv, profs[idx]), color=UT.DO)
    keep_font(ws["B50"], "II. Tự đề xuất xếp loại mức chất lượng: %s" % XEP_LOAI)
    note = ws.cell(row=65, column=1)
    if note.value and isinstance(note.value, str):
        keep_font(note, "")
    if KPI_SHEET in wb.sheetnames:
        k = wb[KPI_SHEET]
        for idx, (tt, diem_td, ten_tc) in enumerate(TIEU_CHI_B):
            r = idx + 4
            for off, val in enumerate(UT.o_nhap(nv_theo_tc[idx], profs[idx])):
                keep_font(k.cell(r, 3 + off), val, color=UT.DO)
        for row in k.iter_rows(min_row=1, max_row=40, min_col=1, max_col=2):
            for c in row:
                if not c.value or not isinstance(c.value, str):
                    continue
                if any(x in c.value for x in ("Ghi chú:", "Lưu ý", "Javis", "SHEET THEO DÕI",
                                              "Sửa ô", "CÁC SỐ MÀU ĐỎ", "Nhập đủ 8 ô")):
                    keep_font(c, "")
    wb.save(path)
    log("patched TDG " + os.path.basename(path))


def patch_checklist(path):
    txt = open(path, encoding="utf-8").read()
    repl = [
        ("4. Mức xếp loại đề xuất: **Hoàn thành xuất sắc nhiệm vụ cho cả 7 đồng chí**.",
         "4. Mức xếp loại đề xuất: **Hoàn thành tốt nhiệm vụ cho cả 7 đồng chí**."),
        ("| 2 | Trần Ngọc Luận | Phó Bí thư Chi bộ 13 | Công tác tổ chức - đảng viên |",
         "| 2 | Trần Ngọc Luận | Phó Bí thư Chi bộ 13 | Tuyên giáo; điều hành thay Bí thư |"),
        ("| 3 | Nguyễn Khắc Trung | Chi ủy viên, Tổ trưởng Tổ đảng 1 | Công tác kiểm tra, giám sát (UBKT) |",
         "| 3 | Nguyễn Khắc Trung | Chi ủy viên, Tổ trưởng Tổ đảng 1 | Kiểm tra, giám sát |"),
        ("| 4 | Nguyễn Đăng Phương | Chi ủy viên, Tổ trưởng Tổ đảng 2, Trợ lý Bí thư | Trợ lý, tổng hợp - báo cáo |",
         "| 4 | Nguyễn Đăng Phương | Chi ủy viên, Tổ trưởng Tổ đảng 2 | Đảng vụ, hồ sơ đảng viên |"),
        ("| 5 | Lê Thị Lan | Chi ủy viên | Công tác đảng phí, tài chính Chi bộ |",
         "| 5 | Lê Thị Lan | Chi ủy viên | Tài chính Đảng, đảng phí, biên bản Chi ủy |"),
        ("| 6 | Phạm Đồng Xuân | Chi ủy viên | Tuyên giáo, công tác đoàn thể |",
         "| 6 | Phạm Đồng Xuân | Chi ủy viên | Dân vận, Đoàn Thanh niên |"),
        ("| 7 | Nguyễn Thị Lệ An | Chi ủy viên | Dân vận, thi đua khen thưởng |",
         "| 7 | Nguyễn Thị Lệ An | Chi ủy viên | Công tác Đảng, quần chúng ưu tú |"),
        ("đã ghi HTXSNV", "đã ghi HTTNV"),
        ("7 HTXSNV", "7 HTTNV"),
        ("**HTXSNV cho 100% cấp ủy.** Công văn 4988 đặt điều kiện HTXSNV là hoàn thành 100% nhiệm vụ, trong đó từ 30% nhiệm vụ vượt mức. Hồ sơ phải chứng minh được tỷ lệ vượt mức này bằng số liệu KPI, nếu không Đảng ủy Thuế thành phố có cơ sở hạ mức. Nhiều cấp có thêm quy định khống chế tỷ lệ HTXSNV trong tổng số được đánh giá, cần kiểm lại Quy định 84-QĐ/TU và Hướng dẫn 07-HD/TU bản đầy đủ.",
         "**HTTNV cho 100% cấp ủy.** User chốt 03/11/2026: 7 đồng chí chỉ hoàn thành tốt nhiệm vụ. Hồ sơ không đặt tiêu chí vượt mức 30% của HTXSNV. Số KPI ước tính (chữ đỏ) đã hạ về mức tốt/đạt."),
    ]
    for a, b in repl:
        txt = txt.replace(a, b)
    open(path, "w", encoding="utf-8").write(txt)
    log("patched checklist")


KH_FILES = [
    "KH_QIII-2026_1_Le-Huu-Nghia.docx",
    "KH_QIII-2026_2_Tran-Ngoc-Luan.docx",
    "KH_QIII-2026_3_Nguyen-Khac-Trung.docx",
    "KH_QIII-2026_4_Nguyen-Dang-Phuong.docx",
    "KH_QIII-2026_5_Le-Thi-Lan.docx",
    "KH_QIII-2026_6_Pham-Dong-Xuan.docx",
    "KH_QIII-2026_7_Nguyen-Thi-Le-An.docx",
]
TDG_FILES = [
    "TDG_QIII-2026_1_Le-Huu-Nghia.xlsx",
    "TDG_QIII-2026_2_Tran-Ngoc-Luan.xlsx",
    "TDG_QIII-2026_3_Nguyen-Khac-Trung.xlsx",
    "TDG_QIII-2026_4_Nguyen-Dang-Phuong.xlsx",
    "TDG_QIII-2026_5_Le-Thi-Lan.xlsx",
    "TDG_QIII-2026_6_Pham-Dong-Xuan.xlsx",
    "TDG_QIII-2026_7_Nguyen-Thi-Le-An.xlsx",
]


def run_on(root):
    patch_01(os.path.join(root, "01_To_trinh_de_xuat_xep_loai.docx"))
    patch_04(os.path.join(root, "04_Quyet_dinh_phan_cong_va_danh_muc_san_pham.docx"))
    patch_07(os.path.join(root, "07_Bao_cao_ket_qua_QIII_2026.docx"))
    gc = os.path.join(root, "00_GHI_CHU_TRUOC_KHI_IN.docx")
    if os.path.exists(gc):
        os.remove(gc)
        log("removed " + gc)
    cl = os.path.join(root, "00_CHECKLIST_HO_SO.md")
    if os.path.exists(cl):
        patch_checklist(cl)
    for n, fn in zip(NHAN_SU, KH_FILES):
        patch_kh(os.path.join(root, "05_Ke_hoach_ca_nhan", fn), n)
    patch_02(os.path.join(root, "02_Danh_sach_trich_ngang.xlsx"))
    patch_03(os.path.join(root, "03_Bang_tong_hop_de_xuat.xlsx"))
    patch_08(os.path.join(root, "08_Thong_ke_cap_uy_vien.xlsx"))
    for n, fn in zip(NHAN_SU, TDG_FILES):
        patch_tdg(os.path.join(root, "06_Ban_tu_danh_gia", fn), n)


def copy_tree_files(src, dst):
    """Copy file ho so (khong copy __pycache__)."""
    skip = {"__pycache__"}
    for dirpath, dirnames, filenames in os.walk(src):
        dirnames[:] = [x for x in dirnames if x not in skip]
        rel = os.path.relpath(dirpath, src)
        target = dst if rel == "." else os.path.join(dst, rel)
        os.makedirs(target, exist_ok=True)
        for fn in filenames:
            if fn.endswith(".pyc"):
                continue
            shutil.copy2(os.path.join(dirpath, fn), os.path.join(target, fn))


if __name__ == "__main__":
    print("VAULT", VAULT)
    run_on(VAULT)
    print("--- copy sang laptop ---")
    copy_tree_files(VAULT, LAPTOP)
    gc_lt = os.path.join(LAPTOP, "00_GHI_CHU_TRUOC_KHI_IN.docx")
    if os.path.exists(gc_lt):
        os.remove(gc_lt)
        print("removed laptop", gc_lt)
    print("LAPTOP", LAPTOP)
    print("DIEM UOC TINH:")
    for n in NHAN_SU:
        print(" ", n["stt"], n["ten_hoa_thuong"], UT.so(diem(n)), n["linh_vuc_ngan"], XEP_LOAI_VT)
    print("DONE", len(LOG), "buoc")
