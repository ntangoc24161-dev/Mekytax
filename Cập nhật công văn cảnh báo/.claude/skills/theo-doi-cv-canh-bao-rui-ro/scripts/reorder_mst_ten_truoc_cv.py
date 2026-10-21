# -*- coding: utf-8 -*-
"""Doi thu tu cot: dua MST + Ten cong ty (don vi ban) len truoc thong tin cong van.

F0 moi (19 cot):
  STT | MSTBan | TenBan | SoCV | NgayCV | DonViBH | NoiDung | TamNgung | NhomND |
  DiaChi | NgayNhan | TinhTrang | GhiChu | File | NamCV | Trung | KhoaCV | CVDau | MSTDau

F1 moi (25 cot):
  STT | MSTBan | TenBan | SoCV | NgayCV | DonViBH | NoiDung | TenMua | MSTMua | TamNgung |
  KyKK | SoHD | DoanhThu | ThueGTGT | NhomND | DiaChi | NgayNhan | TinhTrang | GhiChu |
  File | NamCV | Trung | KhoaCV | CVDau | MSTDau
"""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SO_THEO_DOI = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO"
    r"\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
)
BACKUP_DIR = SO_THEO_DOI.parent / "Sao lưu"
SYNC_COPIES = [
    Path(r"D:\SECOND BRAIN\KIEM SOAT RUI RO\QUAN LY HD RUI RO\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"),
    Path(r"D:\Theo doi CVCBRR-ok\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"),
]

HEADER_ROWS = 7

FONT_TNR = Font(name="Times New Roman", size=10)
FONT_TITLE = Font(name="Times New Roman", size=15, bold=True, color="FFFFFF")
FONT_SUB = Font(name="Times New Roman", size=10, color="FFFFFF")
FONT_HINT = Font(name="Times New Roman", size=9, color="595959")
FONT_GRP = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_H1 = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_H2 = Font(name="Times New Roman", size=8, bold=True, color="E0E0E0")
ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)

FILL_TITLE = PatternFill("solid", fgColor="1B3A4B")
FILL_SUB = PatternFill("solid", fgColor="24556E")
FILL_CV = PatternFill("solid", fgColor="1F4E79")
FILL_BAN = PatternFill("solid", fgColor="C65911")
FILL_ND = PatternFill("solid", fgColor="375623")
FILL_MUA = PatternFill("solid", fgColor="0F6E56")
FILL_TN = PatternFill("solid", fgColor="8A6D1B")
FILL_HD = PatternFill("solid", fgColor="2E75B6")
FILL_XL = PatternFill("solid", fgColor="41719C")
FILL_DS = PatternFill("solid", fgColor="5B9BD5")
FILL_CODE = PatternFill("solid", fgColor="333F48")
FILL_STT = PatternFill("solid", fgColor="1B3A4B")


def log(msg: str) -> None:
    print(msg, flush=True)


def apply_fill(cell, fill, font, align=ALIGN_C):
    cell.fill = fill
    cell.font = font
    cell.alignment = align
    cell.border = THIN


def backup_file() -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = BACKUP_DIR / f"{SO_THEO_DOI.stem}.bak_reorder_mst_{stamp}{SO_THEO_DOI.suffix}"
    shutil.copy2(SO_THEO_DOI, bak)
    log(f"[OK] Backup: {bak}")
    return bak


def read_f0_rows(ws):
    """Doc F0 theo layout hien tai (co cot 2 cong thuc =G hack)."""
    rows = []
    for r in range(HEADER_ROWS + 1, ws.max_row + 1):
        stt = ws.cell(r, 1).value
        if stt is None and all(ws.cell(r, c).value in (None, "") for c in range(1, 21)):
            continue
        # Physical: 1=STT, 2=hack, 3=SoCV, 4=NgayCV, 5=DonViBH, 6=TenBan, 7=MSTBan,
        # 8=NoiDung, 9=TamNgung, 10=NhomND, 11=DiaChi, 12=NgayNhan, 13=TinhTrang,
        # 14=GhiChu, 15=File, 16=NamCV, 17=Trung, 18=KhoaCV, 19=CVDau, 20=MSTDau
        mst = ws.cell(r, 7).value
        ten = ws.cell(r, 6).value
        rows.append(
            {
                "mst": mst,
                "ten": ten,
                "so_cv": ws.cell(r, 3).value,
                "ngay_cv": ws.cell(r, 4).value,
                "don_vi_bh": ws.cell(r, 5).value,
                "noi_dung": ws.cell(r, 8).value,
                "tam_ngung": ws.cell(r, 9).value,
                "nhom_nd": ws.cell(r, 10).value,
                "dia_chi": ws.cell(r, 11).value,
                "ngay_nhan": ws.cell(r, 12).value,
                "tinh_trang": ws.cell(r, 13).value,
                "ghi_chu": ws.cell(r, 14).value,
                "file": ws.cell(r, 15).value,
                "nam_cv": ws.cell(r, 16).value,
                "trung": ws.cell(r, 17).value,
                "khoa_cv": ws.cell(r, 18).value,
                "cv_dau": ws.cell(r, 19).value,
                "mst_dau": ws.cell(r, 20).value,
            }
        )
    return rows


def read_f1_rows(ws):
    """Doc F1 theo layout cu: STT SoCV NgayCV DonViBH TenBan MSTBan ..."""
    rows = []
    for r in range(HEADER_ROWS + 1, ws.max_row + 1):
        stt = ws.cell(r, 1).value
        if stt is None and all(ws.cell(r, c).value in (None, "") for c in range(1, 26)):
            continue
        rows.append(
            {
                "mst_ban": ws.cell(r, 6).value,
                "ten_ban": ws.cell(r, 5).value,
                "so_cv": ws.cell(r, 2).value,
                "ngay_cv": ws.cell(r, 3).value,
                "don_vi_bh": ws.cell(r, 4).value,
                "noi_dung": ws.cell(r, 7).value,
                "ten_mua": ws.cell(r, 8).value,
                "mst_mua": ws.cell(r, 9).value,
                "tam_ngung": ws.cell(r, 10).value,
                "ky_kk": ws.cell(r, 11).value,
                "so_hd": ws.cell(r, 12).value,
                "doanh_thu": ws.cell(r, 13).value,
                "thue_gtgt": ws.cell(r, 14).value,
                "nhom_nd": ws.cell(r, 15).value,
                "dia_chi": ws.cell(r, 16).value,
                "ngay_nhan": ws.cell(r, 17).value,
                "tinh_trang": ws.cell(r, 18).value,
                "ghi_chu": ws.cell(r, 19).value,
                "file": ws.cell(r, 20).value,
                "nam_cv": ws.cell(r, 21).value,
                "trung": ws.cell(r, 22).value,
                "khoa_cv": ws.cell(r, 23).value,
                "cv_dau": ws.cell(r, 24).value,
                "mst_dau": ws.cell(r, 25).value,
            }
        )
    return rows


def build_f0(ws, recs):
    headers_vn = [
        "STT",
        "MST bên bán",
        "Tên công ty bán",
        "Số công văn",
        "Ngày công văn",
        "Đơn vị ban hành",
        "Nội dung chính",
        "Tạm ngưng kê khai",
        "Nhóm nội dung",
        "Địa chỉ bên bán",
        "Ngày nhận",
        "Tình trạng rà soát",
        "Ghi chú",
        "Đường dẫn file",
        "Năm công văn",
        "Dòng trùng",
        "Khóa công văn (ẩn)",
        "CV đầu tiên (ẩn)",
        "MST đầu tiên (ẩn)",
    ]
    headers_code = [
        "STT",
        "MSTBan",
        "TenBan",
        "SoCV",
        "NgayCV",
        "DonViBH",
        "NoiDung",
        "TamNgung",
        "NhomND",
        "DiaChi",
        "NgayNhan",
        "TinhTrang",
        "GhiChu",
        "File",
        "NamCV",
        "Trung",
        "KhoaCV",
        "CVDau",
        "MSTDau",
    ]
    last_col = len(headers_vn)
    last_letter = get_column_letter(last_col)

    ws.merge_cells(f"A1:{last_letter}1")
    ws.merge_cells(f"A2:{last_letter}2")
    ws.merge_cells(f"A4:{last_letter}4")
    ws.merge_cells("B5:C5")  # ĐƠN VỊ BÁN
    ws.merge_cells("D5:F5")  # THÔNG TIN CÔNG VĂN
    ws.merge_cells("I5:N5")  # THEO DÕI XỬ LÝ
    ws.merge_cells("O5:S5")  # ĐỐI SOÁT

    c = ws.cell(1, 1, "SỔ THEO DÕI CÔNG VĂN CẢNH BÁO RỦI RO — DANH SÁCH BÊN BÁN (F0)")
    apply_fill(c, FILL_TITLE, FONT_TITLE, ALIGN_L)
    c = ws.cell(2, 1, "Danh sách doanh nghiệp rủi ro bán hàng (F0) do cơ quan thuế và công an cảnh báo.")
    apply_fill(c, FILL_SUB, FONT_SUB, ALIGN_L)
    ws.cell(3, 1, "Đơn vị sử dụng sổ: Chi cục Thuế").font = FONT_TNR
    ws.cell(3, 4, "Người lập:").font = FONT_TNR
    c = ws.cell(
        4,
        1,
        "Đọc từ trái sang phải: MST → tên công ty bán → số công văn → ngày → đơn vị ban hành → nội dung → xử lý. "
        "Đã bỏ cột thông tin bên mua và chi tiết hóa đơn. Phông Times New Roman.",
    )
    c.font = FONT_HINT

    groups = [
        (1, "STT", FILL_STT),
        (2, "ĐƠN VỊ BÁN (F0)", FILL_BAN),
        (4, "THÔNG TIN CÔNG VĂN", FILL_CV),
        (7, "NỘI DUNG", FILL_ND),
        (8, "TẠM NGƯNG KÊ KHAI", FILL_TN),
        (9, "THEO DÕI XỬ LÝ & LƯU TRỮ", FILL_XL),
        (15, "THÔNG TIN ĐỐI SOÁT", FILL_DS),
    ]
    for col, val, fill in groups:
        apply_fill(ws.cell(5, col, val), fill, FONT_GRP)

    fills_row6 = (
        [FILL_STT]
        + [FILL_BAN] * 2
        + [FILL_CV] * 3
        + [FILL_ND]
        + [FILL_TN]
        + [FILL_XL] * 6
        + [FILL_DS] * 5
    )
    for i, (vn, code) in enumerate(zip(headers_vn, headers_code), 1):
        apply_fill(ws.cell(6, i, vn), fills_row6[i - 1], FONT_H1)
        apply_fill(ws.cell(7, i, code), FILL_CODE, FONT_H2)

    widths = {
        1: 8,
        2: 17,
        3: 55,
        4: 23,
        5: 15,
        6: 29,
        7: 45,
        8: 21,
        9: 29,
        10: 36,
        11: 15,
        12: 23,
        13: 46,
        14: 18,
        15: 11,
        16: 13,
        17: 17,
        18: 11,
        19: 11,
    }
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = w

    for idx, r in enumerate(recs, 1):
        row = idx + HEADER_ROWS
        ngay = r.get("ngay_cv")
        ngay_nhan = r.get("ngay_nhan")
        vals = [
            idx,
            r.get("mst"),
            r.get("ten"),
            r.get("so_cv"),
            ngay if isinstance(ngay, datetime) else ngay,
            r.get("don_vi_bh"),
            r.get("noi_dung"),
            r.get("tam_ngung"),
            r.get("nhom_nd"),
            r.get("dia_chi"),
            ngay_nhan if isinstance(ngay_nhan, datetime) else ngay_nhan,
            r.get("tinh_trang"),
            r.get("ghi_chu"),
            r.get("file"),
            r.get("nam_cv"),
            r.get("trung"),
            r.get("khoa_cv"),
            r.get("cv_dau"),
            r.get("mst_dau"),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row, c, v)
            cell.font = FONT_TNR
            if c in (5, 11) and isinstance(v, datetime):
                cell.number_format = "DD/MM/YYYY"
            if c in (1, 2, 15, 18, 19):
                cell.alignment = ALIGN_C
        if idx % 4000 == 0:
            log(f"  F0 wrote {idx}/{len(recs)}")

    ws.freeze_panes = "A8"
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 26
    ws.auto_filter.ref = f"A7:{last_letter}{HEADER_ROWS + len(recs)}"


def build_f1(ws, recs):
    headers_vn = [
        "STT",
        "MST bên bán",
        "Tên công ty bán",
        "Số công văn",
        "Ngày công văn",
        "Đơn vị ban hành",
        "Nội dung chính",
        "Tên công ty mua",
        "MST bên mua",
        "Tạm ngưng kê khai",
        "Kỳ kê khai",
        "Số hóa đơn",
        "Doanh thu (VNĐ)",
        "Tiền thuế GTGT (VNĐ)",
        "Nhóm nội dung",
        "Địa chỉ bên mua",
        "Ngày nhận",
        "Tình trạng rà soát",
        "Ghi chú",
        "Đường dẫn file",
        "Năm công văn",
        "Dòng trùng",
        "Khóa công văn (ẩn)",
        "CV đầu tiên (ẩn)",
        "MST đầu tiên (ẩn)",
    ]
    headers_code = [
        "STT",
        "MSTBan",
        "TenBan",
        "SoCV",
        "NgayCV",
        "DonViBH",
        "NoiDung",
        "TenMua",
        "MSTMua",
        "TamNgung",
        "KyKK",
        "SoHD",
        "DoanhThu",
        "ThueGTGT",
        "NhomND",
        "DiaChi",
        "NgayNhan",
        "TinhTrang",
        "GhiChu",
        "File",
        "NamCV",
        "Trung",
        "KhoaCV",
        "CVDau",
        "MSTDau",
    ]
    last_letter = "Y"
    ws.merge_cells("A1:Y1")
    ws.merge_cells("A2:Y2")
    ws.merge_cells("A4:Y4")
    ws.merge_cells("B5:C5")  # ĐƠN VỊ BÁN
    ws.merge_cells("D5:F5")  # THÔNG TIN CÔNG VĂN
    ws.merge_cells("H5:I5")  # ĐƠN VỊ MUA
    ws.merge_cells("K5:N5")  # CHI TIẾT HĐ
    ws.merge_cells("O5:T5")  # XỬ LÝ
    ws.merge_cells("U5:Y5")  # ĐỐI SOÁT

    c = ws.cell(1, 1, "SỔ THEO DÕI CÔNG VĂN CẢNH BÁO RỦI RO — DANH SÁCH BÊN MUA (F1)")
    apply_fill(c, FILL_TITLE, FONT_TITLE, ALIGN_L)
    c = ws.cell(
        2,
        1,
        "Danh sách doanh nghiệp mua hàng liên quan đến hóa đơn rủi ro (F1) cần rà soát, kiểm tra.",
    )
    apply_fill(c, FILL_SUB, FONT_SUB, ALIGN_L)
    ws.cell(3, 1, "Đơn vị sử dụng sổ: Chi cục Thuế").font = FONT_TNR
    ws.cell(3, 4, "Người lập:").font = FONT_TNR
    c = ws.cell(
        4,
        1,
        "Đọc từ trái sang phải: MST/tên bên bán → số công văn → ngày → đơn vị ban hành → nội dung → bên mua rà soát → xử lý. "
        "Phông Times New Roman.",
    )
    c.font = FONT_HINT

    groups = [
        (1, "STT", FILL_STT),
        (2, "ĐƠN VỊ BÁN LIÊN QUAN (F0)", FILL_BAN),
        (4, "THÔNG TIN CÔNG VĂN", FILL_CV),
        (7, "NỘI DUNG", FILL_ND),
        (8, "ĐƠN VỊ MUA — ĐỐI TƯỢNG RÀ SOÁT (F1)", FILL_MUA),
        (10, "TẠM NGƯNG KÊ KHAI", FILL_TN),
        (11, "CHI TIẾT HÓA ĐƠN", FILL_HD),
        (15, "THEO DÕI XỬ LÝ & LƯU TRỮ", FILL_XL),
        (21, "THÔNG TIN ĐỐI SOÁT", FILL_DS),
    ]
    for col, val, fill in groups:
        apply_fill(ws.cell(5, col, val), fill, FONT_GRP)

    fills_row6 = (
        [FILL_STT]
        + [FILL_BAN] * 2
        + [FILL_CV] * 3
        + [FILL_ND]
        + [FILL_MUA] * 2
        + [FILL_TN]
        + [FILL_HD] * 4
        + [FILL_XL] * 6
        + [FILL_DS] * 5
    )
    for i, (vn, code) in enumerate(zip(headers_vn, headers_code), 1):
        apply_fill(ws.cell(6, i, vn), fills_row6[i - 1], FONT_H1)
        apply_fill(ws.cell(7, i, code), FILL_CODE, FONT_H2)

    widths = {
        1: 8,
        2: 17,
        3: 45,
        4: 23,
        5: 15,
        6: 29,
        7: 40,
        8: 46,
        9: 17,
        10: 21,
        11: 15,
        12: 19,
        13: 16,
        14: 18,
        15: 29,
        16: 36,
        17: 15,
        18: 23,
        19: 46,
        20: 18,
        21: 11,
        22: 13,
        23: 17,
        24: 11,
        25: 11,
    }
    for i, w in widths.items():
        ws.column_dimensions[get_column_letter(i)].width = w

    for idx, r in enumerate(recs, 1):
        row = idx + HEADER_ROWS
        ngay = r.get("ngay_cv")
        ngay_nhan = r.get("ngay_nhan")
        vals = [
            idx,
            r.get("mst_ban"),
            r.get("ten_ban"),
            r.get("so_cv"),
            ngay if isinstance(ngay, datetime) else ngay,
            r.get("don_vi_bh"),
            r.get("noi_dung"),
            r.get("ten_mua"),
            r.get("mst_mua"),
            r.get("tam_ngung"),
            r.get("ky_kk"),
            r.get("so_hd"),
            r.get("doanh_thu"),
            r.get("thue_gtgt"),
            r.get("nhom_nd"),
            r.get("dia_chi"),
            ngay_nhan if isinstance(ngay_nhan, datetime) else ngay_nhan,
            r.get("tinh_trang"),
            r.get("ghi_chu"),
            r.get("file"),
            r.get("nam_cv"),
            r.get("trung"),
            r.get("khoa_cv"),
            r.get("cv_dau"),
            r.get("mst_dau"),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row, c, v)
            cell.font = FONT_TNR
            if c in (5, 17) and isinstance(v, datetime):
                cell.number_format = "DD/MM/YYYY"
            if c in (1, 2, 9, 21, 24, 25):
                cell.alignment = ALIGN_C
        if idx % 2000 == 0:
            log(f"  F1 wrote {idx}/{len(recs)}")

    ws.freeze_panes = "A8"
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 26
    ws.auto_filter.ref = f"A7:Y{HEADER_ROWS + len(recs)}"


def sync_copies(src: Path):
    for dest in SYNC_COPIES:
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
            log(f"[OK] Sync -> {dest}")
        except Exception as e:
            log(f"[WARN] Khong sync duoc {dest}: {e}")


def main():
    if not SO_THEO_DOI.is_file():
        raise FileNotFoundError(SO_THEO_DOI)

    backup_file()
    log(f"[READ] {SO_THEO_DOI}")
    wb_old = load_workbook(SO_THEO_DOI, data_only=False)
    f0_recs = read_f0_rows(wb_old["F0"])
    f1_recs = read_f1_rows(wb_old["F1"])
    wb_old.close()
    log(f"[READ] F0={len(f0_recs)} dong, F1={len(f1_recs)} dong")

    wb = Workbook()
    # Remove default sheet
    default = wb.active
    wb.remove(default)

    ws0 = wb.create_sheet("F0")
    log("[BUILD] F0 ...")
    build_f0(ws0, f0_recs)

    ws1 = wb.create_sheet("F1")
    log("[BUILD] F1 ...")
    build_f1(ws1, f1_recs)

    log(f"[SAVE] {SO_THEO_DOI}")
    wb.save(SO_THEO_DOI)
    wb.close()

    # Verify
    wb2 = load_workbook(SO_THEO_DOI, read_only=True, data_only=True)
    for name in ("F0", "F1"):
        ws = wb2[name]
        headers = [ws.cell(6, c).value for c in range(1, 8)]
        sample = [ws.cell(8, c).value for c in range(1, 8)]
        log(f"[CHECK] {name} headers6={headers}")
        log(f"[CHECK] {name} row8={sample} max_row={ws.max_row}")
    wb2.close()

    sync_copies(SO_THEO_DOI)
    log("[DONE] Da doi cau truc: MST + Ten cong ty len truoc thong tin cong van.")


if __name__ == "__main__":
    main()
