# -*- coding: utf-8 -*-
"""Bo sung va chuan hoa thong tin 67 MST tu anh vao sheet F0 SoTheoDoi."""
from __future__ import annotations

import re
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
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

FONT_TNR = Font(name="Times New Roman", size=10)
HEADER_ROWS = 7
NCOLS_F0 = 19


def log(msg: str) -> None:
    print(msg, flush=True)


NEW_RECORDS = [
    {
        "mst": "33603967272",
        "ten": "CÔNG TY TNHH XÂY DỰNG VƯỢNG THỊNH PHÁT",
        "so_cv": "5743/TCS5-KTr1",
        "ngay_cv": datetime(2025, 11, 26),
        "don_vi_bh": "TCS5 Đồng Nai",
        "dia_chi": "Số 375, đường Tỉnh 769, ấp 3, xã Long Thành, tỉnh Đồng Nai",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. TB người nộp thuế không hoạt động tại địa chỉ đã đăng ký",
    },
    {
        "mst": "03154432405",
        "ten": "Công ty TNHH thiết kế và Xây Dựng Đại Phước Đạt",
        "so_cv": "6370/CCTQ8-KTr3",
        "ngay_cv": datetime(2023, 11, 29),
        "don_vi_bh": "Chi cục thuế Quận 8",
        "dia_chi": "44 Tạ quang Bửu, phường 1, quận 8, Tp. HCM",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Cảnh báo DN có dấu hiệu rủi ro cao về thuế, hóa đơn.",
    },
    {
        "mst": "0316347086",
        "ten": "Cong Ty TNHH CASSAVA Viet Nam.",
        "so_cv": "896/TB-CCT",
        "ngay_cv": datetime(2023, 8, 16),
        "don_vi_bh": "Chi Cục thuế huyện Xuân Lộc",
        "dia_chi": "tổ 22, ấp Suối Cát 1, xã Suối Cát, huyện Xuân Lộc, tỉnh Đồng Nai",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. BDCKD",
    },
    {
        "mst": "0316420970",
        "ten": "Công ty TNHH Thương Mạí Dịch Vụ Đầu Tư Kim Hoa",
        "so_cv": "24319/CCTGV-Kt3",
        "ngay_cv": datetime(2024, 10, 30),
        "don_vi_bh": "CCT Quận Gò Vấp",
        "dia_chi": "87 Đường 51, Phường 14, quận Gò vấp, Thành phố Hồ Chí Minh.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. DN rủi ro hóa đơn",
    },
    {
        "mst": "03165346",
        "ten": "Công ty TNHH Thương mại Dịch vụ và Xây dựng",
        "so_cv": "8184/CCTQ11-KT1",
        "ngay_cv": datetime(2023, 10, 3),
        "don_vi_bh": "CCT Quận 11",
        "dia_chi": "133/3/5, đường Bình Thới, Phường 11, Quận 11, TPHCM",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Cảnh báo rủi ro",
    },
    {
        "mst": "03166301980",
        "ten": "CÔNG TY TNHH ĐẦU TƯNVÀ SẢN XUẤT TM MỸ LINH",
        "so_cv": "7595/CCTKVQ7NB-KTr2",
        "ngay_cv": datetime(2023, 10, 18),
        "don_vi_bh": "Quận 7 - KV QUB-NBE",
        "dia_chi": "",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: CBRR.",
    },
    {
        "mst": "03171618620",
        "ten": "CÔNG TY TNHH TMDV VẬN CHUYỂN HÀNG HÓA VÀ KINH DOANH PHỤ TÙNG Ô TÔ VIỆT NAM",
        "so_cv": "5370/CCTQ8-KTr1",
        "ngay_cv": datetime(2023, 10, 19),
        "don_vi_bh": "CCT Quận 8",
        "dia_chi": "",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: CBRR.",
    },
    {
        "mst": "03178148852",
        "ten": "CÔNG TY TNHH TM DV ĐẦU TƯ ĐẠI NAM TIẾN",
        "so_cv": "707/CCTBT-KTT3",
        "ngay_cv": datetime(2024, 2, 2),
        "don_vi_bh": "CCTBT-KTT3",
        "dia_chi": "",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: CBRR.",
    },
    {
        "mst": "03505430631",
        "ten": "CTY TNHH TMDV XD TOÀN Á",
        "so_cv": "595/TCCS7-QLDN1",
        "ngay_cv": datetime(2026, 3, 16),
        "don_vi_bh": "TCS7-HCM",
        "dia_chi": "",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: CBRR.",
    },
    {
        "mst": "3600961850",
        "ten": "CNKD Nguyễn Quốc Bình",
        "so_cv": "2749/TB-CCTKV",
        "ngay_cv": datetime(2023, 12, 15),
        "don_vi_bh": "Chi Cục Thuế khu vực Định Quán - Tân Phú (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "Khu phố 114, TT. Định Quán, huyện Định Quán, tỉnh Đồng Nai.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Thông báo vv NNT không HĐ tại địa chỉ đã đăng ký (Công ty TNHH TM ATP VN)",
    },
    {
        "mst": "36033455088",
        "ten": "Công ty TNHH Sản xuất Thương mại Hùng Lâm",
        "so_cv": "2526/PC03-Đ9",
        "ngay_cv": datetime(2023, 3, 24),
        "don_vi_bh": "Công an tỉnh Đồng Nai (PC03)",
        "dia_chi": "Số 54/3 ấp Lộ Đức Xã Hố Nai 3 huyện Trảng Bom Đồng Nai",
        "noi_dung": "Rủi ro do Công an chuyển",
        "nhom_nd": "Rủi ro do Công an chuyển",
        "ghi_chu": "Nguồn: HDRR -cuc. PC 03 - 59 cty",
    },
    {
        "mst": "3603790716",
        "ten": "Công Ty TNHH Cung ứng Lao Động Thành Nam Phát.",
        "so_cv": "20939/CCTKV-KT1",
        "ngay_cv": datetime(2023, 10, 3),
        "don_vi_bh": "Chi Cục thuế khu vực Biên Hòa - Vĩnh Cửu (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc.",
    },
    {
        "mst": "3702437614",
        "ten": "Công ty TNHH Một thành viên Tiệp Cường.",
        "so_cv": "7277/CCT-KTR1",
        "ngay_cv": datetime(2023, 9, 28),
        "don_vi_bh": "Chi cục thuế Thành phố Thủ Dầu Một",
        "dia_chi": "số 1350 , Đại Lộ Bình Dương, Tổ 40, phường Định Hòa, thành phố Thủ Dầu Một, tỉnh Bình Dương",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc.",
    },
    {
        "mst": "5100488855",
        "ten": "Công ty TNHH Khai thác khoáng sản Sông Lô.",
        "so_cv": "680/TB-CTHGI",
        "ngay_cv": datetime(2023, 11, 22),
        "don_vi_bh": "Cục Thuế tỉnh Hà Giang",
        "dia_chi": "Tổ 6, TT. Vị Xuyên, huyện Vị Xuyên, tỉnh Hà Giang",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Bỏ ĐCKD",
    },
    {
        "mst": "8394832609-001",
        "ten": "CNKD Lê Hoàng Mỹ",
        "so_cv": "2722/TB-CCTKV",
        "ngay_cv": datetime(2023, 12, 14),
        "don_vi_bh": "Chi Cục Thuế khu vực Định Quán - Tân Phú (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "Khu phố Hiệp Tâm 1, TT. Định Quán, huyện Định Quán, tỉnh Đồng Nai.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Thông báo vv NNT không HĐ tại địa chỉ đã đăng ký (Công ty TNHH TM ATP VN)",
    },
    {
        "mst": "8518974207-001",
        "ten": "CNKD Nguyễn Thị Hóa",
        "so_cv": "2726/TB-CCTKV",
        "ngay_cv": datetime(2023, 12, 14),
        "don_vi_bh": "Chi Cục Thuế khu vực Định Quán - Tân Phú (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "Khu phố Hiệp Tâm 1, TT. Định Quán, huyện Định Quán, tỉnh Đồng Nai.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Thông báo vv NNT không HĐ tại địa chỉ đã đăng ký (Công ty TNHH TM ATP VN)",
    },
    {
        "mst": "8543610154-001",
        "ten": "CNKD Trần Anh Kỳ.",
        "so_cv": "2724/TB-CCTKV",
        "ngay_cv": datetime(2023, 12, 14),
        "don_vi_bh": "Chi Cục Thuế khu vực Định Quán - Tân Phú (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "Khu phố Hiệp Tâm 1, TT. Định Quán, huyện Định Quán, tỉnh Đồng Nai.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Thông báo vv NNT không HĐ tại địa chỉ đã đăng ký (Công ty TNHH TM ATP VN)",
    },
    {
        "mst": "8552597069",
        "ten": "CNKD Phạm Thành Phát",
        "so_cv": "2719/TB-CCTKV",
        "ngay_cv": datetime(2023, 12, 14),
        "don_vi_bh": "Chi Cục Thuế khu vực Định Quán - Tân Phú (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "Khu phố Hiệp Tâm 1, TT. Định Quán, huyện Định Quán, tỉnh Đồng Nai.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Thông báo vv NNT không HĐ tại địa chỉ đã đăng ký (Công ty TNHH TM ATP VN)",
    },
    {
        "mst": "8615961588-001",
        "ten": "CNKD Thiều Thị Thu",
        "so_cv": "2747/TB-CCTKV",
        "ngay_cv": datetime(2023, 12, 15),
        "don_vi_bh": "Chi Cục Thuế khu vực Định Quán - Tân Phú (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "Khu phố Hiệp Lợi, TT. Định Quán, huyện Định Quán, tỉnh Đồng Nai.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Thông báo vv NNT không HĐ tại địa chỉ đã đăng ký (Công ty TNHH TM ATP VN)",
    },
    {
        "mst": "8794510927-001",
        "ten": "CNKD Nguyễn Thị Ngọc Quý",
        "so_cv": "2723/TB-CCTKV",
        "ngay_cv": datetime(2023, 12, 14),
        "don_vi_bh": "Chi Cục Thuế khu vực Định Quán - Tân Phú (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "Khu phố Hiệp Tâm 1, TT. Định Quán, huyện Định Quán, tỉnh Đồng Nai.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Thông báo vv NNT không HĐ tại địa chỉ đã đăng ký (Công ty TNHH TM ATP VN)",
    },
    {
        "mst": "8794520650-001",
        "ten": "CNKD Vũ Thị Kim Loan",
        "so_cv": "2720/TB-CCTKV",
        "ngay_cv": datetime(2023, 12, 14),
        "don_vi_bh": "Chi Cục Thuế khu vực Định Quán - Tân Phú (tỉnh Đồng Nai - Mã Ủy ban)",
        "dia_chi": "Khu phố Hiệp Tâm 1, TT. Định Quán, huyện Định Quán, tỉnh Đồng Nai.",
        "noi_dung": "Cảnh báo rủi ro về thuế và hóa đơn",
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "ghi_chu": "Nguồn: HDRR -cuc. Thông báo vv NNT không HĐ tại địa chỉ đã đăng ký (Công ty TNHH TM ATP VN)",
    },
]

PATCH_MAP = {
    "2301106934": {
        "ten": "Công ty TNHH Manna Vina",
        "so_cv": "7550/CCTKV-KTR2",
        "don_vi_bh": "CCT Khu Vực Từ Sơn-Yên Phong",
        "dia_chi": "Thôn Nghiêm Xá-Thị trấn Chờ- huyện Yên Phong, tỉnh Bắc Ninh",
    },
    "2301133663": {
        "ten": "Công ty TNHH TM&DV VNC Bắc Ninh",
        "so_cv": "7552/CCTKV-KTR2",
        "don_vi_bh": "CCT Khu Vực Từ Sơn-Yên Phong",
        "dia_chi": "Thôn Nghiêm Xá-Thị trấn Chờ- huyện Yên Phong, tỉnh Bắc Ninh",
    },
    "4101614268": {
        "ten": "Công ty TNHH TM DV TH Anh Khoa",
        "so_cv": "575/CCTKV-KTr",
        "don_vi_bh": "CCT TP Thủ Đức",
        "dia_chi": "05 Đường 17, khu phố 1, Phường Bình Trưng Đông, TP. Thủ Đức, TP. HCM",
    },
    "6001434956": {
        "so_cv": "1402/CCTKV-KTR",
        "don_vi_bh": "CHI CỤC THUẾ KHU VỰC KRÔNG ANA - CƯ KUIN",
        "dia_chi": "Thôn Quỳnh Tân 2, Thị trấn Buôn Trấp, huyện Krông Ana, tỉnh Đắk Lắk",
    },
    "6300349875": {
        "so_cv": "618/CCTKV-Ktra",
        "don_vi_bh": "CCT KV IV",
        "dia_chi": "Ấp Phú Nhơn, xã Đông Phú, huyện Châu Thành, tỉnh Hậu Giang",
    },
    "0402243404-002": {
        "so_cv": "4330/CTNTH-TTKT",
        "don_vi_bh": "Cục thuế Ninh Thuận",
        "dia_chi": "số 72 Đường Lê Duân, Thị trấn Tân Sơn, Huyện Ninh Sơn, tỉnh Ninh Thuận",
    },
    "0313936710": {
        "so_cv": "8203/CCTKVQ7NB-KTr4",
        "don_vi_bh": "CCT Nhà Bè Quận 7",
        "dia_chi": "114 Phan Văn Bảy, ấp 1, xã Hiệp Phước, huyện Nhà Bè",
    },
    "2802532203": {
        "so_cv": "1380/CCT-KT",
        "don_vi_bh": "CCT Thọ Xuân -Thường Xuân",
        "dia_chi": "Thôn 3, xã Thọ Lâm, huyện Thọ Xuân, tỉnh Thanh Hóa",
    },
    "3603431178": {
        "so_cv": "27396/CCTKV-Kt1",
        "don_vi_bh": "CCT Biên Hòa -Vĩnh Cửu",
        "dia_chi": "sổ D28-29, Khu tái định cư Bửu Long, Phường Bửu Long, Tp.Biên Hoà, tỉnh Đồng Nai.",
    },
}


def backup_current(excel_path: Path) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = BACKUP_DIR / f"{excel_path.stem}.bak_{stamp}{excel_path.suffix}"
    shutil.copy2(excel_path, dest)
    return dest


def main():
    log(f"So theo doi: {SO_THEO_DOI}")
    local_source = SYNC_COPIES[0] if SYNC_COPIES[0].is_file() else SO_THEO_DOI
    if not local_source.is_file():
        log(f"[ERROR] Khong tim thay file so theo doi: {local_source}")
        return 1

    bak = backup_current(local_source)
    log(f"[1] Da tao ban sao luu tai: {bak}")

    tmp_working = Path(r"C:\Users\NGOC\AppData\Local\Temp\SoTheoDoi_updating.xlsx")
    shutil.copy2(local_source, tmp_working)
    log(f"[2] Sao chep sang temp de xu ly: {tmp_working}")

    wb = openpyxl.load_workbook(tmp_working)
    ws = wb["F0"]

    log(f"[3] Dang doc du lieu F0 hien tai (max_row = {ws.max_row})...")
    existing_rows = []
    for r in range(HEADER_ROWS + 1, ws.max_row + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, NCOLS_F0 + 1)]
        if all(v is None or str(v).strip() == "" for v in row_vals):
            continue
        existing_rows.append(row_vals)

    log(f"    So dong F0 hien tai: {len(existing_rows)}")

    # Apply patches to existing rows
    patched_count = 0
    for r in existing_rows:
        mst = str(r[5] or "").strip()
        if mst.endswith("-") and len(mst) == 11:
            mst = mst[:-1]
            r[5] = mst
            patched_count += 1
        if mst in PATCH_MAP:
            p = PATCH_MAP[mst]
            if "ten" in p and (r[4] in (None, "", "0x2a") or p["ten"]):
                r[4] = p["ten"]
            if "so_cv" in p:
                r[1] = p["so_cv"]
            if "don_vi_bh" in p and (r[3] in (None, "") or p["don_vi_bh"]):
                r[3] = p["don_vi_bh"]
            if "dia_chi" in p and not r[9]:
                r[9] = p["dia_chi"]
            patched_count += 1

    log(f"    Da patch/chuan hoa {patched_count} thong tin tren cac dong F0 co san.")

    # Convert new records to row format
    added_rows = []
    for nr in NEW_RECORDS:
        r = [None] * NCOLS_F0
        r[1] = nr.get("so_cv")
        r[2] = nr.get("ngay_cv")
        r[3] = nr.get("don_vi_bh")
        r[4] = nr.get("ten")
        r[5] = nr.get("mst")
        r[6] = nr.get("noi_dung")
        r[8] = nr.get("nhom_nd")
        r[9] = nr.get("dia_chi")
        r[12] = nr.get("ghi_chu")
        added_rows.append(r)

    log(f"[4] Da chuan bi {len(added_rows)} dong moi de bo sung.")

    all_rows = existing_rows + added_rows

    # Sort records by NgayCV, SoCV, MST
    def sort_key(row):
        ngay = row[2]
        if not isinstance(ngay, datetime):
            ngay = datetime(1900, 1, 1)
        socv = str(row[1] or "")
        mst = str(row[5] or "")
        return (socv, ngay, mst)

    all_rows.sort(key=sort_key)

    # Recompute flags and STT
    seen_cv = set()
    seen_mst = set()
    for idx, row in enumerate(all_rows, 1):
        row[0] = idx  # STT
        socv = str(row[1] or "").strip().upper()
        mst = str(row[5] or "").strip()
        ngay = row[2]
        nam = ngay.year if isinstance(ngay, datetime) else None
        row[14] = nam  # NamCV
        if socv and nam:
            row[16] = f"{str(row[1]).strip()}_{nam}"  # KhoaCV
        elif socv:
            row[16] = str(row[1]).strip()

        if socv and socv not in seen_cv:
            seen_cv.add(socv)
            row[17] = 1  # CVDau
        else:
            row[17] = None

        if mst and mst not in seen_mst:
            seen_mst.add(mst)
            row[18] = 1  # MSTDau
        else:
            row[18] = None

    log(f"[5] Ghi lai toan bo sheet F0 ({len(all_rows)} dong)...")
    for idx, row in enumerate(all_rows, 1):
        r_idx = HEADER_ROWS + idx
        for c_idx in range(1, NCOLS_F0 + 1):
            val = row[c_idx - 1]
            cell = ws.cell(r_idx, c_idx, val)
            cell.font = FONT_TNR
            if c_idx in (3, 11) and isinstance(val, datetime):
                cell.number_format = "DD/MM/YYYY"
            elif c_idx in (1, 6, 15, 18, 19):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in (2, 4, 5, 7, 8, 9, 10, 12, 13, 14, 16, 17):
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Update auto filter
    last_row = HEADER_ROWS + len(all_rows)
    ws.auto_filter.ref = f"A6:S{last_row}"

    log(f"[6] Luu file tam: {tmp_working}")
    wb.save(tmp_working)
    wb.close()

    # Sync to primary destination and sync copies
    log(f"[7] Dong bo sang cac vi tri luu tru...")
    try:
        SO_THEO_DOI.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(tmp_working, SO_THEO_DOI)
        log(f"    [OK] Da ghi sang {SO_THEO_DOI} ({SO_THEO_DOI.stat().st_size} bytes)")
    except Exception as e:
        log(f"    [WARN] Khong ghi duoc sang Google Drive ({SO_THEO_DOI}): {e}")

    for dest in SYNC_COPIES:
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(tmp_working, dest)
            log(f"    [OK] Da dong bo sang {dest}")
        except Exception as e:
            log(f"    [WARN] Khong dong bo duoc {dest}: {e}")

    log(f"\n[HOAN TAT] Tong so dong F0 sau cap nhat: {len(all_rows)} (tang {len(added_rows)} dong).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
