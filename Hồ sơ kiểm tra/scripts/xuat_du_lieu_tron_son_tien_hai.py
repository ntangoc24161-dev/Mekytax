# -*- coding: utf-8 -*-
"""
Xuat file du lieu tron (mail-merge data source) cho ho so ve vu Son Tien Hai
va sao ban TONG HOP da cap nhat vao output cua brain.

Chay: python scripts/xuat_du_lieu_tron_son_tien_hai.py
"""
from __future__ import annotations

import collections
import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

sys.stdout.reconfigure(encoding="utf-8")

VU = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
          r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\11. SON TIEN HAI\0. BIEN BAN")
LIVE = VU / "TONG HOP.docx"
BRAIN = Path(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra")
OUT = BRAIN / "output" / "3603859012_HO-SO-VE_son-tien-hai"

FIELD_RE = re.compile(r'MERGEFIELD\s+"?([^"\s\\]+)')
WT = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)

# Gia tri du lieu tron cua vu Son Tien Hai (MST 3603859012).
# Nguon: DB NNT.xls, so JAVIS PRO (dong 16), Bang tinh.xlsx, DI-01..DI-07.
DU_LIEU: dict[str, str] = {
    "Tên_công_ty": "Công ty TNHH Thương mại Sản xuất Sơn Tiền Hải",
    "MST": "3603859012",
    "địa_chỉ": "Số 23, Tổ 1, KP 3A, Phường Long Hưng, Thành phố Đồng Nai",
    "NNKD": "Bán buôn vật liệu, thiết bị lắp đặt khác trong xây dựng",
    "GĐ": "Phan Thế Hanh",
    "CCCD": "034083000616",
    "ngày_cấp": "",
    "Nơi_cấp": "",
    "KT": "",
    "CCCDKT": "",
    "TĐ": "Nguyễn Thị Ánh Ngọc",
    "CVTĐ": "Chuyên viên",
    "TV": "Nguyễn Thành Đồng",
    "CVTV": "Kiểm tra viên thuế",
    "Thời_kỳ_kiểm_tra": "Từ năm 2022 đến tháng 10/2026",
    "Thời_hạn_kiểm_tra_": "20 ngày",
    "Nội_dung_kiểm_tra": ("- Giá vốn hàng bán; - Chi phí quản lý kinh doanh; "
                          "- Doanh thu bán hàng và cung cấp dịch vụ; "
                          "- Thuế GTGT phải nộp."),
    "Phân_công_trưởng_đoàn": ("- Giá vốn hàng bán\n"
                              "- Chi phí quản lý kinh doanh"),
    "Phân_công_thành_viên": ("- Doanh thu bán hàng và cung cấp dịch vụ\n"
                             "- Thuế GTGT phải nộp"),
    "QĐKT": "",
    "QĐGS": "",
    "ngày_QĐ": "",
    "Ngày_báo_cáo": "09/11/2026",
    "Tài_khoản_số": "",
    "Ma_CQT": "1140968",
    "CQT": "Phường Long Hưng - Thuế cơ sở 3 Thành phố Đồng Nai",
    "Giảm_lỗ": "809.868.861",
}

# Cac moc va so lieu chot cua vu, de ra ngoai file du lieu cho de doi chieu
CHOT = [
    ("Công bố / bắt đầu kiểm tra", "30/10/2026", "sổ JAVIS PRO dòng 16"),
    ("Kết thúc kiểm tra", "03/11/2026", "sổ JAVIS PRO dòng 16"),
    ("Ký Biên bản kiểm tra", "04/11/2026", "sổ JAVIS PRO dòng 16"),
    ("Báo cáo kết quả", "09/11/2026", "sổ JAVIS PRO dòng 16"),
    ("Thời hạn ghi trên QĐKT", "20 ngày", "sổ JAVIS PRO (CẦN CHỐT LẠI)"),
    ("Truy thu thuế TNDN", "13.052.423", "Bảng tính, sheet DTCHEP"),
    ("Trong đó năm 2025", "10.180.180", "chênh giá vốn 59.883.415 x 17%"),
    ("Trong đó năm 2026", "2.872.243", "chênh 826.764.409 trừ giảm lỗ, x 17%"),
    ("Truy thu thuế GTGT", "0", "chênh lệch GTGT bằng 0 cả 5 kỳ"),
    ("Tiền chậm nộp", "662.730", "0,03%/ngày x 217 ngày, chỉ năm 2025"),
    ("Phạt khai sai 20%", "2.610.485", "khoản 1 Điều 16 NĐ 125/2020"),
    ("Phạt thủ tục thuế", "0", "không phát sinh"),
    ("Tổng tiền phạt", "3.273.215", "662.730 + 2.610.485"),
    ("Tổng truy thu và phạt", "16.325.638", "13.052.423 + 3.273.215"),
    ("Giảm lỗ", "809.868.861", "lỗ lũy kế bị hấp thụ năm 2026"),
    ("Giảm khấu trừ", "0", "không phát sinh"),
]


def main() -> None:
    assert LIVE.exists(), f"khong thay {LIVE}"
    OUT.mkdir(parents=True, exist_ok=True)

    xml = zipfile.ZipFile(LIVE).read("word/document.xml").decode("utf8")
    dem = collections.Counter(FIELD_RE.findall(xml))
    print(f"[doc] {sum(dem.values())} the MERGEFIELD, {len(dem)} truong khac nhau")

    thieu = [f for f in dem if f not in DU_LIEU]
    if thieu:
        print("[!] truong co trong phoi ma chua co du lieu:", thieu)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DU_LIEU_TRON"
    head = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="2F5496")
    cols = sorted(dem, key=lambda f: -dem[f])
    for j, f in enumerate(cols, start=1):
        c = ws.cell(row=1, column=j, value=f)
        c.font, c.fill = head, fill
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.cell(row=2, column=j, value=DU_LIEU.get(f, ""))
        ws.column_dimensions[c.column_letter].width = max(14, min(42, len(f) + 8))
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("SO_LIEU_CHOT")
    for j, t in enumerate(("Chỉ tiêu", "Giá trị (đồng / mốc)", "Nguồn"), start=1):
        c = ws2.cell(row=1, column=j, value=t)
        c.font, c.fill = head, fill
    for i, (a, b, c3) in enumerate(CHOT, start=2):
        ws2.cell(row=i, column=1, value=a)
        ws2.cell(row=i, column=2, value=b)
        ws2.cell(row=i, column=3, value=c3)
    for col, w in zip("ABC", (34, 24, 48)):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"

    ws3 = wb.create_sheet("SO_LAN_DUNG")
    ws3.append(["Trường MERGEFIELD", "Số lần xuất hiện trong TONG HOP"])
    for f, n in dem.most_common():
        ws3.append([f, n])
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 34

    xlsx = OUT / "du-lieu-tron_son-tien-hai_3603859012.xlsx"
    wb.save(xlsx)
    print(f"[out] {xlsx.name}")

    ban_sao = OUT / "TONG HOP.docx"
    shutil.copy2(LIVE, ban_sao)
    print(f"[out] {ban_sao.name}")
    for ten in ("QD XPHC Cong ty Son Tien Hai.docx",):
        src = VU / ten
        if src.exists():
            shutil.copy2(src, OUT / ten)
            print(f"[out] {ten}")

    (OUT / "du-lieu-tron.json").write_text(
        json.dumps({"mst": "3603859012", "truong_tron": DU_LIEU,
                    "so_lieu_chot": {a: b for a, b, _ in CHOT},
                    "so_lan_dung": dict(dem)},
                   ensure_ascii=False, indent=2), encoding="utf8")
    print("[out] du-lieu-tron.json")


if __name__ == "__main__":
    main()
