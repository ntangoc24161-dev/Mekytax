# -*- coding: utf-8 -*-
"""Cap nhat ten/MST/dia chi/noi dung cho 24 dong lo 22-26/09 (ghi de, khong chen them)."""
from __future__ import annotations

import json
import re
from pathlib import Path

import openpyxl

from update_so_theo_doi import DEFAULT_EXCEL_PATH

SCRIPTS = Path(__file__).resolve().parent
RECORDS_PATH = SCRIPTS / "records_batch_22_26.json"

# Key: SoCV chuan hoa
MAP = {
    "18811/TCS3-KTR1": {
        "ten_ban": "Công ty TNHH Thương mại và Dịch vụ Cung ứng nguồn nhân lực M D",
        "mst_ban": "",
        "dia_chi": "",
        "noi_dung": "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn",
    },
    "3213/TCS8-QLDN": {
        "ten_ban": "Công ty TNHH Thương mại Dịch vụ Đầu tư Đình Long",
        "mst_ban": "3801337113",
        "dia_chi": "",
        "noi_dung": "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn",
    },
    "3212/TCS8-QLDN": {
        "ten_ban": "Công ty TNHH TM DV KD Nhật Trường",
        "mst_ban": "3801334440",
        "dia_chi": "",
        "noi_dung": "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn",
    },
    "5977/TCS2-QLDN": {
        "ten_ban": "Công ty TNHH Nông Sản Thái Hưng Long",
        "mst_ban": "3600878225",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT",
    },
    "18712/TCS3-KTR1": {
        "ten_ban": "Công ty TNHH Thương mại Dịch vụ Sản xuất Ánh Thương",
        "mst_ban": "3603863354",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "4280/TCS3-QLDN-KT": {
        "ten_ban": "Công ty TNHH TM SX & DV vận tải Thành Long",
        "mst_ban": "4101430743",
        "dia_chi": "Thôn Tân Hóa Nam, xã Hòa Hội, tỉnh Gia Lai",
        "noi_dung": "Nghi vấn gian lận hóa đơn",
    },
    "9036/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH MTV Xây dựng Thương mại Dịch vụ Nguyễn Duy",
        "mst_ban": "3801267106",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chưa nộp tờ khai thuế; Chênh lệch số liệu tờ khai và HĐĐT",
    },
    "8989/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH Kỹ thuật Chín Vinh",
        "mst_ban": "3801183054",
        "dia_chi": "76/1330, Tổ 3, Khu phố 3, Phường Bình Phước, thành phố Đồng Nai",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "8998/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH Xuất nhập khẩu Hoàng Sáng",
        "mst_ban": "3801302336",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "6621/TCS15-QLDN3": {
        "ten_ban": "Công ty TNHH MTV Tống Minh Luân",
        "mst_ban": "0319617301",
        "dia_chi": "Số nhà 405/6 Thống Nhất, phường Thông Tây Hội, TP. Hồ Chí Minh",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký",
    },
    "612/TCS4-QLDN": {
        "ten_ban": "Công ty TNHH TM DV Đầu tư Buôn bán VLXD và Thi công Diệp Hiền",
        "mst_ban": "1102139429",
        "dia_chi": "277, ấp An Hiệp, Xã An Ninh, Tây Ninh",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký",
    },
    "19138/TCS28-QLDN": {
        "ten_ban": "Công ty TNHH Thương mại Sản xuất Gỗ Văn Phúc",
        "mst_ban": "3703254589",
        "dia_chi": "Số 35 Đường số 4, KDC Thái Bình Dương, khu phố Khánh Hội, P. Tân Khánh, TP. Hồ Chí Minh",
        "noi_dung": "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn",
    },
    "613/TCS4-QLDN": {
        "ten_ban": "Công ty TNHH TMDV KD Diễm Uyên",
        "mst_ban": "1102138249",
        "dia_chi": "99 Bình Lợi, Xã Hòa Khánh, Tây Ninh",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký",
    },
    "17960/DON-QLDN2": {
        "ten_ban": "Công ty TNHH Sản xuất và Thương mại Xuất nhập khẩu Hưng Mạnh Phát",
        "mst_ban": "3603938521",
        "dia_chi": "Tổ 7A, Khu phố Tân Cang, Phường Phước Tân, TP. Đồng Nai",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký",
    },
    "8831/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH Công nghệ Đồng Sáng Việt Nam",
        "mst_ban": "3702795465",
        "dia_chi": "Số 39 Đường Lê Duẩn, Khu phố Tân Phú 1, phường Bình Phước, tỉnh Đồng Nai",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "8836/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH Phú Mỹ Tân",
        "mst_ban": "3801160258",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "8838/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH MTV Sản xuất Thương mại Dịch vụ và Xây dựng Thắng Đức",
        "mst_ban": "3801281245",
        "dia_chi": "KP Tân Thiện, phường Bình Phước, tỉnh Đồng Nai",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "8840/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH Thiết kế Xây dựng Nội ngoại thất Hưng Tín Phát",
        "mst_ban": "3801277016",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "8839/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH Cơ khí Tự động Gia Nguyên",
        "mst_ban": "3801270081",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "8837/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH Thương mại Dịch vụ Sản xuất Thành Công",
        "mst_ban": "3801201296",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "8841/TCS10-QLDN1": {
        "ten_ban": "Công ty TNHH Phú Mỹ Tân",
        "mst_ban": "3801160258",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "8844/TB-TCS10": {
        "ten_ban": "Công ty TNHH MTV Gỗ Vũ Lâm Tiến",
        "mst_ban": "3801237239",
        "dia_chi": "",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT; Nghi vấn gian lận hóa đơn",
    },
    "1168/TCS26-QLDN": {
        "ten_ban": "Công ty TNHH TM DV Toàn Cầu VN",
        "mst_ban": "3703293644",
        "dia_chi": "60 đường Tân Triều, khu phố Tân Hiệp, P. Tân Đông Hiệp, TP. Hồ Chí Minh",
        "noi_dung": "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn",
    },
    "608/TCS4-QLDN": {
        "ten_ban": "Công ty TNHH MTV Buôn bán và Sửa chữa Máy móc Cơ khí Hùng Cường",
        "mst_ban": "1102138087",
        "dia_chi": "Số 129E/12, Đường Nguyễn Thị Thọ, Khu vực 5, Xã Đức Hòa, Tây Ninh",
        "noi_dung": "Không hoạt động tại địa chỉ đăng ký",
    },
}


def norm(s: str) -> str:
    return re.sub(r"\s+", "", (s or "")).upper()


def main() -> int:
    records = json.loads(RECORDS_PATH.read_text(encoding="utf-8"))
    for rec in records:
        key = norm(rec.get("so_cv"))
        info = MAP.get(key)
        if not info:
            print("NO MAP", rec.get("so_cv"))
            continue
        rec.update({k: v for k, v in info.items() if v or k == "mst_ban"})
        print("JSON", rec["so_cv"], rec["ten_ban"], rec.get("mst_ban"))
    RECORDS_PATH.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")

    wb = openpyxl.load_workbook(DEFAULT_EXCEL_PATH)
    ws = wb["THEO_DOI"]
    n = 0
    for r in range(28336, min(ws.max_row, 28359) + 1):
        key = norm(str(ws.cell(r, 2).value or ""))
        info = MAP.get(key)
        if not info:
            print("ROW skip", r, ws.cell(r, 2).value)
            continue
        ws.cell(r, 5, info["ten_ban"])
        ws.cell(r, 6, info["mst_ban"] or None)
        ws.cell(r, 7, info["noi_dung"])
        if info.get("dia_chi"):
            ws.cell(r, 16, info["dia_chi"])
        n += 1
        print(f"ROW {r} {key} -> {info['ten_ban']}")
    wb.save(DEFAULT_EXCEL_PATH)
    print(f"[DONE] patched {n} rows")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
