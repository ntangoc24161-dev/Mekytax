# -*- coding: utf-8 -*-
"""Doc va bao cao file Danh sach kiem tra quyet toan pro.xlsm."""

import os
import argparse
from datetime import datetime

import openpyxl

_SO_DIR = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung"
EXCEL_XLSM = os.path.join(_SO_DIR, "Danh sach kiem tra quyet toan pro.xlsm")
FALLBACK_PATH = os.path.join(_SO_DIR, "Danh sach kiem tra quyet toan.xlsx")
EXCEL_PATH = EXCEL_XLSM  # ten cu EXCEL_PATH giu de tuong thich; chi dung pro.xlsm

TRACK_SHEETS = ("01_DA_BAN_HANH", "02_DANG_THUC_HIEN")


def get_workbook():
    if os.path.exists(EXCEL_XLSM):
        target = EXCEL_XLSM
    else:
        target = FALLBACK_PATH
    return openpyxl.load_workbook(target, data_only=False), target


def is_new_layout(wb):
    return "01_DA_BAN_HANH" in wb.sheetnames or "02_DANG_THUC_HIEN" in wb.sheetnames


def detect_layout(ws):
    """new50 = co cot Thong tin cong ty (Thoi ky o cot 9); new37 = layout cu 37 cot."""
    h9 = str(ws.cell(4, 9).value or "")
    if "Thời kỳ" in h9:
        return "new50"
    if ws.max_column >= 37 and str(ws.cell(4, 2).value or "").startswith("Trạng"):
        return "new37"
    return "new37"


def iter_track_rows(wb):
    """Yield (sheet_name, row_idx, layout) cho moi dong co MST/Ten."""
    if is_new_layout(wb):
        for sheet_name in TRACK_SHEETS:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            layout = detect_layout(ws)
            for r in range(5, ws.max_row + 1):
                mst = ws.cell(r, 4).value
                ten = ws.cell(r, 5).value
                if mst or ten:
                    yield sheet_name, ws, r, layout
        return

    # Fallback cu
    sheet_name = "SO_THEO_DOI_2026" if "SO_THEO_DOI_2026" in wb.sheetnames else "2026"
    ws = wb[sheet_name]
    start = 4 if sheet_name == "SO_THEO_DOI_2026" else 2
    layout = "pro_old" if sheet_name == "SO_THEO_DOI_2026" else "source"
    for r in range(start, ws.max_row + 1):
        mst_col = 4 if layout == "pro_old" else 3
        ten_col = 5 if layout == "pro_old" else 4
        mst = ws.cell(r, mst_col).value
        ten = ws.cell(r, ten_col).value
        if mst or ten:
            yield sheet_name, ws, r, layout


def cell_num(ws, r, c):
    val = ws.cell(r, c).value
    if val is None or val == "":
        return 0.0
    if isinstance(val, str) and val.startswith("="):
        return 0.0
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def finance_from_row(ws, r, layout):
    if layout == "new50":
        tndn = cell_num(ws, r, 38)
        gtgt = cell_num(ws, r, 39)
        p_gtgt = cell_num(ws, r, 41)
        p_tndn = cell_num(ws, r, 42)
        p_20 = cell_num(ws, r, 44)
        p_vphc = cell_num(ws, r, 45)
        truythu = tndn + gtgt
        phat = p_gtgt + p_tndn + p_20 + p_vphc
        return truythu, phat, truythu + phat
    if layout in ("new", "new37"):
        tndn = cell_num(ws, r, 25)
        gtgt = cell_num(ws, r, 26)
        p_gtgt = cell_num(ws, r, 28)
        p_tndn = cell_num(ws, r, 29)
        p_20 = cell_num(ws, r, 31)
        p_vphc = cell_num(ws, r, 32)
        truythu = tndn + gtgt
        phat = p_gtgt + p_tndn + p_20 + p_vphc
        return truythu, phat, truythu + phat
    if layout == "pro_old":
        tndn = cell_num(ws, r, 29)
        gtgt = cell_num(ws, r, 30)
        p_gtgt = cell_num(ws, r, 32)
        p_tndn = cell_num(ws, r, 33)
        p_20 = cell_num(ws, r, 35)
        p_vphc = cell_num(ws, r, 36)
        truythu = tndn + gtgt
        phat = p_gtgt + p_tndn + p_20 + p_vphc
        return truythu, phat, truythu + phat
    tong = cell_num(ws, r, 59)
    return tong, 0.0, tong


def num_to_vietnamese_words(num):
    """Chuyen so tien thanh chu tieng Viet chuan ke toan/thue."""
    if num is None:
        return ""
    try:
        n = int(round(float(num)))
    except (TypeError, ValueError):
        return str(num)

    if n == 0:
        return "Không đồng."

    units = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"]
    digits = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]

    def read_three_digits(val, has_higher):
        h = val // 100
        t = (val % 100) // 10
        u = val % 10
        res = []
        if h > 0 or has_higher:
            res.append(digits[h] + " trăm")
        if t > 1:
            res.append(digits[t] + " mươi")
            if u == 1:
                res.append("mốt")
            elif u == 5:
                res.append("lăm")
            elif u > 0:
                res.append(digits[u])
        elif t == 1:
            res.append("mười")
            if u == 1:
                res.append("một")
            elif u == 5:
                res.append("lăm")
            elif u > 0:
                res.append(digits[u])
        elif t == 0:
            if u > 0:
                if h > 0 or has_higher:
                    res.append("lẻ " + digits[u])
                else:
                    res.append(digits[u])
        return " ".join(res)

    chunks = []
    temp = n
    while temp > 0:
        chunks.append(temp % 1000)
        temp //= 1000

    res_chunks = []
    for i in range(len(chunks) - 1, -1, -1):
        c = chunks[i]
        if c > 0:
            part = read_three_digits(c, i < len(chunks) - 1)
            unit = units[i]
            if unit:
                part += " " + unit
            res_chunks.append(part)

    full_str = " ".join(res_chunks).strip()
    full_str = full_str[0].upper() + full_str[1:] + " đồng."
    return full_str


def cmd_summary():
    wb, path = get_workbook()
    total_audits = 0
    ban_hanh = 0
    dang_th = 0
    overdue = 0
    total_remedy = 0.0
    status_counts = {}
    sheets_used = set()

    for sheet_name, ws, r, layout in iter_track_rows(wb):
        sheets_used.add(sheet_name)
        total_audits += 1
        status = str(ws.cell(r, 2).value or "")
        status_counts[status] = status_counts.get(status, 0) + 1
        if "Đã ban hành" in status:
            ban_hanh += 1
        elif "Đang thực hiện" in status:
            dang_th += 1

        if layout == "new50":
            alert = str(ws.cell(r, 50).value or "")
            if "QUÁ HẠN" in alert:
                overdue += 1
        elif layout in ("new", "new37"):
            alert = str(ws.cell(r, 37).value or "")
            if "QUÁ HẠN" in alert:
                overdue += 1
        _, _, tong = finance_from_row(ws, r, layout)
        total_remedy += tong

    print("### BÁO CÁO NHANH TIẾN ĐỘ KIỂM TRA THUẾ NĂM 2026")
    print(f"- Nguồn dữ liệu: `{os.path.basename(path)}`")
    print(f"- Sheet đọc: {', '.join(sorted(sheets_used)) or '(không có)'}")
    print(f"- **Tổng số hồ sơ:** **{total_audits}**")
    print(f"- **Đã ban hành:** **{ban_hanh}** | **Đang thực hiện:** **{dang_th}**")
    print(f"- **Cảnh báo quá hạn:** **{overdue}**")
    print(f"- **Tổng kiến nghị xử lý:** **{total_remedy:,.0f} đ**")
    print("\n#### Phân bổ theo trạng thái:")
    for st, count in status_counts.items():
        print(f"- {st}: **{count} hồ sơ**")


def cmd_lookup(keyword):
    wb, path = get_workbook()
    found = []
    kw = keyword.lower().strip()

    for sheet_name, ws, r, layout in iter_track_rows(wb):
        mst_col = 4 if layout in ("new50", "new37", "new", "pro_old") else 3
        ten_col = 5 if layout in ("new50", "new37", "new", "pro_old") else 4
        mst = str(ws.cell(r, mst_col).value or "").strip()
        ten = str(ws.cell(r, ten_col).value or "").strip()
        if kw in mst.lower() or kw in ten.lower():
            found.append((sheet_name, ws, r, layout, mst, ten))

    if not found:
        print(f"Không tìm thấy doanh nghiệp nào khớp với từ khóa: `{keyword}`")
        return

    print(f"### KẾT QUẢ TRA CỨU HỒ SƠ KIỂM TRA ({len(found)} kết quả)")
    for sheet_name, ws, r, layout, mst, ten in found:
        status = ws.cell(r, 2).value
        if layout == "new50":
            td = ws.cell(r, 22).value
            tv = ws.cell(r, 24).value
            qdkt = ws.cell(r, 27).value
            ngay_cb = ws.cell(r, 30).value
            ngay_bb = ws.cell(r, 34).value
            qdxphc = ws.cell(r, 36).value
            alert = ws.cell(r, 50).value
            thoi_ky = ws.cell(r, 9).value
            noi_dung = ws.cell(r, 26).value
        elif layout in ("new", "new37"):
            td = ws.cell(r, 9).value
            tv = ws.cell(r, 11).value
            qdkt = ws.cell(r, 14).value
            ngay_cb = ws.cell(r, 17).value
            ngay_bb = ws.cell(r, 21).value
            qdxphc = ws.cell(r, 23).value
            alert = ws.cell(r, 37).value
            thoi_ky = ""
            noi_dung = ws.cell(r, 13).value
        elif layout == "pro_old":
            td = ws.cell(r, 9).value
            tv = ws.cell(r, 11).value
            qdkt = ws.cell(r, 14).value
            ngay_cb = ws.cell(r, 18).value
            ngay_bb = ws.cell(r, 23).value
            qdxphc = ws.cell(r, 25).value
            alert = ws.cell(r, 48).value
            thoi_ky = ""
            noi_dung = ""
        else:
            td = ws.cell(r, 8).value
            tv = ws.cell(r, 10).value
            qdkt = ws.cell(r, 14).value
            ngay_cb = ws.cell(r, 30).value
            ngay_bb = ws.cell(r, 33).value
            qdxphc = ws.cell(r, 35).value
            alert = ""
            thoi_ky = ws.cell(r, 7).value
            noi_dung = ws.cell(r, 6).value

        truythu, phat, tong = finance_from_row(ws, r, layout)
        print(f"\n#### Doanh nghiệp: {ten} (MST: `{mst}`)")
        print(f"- **Sheet / Trạng thái:** `{sheet_name}` / {status}")
        if thoi_ky:
            print(f"- **Thời kỳ kiểm tra:** {thoi_ky}")
        if noi_dung:
            print(f"- **Nội dung:** {noi_dung}")
        print(f"- **Đoàn kiểm tra:** Trưởng đoàn **{td}** | Thành viên **{tv}**")
        print(f"- **Quyết định KT:** `{qdkt}` | **Ngày công bố:** `{ngay_cb}`")
        print(f"- **Biên bản KT:** Ngày `{ngay_bb}` | **QĐ XPHC:** `{qdxphc or 'Chưa ban hành'}`")
        print(f"- **Thông báo quá hạn:** {alert or 'Không'}")
        print(f"- **Số tiền truy thu:** **{truythu:,.0f} đ**")
        print(f"- **Số tiền phạt & chậm nộp:** **{phat:,.0f} đ**")
        print(f"- **Tổng cộng nghĩa vụ:** **{tong:,.0f} đ**")
        print(f"- *Bằng chữ:* {num_to_vietnamese_words(tong)}")


def cmd_alerts():
    wb, path = get_workbook()
    now = datetime.now().strftime("%H:%M %d/%m/%Y")
    print("### RÀ SOÁT CẢNH BÁO QUÁ HẠN (CHƯA THỰC HIỆN ĐÚNG HẠN)")
    print(f"- Thời điểm rà soát: {now}")
    print(f"- Nguồn: `{os.path.basename(path)}`")

    alerts = []
    for sheet_name, ws, r, layout in iter_track_rows(wb):
        mst_col = 4 if layout in ("new50", "new37", "new", "pro_old") else 3
        ten_col = 5 if layout in ("new50", "new37", "new", "pro_old") else 4
        mst = str(ws.cell(r, mst_col).value or "").strip()
        ten = str(ws.cell(r, ten_col).value or "").strip()

        if layout == "new50":
            alert = str(ws.cell(r, 50).value or "")
            if "QUÁ HẠN" in alert or "CẬN HẠN" in alert:
                alerts.append(f"- **{ten}** (MST: `{mst}` | `{sheet_name}`): {alert}")
            continue
        if layout in ("new", "new37"):
            alert = str(ws.cell(r, 37).value or "")
            if "QUÁ HẠN" in alert or "CẬN HẠN" in alert:
                alerts.append(f"- **{ten}** (MST: `{mst}` | `{sheet_name}`): {alert}")
            continue

        # Fallback logic cu (khong theo doi nop NSNN nua)
        qdkt = ws.cell(r, 14).value
        ngay_cb = ws.cell(r, 18 if layout == "pro_old" else 30).value
        ngay_bb = ws.cell(r, 23 if layout == "pro_old" else 33).value
        qdxphc = ws.cell(r, 25 if layout == "pro_old" else 35).value
        if not ngay_cb and not qdkt:
            alerts.append(f"- **Chưa ban hành QĐKT:** `{ten}` (MST: `{mst}`)")
        elif ngay_bb and not qdxphc:
            alerts.append(f"- **Chờ ban hành QĐXPHC:** `{ten}` (MST: `{mst}`) - Đã ký BBKT ngày `{ngay_bb}`.")

    if not alerts:
        print("- Không có cảnh báo quá hạn / cận hạn.")
    else:
        for a in alerts:
            print(a)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Javis Tax Audit Tracker")
    parser.add_argument("--summary", action="store_true", help="Báo cáo tóm tắt tiến độ")
    parser.add_argument("--alerts", action="store_true", help="Quét cảnh báo quá hạn")
    parser.add_argument("--lookup", type=str, help="Tra cứu theo MST hoặc tên DN")
    parser.add_argument("--num2words", type=float, help="Chuyển số thành chữ tiếng Việt")
    args = parser.parse_args()

    if args.summary:
        cmd_summary()
    elif args.alerts:
        cmd_alerts()
    elif args.lookup:
        cmd_lookup(args.lookup)
    elif args.num2words is not None:
        print(num_to_vietnamese_words(args.num2words))
    else:
        cmd_summary()
