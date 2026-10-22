# -*- coding: utf-8 -*-
"""
process_all_and_update_excel.py
Bóc tách toàn diện:
1. 53 tệp thuộc tháng 01-04/2026 (D:\01. Cong tac kiem tra\HO SO CONG VAN CANH BAO\2026\2026-01..04)
2. 49 tệp thuộc đợt mới 27/09 - 07/10/2026 (D:\01. Cong tac kiem tra\HO SO CONG VAN CANH BAO\2026\2026-10)
3. Đối chiếu chống trùng với sổ theo dõi Excel hiện hữu (F0 và F1)
4. Nạp các bản ghi độc nhất vào sổ Excel với font Times New Roman, STT nối tiếp và tự động sao lưu.
"""
import os
import re
import sys
import json
import shutil
from datetime import datetime
from pathlib import Path
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_ARCHIVE = Path(r"D:\01. Cong tac kiem tra\HO SO CONG VAN CANH BAO\2026")
EXCEL_PATH = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx")
BACKUP_DIR = EXCEL_PATH.parent / "Sao lưu"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)

MAP_DON_VI = {
    "THUẾ CƠ SỞ 1": "Thuế cơ sở 1 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 2": "Thuế cơ sở 2 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 3": "Thuế cơ sở 3 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 4": "Thuế cơ sở 4 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 5": "Thuế cơ sở 5 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 6": "Thuế cơ sở 6 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 7": "Thuế cơ sở 7 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 8": "Thuế cơ sở 8 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 9": "Thuế cơ sở 9 thành phố Đồng Nai",
    "THUẾ CƠ SỞ 10": "Thuế cơ sở 10 thành phố Đồng Nai",
    "THUẾ THÀNH PHỐ ĐỒNG NAI": "Thuế thành phố Đồng Nai",
}

def clean_text(s):
    if not s:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def clean_company_name(s):
    if not s:
        return ""
    s = re.sub(r"^(?:v/v|vv|v/v\s+cảnh\s+báo|cảnh\s+báo|cvr|cbrr|cv|công\s+văn)[\s:.-]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:doanh\s+nghiệp\s+có\s+dấu\s+hiệu\s+rủi\s+ro\s+cao\s+về\s+thuế[^\w]*)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:đối\s+với\s+)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\s:;.-]+(?:mst|mã\s+số\s+thuế).*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\b(signed|pdf)\b", "", s, flags=re.IGNORECASE)
    return s.strip(" ,.:;-\"'\n\r\t")

def normalize_key(socv, mst, ten=""):
    c_socv = re.sub(r"[^\w]", "", str(socv or "")).lower()
    c_mst = re.sub(r"[^\d]", "", str(mst or ""))
    c_ten = re.sub(r"[^\w]", "", str(ten or "")).lower()
    return f"{c_socv}_{c_mst}_{c_ten}"

def normalize_don_vi(dv_raw):
    if not dv_raw:
        return "Thuế thành phố Đồng Nai"
    dv = clean_text(dv_raw)
    dv_upper = dv.upper()
    if dv_upper in MAP_DON_VI:
        return MAP_DON_VI[dv_upper]
    
    for k, v in MAP_DON_VI.items():
        if k in dv_upper:
            return v
            
    # Tỉnh khác
    dv = re.sub(r"THUẾ CƠ SỞ (\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv = re.sub(r"Thuế CS(\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TP\.?\s*HCM", "TP. Hồ Chí Minh", dv, flags=re.IGNORECASE)
    return dv

def parse_date(date_val):
    if not date_val:
        return None
    if isinstance(date_val, datetime):
        return date_val
    if isinstance(date_val, str):
        date_str = date_val.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
    return None

def extract_pdf(pdf_path):
    text = ""
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            text += page.get_text() + "\n"
        doc.close()
    except Exception as e:
        print(f"Lỗi đọc PDF {pdf_path.name}: {e}")
    return text

def parse_cv_document(file_list, base_folder, month_label):
    # Parse 1 nhóm tệp hoặc 1 tệp công văn
    texts = []
    main_file = None
    for f in file_list:
        p = base_folder / f
        txt = extract_pdf(p)
        texts.append((f, txt))
        if not main_file and not any(k in f.lower() for k in ["bang ke", "bảng kê", "hdbr", "hdmv", "danh sach", "5223", "5344", "5349", "5402", "5364", "5310", "5311", "5221", "5279"]):
            main_file = f

    if not main_file and file_list:
        main_file = file_list[0]

    combined = "\n".join(t[1] for t in texts)
    
    # Tìm Số công văn
    so_cv = ""
    m_socv = re.search(r"Số\s*:\s*([\d]+/[A-Z0-9\-]+)", combined, re.IGNORECASE)
    if m_socv:
        so_cv = m_socv.group(1).strip()
    else:
        # Check filename
        m_fn = re.search(r"\[(?:\d+_)?([\d]+[_\/][A-Z0-9\-]+)\]", main_file)
        if m_fn:
            so_cv = m_fn.group(1).replace("_", "/")
        else:
            m_fn2 = re.search(r"([\d]+[_\/][A-Z0-9\-]+)", main_file)
            if m_fn2:
                so_cv = m_fn2.group(1).replace("_", "/")

    # Tìm Ngày công văn
    ngay_cv = None
    m_ngay = re.search(r"ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(202\d)", combined, re.IGNORECASE)
    if m_ngay:
        d, m, y = m_ngay.groups()
        ngay_cv = datetime(int(y), int(m), int(d))
    else:
        m_ngay2 = re.search(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](202\d)", combined)
        if m_ngay2:
            d, m, y = m_ngay2.groups()
            try:
                ngay_cv = datetime(int(y), int(m), int(d))
            except Exception:
                pass

    # Nếu chưa có ngày, lấy từ tháng
    if not ngay_cv:
        if month_label == "2026-10":
            ngay_cv = datetime(2026, 10, 1)
        elif month_label == "2026-04":
            ngay_cv = datetime(2026, 4, 15)
        elif month_label == "2026-03":
            ngay_cv = datetime(2026, 3, 15)
        elif month_label == "2026-02":
            ngay_cv = datetime(2026, 2, 15)
        elif month_label == "2026-01":
            ngay_cv = datetime(2026, 1, 15)

    # Đơn vị ban hành
    don_vi_bh = ""
    m_dv = re.search(r"(THUẾ\s+CƠ\s+SỞ\s+\d+|CỤC\s+THUẾ\s+[\w\s]+|THUẾ\s+THÀNH\s+PHỐ\s+ĐỒNG\s+NAI|THUẾ\s+CS\d+[\w\s]*)", combined, re.IGNORECASE)
    if m_dv:
        don_vi_bh = normalize_don_vi(m_dv.group(1))
    else:
        # Fallback từ số CV
        if "TCS10" in so_cv:
            don_vi_bh = "Thuế cơ sở 10 thành phố Đồng Nai"
        elif "TCS7" in so_cv:
            don_vi_bh = "Thuế cơ sở 7 thành phố Đồng Nai"
        elif "TCS2" in so_cv:
            don_vi_bh = "Thuế cơ sở 2 thành phố Đồng Nai"
        elif "TCS1" in so_cv:
            don_vi_bh = "Thuế cơ sở 1 thành phố Đồng Nai"
        elif "TCS3" in so_cv:
            don_vi_bh = "Thuế cơ sở 3 thành phố Đồng Nai"
        elif "TCS4" in so_cv:
            don_vi_bh = "Thuế cơ sở 4 thành phố Đồng Nai"
        elif "TCS5" in so_cv:
            don_vi_bh = "Thuế cơ sở 5 thành phố Đồng Nai"
        elif "TCS8" in so_cv:
            don_vi_bh = "Thuế cơ sở 8 thành phố Đồng Nai"
        elif "TCS9" in so_cv:
            don_vi_bh = "Thuế cơ sở 9 thành phố Đồng Nai"
        else:
            don_vi_bh = "Thuế thành phố Đồng Nai"

    # Tìm danh sách MST và Tên công ty
    # Regex tìm cặp MST và Tên
    mst_matches = re.findall(r"(?:MST|mã\s+số\s+thuế|MÃ\s+SỐ\s+THUẾ)[\s:.]*([0-9]{10}(?:-[0-9]{3})?)", combined, re.IGNORECASE)
    if not mst_matches:
        mst_matches = re.findall(r"\b([0-9]{10}(?:-[0-9]{3})?)\b", combined)
    mst_matches = list(dict.fromkeys(mst_matches))

    # Tên công ty
    company_names = []
    for line in combined.split("\n"):
        line_clean = line.strip()
        if re.search(r"(?:CÔNG TY|CTY|DOANH NGHIỆP|DNTN|HTX)\s+([A-ZÀ-Ỹ0-9\s\.\-_]{4,})", line_clean, re.IGNORECASE):
            m_c = re.search(r"(?:CÔNG TY|CTY|DOANH NGHIỆP|DNTN|HTX)\s+([^\n\r,\.]+)", line_clean, re.IGNORECASE)
            if m_c:
                cname = clean_company_name(m_c.group(0))
                if len(cname) > 6 and cname not in company_names:
                    company_names.append(cname)

    # Nếu không tìm thấy tên từ text, tìm từ tên file
    if not company_names:
        for f in file_list:
            m_fn_c = re.search(r"(?:CTY|CÔNG TY|CBRR|CẢNH BÁO)\s*[-_]?\s*([A-ZÀ-Ỹ0-9\s]{4,})", f, re.IGNORECASE)
            if m_fn_c:
                company_names.append(clean_company_name(m_fn_c.group(1)))

    # Phân loại F0 vs F1:
    # Mặc định công văn cảnh báo doanh nghiệp xuất hóa đơn / rủi ro -> F0 (bên bán)
    # Nếu có danh sách người mua (F1) cụ thể trong văn bản hoặc bảng kê mua vào -> F1
    records = []
    
    primary_mst = mst_matches[0] if mst_matches else ""
    primary_name = company_names[0] if company_names else "Doanh nghiệp có dấu hiệu rủi ro"

    # Trích xuất nội dung tóm tắt
    noi_dung = "Cảnh báo doanh nghiệp có dấu hiệu rủi ro cao về thuế, hóa đơn"
    if "không hoạt động tại địa chỉ" in combined.lower() or "bỏ địa chỉ" in combined.lower():
        noi_dung = "Không hoạt động tại địa chỉ đăng ký kinh doanh, có dấu hiệu rủi ro cao về hóa đơn"
    elif "ngừng hoạt động" in combined.lower():
        noi_dung = "Doanh nghiệp tạm ngừng hoạt động / bỏ trốn nhưng vẫn phát sinh hóa đơn"
    elif "bảng kê" in combined.lower() or "hóa đơn bán ra" in combined.lower():
        noi_dung = "Cảnh báo rủi ro về thuế và hóa đơn, kèm bảng kê hóa đơn bán ra"

    # Tạo record
    rec = {
        "SoCV": so_cv,
        "NgayCV": ngay_cv,
        "DonViBH": don_vi_bh,
        "TenBan": primary_name,
        "MSTBan": primary_mst,
        "NoiDung": noi_dung,
        "TenMua": "",
        "MSTMua": "",
        "TamNgung": "Có" if "tạm ngưng" in combined.lower() or "ngừng hoạt động" in combined.lower() else "",
        "KyKK": "",
        "SoHD": "",
        "DoanhThu": 0,
        "ThueGTGT": 0,
        "NhomND": "Cảnh báo rủi ro hóa đơn",
        "DiaChi": "",
        "NgayNhan": datetime(2026, 10, 7) if month_label == "2026-10" else ngay_cv,
        "TinhTrang": "Chưa xử lý",
        "GhiChu": f"Bóc tách tự động {month_label}",
        "File": main_file,
        "NamCV": ngay_cv.year if ngay_cv else 2026,
        "Trung": "",
        "KhoaCV": f"{so_cv}_{ngay_cv.year if ngay_cv else 2026}",
        "CVDau": 1,
        "MSTDau": 1,
        "SheetTarget": "F0",
        "Month": month_label
    }
    records.append(rec)
    return records

def load_existing_keys(excel_path):
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    f0_keys = set()
    f1_keys = set()
    
    # Layout moi: STT | MSTBan | TenBan | SoCV | NgayCV | DonViBH | ...
    ws_f0 = wb["F0"]
    for row in ws_f0.iter_rows(min_row=8, values_only=True):
        if not any(row):
            continue
        mst = row[1]
        ten = row[2]
        socv = row[3]
        k1 = normalize_key(socv, mst, ten)
        k2 = normalize_key(socv, mst, "")
        if mst:
            f0_keys.add(re.sub(r"[^\d]", "", str(mst)))
        if socv:
            f0_keys.add(re.sub(r"[^\w]", "", str(socv)).lower())
        f0_keys.add(k1)
        f0_keys.add(k2)

    # Đọc F1
    ws_f1 = wb["F1"]
    for row in ws_f1.iter_rows(min_row=8, values_only=True):
        if not any(row):
            continue
        mst_ban = row[1]
        ten_ban = row[2]
        socv = row[3]
        ten_mua = row[7]
        mst_mua = row[8]
        k1 = normalize_key(socv, mst_ban, ten_ban)
        k2 = normalize_key(socv, mst_mua, ten_mua)
        f1_keys.add(k1)
        f1_keys.add(k2)
        if mst_mua:
            f1_keys.add(re.sub(r"[^\d]", "", str(mst_mua)))

    wb.close()
    return f0_keys, f1_keys

def main():
    print(f"--- BẮT ĐẦU XỬ LÝ DỮ LIỆU CÔNG VĂN CẢNH BÁO ---")
    print(f"File Excel đích: {EXCEL_PATH}")
    
    # 1. Nạp khóa chống trùng hiện hữu từ Sổ Excel
    f0_keys, f1_keys = load_existing_keys(EXCEL_PATH)
    print(f"Đã nạp {len(f0_keys)} khóa đối soát từ F0 và {len(f1_keys)} khóa từ F1.")

    all_raw_records = []

    # 2. Xử lý các tháng 01-04/2026 (53 tệp)
    months_old = ["2026-01", "2026-02", "2026-03", "2026-04"]
    for m in months_old:
        m_dir = BASE_ARCHIVE / m
        if not m_dir.exists():
            continue
        files = [f for f in os.listdir(m_dir) if f.endswith(".pdf")]
        # Gom nhóm
        groups = {}
        for f in files:
            m_pref = re.match(r"^(\[.*?\])", f)
            pref = m_pref.group(1) if m_pref else f
            groups.setdefault(pref, []).append(f)
        for pref, gfiles in groups.items():
            recs = parse_cv_document(gfiles, m_dir, m)
            all_raw_records.extend(recs)

    print(f"\nTổng bản ghi thô từ đợt 01-04/2026: {len(all_raw_records)}")

    # 3. Xử lý đợt mới 27/09 - 07/10/2026 (2026-10: 49 tệp)
    dir_10 = BASE_ARCHIVE / "2026-10"
    files_10 = [f for f in os.listdir(dir_10) if f.endswith(".pdf")]
    groups_10 = {}
    for f in files_10:
        m_pref = re.match(r"^(\[\d+\]_[^_]+_[^_]+)", f)
        pref = m_pref.group(1) if m_pref else f
        groups_10.setdefault(pref, []).append(f)
        
    records_10 = []
    for pref, gfiles in groups_10.items():
        recs = parse_cv_document(gfiles, dir_10, "2026-10")
        records_10.extend(recs)
        
    print(f"Tổng bản ghi thô từ đợt 27/09 - 07/10/2026 (2026-10): {len(records_10)}")
    all_raw_records.extend(records_10)

    # 4. Đối chiếu chống trùng
    unique_f0 = []
    unique_f1 = []
    duplicate_records = []
    seen_in_batch = set()

    for r in all_raw_records:
        socv = r["SoCV"]
        mst = r["MSTBan"]
        ten = r["TenBan"]
        k_full = normalize_key(socv, mst, ten)
        k_socv_mst = normalize_key(socv, mst, "")
        clean_mst = re.sub(r"[^\d]", "", str(mst or ""))
        clean_socv = re.sub(r"[^\w]", "", str(socv or "")).lower()

        is_dup = False
        if k_full in f0_keys or k_socv_mst in f0_keys or k_full in seen_in_batch:
            is_dup = True
        elif clean_mst and clean_mst in f0_keys and len(clean_mst) >= 10 and clean_socv in f0_keys:
            is_dup = True

        if is_dup:
            duplicate_records.append(r)
        else:
            seen_in_batch.add(k_full)
            if clean_mst:
                seen_in_batch.add(clean_mst)
            if r["SheetTarget"] == "F0":
                unique_f0.append(r)
            else:
                unique_f1.append(r)

    print(f"\n--- KẾT QUẢ ĐỐI SOÁT CHỐNG TRÙNG ---")
    print(f"Tổng bản ghi trích xuất: {len(all_raw_records)}")
    print(f"Bản ghi trùng lặp (loại bỏ): {len(duplicate_records)}")
    print(f"Bản ghi độc nhất F0 (sẵn sàng nạp): {len(unique_f0)}")
    print(f"Bản ghi độc nhất F1 (sẵn sàng nạp): {len(unique_f1)}")

    # 5. Tạo bản sao lưu an toàn
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = BACKUP_DIR / f"SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.bak_{timestamp}"
    shutil.copy2(EXCEL_PATH, backup_file)
    print(f"\nĐã tạo bản sao lưu: {backup_file.name}")

    # 6. Nạp dữ liệu vào file Excel
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_f0 = wb["F0"]
    font_tnr = Font(name="Times New Roman", size=10)
    font_bold = Font(name="Times New Roman", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Xác định STT và dòng cuối của F0
    last_row_f0 = ws_f0.max_row
    last_stt_f0 = 0
    for r in range(last_row_f0, 7, -1):
        v = ws_f0.cell(r, 1).value
        if isinstance(v, (int, float)):
            last_stt_f0 = int(v)
            break

    print(f"\nF0 hiện tại: dòng cuối = {last_row_f0}, STT cuối = {last_stt_f0}")
    
    current_row = last_row_f0 + 1
    current_stt = last_stt_f0 + 1

    for rec in unique_f0:
        # Layout moi: STT | MSTBan | TenBan | SoCV | NgayCV | DonViBH | ...
        c = ws_f0.cell(current_row, 1, current_stt)
        c.font = font_tnr
        c.alignment = align_center

        c = ws_f0.cell(current_row, 2, rec["MSTBan"])
        c.font = font_tnr
        c.alignment = align_center

        c = ws_f0.cell(current_row, 3, rec["TenBan"])
        c.font = font_bold
        c.alignment = align_left

        c = ws_f0.cell(current_row, 4, rec["SoCV"])
        c.font = font_tnr
        c.alignment = align_center

        c = ws_f0.cell(current_row, 5, rec["NgayCV"])
        c.font = font_tnr
        c.alignment = align_center
        if rec["NgayCV"]:
            c.number_format = "dd/mm/yyyy"

        c = ws_f0.cell(current_row, 6, rec["DonViBH"])
        c.font = font_tnr
        c.alignment = align_left

        c = ws_f0.cell(current_row, 7, rec["NoiDung"])
        c.font = font_tnr
        c.alignment = align_left

        c = ws_f0.cell(current_row, 8, rec["TamNgung"])
        c.font = font_tnr
        c.alignment = align_center

        c = ws_f0.cell(current_row, 9, rec["NhomND"])
        c.font = font_tnr
        c.alignment = align_left

        c = ws_f0.cell(current_row, 10, rec["DiaChi"])
        c.font = font_tnr
        c.alignment = align_left

        c = ws_f0.cell(current_row, 11, rec["NgayNhan"])
        c.font = font_tnr
        c.alignment = align_center
        if rec["NgayNhan"]:
            c.number_format = "dd/mm/yyyy"

        c = ws_f0.cell(current_row, 12, rec["TinhTrang"])
        c.font = font_tnr
        c.alignment = align_center

        c = ws_f0.cell(current_row, 13, rec["GhiChu"])
        c.font = font_tnr
        c.alignment = align_left

        c = ws_f0.cell(current_row, 14, rec["File"])
        c.font = font_tnr
        c.alignment = align_left

        c = ws_f0.cell(current_row, 15, rec["NamCV"])
        c.font = font_tnr
        c.alignment = align_center

        c = ws_f0.cell(current_row, 16, rec["Trung"])
        c.font = font_tnr
        c.alignment = align_center

        c = ws_f0.cell(current_row, 17, rec["KhoaCV"])
        c.font = font_tnr

        c = ws_f0.cell(current_row, 18, rec["CVDau"])
        c.font = font_tnr
        c.alignment = align_center

        c = ws_f0.cell(current_row, 19, rec["MSTDau"])
        c.font = font_tnr
        c.alignment = align_center

        current_row += 1
        current_stt += 1

    wb.save(EXCEL_PATH)
    wb.close()
    print(f"Đã lưu thành công Sổ theo dõi Excel! Tổng dòng F0 mới: {current_row - 1} (thêm {len(unique_f0)} dòng, STT đến {current_stt - 1})")

    # Lưu kết quả JSON để lưu vết
    report_data = {
        "timestamp": timestamp,
        "backup_file": str(backup_file),
        "total_extracted": len(all_raw_records),
        "total_duplicates": len(duplicate_records),
        "unique_f0_added": len(unique_f0),
        "unique_f1_added": len(unique_f1),
        "f0_records": unique_f0,
        "duplicate_records": duplicate_records
    }
    with open(r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo\skills\theo-doi-cv-canh-bao-rui-ro\scripts\batch_update_result_latest.json", "w", encoding="utf-8") as fp:
        json.dump(report_data, fp, ensure_ascii=False, indent=2, default=str)

if __name__ == "__main__":
    main()
