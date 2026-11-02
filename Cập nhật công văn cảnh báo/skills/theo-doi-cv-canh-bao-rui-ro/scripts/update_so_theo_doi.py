"""
update_so_theo_doi.py - Tu dong cap nhat sheet F0 (Ben ban) va F1 (Ben mua)
trong file SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.
"""
import sys
import os
import json
import shutil
import re
import argparse
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font

DEFAULT_EXCEL_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
if not Path(DEFAULT_EXCEL_PATH).is_file():
    DEFAULT_EXCEL_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"

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
    return date_val

def parse_number(num_val):
    if num_val is None or num_val == "":
        return 0
    if isinstance(num_val, (int, float)):
        return num_val
    try:
        clean = str(num_val).replace(".", "").replace(",", "").strip()
        if not clean:
            return 0
        return int(float(clean))
    except Exception:
        return 0

def append_records_to_sheet(ws, records, font_tnr):
    old_max_row = ws.max_row
    last_stt = 0
    for r in range(old_max_row, 1, -1):
        v = ws.cell(r, 1).value
        if isinstance(v, (int, float)):
            last_stt = int(v)
            break

    current_row = old_max_row + 1
    current_stt = last_stt + 1

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

    for rec in records:
        # Cột 1: STT
        ws.cell(current_row, 1, current_stt)
        # Cột 2: Số công văn (SoCV)
        ws.cell(current_row, 2, rec.get("so_cv") or rec.get("SoCV"))
        # Cột 3: Ngày công văn (NgayCV)
        ngay_cv = parse_date(rec.get("ngay_cv") or rec.get("NgayCV"))
        ws.cell(current_row, 3, ngay_cv)
        # Cột 4: Đơn vị ban hành (DonViBH)
        dv_raw = str(rec.get("don_vi_bh") or rec.get("DonViBH") or "").strip()
        dv_val = MAP_DON_VI.get(dv_raw.upper(), dv_raw)
        ws.cell(current_row, 4, dv_val)
        # Cột 5: Tên công ty bán (TenBan)
        ws.cell(current_row, 5, rec.get("ten_ban") or rec.get("TenBan"))
        # Cột 6: MST bên bán (MSTBan)
        ws.cell(current_row, 6, rec.get("mst_ban") or rec.get("MSTBan"))
        # Cột 7: Nội dung chính (NoiDung)
        ws.cell(current_row, 7, rec.get("noi_dung") or rec.get("NoiDung"))
        # Cột 8: Tên công ty mua (TenMua)
        ws.cell(current_row, 8, rec.get("ten_mua") or rec.get("TenMua"))
        # Cột 9: MST bên mua (MSTMua)
        ws.cell(current_row, 9, rec.get("mst_mua") or rec.get("MSTMua"))
        # Cột 10: Tạm ngưng kê khai (TamNgung)
        ws.cell(current_row, 10, rec.get("tam_ngung") or rec.get("TamNgung"))
        # Cột 11: Kỳ kê khai (KyKK)
        ws.cell(current_row, 11, rec.get("ky_kk") or rec.get("KyKK"))
        # Cột 12: Số hóa đơn (SoHD)
        ws.cell(current_row, 12, rec.get("so_hd") or rec.get("SoHD"))
        # Cột 13: Doanh thu (DoanhThu)
        ws.cell(current_row, 13, parse_number(rec.get("doanh_thu") or rec.get("DoanhThu")))
        # Cột 14: Tiền thuế GTGT (ThueGTGT)
        ws.cell(current_row, 14, parse_number(rec.get("thue_gtgt") or rec.get("ThueGTGT")))
        # Cột 15: Nhóm nội dung (NhomND)
        ws.cell(current_row, 15, rec.get("nhom_nd") or rec.get("NhomND") or "Dấu hiệu rủi ro về thuế")
        # Cột 16: Địa chỉ bên bán (DiaChi)
        ws.cell(current_row, 16, rec.get("dia_chi") or rec.get("DiaChi"))
        # Cột 17: Ngày nhận (NgayNhan)
        ngay_nhan = parse_date(rec.get("ngay_nhan") or rec.get("NgayNhan"))
        ws.cell(current_row, 17, ngay_nhan)
        # Cột 18: Tình trạng rà soát (TinhTrang)
        ws.cell(current_row, 18, rec.get("tinh_trang") or rec.get("TinhTrang"))
        # Cột 19: Ghi chú (GhiChu)
        ws.cell(current_row, 19, rec.get("ghi_chu") or rec.get("GhiChu"))
        # Cột 20: Đường dẫn file (File)
        ws.cell(current_row, 20, rec.get("file") or rec.get("File"))
        # Cột 21: Năm công văn (NamCV)
        nam_cv = rec.get("nam_cv") or rec.get("NamCV")
        if not nam_cv and isinstance(ngay_cv, datetime):
            nam_cv = ngay_cv.year
        ws.cell(current_row, 21, nam_cv)
        # Cột 22: Dòng trùng (Trung)
        ws.cell(current_row, 22, rec.get("trung") or rec.get("Trung"))
        # Cột 23: Khóa công văn (KhoaCV)
        ws.cell(current_row, 23, rec.get("khoa_cv") or rec.get("KhoaCV"))
        # Cột 24: CV đầu tiên (CVDau) - 1 hoặc None
        cv_dau = rec.get("cv_dau") if "cv_dau" in rec else rec.get("CVDau")
        ws.cell(current_row, 24, 1 if cv_dau in (1, "1", True) else None)
        # Cột 25: MST đầu tiên (MSTDau) - 1 hoặc None
        mst_dau = rec.get("mst_dau") if "mst_dau" in rec else rec.get("MSTDau")
        ws.cell(current_row, 25, 1 if mst_dau in (1, "1", True) else None)

        # Định dạng phông chữ Times New Roman
        for c in range(1, 26):
            cell = ws.cell(current_row, c)
            if cell.value is not None:
                cell.font = font_tnr

        current_row += 1
        current_stt += 1

    return old_max_row, current_row - 1

def update_workbook(records, excel_path=DEFAULT_EXCEL_PATH, backup=True, sheet_name=None):
    excel_path = Path(excel_path)
    if not excel_path.is_file():
        raise FileNotFoundError(f"Không tìm thấy file Excel tại {excel_path}")

    if backup:
        bak_dir = excel_path.parent / "Sao lưu"
        bak_dir.mkdir(exist_ok=True)
        bak_path = bak_dir / f"{excel_path.stem}.bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}{excel_path.suffix}"
        shutil.copy2(excel_path, bak_path)
        print(f"[OK] Đã tạo bản sao lưu tại: {bak_path}")

    wb = openpyxl.load_workbook(excel_path)
    font_tnr = Font(name="Times New Roman", size=10)

    # Phân loại records theo sheet
    records_f0 = []
    records_f1 = []
    records_unassigned = []

    for r in records:
        target = r.get("sheet") or r.get("Sheet") or sheet_name
        if target:
            target = str(target).strip().upper()
            if target in ("F0", "BAN", "BEN_BAN", "SELLER"):
                records_f0.append(r)
            elif target in ("F1", "MUA", "BEN_MUA", "BUYER"):
                records_f1.append(r)
            else:
                records_unassigned.append((target, r))
        else:
            # Tự động phán đoán: nếu có thông tin người bán mà không có bên mua cụ thể -> F0, ngược lại nếu trọng tâm bên mua -> F1
            loai = str(r.get("loai") or r.get("Loai") or "").lower()
            if "mua" in loai or "f1" in loai:
                records_f1.append(r)
            else:
                records_f0.append(r)

    total_inserted = 0

    if records_f0:
        target_sheet = "F0" if "F0" in wb.sheetnames else ("THEO_DOI" if "THEO_DOI" in wb.sheetnames else None)
        if not target_sheet:
            raise ValueError("Không tìm thấy sheet F0 trong file Excel")
        ws = wb[target_sheet]
        old_r, new_r = append_records_to_sheet(ws, records_f0, font_tnr)
        print(f"[OK] Đã chèn {len(records_f0)} dòng vào sheet [{target_sheet}] (dòng {old_r + 1} -> {new_r})")
        total_inserted += len(records_f0)

    if records_f1:
        target_sheet = "F1" if "F1" in wb.sheetnames else None
        if not target_sheet:
            raise ValueError("Không tìm thấy sheet F1 trong file Excel")
        ws = wb[target_sheet]
        old_r, new_r = append_records_to_sheet(ws, records_f1, font_tnr)
        print(f"[OK] Đã chèn {len(records_f1)} dòng vào sheet [{target_sheet}] (dòng {old_r + 1} -> {new_r})")
        total_inserted += len(records_f1)

    for tgt, r in records_unassigned:
        if tgt in wb.sheetnames:
            ws = wb[tgt]
            old_r, new_r = append_records_to_sheet(ws, [r], font_tnr)
            print(f"[OK] Đã chèn 1 dòng vào sheet [{tgt}]")
            total_inserted += 1

    wb.save(excel_path)
    print(f"[OK] Đã lưu thành công file Excel: {excel_path} (Tổng cộng: {total_inserted} dòng mới)")
    return total_inserted

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Cập nhật dữ liệu công văn cảnh báo vào sheet F0/F1 sổ Excel")
    parser.add_argument("json_path", type=str, help="Đường dẫn file JSON chứa danh sách records")
    parser.add_argument("--excel", type=str, default=DEFAULT_EXCEL_PATH, help="Đường dẫn file Excel đích")
    parser.add_argument("--sheet", type=str, default=None, choices=["F0", "F1"], help="Chỉ định sheet đích (F0 hoặc F1)")
    parser.add_argument("--no-backup", action="store_true", help="Bỏ qua bước sao lưu trước khi ghi")
    args = parser.parse_args()

    json_p = Path(args.json_path)
    if not json_p.is_file():
        print(f"[ERROR] Không tìm thấy file JSON: {json_p}")
        sys.exit(1)

    with open(json_p, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict) and "records" in data:
        data = data["records"]
    if not isinstance(data, list):
        data = [data]

    update_workbook(data, excel_path=args.excel, backup=not args.no_backup, sheet_name=args.sheet)
