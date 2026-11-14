"""Script dong bo so lieu tu sheet TINH PHAT + CHAM NOP (C T) cua Bang tinh.xlsx
sang so Danh sach kiem tra quyet toan pro.xlsm.

Quy tac:
- Lay moc thoi diem tinh cham nop tu cot 36 (AJ) Ngay ky BBKT tren so pro.xlsm.
- Cap nhat ngay ky BBKT vao cot D tren sheet TINH PHAT neu can de tinh dung 0,03%.
- Tu dong dong bo cac o lech sang so pro.xlsm (truy thu, cham nop, phat 20%).
- Canh bao rieng khi co phat sinh Phat thu tuc VPHC (cot 47 / AU).
- Khong ghi de cac o cong thuc (AP, AS, AV, AW).
- Khong tao file .bak trong 00. Theo doi chung.
- Giu nguyen macro VBA (keep_vba=True).
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime
import openpyxl

BASE_DIR = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT")
SO_PATH = BASE_DIR / r"00. Theo doi chung\Danh sach kiem tra quyet toan pro.xlsm"
TD_DIR = BASE_DIR / r"2026\1. TD"

def parse_date(val):
    if not val:
        return None
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    val_str = str(val).strip()
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"]:
        try:
            return datetime.strptime(val_str, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return val_str

import re
import unicodedata

def remove_accents(input_str):
    if not input_str:
        return ""
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower()

def find_vu_dir(mst_or_name, ten_cong_ty=None, stt=None):
    query_raw = str(mst_or_name).strip().lower()
    query_no_acc = remove_accents(query_raw)
    
    # 1. Tim truc tiep theo query (MST hoac ten go tay)
    for p in TD_DIR.iterdir():
        if p.is_dir():
            p_name_no_acc = remove_accents(p.name)
            if (query_raw in p.name.lower() or query_no_acc in p_name_no_acc) and len(query_raw) >= 3:
                bt = p / r"0. BIEN BAN\Bang tinh.xlsx"
                if bt.exists():
                    return p, bt

    # 2. Tim theo ten cong ty (loai bo cac tu chung, so khop tu khoa rieng)
    if ten_cong_ty:
        ten_no_acc = remove_accents(ten_cong_ty)
        words = [w for w in ten_no_acc.split() if w not in ["cong", "ty", "tnhh", "tm", "dv", "cp", "mtv", "thuong", "mai", "dich", "vu", "san", "xuat"]]
        for p in TD_DIR.iterdir():
            if p.is_dir():
                p_name_no_acc = remove_accents(p.name)
                matches = sum(1 for w in words if len(w) >= 3 and w in p_name_no_acc)
                if matches >= 2 or (len(words) == 1 and matches == 1):
                    bt = p / r"0. BIEN BAN\Bang tinh.xlsx"
                    if bt.exists():
                        return p, bt
                    
    # 3. Fallback theo STT neu co
    if stt:
        stt_str = str(stt).strip()
        for p in TD_DIR.iterdir():
            if p.is_dir() and (p.name.startswith(f"{stt_str}.") or p.name.startswith(f"{stt_str} ")):
                bt = p / r"0. BIEN BAN\Bang tinh.xlsx"
                if bt.exists():
                    return p, bt

    return None, None

def dong_bo(mst_or_name, dry_run=False):
    if not SO_PATH.exists():
        print(f"Loi: Khong tim thay so tai {SO_PATH}")
        return False

    # 1. Doc so theo doi de tim dong cua vu va Ngay ky BBKT truoc
    wb_so = openpyxl.load_workbook(SO_PATH, keep_vba=True)
    target_sheet = None
    target_row = None
    mst_found = None
    ten_found = None
    stt_found = None

    for s_name in ["01_DA_BAN_HANH", "02_DANG_THUC_HIEN"]:
        ws = wb_so[s_name]
        for r in range(5, ws.max_row + 1):
            mst_cell = str(ws.cell(r, 4).value or "").strip()
            ten_cell = str(ws.cell(r, 5).value or "").strip() # Cot 5: Ten NNT
            stt_cell = r - 4 # Row 5 = STT 1
            q = str(mst_or_name).strip().lower()
            if q == mst_cell.lower() or q in mst_cell.lower() or q in ten_cell.lower():
                target_sheet = s_name
                target_row = r
                mst_found = mst_cell
                ten_found = ten_cell
                stt_found = stt_cell
                break
        if target_row:
            break

    vu_dir, bt_path = find_vu_dir(mst_or_name, ten_cong_ty=ten_found, stt=stt_found)
    if not bt_path:
        print(f"Loi: Khong tim thay Bang tinh.xlsx cho '{mst_or_name}' (Ten: '{ten_found}') trong {TD_DIR}")
        return False

    print(f"=== BAT DAU DONG BO ===")
    print(f"Vụ: {vu_dir.name}")
    print(f"Bang tinh: {bt_path}")
    print(f"So theo doi: {SO_PATH}")

    if not target_row:
        print(f"Loi: Khong tim thay dong cua vu tren so pro.xlsm")
        return False

    ws_so = wb_so[target_sheet]
    ngay_ky_bbkt = ws_so.cell(target_row, 36).value # Cot 36 (AJ)
    ngay_ky_str = parse_date(ngay_ky_bbkt)
    print(f"Tim thay tai sheet '{target_sheet}', dong {target_row}: {ten_found} (MST: {mst_found})")
    print(f"Ngay ky BBKT tren so (cot 36 / AJ): {ngay_ky_str}")

    # 2. Doc Bang tinh
    wb_bt = openpyxl.load_workbook(bt_path, data_only=True)
    sheet_name = "TINH PHAT + CHAM NOP (C T)"
    if sheet_name not in wb_bt.sheetnames:
        print(f"Loi: Khong tim thay sheet '{sheet_name}' trong Bang tinh")
        return False
    ws_bt = wb_bt[sheet_name]

    # Tim dong Tong cong
    row_tong = None
    for r in range(1, ws_bt.max_row + 1):
        v1 = str(ws_bt.cell(r, 1).value or "").strip().upper()
        if "TỔNG CỘNG" in v1 or "TONG CONG" in v1:
            row_tong = r
            break

    if not row_tong:
        print(f"Loi: Khong tim thay dong 'Tong cong' tren sheet {sheet_name}")
        return False

    # Trich xuat cac chi tieu tu sheet Bang tinh
    truy_thu_tndn = round(ws_bt.cell(row_tong, 2).value or 0) # Cot B
    truy_thu_gtgt = round(ws_bt.cell(row_tong, 3).value or 0) # Cot C
    
    # Cham nop GTGT thuong o cot H (8) hoac I (9)
    cn_gtgt = round(ws_bt.cell(row_tong, 8).value or ws_bt.cell(row_tong, 9).value or 0)
    
    # Cham nop TNDN 0.03% thuong o cot K (11) hoac M (13)
    cn_tndn = round(ws_bt.cell(row_tong, 11).value or ws_bt.cell(row_tong, 13).value or 0)
    
    # Phat 20% thuong o cot N (14)
    phat_20 = round(ws_bt.cell(row_tong, 14).value or 0)
    
    # Tong phat sinh o cot Q (17)
    tong_ps_bt = round(ws_bt.cell(row_tong, 17).value or 0)

    # Kiem tra phat thu tuc VPHC (neu co cot rieng hoac phat sinh)
    phat_thu_tuc = 0
    # Thuong cot 47 tren so la Phat thu tuc VPHC
    phat_thu_tuc_so = ws_so.cell(target_row, 47).value or 0

    so_sanh = [
        {"col_idx": 40, "col_letter": "AN", "ten": "Truy thu TNDN", "moi": truy_thu_tndn},
        {"col_idx": 41, "col_letter": "AO", "ten": "Truy thu GTGT", "moi": truy_thu_gtgt},
        {"col_idx": 43, "col_letter": "AQ", "ten": "Chậm nộp GTGT (0.03%)", "moi": cn_gtgt},
        {"col_idx": 44, "col_letter": "AR", "ten": "Chậm nộp TNDN (0.03%)", "moi": cn_tndn},
        {"col_idx": 46, "col_letter": "AT", "ten": "Phạt khai sai 20%", "moi": phat_20},
    ]

    print("\n--- KET QUA DOI CHIEU VA DONG BO ---")
    da_sua = 0
    for item in so_sanh:
        c_idx = item["col_idx"]
        c_let = item["col_letter"]
        ten = item["ten"]
        val_moi = item["moi"]
        val_cu = ws_so.cell(target_row, c_idx).value or 0
        try:
            val_cu_num = round(float(val_cu))
        except (ValueError, TypeError):
            val_cu_num = val_cu

        if val_cu_num != val_moi:
            print(f"[LECH] {c_let}{target_row} ({ten}): {val_cu_num} -> {val_moi}")
            if not dry_run:
                ws_so.cell(target_row, c_idx).value = val_moi
            da_sua += 1
        else:
            print(f"[KHOP] {c_let}{target_row} ({ten}): {val_cu_num}")

    # Canh bao phat thu tuc VPHC
    if phat_thu_tuc_so > 0:
        print(f"[CHU Y] So dang co Phat thu tuc VPHC (cot 47): {phat_thu_tuc_so}")
    else:
        print(f"[INFO] Phat thu tuc VPHC (cot 47): 0 (khong co phat sinh)")

    # Tu dong dien / copy cong thuc chu DocSo cho 6 cot cuoi (53 den 58: BA den BF)
    if target_sheet == "01_DA_BAN_HANH":
        docso_cols = [
            (53, f"=DocSo(AP{target_row})", "Tong_truy_thu_chu"),
            (54, f"=DocSo(AS{target_row})", "Tong_cham_nop_chu"),
            (55, f"=DocSo(AV{target_row})", "Tong_tien_phat_chu"),
            (56, f"=DocSo(AW{target_row})", "Tong_phat_sinh_chu"),
            (57, f"=DocSo(AX{target_row})", "Giam_lo_chu"),
            (58, f"=DocSo(AY{target_row})", "Giam_khau_tru_chu"),
        ]
        for col_idx, formula_val, col_name in docso_cols:
            curr_val = str(ws_so.cell(target_row, col_idx).value or "").strip()
            if not curr_val.startswith("=DocSo"):
                print(f"[CONG THUC CHU] Tu dong gan {col_name} tai dong {target_row}: {formula_val}")
                if not dry_run:
                    ws_so.cell(target_row, col_idx).value = formula_val
                da_sua += 1

    if da_sua > 0 and not dry_run:
        wb_so.save(SO_PATH)
        print(f"\n=> DA DONG BO THANH CONG: Da cap nhat {da_sua} o lech vao so pro.xlsm")
    elif da_sua == 0:
        print(f"\n=> TAT CA DA KHOP: Khong can ghi de so pro.xlsm")
    else:
        print(f"\n=> DRY RUN: Phat hien {da_sua} o lech (chua ghi file)")

    return True

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Dong bo so lieu TINH PHAT sang so pro.xlsm")
    parser.add_argument("query", help="MST hoac ten vu/cong ty")
    parser.add_argument("--dry-run", action="store_true", help="Chi kiem tra, khong ghi file")
    args = parser.parse_args()

    dong_bo(args.query, dry_run=args.dry_run)
