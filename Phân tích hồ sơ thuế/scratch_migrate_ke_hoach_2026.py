#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
KỊCH BẢN THỰC THI CHUẨN HÓA VÀ QUY HOẠCH HỒ SƠ 13 DOANH NGHIỆP KẾ HOẠCH 2026
Tự động sao chép cấu trúc mẫu 4 bước, gom tờ khai XML, phân loại sổ sách kế toán,
và xuất phiếu kết quả rà soát Excel A4 cho từng doanh nghiệp.
"""

import os
import sys
import tempfile
import zipfile
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH")
TEMPLATE_DIR = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\thu-muc-mau-ho-so-doanh-nghiep")

COMPANIES = [
    {
        "stt": 1,
        "name": "CÔNG TY TNHH THƯƠNG MẠI SẢN XUẤT SƠN TIỀN HẢI",
        "short_name": "SƠN TIỀN HẢI",
        "mst": "3603859012",
        "target_folder": "SON_TIEN_HAI_3603859012",
        "sources": [
            BASE_DIR / "1. SON TIEN HAI"
        ]
    },
    {
        "stt": 2,
        "name": "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ T&T BIÊN HÒA",
        "short_name": "T&T BIÊN HÒA",
        "mst": "3603883505",
        "target_folder": "TT_BIEN_HOA_3603883505",
        "sources": [
            BASE_DIR / "2. T&T BHOA",
            BASE_DIR / "To khai" / "3603883505",
            BASE_DIR / "To khai" / "3603883505-TT"
        ]
    },
    {
        "stt": 3,
        "name": "CÔNG TY TNHH PHÚC NGUYÊN TIẾN",
        "short_name": "PHÚC NGUYÊN TIẾN",
        "mst": "3603698904",
        "target_folder": "PHUC_NGUYEN_TIEN_3603698904",
        "sources": [
            BASE_DIR / "3. PHUC NGUYEN TIEN",
            BASE_DIR / "To khai" / "3603698904",
            BASE_DIR / "To khai" / "3603698904-PHUC NGUYEN TIEN"
        ]
    },
    {
        "stt": 4,
        "name": "CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ TÂM HOÀNG PHÁT",
        "short_name": "TÂM HOÀNG PHÁT",
        "mst": "3603721832",
        "target_folder": "TAM_HOANG_PHAT_3603721832",
        "sources": [
            BASE_DIR / "4. TAM HOANG PHAT"
        ]
    },
    {
        "stt": 5,
        "name": "CÔNG TY TNHH MTV SX-TM-DV HÙNG ANH PHÁT",
        "short_name": "HÙNG ANH PHÁT",
        "mst": "3602451559",
        "target_folder": "HUNG_ANH_PHAT_3602451559",
        "sources": [
            BASE_DIR / "5. HUNG ANH PHAT"
        ]
    },
    {
        "stt": 6,
        "name": "CÔNG TY TNHH ĐẠI HUY HOÀNG",
        "short_name": "ĐẠI HUY HOÀNG",
        "mst": "3602179712",
        "target_folder": "DAI_HUY_HOANG_3602179712",
        "sources": [
            BASE_DIR / "6. DAI HUY HOANG",
            BASE_DIR / "To khai" / "3602179712"
        ]
    },
    {
        "stt": 7,
        "name": "CÔNG TY TNHH MTV NHẤT VINH",
        "short_name": "NHẤT VINH",
        "mst": "3601946069",
        "target_folder": "NHAT_VINH_3601946069",
        "sources": [
            BASE_DIR / "7. NHAT VINH-PASS LAM",
            BASE_DIR / "To khai" / "3601946069"
        ]
    },
    {
        "stt": 8,
        "name": "CÔNG TY TNHH TM DV VÀ SX HUY THÀNH PHÁT",
        "short_name": "HUY THÀNH PHÁT",
        "mst": "3603820015",
        "target_folder": "HUY_THANH_PHAT_3603820015",
        "sources": [
            BASE_DIR / "8. HUY THANH DAT",
            BASE_DIR / "To khai" / "3603820015",
            BASE_DIR / "To khai" / "3603820015-HUY THANH PHAT"
        ]
    },
    {
        "stt": 9,
        "name": "CÔNG TY TNHH TM VÀ DV TRÍ QUYÊN",
        "short_name": "TRÍ QUYÊN",
        "mst": "3603157158",
        "target_folder": "TRI_QUYEN_3603157158",
        "sources": [
            BASE_DIR / "9. TRIQUYEN"
        ]
    },
    {
        "stt": 10,
        "name": "CÔNG TY TNHH SẢN XUẤT VÀ THƯƠNG MẠI PHÁT HOÀNG LONG",
        "short_name": "PHÁT HOÀNG LONG",
        "mst": "3600974507",
        "target_folder": "PHAT_HOANG_LONG_3600974507",
        "sources": [
            BASE_DIR / ". Phat hoang long"
        ]
    },
    {
        "stt": 11,
        "name": "CÔNG TY TNHH MỘT THÀNH VIÊN AN PHƯƠNG TÂM",
        "short_name": "AN PHƯƠNG TÂM",
        "mst": "3602694223",
        "target_folder": "AN_PHUONG_TAM_3602694223",
        "sources": [
            BASE_DIR / "HO SO KE KOACH" / "2. An Phuong Tam",
            BASE_DIR / "HO SO KE KOACH" / "Ra soat an phuong tam- dong khanh.docx",
            BASE_DIR / "Tai khoan.docx",
            BASE_DIR / "Tai khoan ngan hang APT, DK.xlsx"
        ]
    },
    {
        "stt": 12,
        "name": "CÔNG TY TNHH THANH THANH PHƯƠNG",
        "short_name": "THANH THANH PHƯƠNG",
        "mst": "3600840461",
        "target_folder": "THANH_THANH_PHUONG_3600840461",
        "sources": [
            BASE_DIR / "To khai" / "3600840461",
            BASE_DIR / "To khai" / "3600840461-CÔNG TY TNHH THANH THANH PHƯƠNG"
        ]
    },
    {
        "stt": 13,
        "name": "CÔNG TY TNHH SX - TM VĨNH TRƯỜNG PHÁT",
        "short_name": "VĨNH TRƯỜNG PHÁT",
        "mst": "3600854351",
        "target_folder": "VINH_TRUONG_PHAT_3600854351",
        "sources": [
            BASE_DIR / "To khai" / "3600854351"
        ]
    }
]

def safe_copy(src, dst):
    """Sao chép an toàn tránh lỗi CopyFile2 trên ổ đĩa ảo/mạng Google Drive"""
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists() and dst.stat().st_size == src.stat().st_size:
        return
    with open(src, "rb") as fsrc:
        with open(dst, "wb") as fdst:
            while True:
                buf = fsrc.read(1024 * 1024)
                if not buf:
                    break
                fdst.write(buf)

def classify_target_subfolder(file_path):
    fn = file_path.name.lower()
    parents = [p.name.lower() for p in file_path.parents]
    pf_str = " ".join(parents)
    
    # 1.1 To khai XML
    if fn.endswith(".xml"):
        return "1. DU LIEU THO NGANH THUE/1.1. To khai XML"
        
    # 1.3 PTDN / Nghia vu / Thong bao
    if any(k in fn for k in ["ptdn01", "ptdn02", "thong bao", "etax", "nghia vu", "nop thue"]):
        return "1. DU LIEU THO NGANH THUE/1.3. Nghia vu nop thue & Thong bao"
        
    # 1.2 HDDT & Bang ke
    if any(k in fn for k in ["bang ke", "bang_ke", "thbk", "hddt", "hoa don", "data-hddt", "tieu_de", "th hd", "thmv-br", "ban ra", "mua vao"]):
        return "1. DU LIEU THO NGANH THUE/1.2. Hoa don dien tu & Bang ke"
    if any(k in pf_str for k in ["bang ke", "ds hóa đơn", "th hd"]):
        return "1. DU LIEU THO NGANH THUE/1.2. Hoa don dien tu & Bang ke"

    # 3.2 Danh ba NNT
    if any(k in fn for k in ["danh ba", "danh sach ncc"]):
        return "3. DU LIEU BO SUNG & TRA CUU QLT/3.2. Danh ba trang thai NNT"
        
    # 3.3 TNCN da nguon
    if "tncn" in fn and not fn.endswith(".xml"):
        return "3. DU LIEU BO SUNG & TRA CUU QLT/3.3. Tra cuu TNCN da nguon & Nhan su"
        
    # 2.3 Ngan hang
    if any(k in fn for k in ["sao ke", "ngan hang", "so phu", "vcb", "bidv", "hdbank", "publicbank", "eximbank", "acb", "vietinbank", "vietcombank", "tai khoan", "unc"]):
        return "2. HO SO NNT CUNG CAP/2.3. Ngan hang & Dong tien"
    if any(k in pf_str for k in ["sao ke", "ngan hang", "so phu"]):
        return "2. HO SO NNT CUNG CAP/2.3. Ngan hang & Dong tien"

    # 2.4 Lao dong, Tien luong, TSCD
    if any(k in fn for k in ["luong", "cham cong", "hop dong lao dong", "hdld", "khau hao", "tscd", "ccdc", "214", "242"]):
        return "2. HO SO NNT CUNG CAP/2.4. Lao dong - Tien luong - TSCD"
    if any(k in pf_str for k in ["luong", "khau hao"]):
        return "2. HO SO NNT CUNG CAP/2.4. Lao dong - Tien luong - TSCD"

    # 2.2 Kho, Gia thanh
    if any(k in fn for k in ["xnt", "ton kho", "gia thanh", "dinh muc", "vat tu", "152", "155", "156", "154", "tram xe", "nguyen lieu"]):
        return "2. HO SO NNT CUNG CAP/2.2. Kho - Chi phi - Gia thanh"
    if any(k in pf_str for k in ["nxt", "kho", "gia thanh", "nguyen lieu"]):
        return "2. HO SO NNT CUNG CAP/2.2. Kho - Chi phi - Gia thanh"

    # 2.1 So ke toan
    if any(k in fn for k in ["nkc", "nhat ky chung", "so cai", "sctk", "so chi tiet", "cong no", "phai thu", "phai tra", "can doi", "cdps"]):
        return "2. HO SO NNT CUNG CAP/2.1. So ke toan tong hop va chi tiet"
    if "so kt" in pf_str:
        return "2. HO SO NNT CUNG CAP/2.1. So ke toan tong hop va chi tiet"
        
    # 3.4 Giai trinh, QĐ, BB, Ho so bo sung
    if any(k in fn for k in ["quyet dinh", "bien ban", "giai trinh", "ra soat", "hop dong", "ghi chu", "gui cty", "thp.docx"]):
        return "3. DU LIEU BO SUNG & TRA CUU QLT/3.4. Giai trinh & Ho so bo sung"

    return "2. HO SO NNT CUNG CAP/2.1. So ke toan tong hop va chi tiet"

def copy_template_to_company(target_path):
    target_path.mkdir(parents=True, exist_ok=True)
    
    # 4 Main folders
    folders = [
        "1. DU LIEU THO NGANH THUE/1.1. To khai XML",
        "1. DU LIEU THO NGANH THUE/1.2. Hoa don dien tu & Bang ke",
        "1. DU LIEU THO NGANH THUE/1.3. Nghia vu nop thue & Thong bao",
        "2. HO SO NNT CUNG CAP/2.1. So ke toan tong hop va chi tiet",
        "2. HO SO NNT CUNG CAP/2.2. Kho - Chi phi - Gia thanh",
        "2. HO SO NNT CUNG CAP/2.3. Ngan hang & Dong tien",
        "2. HO SO NNT CUNG CAP/2.4. Lao dong - Tien luong - TSCD",
        "3. DU LIEU BO SUNG & TRA CUU QLT/3.1. Canh bao rui ro CVCBRR & Cong an",
        "3. DU LIEU BO SUNG & TRA CUU QLT/3.2. Danh ba trang thai NNT",
        "3. DU LIEU BO SUNG & TRA CUU QLT/3.3. Tra cuu TNCN da nguon & Nhan su",
        "3. DU LIEU BO SUNG & TRA CUU QLT/3.4. Giai trinh & Ho so bo sung",
        "4. KET QUA PHAN TICH/cau-thanh",
        "4. KET QUA PHAN TICH/sao-luu"
    ]
    for sub in folders:
        (target_path / sub).mkdir(parents=True, exist_ok=True)
        
    # Copy template files
    template_files = [
        "CHECKLIST_TONG_HOP_HO_SO.md",
        "CHECKLIST_TONG_HOP_HO_SO.xlsx",
        "kiem_tra_tinh_day_du_ho_so.py",
        "README.md",
        "1. DU LIEU THO NGANH THUE/CHECKLIST_BUOC_1_DU_LIEU_THO_CQT.xlsx",
        "1. DU LIEU THO NGANH THUE/README_CHECKLIST_BUOC_1.md",
        "2. HO SO NNT CUNG CAP/CHECKLIST_BUOC_2_HO_SO_NNT_CUNG_CAP.xlsx",
        "2. HO SO NNT CUNG CAP/README_CHECKLIST_BUOC_2.md",
        "3. DU LIEU BO SUNG & TRA CUU QLT/CHECKLIST_BUOC_3_DU_LIEU_BO_SUNG_QLT.xlsx",
        "3. DU LIEU BO SUNG & TRA CUU QLT/README_CHECKLIST_BUOC_3.md",
        "4. KET QUA PHAN TICH/CHECKLIST_BUOC_4_KET_QUA_PHAN_TICH.xlsx",
        "4. KET QUA PHAN TICH/README_CHECKLIST_BUOC_4.md"
    ]
    for tf in template_files:
        src = TEMPLATE_DIR / tf
        dst = target_path / tf
        if src.exists():
            safe_copy(src, dst)

def process_company(comp):
    print(f"\n=======================================================")
    print(f"[{comp['stt']}/13] Đang xử lý: {comp['name']} (MST: {comp['mst']})")
    print(f"Thư mục đích: {comp['target_folder']}")
    
    target_path = BASE_DIR / comp["target_folder"]
    copy_template_to_company(target_path)
    
    total_copied = 0
    
    for src_entry in comp["sources"]:
        if not src_entry.exists():
            print(f"  [CẢNH BÁO] Nguồn không tồn tại: {src_entry}")
            continue
            
        if src_entry.is_file():
            # Xử lý file đơn
            if src_entry.suffix.lower() == ".zip" and comp["mst"] == "3603820015":
                print(f"  -> Giải nén file zip tài liệu Huy Thành Phát...")
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_p = Path(temp_dir)
                    with zipfile.ZipFile(src_entry, 'r') as zf:
                        zf.extractall(temp_p)
                    for extracted_file in temp_p.iterdir():
                        sub = classify_target_subfolder(extracted_file)
                        dst_file = target_path / sub / extracted_file.name
                        safe_copy(extracted_file, dst_file)
                        total_copied += 1
            else:
                sub = classify_target_subfolder(src_entry)
                dst_file = target_path / sub / src_entry.name
                safe_copy(src_entry, dst_file)
                total_copied += 1
        elif src_entry.is_dir():
            for root, dirs, files in os.walk(src_entry):
                for f in files:
                    file_p = Path(root) / f
                    if file_p.name.startswith("CHECKLIST") or file_p.name.startswith("README") or file_p.name.startswith("kiem_tra_"):
                        continue
                    sub = classify_target_subfolder(file_p)
                    dst_file = target_path / sub / f
                    safe_copy(file_p, dst_file)
                    total_copied += 1
                        
    print(f"  -> Đã phân loại và sao chép: {total_copied} file vào 4 cấp thư mục chuẩn.")
    return total_copied

# Import logic quét và xuất báo cáo từ module kiem_tra_tinh_day_du_ho_so
sys.path.append(str(TEMPLATE_DIR))
import kiem_tra_tinh_day_du_ho_so as checker

def run_checks_and_export_reports():
    print("\n=======================================================")
    print("Bắt đầu chạy rà soát tự động và xuất báo cáo Checklist cho 13 DN...")
    
    all_reports = []
    
    for comp in COMPANIES:
        target_path = BASE_DIR / comp["target_folder"]
        report = checker.scan_directory(str(target_path))
        if report:
            report["comp_info"] = comp
            all_reports.append(report)
            
            # Xuất file markdown và Excel checklist 1 trang A4
            md_path = target_path / "CHECKLIST_TONG_HOP_HO_SO.md"
            excel_path = target_path / "Phieu_ket_qua_ra_soat.xlsx"
            
            md_text = checker.format_markdown_report(report)
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_text)
                
            checker.export_excel_report(report, str(excel_path))
            print(f"  -> Đã xuất Checklist cho {comp['short_name']}: {report['so_muc_dat']}/{report['tong_so_muc']} mục đạt.")

    # Xuất Master Workbook tổng hợp tiến độ 13 doanh nghiệp
    export_master_overview(all_reports)

def export_master_overview(all_reports):
    master_excel_path = BASE_DIR / "TONG_HOP_TIEN_DO_13_DOANH_NGHIEP_2026.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tong_hop_13_DN_Ke_hoach_2026"
    
    font_name = 'Times New Roman'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.views.sheetView[0].showGridLines = True

    f_org = Font(name=font_name, size=9.5, bold=True)
    f_title = Font(name=font_name, size=14, bold=True, color='1F4E79')
    f_sub = Font(name=font_name, size=9.5, italic=True, color='595959')
    f_tbl_hdr = Font(name=font_name, size=9, bold=True, color='FFFFFF')
    f_cell = Font(name=font_name, size=8.5)
    f_cell_bold = Font(name=font_name, size=8.5, bold=True)
    f_pass = Font(name=font_name, size=8.5, bold=True, color='006100')
    f_fail = Font(name=font_name, size=8.5, bold=True, color='9C0006')

    fill_hdr = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
    fill_pass = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    fill_fail = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    fill_alt = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')
    fill_highlight = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )

    al_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    al_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Header
    ws.merge_cells("A1:C1"); ws["A1"] = "CỤC THUẾ TỈNH ĐỒNG NAI"; ws["A1"].font = f_org; ws["A1"].alignment = al_center
    ws.merge_cells("A2:C2"); ws["A2"] = "ĐOÀN KIỂM TRA THUẾ NĂM 2026"; ws["A2"].font = f_org; ws["A2"].alignment = al_center
    ws.merge_cells("E1:H1"); ws["E1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"; ws["E1"].font = f_org; ws["E1"].alignment = al_center
    ws.merge_cells("E2:H2"); ws["E2"] = "Độc lập - Tự do - Hạnh phúc"; ws["E2"].font = Font(name=font_name, size=8.5, italic=True); ws["E2"].alignment = al_center

    ws.merge_cells("A4:H4"); ws["A4"] = "BẢNG TỔNG HỢP TIẾN ĐỘ THU THẬP & RÀ SOÁT HỒ SƠ 13 DOANH NGHIỆP KẾ HOẠCH 2026"
    ws["A4"].font = f_title; ws["A4"].alignment = al_center; ws.row_dimensions[4].height = 22

    ws.merge_cells("A5:H5"); ws["A5"] = "(Quy chuẩn theo Bộ thư mục mẫu 4 bước và Kịch bản kiểm soát tính đầy đủ hồ sơ)"
    ws["A5"].font = f_sub; ws["A5"].alignment = al_center; ws.row_dimensions[5].height = 14

    headers = [
        ("STT", 4.5),
        ("Tên Doanh nghiệp", 28),
        ("Mã số thuế", 13),
        ("Thư mục quy hoạch chuẩn", 26),
        ("B1. Dữ liệu Thuế", 14),
        ("B2. Hồ sơ NNT", 14),
        ("B3. Dữ liệu QLT", 14),
        ("Đánh giá tình trạng", 18)
    ]
    for col_idx, (h_text, w) in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col_idx, value=h_text)
        cell.font = f_tbl_hdr; cell.fill = fill_hdr; cell.alignment = al_center
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[7].height = 20

    row_idx = 8
    for r in all_reports:
        comp = r["comp_info"]
        ws.row_dimensions[row_idx].height = 18
        
        # Đếm file từng bước
        b1_files = sum(mc["so_luong"] for mc in r["steps"]["BUOC_1"]["muc_con"])
        b2_files = sum(mc["so_luong"] for mc in r["steps"]["BUOC_2"]["muc_con"])
        b3_files = sum(mc["so_luong"] for mc in r["steps"]["BUOC_3"]["muc_con"])
        
        c1 = ws.cell(row=row_idx, column=1, value=comp["stt"])
        c2 = ws.cell(row=row_idx, column=2, value=comp["name"])
        c3 = ws.cell(row=row_idx, column=3, value=comp["mst"])
        c4 = ws.cell(row=row_idx, column=4, value=comp["target_folder"])
        c5 = ws.cell(row=row_idx, column=5, value=f"{b1_files} file")
        c6 = ws.cell(row=row_idx, column=6, value=f"{b2_files} file")
        c7 = ws.cell(row=row_idx, column=7, value=f"{b3_files} file")
        
        status_str = f"Thiếu {r['so_muc_thieu_bat_buoc']} mục bắt buộc" if r["so_muc_thieu_bat_buoc"] > 0 else "ĐỦ ĐIỀU KIỆN"
        c8 = ws.cell(row=row_idx, column=8, value=status_str)
        
        c1.alignment = al_center; c2.alignment = al_left; c3.alignment = al_center
        c4.alignment = al_left; c5.alignment = al_center; c6.alignment = al_center
        c7.alignment = al_center; c8.alignment = al_center

        c1.font = f_cell_bold; c2.font = f_cell_bold; c3.font = f_cell_bold; c4.font = f_cell
        c5.font = f_cell; c6.font = f_cell; c7.font = f_cell
        c8.font = f_pass if r["so_muc_thieu_bat_buoc"] == 0 else f_fail
        c8.fill = fill_pass if r["so_muc_thieu_bat_buoc"] == 0 else fill_fail

        for col_i in range(1, 9):
            ws.cell(row=row_idx, column=col_i).border = border_thin
            if row_idx % 2 == 1 and col_i != 8:
                ws.cell(row=row_idx, column=col_i).fill = fill_alt

        row_idx += 1

    wb.save(master_excel_path)
    print(f"\n-> Đã xuất Bảng tổng hợp tiến độ 13 Doanh nghiệp: {master_excel_path}")

def main():
    print("Bắt đầu thực thi quy hoạch và phân loại 13 doanh nghiệp...")
    for comp in COMPANIES:
        process_company(comp)
    print("\nHoàn tất sao chép và phân loại file!")
    
    run_checks_and_export_reports()
    print("\n>>> TOÀN BỘ QUY TRÌNH THỰC THI HOÀN TẤT THÀNH CÔNG! <<<")

if __name__ == "__main__":
    main()
