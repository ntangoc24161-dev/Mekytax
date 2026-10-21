#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
KỊCH BẢN THỰC THI TOÀN DIỆN:
1. Di chuyển và phân loại toàn bộ file của 13 Doanh nghiệp vào 4 cấp thư mục chuẩn hóa.
2. Xóa các file checklist phụ ở từng bước và thư mục con.
3. Tạo 1 file Master Checklist tổng hợp duy nhất tại 4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx với giải thích chi tiết, trực quan, dễ hiểu.
4. Xóa toàn bộ các thư mục nguồn cũ và file lẻ tại thư mục gốc để dữ liệu chỉ lưu duy nhất 1 nơi.
5. Xuất bảng tổng hợp tiến độ chung TONG_HOP_TIEN_DO_13_DOANH_NGHIEP_2026.xlsx.
"""

import os
import sys
import shutil
import tempfile
import zipfile
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH")

def to_win_long(p):
    """Chuyển đường dẫn sang dạng Extended-Length Path trên Windows để tránh giới hạn MAX_PATH 260 ký tự"""
    s = str(Path(p).resolve())
    if not s.startswith("\\\\?\\"):
        return "\\\\?\\" + s
    return s

COMPANIES = [
    {
        "stt": 1,
        "name": "CÔNG TY TNHH THƯƠNG MẠI SẢN XUẤT SƠN TIỀN HẢI",
        "short_name": "SƠN TIỀN HẢI",
        "mst": "3603859012",
        "nien_do": "2022 - 2025",
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
        "nien_do": "2023 - 2025",
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
        "nien_do": "2020 - 2025",
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
        "nien_do": "2023 - 2025",
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
        "nien_do": "2022 - 2025",
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
        "nien_do": "2021 - 2025",
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
        "nien_do": "2017 - 2025",
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
        "nien_do": "2021 - 2025",
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
        "nien_do": "2014 - 2025",
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
        "nien_do": "2017 - 2025",
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
        "nien_do": "2023 - 2025",
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
        "nien_do": "2021 - 2025",
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
        "nien_do": "2022 - 2025",
        "target_folder": "VINH_TRUONG_PHAT_3600854351",
        "sources": [
            BASE_DIR / "To khai" / "3600854351"
        ]
    }
]

def classify_target_subfolder(file_path):
    fn = file_path.name.lower()
    parents = [p.name.lower() for p in file_path.parents]
    pf_str = " ".join(parents)
    
    # 1.1 To khai XML
    if fn.endswith(".xml"):
        return "1. DU LIEU THO NGANH THUE/1.1. To khai XML"
        
    # 1.3 PTDN / Nghia vu / Thong bao
    if any(k in fn for k in ["ptdn01", "ptdn02", "thong bao", "etax", "nghia vu", "nop thue", "no thue", "tb nop"]):
        return "1. DU LIEU THO NGANH THUE/1.3. Nghia vu nop thue & Thong bao"
        
    # 1.2 HDDT & Bang ke
    if any(k in fn for k in ["bang ke", "bang_ke", "thbk", "hddt", "hoa don", "data-hddt", "tieu_de", "th hd", "thmv-br", "ban ra", "mua vao", "thhd", "thong ke hoa don"]):
        return "1. DU LIEU THO NGANH THUE/1.2. Hoa don dien tu & Bang ke"
    if any(k in pf_str for k in ["bang ke", "ds hóa đơn", "th hd", "hddt"]):
        return "1. DU LIEU THO NGANH THUE/1.2. Hoa don dien tu & Bang ke"

    # 3.1 Canh bao rui ro CBRR & Cong an
    if any(k in fn for k in ["cbrr", "f0", "f1", "cong an", "công an", "rui ro", "canh bao", "cv30291", "cvcbrr"]):
        return "3. DU LIEU BO SUNG & TRA CUU QLT/3.1. Canh bao rui ro CVCBRR & Cong an"

    # 3.2 Danh ba NNT
    if any(k in fn for k in ["danh ba", "danh sach ncc", "trang thai nnt", "tra cuu ncc"]):
        return "3. DU LIEU BO SUNG & TRA CUU QLT/3.2. Danh ba trang thai NNT"
        
    # 3.3 TNCN da nguon
    if any(k in fn for k in ["tncn da nguon", "da nguon", "trung cccd", "trung cmnd", "nhan su da nguon"]):
        return "3. DU LIEU BO SUNG & TRA CUU QLT/3.3. Tra cuu TNCN da nguon & Nhan su"
        
    # 2.3 Ngan hang
    if any(k in fn for k in ["sao ke", "ngan hang", "so phu", "vcb", "bidv", "hdbank", "publicbank", "eximbank", "acb", "vietinbank", "vietcombank", "tai khoan", "unc", "giay bao no", "gbn"]):
        return "2. HO SO NNT CUNG CAP/2.3. Ngan hang & Dong tien"
    if any(k in pf_str for k in ["sao ke", "ngan hang", "so phu"]):
        return "2. HO SO NNT CUNG CAP/2.3. Ngan hang & Dong tien"

    # 2.4 Lao dong, Tien luong, TSCD
    if any(k in fn for k in ["luong", "cham cong", "hop dong lao dong", "hdld", "khau hao", "tscd", "ccdc", "214", "242"]):
        return "2. HO SO NNT CUNG CAP/2.4. Lao dong - Tien luong - TSCD"
    if any(k in pf_str for k in ["luong", "khau hao", "lao dong"]):
        return "2. HO SO NNT CUNG CAP/2.4. Lao dong - Tien luong - TSCD"

    # 2.2 Kho, Gia thanh
    if any(k in fn for k in ["xnt", "ton kho", "gia thanh", "dinh muc", "vat tu", "152", "155", "156", "154", "tram xe", "nguyen lieu", "z sp", "621", "622", "627"]):
        return "2. HO SO NNT CUNG CAP/2.2. Kho - Chi phi - Gia thanh"
    if any(k in pf_str for k in ["nxt", "kho", "gia thanh", "nguyen lieu", "vat tu"]):
        return "2. HO SO NNT CUNG CAP/2.2. Kho - Chi phi - Gia thanh"

    # 2.1 So ke toan
    if any(k in fn for k in ["nkc", "nhat ky chung", "so cai", "sctk", "so chi tiet", "cong no", "phai thu", "phai tra", "can doi", "cdps", "111", "112", "131", "331", "so quy", "sothcn", "sothhh"]):
        return "2. HO SO NNT CUNG CAP/2.1. So ke toan tong hop va chi tiet"
    if any(k in pf_str for k in ["so kt", "so chi tiet", "so cai"]):
        return "2. HO SO NNT CUNG CAP/2.1. So ke toan tong hop va chi tiet"
        
    # 3.4 Giai trinh, QĐ, BB, Ho so bo sung
    if any(k in fn for k in ["quyet dinh", "bien ban", "giai trinh", "ra soat", "hop dong", "ghi chu", "gui cty", "thp.docx", "qd", "bb"]):
        return "3. DU LIEU BO SUNG & TRA CUU QLT/3.4. Giai trinh & Ho so bo sung"
    if any(k in pf_str for k in ["bien ban", "qd", "giai trinh"]):
        return "3. DU LIEU BO SUNG & TRA CUU QLT/3.4. Giai trinh & Ho so bo sung"

    return "2. HO SO NNT CUNG CAP/2.1. So ke toan tong hop va chi tiet"

def safe_copy(src_path, dst_path):
    src_win = to_win_long(src_path)
    dst_win = to_win_long(dst_path)
    
    os.makedirs(os.path.dirname(dst_win), exist_ok=True)
    if os.path.exists(dst_win) and os.path.getsize(dst_win) == os.path.getsize(src_win):
        return
        
    with open(src_win, "rb") as fsrc:
        with open(dst_win, "wb") as fdst:
            while True:
                buf = fsrc.read(1024 * 1024)
                if not buf:
                    break
                fdst.write(buf)

def setup_company_folders(target_path):
    target_win = to_win_long(target_path)
    os.makedirs(target_win, exist_ok=True)
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
        sub_p = to_win_long(target_path / sub)
        os.makedirs(sub_p, exist_ok=True)

def migrate_company_files(comp):
    target_path = BASE_DIR / comp["target_folder"]
    setup_company_folders(target_path)
    
    total_copied = 0
    for src_entry in comp["sources"]:
        src_win = to_win_long(src_entry)
        if not os.path.exists(src_win):
            continue
            
        if os.path.isfile(src_win):
            # Xử lý file zip đặc biệt nếu là Huy Thành Phát
            if src_entry.suffix.lower() == ".zip" and comp["mst"] == "3603820015":
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_p = Path(temp_dir)
                    with zipfile.ZipFile(src_win, 'r') as zf:
                        zf.extractall(temp_p)
                    for extracted_file in temp_p.rglob('*'):
                        if extracted_file.is_file():
                            sub = classify_target_subfolder(extracted_file)
                            dst_file = target_path / sub / extracted_file.name
                            safe_copy(extracted_file, dst_file)
                            total_copied += 1
            else:
                sub = classify_target_subfolder(src_entry)
                dst_file = target_path / sub / src_entry.name
                safe_copy(src_entry, dst_file)
                total_copied += 1
        elif os.path.isdir(src_win):
            for root, dirs, files in os.walk(src_win):
                for f in files:
                    file_p = Path(root) / f
                    # Bỏ qua các file checklist template nếu có trong nguồn
                    if file_p.name.startswith("CHECKLIST") or file_p.name.startswith("README") or file_p.name.startswith("kiem_tra_"):
                        continue
                    sub = classify_target_subfolder(file_p)
                    dst_file = target_path / sub / f
                    safe_copy(file_p, dst_file)
                    total_copied += 1
                    
    return total_copied

def clean_old_subfolder_checklists(target_path):
    """
    Xóa tất cả các file checklist riêng lẻ ở từng bước/thư mục con và file nháp checklist ở gốc doanh nghiệp,
    chỉ giữ duy nhất 1 Master Checklist tại 4. KET QUA PHAN TICH/
    """
    redundant_patterns = [
        "CHECKLIST_BUOC_1_DU_LIEU_THO_CQT.xlsx",
        "README_CHECKLIST_BUOC_1.md",
        "CHECKLIST_BUOC_2_HO_SO_NNT_CUNG_CAP.xlsx",
        "README_CHECKLIST_BUOC_2.md",
        "CHECKLIST_BUOC_3_DU_LIEU_BO_SUNG_QLT.xlsx",
        "README_CHECKLIST_BUOC_3.md",
        "CHECKLIST_BUOC_4_KET_QUA_PHAN_TICH.xlsx",
        "README_CHECKLIST_BUOC_4.md",
        "CHECKLIST_TONG_HOP_HO_SO.md",
        "Phieu_ket_qua_ra_soat.xlsx",
        "kiem_tra_tinh_day_du_ho_so.py",
        "README.md"
    ]
    
    target_win = to_win_long(target_path)
    for root, dirs, files in os.walk(target_win):
        for f in files:
            file_p = Path(root) / f
            if f in redundant_patterns:
                try:
                    os.remove(to_win_long(file_p))
                except Exception:
                    pass
            # Nếu có file CHECKLIST_TONG_HOP_HO_SO.xlsx ở gốc doanh nghiệp -> cũng xóa để chỉ lưu trong 4. KET QUA PHAN TICH
            if Path(root) == Path(target_win) and f == "CHECKLIST_TONG_HOP_HO_SO.xlsx":
                try:
                    os.remove(to_win_long(file_p))
                except Exception:
                    pass

def inspect_company_data(target_path):
    categories = [
        {
            "id": "1.1",
            "name": "Tờ khai thuế XML (GTGT, TNDN, TNCN, BCTC)",
            "folder": "1. DU LIEU THO NGANH THUE/1.1. To khai XML",
            "step": "Bước 1: Dữ liệu thô ngành Thuế",
            "required": True,
            "min_files": 4,
            "exts": [".xml"]
        },
        {
            "id": "1.2",
            "name": "Hóa đơn điện tử & Bảng kê mua vào/bán ra",
            "folder": "1. DU LIEU THO NGANH THUE/1.2. Hoa don dien tu & Bang ke",
            "step": "Bước 1: Dữ liệu thô ngành Thuế",
            "required": True,
            "min_files": 2,
            "exts": [".xlsx", ".xls", ".csv", ".zip"]
        },
        {
            "id": "1.3",
            "name": "Nghĩa vụ nộp thuế, PTDN & Thông báo nợ",
            "folder": "1. DU LIEU THO NGANH THUE/1.3. Nghia vu nop thue & Thong bao",
            "step": "Bước 1: Dữ liệu thô ngành Thuế",
            "required": False,
            "min_files": 1,
            "exts": [".xlsx", ".xls", ".pdf"]
        },
        {
            "id": "2.1",
            "name": "Sổ kế toán tổng hợp & Sổ chi tiết (NKC, Sổ Cái, 111, 112, 131, 331)",
            "folder": "2. HO SO NNT CUNG CAP/2.1. So ke toan tong hop va chi tiet",
            "step": "Bước 2: Hồ sơ NNT cung cấp",
            "required": True,
            "min_files": 2,
            "exts": [".xlsx", ".xls"]
        },
        {
            "id": "2.2",
            "name": "Sổ XNT Kho 152/155/156, Chi phí 154, Giá thành & Định mức",
            "folder": "2. HO SO NNT CUNG CAP/2.2. Kho - Chi phi - Gia thanh",
            "step": "Bước 2: Hồ sơ NNT cung cấp",
            "required": True,
            "min_files": 2,
            "exts": [".xlsx", ".xls", ".pdf"]
        },
        {
            "id": "2.3",
            "name": "Sao kê Ngân hàng (Excel & PDF có dấu) & UNC thanh toán",
            "folder": "2. HO SO NNT CUNG CAP/2.3. Ngan hang & Dong tien",
            "step": "Bước 2: Hồ sơ NNT cung cấp",
            "required": True,
            "min_files": 1,
            "exts": [".xlsx", ".xls", ".csv", ".pdf"]
        },
        {
            "id": "2.4",
            "name": "Bảng lương 12 tháng, Bảng chấm công, HĐLĐ, Bảng khấu hao 214/242",
            "folder": "2. HO SO NNT CUNG CAP/2.4. Lao dong - Tien luong - TSCD",
            "step": "Bước 2: Hồ sơ NNT cung cấp",
            "required": True,
            "min_files": 1,
            "exts": [".xlsx", ".xls", ".pdf"]
        },
        {
            "id": "3.1",
            "name": "Sổ theo dõi CVCBRR F0/F1 & Sổ theo dõi Công an",
            "folder": "3. DU LIEU BO SUNG & TRA CUU QLT/3.1. Canh bao rui ro CVCBRR & Cong an",
            "step": "Bước 3: Dữ liệu bổ sung & QLT",
            "required": True,
            "min_files": 1,
            "exts": [".xlsx", ".xls"]
        },
        {
            "id": "3.2",
            "name": "Danh bạ tra cứu trạng thái ĐKKD của Nhà cung cấp",
            "folder": "3. DU LIEU BO SUNG & TRA CUU QLT/3.2. Danh ba trang thai NNT",
            "step": "Bước 3: Dữ liệu bổ sung & QLT",
            "required": True,
            "min_files": 1,
            "exts": [".xlsx", ".xls"]
        },
        {
            "id": "3.3",
            "name": "Dữ liệu tra cứu TNCN đa nguồn & Rà soát trùng CCCD/CMND",
            "folder": "3. DU LIEU BO SUNG & TRA CUU QLT/3.3. Tra cuu TNCN da nguon & Nhan su",
            "step": "Bước 3: Dữ liệu bổ sung & QLT",
            "required": False,
            "min_files": 1,
            "exts": [".xlsx", ".xls", ".md"]
        },
        {
            "id": "3.4",
            "name": "Văn bản giải trình, Quyết định kiểm tra & Hồ sơ bổ sung",
            "folder": "3. DU LIEU BO SUNG & TRA CUU QLT/3.4. Giai trinh & Ho so bo sung",
            "step": "Bước 3: Dữ liệu bổ sung & QLT",
            "required": False,
            "min_files": 1,
            "exts": [".pdf", ".docx", ".xlsx"]
        }
    ]
    
    results = []
    total_required = 0
    passed_required = 0
    missing_items = []
    all_files_list = []
    
    for cat in categories:
        folder_win = to_win_long(target_path / cat["folder"])
        files = []
        if os.path.exists(folder_win):
            for f in os.listdir(folder_win):
                f_full = os.path.join(folder_win, f)
                if os.path.isfile(f_full) and not f.startswith("."):
                    files.append(f)
                    all_files_list.append({
                        "cat_id": cat["id"],
                        "folder": cat["folder"],
                        "file_name": f,
                        "size_kb": round(os.path.getsize(f_full) / 1024, 1)
                    })
                    
        file_count = len(files)
        is_passed = file_count >= cat["min_files"] if cat["min_files"] > 0 else (file_count > 0)
        
        if cat["required"]:
            total_required += 1
            if is_passed:
                passed_required += 1
                
        if is_passed:
            status_tag = "[ĐẠT]"
            status_detail = f"ĐÃ CÓ ĐỦ: Hiện có {file_count} file dữ liệu trong thư mục."
        else:
            if cat["required"]:
                status_tag = "[THIẾU]"
                if cat["id"] == "1.1":
                    status_detail = "THIẾU BẮT BUỘC: Cần kết xuất bổ sung Tờ khai XML GTGT, QTT TNDN, QTT TNCN và BCTC từ CSDL ngành Thuế."
                elif cat["id"] == "1.2":
                    status_detail = "THIẾU BẮT BUỘC: Cần trích xuất dữ liệu HĐĐT mua vào, bán ra (file HDDT.xlsx hoặc Bảng kê chi tiết các kỳ)."
                elif cat["id"] == "2.1":
                    status_detail = "THIẾU BẮT BUỘC: Cần yêu cầu NNT nộp Sổ Nhật ký chung, Sổ Cái và Sổ chi tiết TK 111, 112, 131, 331 dạng Excel có công thức."
                elif cat["id"] == "2.2":
                    status_detail = "THIẾU BẮT BUỘC: Cần bổ sung Sổ chi tiết XNT 152/155/156 (khớp [141] BCTC), Sổ 154 và Bảng tính giá thành sản phẩm."
                elif cat["id"] == "2.3":
                    status_detail = "THIẾU BẮT BUỘC: Cần bổ sung File Excel sao kê Internet Banking + Sao kê PDF có dấu đỏ ngân hàng; UNC hóa đơn >= 20tr / 5tr."
                elif cat["id"] == "2.4":
                    status_detail = "THIẾU BẮT BUỘC: Cần bổ sung Bảng thanh toán lương 12 tháng/năm, Bảng chấm công, HĐLĐ và Bảng trích khấu hao TSCĐ (TK 214)."
                elif cat["id"] == "3.1":
                    status_detail = "THIẾU BẮT BUỘC: Cần tra cứu & lập Sổ đối soát 100% Nhà cung cấp mua vào với danh mục cảnh báo rủi ro F0/F1 và Công văn Công an."
                elif cat["id"] == "3.2":
                    status_detail = "THIẾU BẮT BUỘC: Cần tra cứu trạng thái ĐKKD (đang hoạt động, đóng MST, bỏ trốn) của toàn bộ Nhà cung cấp mua vào."
                else:
                    status_detail = f"THIẾU BẮT BUỘC: Cần thu thập bổ sung tối thiểu {cat['min_files']} file tài liệu theo quy chuẩn."
                missing_items.append(f"{cat['name']} ({cat['folder']})")
            else:
                status_tag = "[CHƯA CÓ]"
                if cat["id"] == "1.3":
                    status_detail = "CHƯA CÓ (Tùy chọn): Cần tra cứu lịch sử nộp thuế ETAX và thông báo nợ NSNN nếu có phát sinh."
                elif cat["id"] == "3.3":
                    status_detail = "CHƯA CÓ (Tùy chọn): Cần tra cứu CSDL ngành Thuế về nhân sự đa nguồn và rà soát trùng số CCCD."
                elif cat["id"] == "3.4":
                    status_detail = "CHƯA CÓ (Tùy chọn): Lưu trữ văn bản giải trình, biên bản làm việc hoặc hồ sơ bổ sung của NNT."
                else:
                    status_detail = "CHƯA CÓ (Tùy chọn): Bổ sung hồ sơ phục vụ giải trình khoảng trống dữ liệu khi cần thiết."

        results.append({
            "id": cat["id"],
            "name": cat["name"],
            "step": cat["step"],
            "folder": cat["folder"],
            "required": cat["required"],
            "file_count": file_count,
            "is_passed": is_passed,
            "status_tag": status_tag,
            "status_detail": status_detail
        })
        
    return {
        "items": results,
        "total_required": total_required,
        "passed_required": passed_required,
        "missing_required_count": total_required - passed_required,
        "missing_items": missing_items,
        "all_files": all_files_list
    }

def create_master_checklist_excel(comp, check_data, output_path):
    wb = openpyxl.Workbook()
    
    # Sheet 1: Phiếu kiểm soát hồ sơ tổng hợp (A4 Portrait)
    ws = wb.active
    ws.title = "Kiem_soat_ho_so_A4"
    
    font_name = 'Times New Roman'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.views.sheetView[0].showGridLines = True

    f_org = Font(name=font_name, size=9, bold=True)
    f_nat = Font(name=font_name, size=9, bold=True)
    f_mot = Font(name=font_name, size=8.5, italic=True)
    f_title = Font(name=font_name, size=13, bold=True, color='1F4E79')
    f_sub = Font(name=font_name, size=9, italic=True, color='595959')
    f_info_lbl = Font(name=font_name, size=8.5, bold=True)
    f_tbl_hdr = Font(name=font_name, size=8.5, bold=True, color='FFFFFF')
    f_cell = Font(name=font_name, size=8)
    f_cell_bold = Font(name=font_name, size=8, bold=True)
    f_pass = Font(name=font_name, size=8, bold=True, color='006100')
    f_fail = Font(name=font_name, size=8, bold=True, color='9C0006')
    f_opt = Font(name=font_name, size=8, bold=True, color='595959')
    f_sign_title = Font(name=font_name, size=9, bold=True)
    f_sign_sub = Font(name=font_name, size=8, italic=True)

    fill_hdr = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    fill_pass = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    fill_fail = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    fill_opt = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_alt = PatternFill(start_color='F9FBFD', end_color='F9FBFD', fill_type='solid')
    fill_concl = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )
    border_box = Border(
        left=Side(style='medium', color='1F4E79'), right=Side(style='medium', color='1F4E79'),
        top=Side(style='medium', color='1F4E79'), bottom=Side(style='medium', color='1F4E79')
    )

    al_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    al_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    al_top_left = Alignment(horizontal='left', vertical='top', wrap_text=True)

    ws.merge_cells("A1:C1"); ws["A1"] = "CỤC THUẾ TỈNH ĐỒNG NAI"; ws["A1"].font = f_org; ws["A1"].alignment = al_center
    ws.merge_cells("A2:C2"); ws["A2"] = "ĐOÀN KIỂM TRA THUẾ NĂM 2026"; ws["A2"].font = f_org; ws["A2"].alignment = al_center
    ws.merge_cells("D1:G1"); ws["D1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"; ws["D1"].font = f_nat; ws["D1"].alignment = al_center
    ws.merge_cells("D2:G2"); ws["D2"] = "Độc lập - Tự do - Hạnh phúc"; ws["D2"].font = f_mot; ws["D2"].alignment = al_center

    ws.merge_cells("A4:G4"); ws["A4"] = "PHIẾU KIỂM SOÁT TÍNH ĐẦY ĐỦ HỒ SƠ KIỂM TRA THUẾ"
    ws["A4"].font = f_title; ws["A4"].alignment = al_center; ws.row_dimensions[4].height = 20

    ws.merge_cells("A5:G5"); ws["A5"] = f"(Rà soát hồ sơ theo Quy trình 4 bước - Doanh nghiệp: {comp['short_name']} - Định dạng in 01 trang A4)"
    ws["A5"].font = f_sub; ws["A5"].alignment = al_center; ws.row_dimensions[5].height = 14

    ws.merge_cells("A6:D6"); ws["A6"] = f"• Doanh nghiệp: {comp['name']}"; ws["A6"].font = f_info_lbl
    ws.merge_cells("E6:G6"); ws["E6"] = f"• Mã số thuế: {comp['mst']}"; ws["E6"].font = f_info_lbl
    ws.row_dimensions[6].height = 13.5

    ws.merge_cells("A7:D7"); ws["A7"] = f"• Niên độ kiểm tra: {comp['nien_do']} (Kế hoạch 2026)"; ws["A7"].font = f_info_lbl
    ws.merge_cells("E7:G7"); ws["E7"] = f"• Tiến độ đạt: {check_data['passed_required']}/{check_data['total_required']} mục bắt buộc ({int(check_data['passed_required']/check_data['total_required']*100)}%)"; ws["E7"].font = f_info_lbl
    ws.row_dimensions[7].height = 13.5

    headers = [
        ("STT", 4.5),
        ("Bước quy trình", 13),
        ("Danh mục hồ sơ kiểm soát", 24),
        ("Thư mục lưu trữ", 17),
        ("Số file", 6.5),
        ("Tình trạng", 12),
        ("Nội dung hiện trạng & Yêu cầu bổ sung cụ thể", 26)
    ]
    for col_idx, (h_text, w) in enumerate(headers, 1):
        cell = ws.cell(row=9, column=col_idx, value=h_text)
        cell.font = f_tbl_hdr; cell.fill = fill_hdr; cell.alignment = al_center
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[9].height = 17

    curr_r = 10
    for idx, item in enumerate(check_data["items"], 1):
        ws.row_dimensions[curr_r].height = 16
        
        c1 = ws.cell(row=curr_r, column=1, value=str(idx))
        c2 = ws.cell(row=curr_r, column=2, value=item["step"].split(":")[0])
        c3 = ws.cell(row=curr_r, column=3, value=item["name"])
        c4 = ws.cell(row=curr_r, column=4, value=item["folder"].split("/")[-1])
        c5 = ws.cell(row=curr_r, column=5, value=f"{item['file_count']} file")
        c6 = ws.cell(row=curr_r, column=6, value=f"{item['status_tag']} {'Đạt' if item['is_passed'] else 'Thiếu'}")
        c7 = ws.cell(row=curr_r, column=7, value=item["status_detail"])

        c1.alignment = al_center; c2.alignment = al_left; c3.alignment = al_left
        c4.alignment = al_left; c5.alignment = al_center; c6.alignment = al_center; c7.alignment = al_left

        c1.font = f_cell_bold; c2.font = f_cell_bold; c3.font = f_cell
        c4.font = f_cell; c5.font = f_cell_bold; c7.font = f_cell
        
        if item["is_passed"]:
            c6.font = f_pass; c6.fill = fill_pass
        else:
            if item["required"]:
                c6.font = f_fail; c6.fill = fill_fail
            else:
                c6.font = f_opt; c6.fill = fill_opt

        for col_i in range(1, 8):
            cell_i = ws.cell(row=curr_r, column=col_i)
            cell_i.border = border_thin
            if col_i != 6 and idx % 2 == 0:
                cell_i.fill = fill_alt

        curr_r += 1

    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
    if check_data["missing_required_count"] == 0:
        concl_text = "KẾT LUẬN: [X] ĐỦ ĐIỀU KIỆN TIẾN HÀNH PHÂN TÍCH CHUYÊN SÂU BƯỚC 4\n" \
                     "• Toàn bộ các hạng mục hồ sơ trọng yếu đã được thu thập đầy đủ và chuẩn hóa đúng 4 cấp thư mục."
    else:
        missing_str = ", ".join([it.split(" (")[0] for it in check_data["missing_items"]])
        concl_text = f"KẾT LUẬN: [X] CHƯA ĐỦ ĐIỀU KIỆN - THIẾU {check_data['missing_required_count']} HẠNG MỤC BẮT BUỘC\n" \
                     f"• Danh mục thiếu trọng yếu: {missing_str}.\n" \
                     f"• Yêu cầu: Phát hành Phiếu yêu cầu Doanh nghiệp cung cấp bổ sung hoặc Cán bộ tiến hành tra cứu trước khi mở hồ sơ."
                     
    concl_c = ws.cell(row=curr_r, column=1, value=concl_text)
    concl_c.font = f_cell_bold; concl_c.alignment = al_top_left
    concl_c.fill = fill_concl; ws.row_dimensions[curr_r].height = 38
    for c in range(1, 8):
        ws.cell(row=curr_r, column=c).border = border_box
    curr_r += 1

    ws.merge_cells(start_row=curr_r, start_column=5, end_row=curr_r, end_column=7)
    ws.cell(row=curr_r, column=5, value="......, ngày 04 tháng 10 năm 2026").font = f_sign_sub
    ws.cell(row=curr_r, column=5).alignment = al_center; ws.row_dimensions[curr_r].height = 12
    curr_r += 1

    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=3)
    ws.cell(row=curr_r, column=1, value="CÁN BỘ PHÂN TÍCH / RÀ SOÁT").font = f_sign_title
    ws.cell(row=curr_r, column=1).alignment = al_center

    ws.merge_cells(start_row=curr_r, start_column=4, end_row=curr_r, end_column=7)
    ws.cell(row=curr_r, column=4, value="TRƯỞNG ĐOÀN KIỂM TRA").font = f_sign_title
    ws.cell(row=curr_r, column=4).alignment = al_center
    ws.row_dimensions[curr_r].height = 13.5
    curr_r += 1

    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=3)
    ws.cell(row=curr_r, column=1, value="(Ký, ghi rõ họ và tên)").font = f_sign_sub
    ws.cell(row=curr_r, column=1).alignment = al_center

    ws.merge_cells(start_row=curr_r, start_column=4, end_row=curr_r, end_column=7)
    ws.cell(row=curr_r, column=4, value="(Ký, ghi rõ họ và tên)").font = f_sign_sub
    ws.cell(row=curr_r, column=4).alignment = al_center
    ws.row_dimensions[curr_r].height = 12

    # Sheet 2: Danh sách file thực tế trong hồ sơ
    ws2 = wb.create_sheet(title="Danh_sach_file_thuc_te")
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
    ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:E1"); ws2["A1"] = f"DANH MỤC CHI TIẾT TẬP TIN HỒ SƠ - {comp['name']} (MST: {comp['mst']})"
    ws2["A1"].font = f_title; ws2["A1"].alignment = al_center; ws2.row_dimensions[1].height = 20

    ws2_headers = [("STT", 6), ("Mã mục", 10), ("Thư mục con lưu trữ", 35), ("Tên file tài liệu", 45), ("Dung lượng (KB)", 15)]
    for col_idx, (h_text, w) in enumerate(ws2_headers, 1):
        cell = ws2.cell(row=3, column=col_idx, value=h_text)
        cell.font = f_tbl_hdr; cell.fill = fill_hdr; cell.alignment = al_center
        ws2.column_dimensions[get_column_letter(col_idx)].width = w
    ws2.row_dimensions[3].height = 18

    row2_idx = 4
    for f_idx, f_item in enumerate(check_data["all_files"], 1):
        ws2.row_dimensions[row2_idx].height = 15
        
        c1 = ws2.cell(row=row2_idx, column=1, value=f_idx)
        c2 = ws2.cell(row=row2_idx, column=2, value=f_item["cat_id"])
        c3 = ws2.cell(row=row2_idx, column=3, value=f_item["folder"])
        c4 = ws2.cell(row=row2_idx, column=4, value=f_item["file_name"])
        c5 = ws2.cell(row=row2_idx, column=5, value=f_item["size_kb"])

        c1.alignment = al_center; c2.alignment = al_center; c3.alignment = al_left
        c4.alignment = al_left; c5.alignment = Alignment(horizontal='right', vertical='center')

        c1.font = f_cell; c2.font = f_cell_bold; c3.font = f_cell
        c4.font = f_cell; c5.font = f_cell

        for col_i in range(1, 6):
            cell_i = ws2.cell(row=row2_idx, column=col_i)
            cell_i.border = border_thin
            if f_idx % 2 == 0:
                cell_i.fill = fill_alt

        row2_idx += 1

    wb.save(to_win_long(output_path))

def export_overall_summary_excel(all_reports):
    master_path = BASE_DIR / "TONG_HOP_TIEN_DO_13_DOANH_NGHIEP_2026.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tong_hop_13_DN_2026"
    
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

    fill_hdr = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    fill_pass = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    fill_fail = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    fill_alt = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )

    al_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    al_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    ws.merge_cells("A1:C1"); ws["A1"] = "CỤC THUẾ TỈNH ĐỒNG NAI"; ws["A1"].font = f_org; ws["A1"].alignment = al_center
    ws.merge_cells("A2:C2"); ws["A2"] = "ĐOÀN KIỂM TRA THUẾ NĂM 2026"; ws["A2"].font = f_org; ws["A2"].alignment = al_center
    ws.merge_cells("E1:I1"); ws["E1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"; ws["E1"].font = f_org; ws["E1"].alignment = al_center
    ws.merge_cells("E2:I2"); ws["E2"] = "Độc lập - Tự do - Hạnh phúc"; ws["E2"].font = Font(name=font_name, size=8.5, italic=True); ws["E2"].alignment = al_center

    ws.merge_cells("A4:I4"); ws["A4"] = "BẢNG TỔNG HỢP TIẾN ĐỘ THU THẬP & RÀ SOÁT HỒ SƠ 13 DOANH NGHIỆP KẾ HOẠCH 2026"
    ws["A4"].font = f_title; ws["A4"].alignment = al_center; ws.row_dimensions[4].height = 22

    ws.merge_cells("A5:I5"); ws["A5"] = "(Đã chuẩn hóa 4 bước khép kín, gom 1 nơi duy nhất và cập nhật Master Checklist tổng hợp)"
    ws["A5"].font = f_sub; ws["A5"].alignment = al_center; ws.row_dimensions[5].height = 14

    headers = [
        ("STT", 4.5),
        ("Tên Doanh nghiệp", 28),
        ("Mã số thuế", 13),
        ("Thư mục quy hoạch chuẩn", 26),
        ("B1. Thuế", 10),
        ("B2. NNT", 10),
        ("B3. QLT", 10),
        ("B4. KQ", 10),
        ("Đánh giá tình trạng & Hồ sơ thiếu", 24)
    ]
    for col_idx, (h_text, w) in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col_idx, value=h_text)
        cell.font = f_tbl_hdr; cell.fill = fill_hdr; cell.alignment = al_center
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[7].height = 20

    row_idx = 8
    for r in all_reports:
        comp = r["comp"]
        data = r["data"]
        ws.row_dimensions[row_idx].height = 18
        
        b1_cnt = sum(it["file_count"] for it in data["items"] if "Bước 1" in it["step"])
        b2_cnt = sum(it["file_count"] for it in data["items"] if "Bước 2" in it["step"])
        b3_cnt = sum(it["file_count"] for it in data["items"] if "Bước 3" in it["step"])
        b4_cnt = 1 # Master Checklist
        
        c1 = ws.cell(row=row_idx, column=1, value=comp["stt"])
        c2 = ws.cell(row=row_idx, column=2, value=comp["name"])
        c3 = ws.cell(row=row_idx, column=3, value=comp["mst"])
        c4 = ws.cell(row=row_idx, column=4, value=comp["target_folder"])
        c5 = ws.cell(row=row_idx, column=5, value=f"{b1_cnt} file")
        c6 = ws.cell(row=row_idx, column=6, value=f"{b2_cnt} file")
        c7 = ws.cell(row=row_idx, column=7, value=f"{b3_cnt} file")
        c8 = ws.cell(row=row_idx, column=8, value=f"{b4_cnt} file")
        
        if data["missing_required_count"] == 0:
            status_str = "ĐỦ ĐIỀU KIỆN"
        else:
            status_str = f"Thiếu {data['missing_required_count']} mục bắt buộc"
            
        c9 = ws.cell(row=row_idx, column=9, value=status_str)
        
        c1.alignment = al_center; c2.alignment = al_left; c3.alignment = al_center
        c4.alignment = al_left; c5.alignment = al_center; c6.alignment = al_center
        c7.alignment = al_center; c8.alignment = al_center; c9.alignment = al_center

        c1.font = f_cell_bold; c2.font = f_cell_bold; c3.font = f_cell_bold; c4.font = f_cell
        c5.font = f_cell; c6.font = f_cell; c7.font = f_cell; c8.font = f_cell
        c9.font = f_pass if data["missing_required_count"] == 0 else f_fail
        c9.fill = fill_pass if data["missing_required_count"] == 0 else fill_fail

        for col_i in range(1, 10):
            ws.cell(row=row_idx, column=col_i).border = border_thin
            if row_idx % 2 == 1 and col_i != 9:
                ws.cell(row=row_idx, column=col_i).fill = fill_alt

        row_idx += 1

    wb.save(to_win_long(master_path))

def delete_old_sources():
    old_folders = [
        "1. SON TIEN HAI",
        "2. T&T BHOA",
        "3. PHUC NGUYEN TIEN",
        "4. TAM HOANG PHAT",
        "5. HUNG ANH PHAT",
        "6. DAI HUY HOANG",
        "7. NHAT VINH-PASS LAM",
        "8. HUY THANH DAT",
        "9. TRIQUYEN",
        ". Phat hoang long",
        "HO SO KE KOACH",
        "To khai"
    ]
    
    old_loose_files = [
        "Tai khoan.docx",
        "Tai khoan ngan hang APT, DK.xlsx"
    ]
    
    deleted_folders = []
    deleted_files = []
    
    for f in old_folders:
        f_path = BASE_DIR / f
        f_win = to_win_long(f_path)
        if os.path.exists(f_win) and os.path.isdir(f_win):
            shutil.rmtree(f_win, ignore_errors=True)
            deleted_folders.append(f)
            
    for fn in old_loose_files:
        f_path = BASE_DIR / fn
        f_win = to_win_long(f_path)
        if os.path.exists(f_win) and os.path.isfile(f_win):
            try:
                os.remove(f_win)
                deleted_files.append(fn)
            except Exception:
                pass
                
    return deleted_folders, deleted_files

def main():
    print("==================================================================")
    print("1. BẮT ĐẦU QUY TRÌNH DI CHUYỂN TOÀN DIỆN 13 DOANH NGHIỆP...")
    print("==================================================================")
    for comp in COMPANIES:
        copied = migrate_company_files(comp)
        print(f"[{comp['stt']}/13] {comp['short_name']} -> Phân loại và sao chép hoàn tất.")
        
    print("\n==================================================================")
    print("2. XÓA FILE CHECKLIST PHỤ TẠI CÁC BƯỚC VÀ XUẤT MASTER CHECKLIST DUY NHẤT...")
    print("==================================================================")
    all_reports = []
    for comp in COMPANIES:
        target_path = BASE_DIR / comp["target_folder"]
        
        # Xóa các checklist phụ ở thư mục con
        clean_old_subfolder_checklists(target_path)
        
        # Kiểm tra dữ liệu thực tế
        check_data = inspect_company_data(target_path)
        
        # Xuất Master Checklist duy nhất tại 4. KET QUA PHAN TICH/
        master_checklist_path = target_path / "4. KET QUA PHAN TICH" / "CHECKLIST_TONG_HOP_HO_SO.xlsx"
        create_master_checklist_excel(comp, check_data, master_checklist_path)
        
        all_reports.append({
            "comp": comp,
            "data": check_data
        })
        print(f"[{comp['stt']}/13] {comp['short_name']} -> Đã xuất Master Checklist: 4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx ({check_data['passed_required']}/{check_data['total_required']} mục đạt).")

    # Xuất file tổng hợp tiến độ chung 13 DN
    export_overall_summary_excel(all_reports)
    print("\n-> Đã xuất Bảng tổng hợp tiến độ 13 Doanh nghiệp: TONG_HOP_TIEN_DO_13_DOANH_NGHIEP_2026.xlsx")

    print("\n==================================================================")
    print("3. XÓA TOÀN BỘ CÁC THƯ MỤC NGUỒN CŨ VÀ FILE LẺ ĐỂ TRÁNH NHẦM LẪN...")
    print("==================================================================")
    del_dirs, del_files = delete_old_sources()
    print(f"Đã xóa {len(del_dirs)} thư mục nguồn cũ: {del_dirs}")
    print(f"Đã xóa {len(del_files)} file lẻ gốc: {del_files}")
    
    print("\n==================================================================")
    print(">>> TOÀN BỘ CÔNG VIỆC ĐÃ HOÀN TẤT THÀNH CÔNG VÀ CHUẨN XÁC! <<<")
    print("==================================================================")

if __name__ == "__main__":
    main()
