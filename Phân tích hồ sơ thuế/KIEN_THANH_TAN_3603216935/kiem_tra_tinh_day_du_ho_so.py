#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
KỊCH BẢN TỰ ĐỘNG QUÉT VÀ KIỂM TRA TÍNH ĐẦY ĐỦ CỦA HỒ SƠ KIỂM TRA THUẾ
Chuẩn hóa theo Quy trình 4 bước và Bộ thư mục mẫu.
Hỗ trợ xuất Báo cáo Markdown và xuất File Excel Checklist 1 trang A4.
"""

import os
import sys
import argparse
from pathlib import Path

# Định nghĩa danh mục hồ sơ chuẩn theo 4 bước
CHECKLIST_DEFINITIONS = {
    "BUOC_1": {
        "ten": "Bước 1: Dữ liệu thô ngành Thuế",
        "thu_muc": "1. DU LIEU THO NGANH THUE",
        "muc_con": [
            {
                "id": "1.1_XML",
                "ten": "Tờ khai XML (GTGT, TNDN, TNCN, BCTC, KHBS)",
                "subfolder": "1.1. To khai XML",
                "extensions": [".xml"],
                "min_files": 4,
                "bat_buoc": True,
                "mo_ta": "Cần có tối thiểu các tờ khai chính: GTGT (842), QTT TNDN (892), QTT TNCN (953), BCTC (402)"
            },
            {
                "id": "1.2_HDDT_BANGKE",
                "ten": "Hóa đơn điện tử & Bảng kê mua vào/bán ra",
                "subfolder": "1.2. Hoa don dien tu & Bang ke",
                "extensions": [".xlsx", ".xls", ".csv"],
                "min_files": 2,
                "bat_buoc": True,
                "mo_ta": "Cần có file HDDT.xlsx và các file bảng kê mua vào, bán ra"
            },
            {
                "id": "1.3_NGHIA_VU",
                "ten": "Lịch sử nộp thuế & Thông báo nghĩa vụ NSNN",
                "subfolder": "1.3. Nghia vu nop thue & Thong bao",
                "extensions": [".xlsx", ".xls", ".pdf"],
                "min_files": 1,
                "bat_buoc": False,
                "mo_ta": "Dữ liệu nộp thuế vào NSNN, thông báo nợ thuế"
            }
        ]
    },
    "BUOC_2": {
        "ten": "Bước 2: Hồ sơ do Người nộp thuế cung cấp",
        "thu_muc": "2. HO SO NNT CUNG CAP",
        "muc_con": [
            {
                "id": "2.1_SO_KE_TOAN",
                "ten": "Sổ kế toán tổng hợp & Sổ chi tiết (NKC, Sổ Cái, 111, 112, 131, 331)",
                "subfolder": "2.1. So ke toan tong hop va chi tiet",
                "extensions": [".xlsx", ".xls"],
                "min_files": 2,
                "bat_buoc": True,
                "mo_ta": "Sổ Nhật ký chung, Sổ Cái, Sổ chi tiết TK 1111, 1121, 131, 331 dạng Excel"
            },
            {
                "id": "2.2_KHO_GIA_THANH",
                "ten": "Sổ XNT Kho, Sổ Chi phí 154, Bảng tính Giá thành & Định mức",
                "subfolder": "2.2. Kho - Chi phi - Gia thanh",
                "extensions": [".xlsx", ".xls", ".pdf"],
                "min_files": 2,
                "bat_buoc": True,
                "mo_ta": "Sổ XNT 152/155/156 (khớp BCTC [141]), Sổ 154/621/622/627, Bảng giá thành, Định mức kỹ thuật"
            },
            {
                "id": "2.3_NGAN_HANG",
                "ten": "Sao kê Ngân hàng (Excel & PDF) & Chứng từ UNC thanh toán",
                "subfolder": "2.3. Ngan hang & Dong tien",
                "extensions": [".xlsx", ".xls", ".csv", ".pdf"],
                "min_files": 1,
                "bat_buoc": True,
                "mo_ta": "File Excel sao kê Internet Banking, PDF sao kê có dấu, UNC hóa đơn >= 20tr / 5tr"
            },
            {
                "id": "2.4_LAO_DONG_TSCD",
                "ten": "Bảng lương, Bảng chấm công, HĐLĐ, Khấu hao TSCĐ (214), Phân bổ (242)",
                "subfolder": "2.4. Lao dong - Tien luong - TSCD",
                "extensions": [".xlsx", ".xls", ".pdf"],
                "min_files": 1,
                "bat_buoc": True,
                "mo_ta": "Bảng thanh toán lương 12 tháng, Bảng chấm công, HĐLĐ, Bảng trích khấu hao 214"
            }
        ]
    },
    "BUOC_3": {
        "ten": "Bước 3: Dữ liệu bổ sung & Tra cứu Quản lý thuế",
        "thu_muc": "3. DU LIEU BO SUNG & TRA CUU QLT",
        "muc_con": [
            {
                "id": "3.1_CBRR_CONGAN",
                "ten": "Sổ theo dõi CVCBRR F0/F1 & Sổ theo dõi Công an",
                "subfolder": "3.1. Canh bao rui ro CVCBRR & Cong an",
                "extensions": [".xlsx", ".xls"],
                "min_files": 1,
                "bat_buoc": True,
                "mo_ta": "Sổ F0 người bán rủi ro cao, Sổ F1 người mua, Sổ Công an theo dõi hóa đơn"
            },
            {
                "id": "3.2_DANH_BA_DKKD",
                "ten": "Danh bạ trạng thái ĐKKD của Nhà cung cấp",
                "subfolder": "3.2. Danh ba trang thai NNT",
                "extensions": [".xlsx", ".xls"],
                "min_files": 1,
                "bat_buoc": True,
                "mo_ta": "Dữ liệu tra cứu trạng thái ĐKKD (đang hoạt động, đóng MST, bỏ trốn...) của 100% NCC"
            },
            {
                "id": "3.3_TNCN_DA_NGUON",
                "ten": "Dữ liệu tra cứu TNCN đa nguồn & Nhân sự trùng CCCD",
                "subfolder": "3.3. Tra cuu TNCN da nguon & Nhan su",
                "extensions": [".xlsx", ".xls", ".md"],
                "min_files": 1,
                "bat_buoc": False,
                "mo_ta": "Kết quả tra cứu CSDL QLT ngành Thuế xem lao động có thu nhập nơi khác không"
            },
            {
                "id": "3.4_GIAI_TRINH",
                "ten": "Phiếu yêu cầu, Công văn giải trình & Hồ sơ bổ sung",
                "subfolder": "3.4. Giai trinh & Ho so bo sung",
                "extensions": [".pdf", ".xlsx", ".docx"],
                "min_files": 0,
                "bat_buoc": False,
                "mo_ta": "Văn bản giải trình của NNT và chứng từ bổ sung sau rà soát khoảng trống"
            }
        ]
    },
    "BUOC_4": {
        "ten": "Bước 4: Kết quả phân tích & Báo cáo đầu ra",
        "thu_muc": "4. KET QUA PHAN TICH",
        "muc_con": [
            {
                "id": "4.1_MASTER_EXCEL",
                "ten": "Master Excel tổng hợp chung 7 Sheet (Phan_tich_ho_so_cong_ty_<TenDN>.xlsx)",
                "subfolder": "",
                "extensions": [".xlsx"],
                "min_files": 1,
                "bat_buoc": True,
                "mo_ta": "File Excel duy nhất tích hợp đầy đủ 7 Sheet phân tích chuyên sâu"
            },
            {
                "id": "4.2_CAU_THANH",
                "ten": "Thư mục cấu thành (chứa các file phân tích chuyên đề độc lập)",
                "subfolder": "cau-thanh",
                "extensions": [".xlsx"],
                "min_files": 1,
                "bat_buoc": False,
                "mo_ta": "Chứa bảng kê gộp, đối chiếu 3 chiều, xét NCC F0, TNCN đa nguồn, XNT, Giá thành, Tiền mặt"
            }
        ]
    }
}

def scan_directory(target_dir):
    target_path = Path(target_dir).resolve()
    if not target_path.exists():
        print(f"Lỗi: Thư mục không tồn tại: {target_path}")
        return None

    report = {
        "target_dir": str(target_path),
        "steps": {},
        "tong_so_muc": 0,
        "so_muc_dat": 0,
        "so_muc_thieu_bat_buoc": 0,
        "thieu_trong_yeu": []
    }

    for buoc_key, buoc_info in CHECKLIST_DEFINITIONS.items():
        step_res = {
            "ten": buoc_info["ten"],
            "thu_muc_chinh": buoc_info["thu_muc"],
            "muc_con": []
        }
        
        buoc_base_path = target_path / buoc_info["thu_muc"]
        
        for mc in buoc_info["muc_con"]:
            report["tong_so_muc"] += 1
            if mc["subfolder"]:
                check_path = buoc_base_path / mc["subfolder"]
            else:
                check_path = buoc_base_path

            found_files = []
            if check_path.exists():
                for f in check_path.iterdir():
                    if f.is_file() and not f.name.startswith("README") and not f.name.startswith(".") and not "CHECKLIST" in f.name:
                        if any(f.suffix.lower() == ext.lower() for ext in mc["extensions"]):
                            found_files.append(f.name)

            so_luong = len(found_files)
            dat_yeu_cau = so_luong >= mc["min_files"] if mc["min_files"] > 0 else True
            
            # Cảnh báo nếu bắt buộc nhưng không có file
            if mc["bat_buoc"] and not dat_yeu_cau:
                report["so_muc_thieu_bat_buoc"] += 1
                report["thieu_trong_yeu"].append({
                    "buoc": buoc_info["ten"],
                    "ten_muc": mc["ten"],
                    "thu_muc": str(check_path.relative_to(target_path) if check_path.is_relative_to(target_path) else check_path),
                    "mo_ta": mc["mo_ta"]
                })
            elif dat_yeu_cau:
                report["so_muc_dat"] += 1

            step_res["muc_con"].append({
                "id": mc["id"],
                "ten": mc["ten"],
                "bat_buoc": mc["bat_buoc"],
                "subfolder": mc["subfolder"],
                "so_luong": so_luong,
                "dat_yeu_cau": dat_yeu_cau,
                "files": found_files,
                "mo_ta": mc["mo_ta"]
            })

        report["steps"][buoc_key] = step_res

    return report

def format_markdown_report(report):
    if not report:
        return "Không có dữ liệu báo cáo."

    lines = []
    lines.append(f"# BÁO CÁO RÀ SOÁT TÍNH ĐẦY ĐỦ HỒ SƠ KIỂM TRA THUẾ")
    lines.append(f"- **Thư mục kiểm tra:** `{report['target_dir']}`")
    lines.append(f"- **Tổng số hạng mục kiểm soát:** {report['tong_so_muc']} mục")
    lines.append(f"- **Số mục đạt yêu cầu:** **{report['so_muc_dat']} / {report['tong_so_muc']}**")
    lines.append(f"- **Số mục bắt buộc còn thiếu:** **{report['so_muc_thieu_bat_buoc']}** mục")
    lines.append("")

    if report["so_muc_thieu_bat_buoc"] == 0:
        lines.append("> [!NOTE]\n> **ĐÁNH GIÁ: ĐỦ ĐIỀU KIỆN.** Hồ sơ đã thu thập đầy đủ các hạng mục bắt buộc để tiến hành phân tích chuyên sâu Bước 4.")
    else:
        lines.append(f"> [!WARNING]\n> **ĐÁNH GIÁ: THIẾU {report['so_muc_thieu_bat_buoc']} HẠNG MỤC BẮT BUỘC.** Cần bổ sung tài liệu trước khi phân tích chuyên sâu.")

    lines.append("\n---")

    for buoc_key, step_data in report["steps"].items():
        lines.append(f"\n### {step_data['ten']}")
        lines.append(f"*(Thư mục chính: `{step_data['thu_muc_chinh']}`)*\n")
        lines.append("| Tình trạng | Hạng mục hồ sơ | Thư mục con | Số file hiện có | Tính chất | Yêu cầu kiểm tra |")
        lines.append("|:---:|:---|:---|:---:|:---:|:---|")

        for mc in step_data["muc_con"]:
            status_icon = "✅ Đạt" if mc["dat_yeu_cau"] else ("❌ **Thiếu**" if mc["bat_buoc"] else "⚠️ Chưa có")
            tinh_chat = "**Bắt buộc**" if mc["bat_buoc"] else "Tùy chọn"
            folder_str = f"`{mc['subfolder']}`" if mc['subfolder'] else "`(Thư mục gốc bước)`"
            lines.append(f"| {status_icon} | {mc['ten']} | {folder_str} | **{mc['so_luong']}** file | {tinh_chat} | {mc['mo_ta']} |")

    if report["thieu_trong_yeu"]:
        lines.append("\n---")
        lines.append("\n### 🚨 Danh mục tài liệu trọng yếu còn thiếu cần yêu cầu bổ sung:")
        for idx, item in enumerate(report["thieu_trong_yeu"], 1):
            lines.append(f"{idx}. **{item['ten_muc']}** (`{item['thu_muc']}`)")
            lines.append(f"   - *Yêu cầu:* {item['mo_ta']}")

    return "\n".join(lines)

def export_excel_report(report, output_path):
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("Lỗi: Không tìm thấy thư viện openpyxl. Vui lòng cài đặt bằng: pip install openpyxl")
        return False

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ket_qua_ra_soat_ho_so"
    
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

    # Styles
    f_org = Font(name=font_name, size=9.5, bold=True)
    f_title = Font(name=font_name, size=13, bold=True, color='1F4E79')
    f_sub = Font(name=font_name, size=9.5, italic=True, color='595959')
    f_info_lbl = Font(name=font_name, size=8.5, bold=True)
    f_tbl_hdr = Font(name=font_name, size=8.5, bold=True, color='FFFFFF')
    f_cell = Font(name=font_name, size=8)
    f_cell_bold = Font(name=font_name, size=8, bold=True)
    f_status_pass = Font(name=font_name, size=8, bold=True, color='006100')
    f_status_fail = Font(name=font_name, size=8, bold=True, color='9C0006')
    f_sign_title = Font(name=font_name, size=9, bold=True)
    f_sign_sub = Font(name=font_name, size=8, italic=True)

    fill_hdr = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
    fill_pass = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    fill_fail = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    fill_highlight = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    fill_alt = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')

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

    # Header block
    ws.merge_cells("A1:C1"); ws["A1"] = "CỤC THUẾ TỈNH / TP ...................."; ws["A1"].font = f_org; ws["A1"].alignment = al_center
    ws.merge_cells("A2:C2"); ws["A2"] = "CHI CỤC THUẾ / ĐOÀN KIỂM TRA"; ws["A2"].font = f_org; ws["A2"].alignment = al_center
    ws.merge_cells("D1:G1"); ws["D1"] = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"; ws["D1"].font = f_org; ws["D1"].alignment = al_center
    ws.merge_cells("D2:G2"); ws["D2"] = "Độc lập - Tự do - Hạnh phúc"; ws["D2"].font = f_sign_sub; ws["D2"].alignment = al_center

    ws.merge_cells("A4:G4"); ws["A4"] = "PHIẾU KẾT QUẢ RÀ SOÁT TÍNH ĐẦY ĐỦ HỒ SƠ KIỂM TRA THUẾ"
    ws["A4"].font = f_title; ws["A4"].alignment = al_center; ws.row_dimensions[4].height = 20

    ws.merge_cells("A5:G5"); ws["A5"] = f"(Thư mục quét: {Path(report['target_dir']).name} - Định dạng in 01 trang A4)"
    ws["A5"].font = f_sub; ws["A5"].alignment = al_center; ws.row_dimensions[5].height = 14

    # Enterprise Info
    target_name = Path(report['target_dir']).name
    ws.merge_cells("A6:D6"); ws["A6"] = f"• Thư mục kiểm tra: {target_name}"; ws["A6"].font = f_info_lbl
    ws.merge_cells("E6:G6"); ws["E6"] = "• Mã số thuế: ..........................................."; ws["E6"].font = f_info_lbl
    ws.row_dimensions[6].height = 14

    ws.merge_cells("A7:D7"); ws["A7"] = f"• Tiến độ rà soát: {report['so_muc_dat']} / {report['tong_so_muc']} mục đạt ({int(report['so_muc_dat']/report['tong_so_muc']*100)}%)"; ws["A7"].font = f_info_lbl
    ws.merge_cells("E7:G7"); ws["E7"] = f"• Số mục thiếu bắt buộc: {report['so_muc_thieu_bat_buoc']} mục"; ws["E7"].font = f_info_lbl
    ws.row_dimensions[7].height = 14

    # Table Header
    headers = [("STT", 4.5), ("Bước quy trình", 13), ("Hạng mục kiểm tra", 25), ("Thư mục", 16), ("File", 5.5), ("Trạng thái", 12), ("Yêu cầu & Cảnh báo rủi ro", 23)]
    for col_idx, (h_text, w) in enumerate(headers, 1):
        cell = ws.cell(row=9, column=col_idx, value=h_text)
        cell.font = f_tbl_hdr; cell.fill = fill_hdr; cell.alignment = al_center
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[9].height = 18

    curr_r = 10
    stt = 1
    for b_key, b_val in report["steps"].items():
        for mc in b_val["muc_con"]:
            ws.row_dimensions[curr_r].height = 15.5
            
            c1 = ws.cell(row=curr_r, column=1, value=str(stt))
            c2 = ws.cell(row=curr_r, column=2, value=b_val["ten"].split(":")[0])
            c3 = ws.cell(row=curr_r, column=3, value=mc["ten"])
            c4 = ws.cell(row=curr_r, column=4, value=mc["subfolder"] if mc["subfolder"] else "(gốc)")
            c5 = ws.cell(row=curr_r, column=5, value=mc["so_luong"])
            
            status_text = "[X] ĐẠT" if mc["dat_yeu_cau"] else ("[X] THIẾU" if mc["bat_buoc"] else "[X] Chưa có")
            c6 = ws.cell(row=curr_r, column=6, value=status_text)
            c7 = ws.cell(row=curr_r, column=7, value=mc["mo_ta"])

            c1.alignment = al_center; c2.alignment = al_left; c3.alignment = al_left
            c4.alignment = al_left; c5.alignment = al_center; c6.alignment = al_center; c7.alignment = al_left

            c1.font = f_cell_bold; c2.font = f_cell_bold; c3.font = f_cell
            c4.font = f_cell; c5.font = f_cell_bold
            c6.font = f_status_pass if mc["dat_yeu_cau"] else f_status_fail
            c7.font = f_cell

            fill = fill_pass if mc["dat_yeu_cau"] else (fill_fail if mc["bat_buoc"] else fill_alt)
            c6.fill = fill

            for col_i in range(1, 8):
                ws.cell(row=curr_r, column=col_i).border = border_thin

            curr_r += 1
            stt += 1

    # Conclusion Box
    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
    if report["so_muc_thieu_bat_buoc"] == 0:
        c_text = "KẾT LUẬN: [X] ĐỦ ĐIỀU KIỆN TIẾN HÀNH PHÂN TÍCH CHUYÊN SÂU BƯỚC 4\n" \
                 "• Toàn bộ các hạng mục hồ sơ trọng yếu đã được thu thập đầy đủ và sẵn sàng phân tích."
    else:
        c_text = f"KẾT LUẬN: [X] CHƯA ĐỦ ĐIỀU KIỆN - THIẾU {report['so_muc_thieu_bat_buoc']} HẠNG MỤC BẮT BUỘC\n" \
                 f"• Đề nghị phát hành Phiếu yêu cầu bổ sung các tài liệu trọng yếu còn thiếu trước khi mở hồ sơ kiểm tra."
    
    concl_c = ws.cell(row=curr_r, column=1, value=c_text)
    concl_c.font = f_cell_bold; concl_c.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    concl_c.fill = fill_highlight; ws.row_dimensions[curr_r].height = 36
    for c in range(1, 8):
        ws.cell(row=curr_r, column=c).border = border_box
    curr_r += 1

    # Signatures
    ws.merge_cells(start_row=curr_r, start_column=5, end_row=curr_r, end_column=7)
    ws.cell(row=curr_r, column=5, value="......, ngày ..... tháng ..... năm 2026").font = f_sign_sub
    ws.cell(row=curr_r, column=5).alignment = al_center; ws.row_dimensions[curr_r].height = 13
    curr_r += 1

    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=3)
    ws.cell(row=curr_r, column=1, value="CÁN BỘ PHÂN TÍCH / RÀ SOÁT").font = f_sign_title
    ws.cell(row=curr_r, column=1).alignment = al_center

    ws.merge_cells(start_row=curr_r, start_column=4, end_row=curr_r, end_column=7)
    ws.cell(row=curr_r, column=4, value="TRƯỞNG ĐOÀN KIỂM TRA").font = f_sign_title
    ws.cell(row=curr_r, column=4).alignment = al_center
    ws.row_dimensions[curr_r].height = 14
    curr_r += 1

    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=3)
    ws.cell(row=curr_r, column=1, value="(Ký, ghi rõ họ và tên)").font = f_sign_sub
    ws.cell(row=curr_r, column=1).alignment = al_center

    ws.merge_cells(start_row=curr_r, start_column=4, end_row=curr_r, end_column=7)
    ws.cell(row=curr_r, column=4, value="(Ký, ghi rõ họ và tên)").font = f_sign_sub
    ws.cell(row=curr_r, column=4).alignment = al_center
    ws.row_dimensions[curr_r].height = 13

    wb.save(output_path)
    print(f"\n-> Đã xuất Bảng kết quả rà soát ra file Excel (A4): {output_path}")
    return True

def main():
    parser = argparse.ArgumentParser(description="Quét và kiểm tra tính đầy đủ hồ sơ thuế theo 4 bước")
    parser.add_argument("dir", nargs="?", default=".", help="Đường dẫn thư mục doanh nghiệp cần kiểm tra (mặc định: thư mục hiện tại)")
    parser.add_argument("--output", "-o", help="Đường dẫn file markdown xuất báo cáo")
    parser.add_argument("--excel", "-e", help="Đường dẫn file Excel xuất kết quả rà soát (chuẩn A4 in ấn)")
    args = parser.parse_args()

    report = scan_directory(args.dir)
    if report:
        md_text = format_markdown_report(report)
        print(md_text)
        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(md_text)
            print(f"\nĐã xuất báo cáo ra file Markdown: {args.output}")
        if args.excel:
            export_excel_report(report, args.excel)

if __name__ == "__main__":
    main()
