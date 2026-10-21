"""
parse_september_cv.py
Trích xuất dữ liệu từ các công văn số hóa tháng 09/2026 tại D:\\HO SO CONG VAN CANH BAO\\2026\\2026-09
và cập nhật vào sổ theo dõi Excel SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.
"""

import os
import re
import csv
import json
import sys
from pathlib import Path

# Thêm đường dẫn scripts để import update_so_theo_doi
sys.path.insert(0, str(Path(__file__).parent))
from update_so_theo_doi import update_workbook, DEFAULT_EXCEL_PATH

TARGET_DIR = r"D:\HO SO CONG VAN CANH BAO\2026\2026-09"
CSV_PATH = r"D:\HO SO CONG VAN CANH BAO\DANH_SACH_CONG_VAN_CANH_BAO.csv"

def clean_company_name(s):
    if not s:
        return ""
    s = re.sub(r"^(?:v/v|vv|v/v\s+cảnh\s+báo|cảnh\s+báo|cvr|cbrr|cv|công\s+văn)[\s:.-]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:doanh\s+nghiệp\s+có\s+dấu\s+hiệu\s+rủi\s+ro\s+cao\s+về\s+thuế[^\w]*)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:đối\s+với\s+)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\s:;.-]+(?:mst|mã\s+số\s+thuế).*$", "", s, flags=re.IGNORECASE)
    return s.strip(" ,.:;-\"'\n\r\t")

def parse_all_digital_cvs():
    # 1. Đọc metadata từ CSV bảng kê
    csv_rows = {}
    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, "r", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                fn = row.get("Tên file lưu", "")
                if os.path.exists(os.path.join(TARGET_DIR, fn)):
                    m = re.match(r"^\[(.*?)\]", fn)
                    p = m.group(1) if m else fn
                    csv_rows.setdefault(p, []).append(row)

    # 2. Gom nhóm các tệp trong thư mục tháng 9/2026
    files = os.listdir(TARGET_DIR)
    groups = {}
    for f in files:
        m = re.match(r"^\[(.*?)\]\s*(.*)", f)
        if m:
            p = m.group(1)
            groups.setdefault(p, {"pdfs": [], "mds": []})
            if f.endswith(".pdf"):
                groups[p]["pdfs"].append(f)
            elif f.endswith(".md"):
                groups[p]["mds"].append(f)

    digital_groups = {k: v for k, v in groups.items() if len(v["mds"]) > 0}
    print(f"Tổng số nhóm công văn số hóa (digital): {len(digital_groups)}")

    records = []
    seen_msts = set()

    MAP_DON_VI = {
        "THUẾ CƠ SỞ 1": "Thuế cơ sở 1 thành phố Đồng Nai",
        "THUẾ CƠ SỞ 2": "Thuế cơ sở 2 thành phố Đồng Nai",
        "THUẾ CƠ SỞ 3": "Thuế cơ sở 3 thành phố Đồng Nai",
        "THUẾ CƠ SỞ 4": "Thuế cơ sở 4 thành phố Đồng Nai",
        "THUẾ CƠ SỞ 5": "Thuế cơ sở 5 thành phố Đồng Nai",
        "THUẾ CƠ SỞ 6": "Thuế cơ sở 6 thành phố Đồng Nai",
        "THUẾ CƠ SỞ 7": "Thuế cơ sở 7 thành phố Đồng Nai",
        "THUẾ CƠ SỞ 8": "Thuế cơ sở 8 thành phố Đồng Nai",
        "THUẾ CƠ SỞ 10": "Thuế cơ sở 10 thành phố Đồng Nai",
        "THUẾ THÀNH PHỐ ĐỒNG NAI": "Thuế thành phố Đồng Nai",
    }

    for p, data in digital_groups.items():
        meta = csv_rows.get(p, [{}])[0]
        socv = meta.get("Số ký hiệu", p.replace("_", "/"))
        ngay_cv = meta.get("Ngày ban hành", "")
        raw_don_vi = meta.get("Đơn vị ban hành", "")
        don_vi_bh = MAP_DON_VI.get(raw_don_vi.strip(), raw_don_vi.strip())
        trich_yeu = meta.get("Trích yếu nội dung", "")

        main_md_text = ""
        bang_ke_md_text = ""
        main_pdf = ""
        bang_ke_files = []

        for mdf in data["mds"]:
            pdf_candidate = mdf[:-3] + ".pdf"
            if any(x in mdf.lower() for x in ["bang ke", "bảng kê", "hdbr", "hdmv", "danh sách hóa đơn"]):
                bang_ke_files.append(pdf_candidate)
                try:
                    with open(os.path.join(TARGET_DIR, mdf), "r", encoding="utf-8", errors="ignore") as fp:
                        bang_ke_md_text += fp.read() + "\n"
                except Exception:
                    pass
            else:
                try:
                    with open(os.path.join(TARGET_DIR, mdf), "r", encoding="utf-8", errors="ignore") as fp:
                        main_md_text += fp.read() + "\n"
                except Exception:
                    pass
                main_pdf = os.path.join(TARGET_DIR, pdf_candidate)

        if not main_pdf and data["pdfs"]:
            main_pdf = os.path.join(TARGET_DIR, data["pdfs"][0])
        if not main_md_text and bang_ke_md_text:
            main_md_text = bang_ke_md_text

        # Bóc tách MST người nộp thuế
        mst_ban = ""
        mst_matches = re.findall(r"(?:mã số thuế|mst|mã số doanh nghiệp)[:\s]*([0-9]{10}(?:-[0-9]{3})?)", main_md_text, re.IGNORECASE)
        if mst_matches:
            mst_ban = mst_matches[0]
        else:
            m_mst = re.search(r"\b(36\d{8}|38\d{8}|03\d{8}|[0-9]{10})\b", trich_yeu + " " + " ".join(data["pdfs"]))
            if m_mst:
                mst_ban = m_mst.group(1)

        # Bóc tách tên người nộp thuế
        ten_ban = ""
        # Cách 1: Trường cụ thể trong text
        m_nnt = re.search(r"(?:Tên\s+người\s+nộp\s+thuế|Tên\s+doanh\s+nghiệp|Tên\s+đơn\s+vị|Người\s+nộp\s+thuế)[:\s]*([^\n\r,]{5,80})", main_md_text, re.IGNORECASE)
        if m_nnt:
            cand = clean_company_name(m_nnt.group(1))
            if len(cand) > 5 and not any(w in cand.lower() for w in ["dấu hiệu", "rủi ro", "không hoạt động", "hóa đơn"]):
                ten_ban = cand

        # Cách 2: Trích yếu trong ngoặc đơn
        if not ten_ban:
            m_ty = re.search(r"\((?:Cty|Công ty|Doanh nghiệp|DNTN|DNNN)\s+([^)]+)\)", trich_yeu, re.IGNORECASE)
            if m_ty:
                cand = m_ty.group(0).strip("()")
                if not re.match(r"^(?:Công ty|Cty|Doanh nghiệp|DNTN)", cand, re.IGNORECASE):
                    cand = "Công ty " + cand
                ten_ban = cand

        # Cách 3: Trích yếu không ngoặc đơn
        if not ten_ban:
            m_ty2 = re.search(r"(?:CÔNG TY TNHH|CÔNG TY CỔ PHẦN|CÔNG TY CP|CÔNG TY MTV|DOANH NGHIỆP TƯ NHÂN|DNTN|CTY TNHH|CTY CP|Cty\s+|Công ty\s+)[^,;:\n\r()]+", trich_yeu, re.IGNORECASE)
            if m_ty2:
                cand = clean_company_name(m_ty2.group(0))
                if len(cand) > 5:
                    ten_ban = cand

        # Cách 4: Quét các khối tên cty chuẩn trong text
        if not ten_ban:
            m_cty = re.findall(r"(?:CÔNG TY TNHH|CÔNG TY CỔ PHẦN|CÔNG TY CP|DOANH NGHIỆP TƯ NHÂN|DNTN)[^\n\r,.:;]{4,80}", main_md_text)
            for c in m_cty:
                if "THÀNH PHỐ" not in c and "ĐỒNG NAI" not in c and "VIỆT NAM" not in c and "HẠNH PHÚC" not in c:
                    ten_ban = c.strip()
                    break

        # Cách 5: Tên file
        if not ten_ban:
            for pdf in data["pdfs"]:
                m_fn = re.search(r"(?:CÔNG TY|CTY|DOANH NGHIỆP|DNTN)[^.]+", pdf, re.IGNORECASE)
                if m_fn:
                    ten_ban = clean_company_name(m_fn.group(0))
                    break

        if not ten_ban:
            ten_ban = "Doanh nghiệp theo CV " + socv

        # Bóc tách địa chỉ
        dia_chi = ""
        m_addr = re.search(r"(?:địa chỉ|trụ sở)[:\s]*([^\n\r]{10,150})", main_md_text, re.IGNORECASE)
        if m_addr:
            dia_chi = m_addr.group(1).strip()
            dia_chi = re.sub(r"^(?:đăng ký kinh doanh|trụ sở chính|trụ sở|đã đăng ký)[:\s]*", "", dia_chi, flags=re.IGNORECASE).strip()

        # Dấu hiệu vi phạm (NoiDung)
        reasons = []
        if re.search(r"không hoạt động tại địa chỉ", main_md_text, re.IGNORECASE):
            reasons.append("Không hoạt động tại địa chỉ đăng ký")
        if re.search(r"chưa nộp tờ khai|không nộp tờ khai", main_md_text, re.IGNORECASE):
            reasons.append("Chưa nộp tờ khai thuế")
        if re.search(r"chênh lệch.*hóa đơn|chênh lệch.*tờ khai", main_md_text, re.IGNORECASE):
            reasons.append("Chênh lệch số liệu tờ khai và HĐĐT")
        if re.search(r"cơ quan điều tra|chuyển tin báo", main_md_text, re.IGNORECASE):
            reasons.append("Chuyển tin báo CQĐT")
        if re.search(r"tạm ngừng kinh doanh", main_md_text, re.IGNORECASE):
            reasons.append("Tạm ngừng kinh doanh")
        if re.search(r"bãi bỏ", main_md_text + " " + trich_yeu, re.IGNORECASE):
            reasons.append("Bãi bỏ văn bản cảnh báo trước")
        if re.search(r"gian lận.*hoá đơn|bán lẻ.*hoá đơn", main_md_text, re.IGNORECASE):
            reasons.append("Nghi vấn gian lận hóa đơn")

        if reasons:
            noi_dung = "; ".join(reasons)
        else:
            noi_dung = trich_yeu if trich_yeu else "Cảnh báo doanh nghiệp có rủi ro cao về thuế, hóa đơn"

        # Ghi chú
        ghi_chu = f"Nguồn: CBRR-ben ban. Trích yếu: {trich_yeu}."
        if bang_ke_files:
            ghi_chu += f" Kèm tệp bảng kê: {', '.join(bang_ke_files)}."

        # Cờ đếm duy nhất
        mst_dau = 1 if mst_ban and (mst_ban not in seen_msts) else None
        if mst_ban:
            seen_msts.add(mst_ban)

        record = {
            "so_cv": socv,
            "ngay_cv": ngay_cv,
            "don_vi_bh": don_vi_bh,
            "ten_ban": ten_ban,
            "mst_ban": mst_ban,
            "dia_chi": dia_chi,
            "noi_dung": noi_dung,
            "nhom_nd": "Dấu hiệu rủi ro về thuế",
            "file": main_pdf,
            "nam_cv": 2026,
            "ghi_chu": ghi_chu,
            "cv_dau": 1,
            "mst_dau": mst_dau
        }
        records.append(record)

    # Sắp xếp theo ngày ký tăng dần hoặc theo số ký hiệu
    records.sort(key=lambda x: (x["ngay_cv"] or "", x["so_cv"]))
    return records

if __name__ == "__main__":
    records = parse_all_digital_cvs()
    print(f"Đã trích xuất hoàn tất {len(records)} công văn số hóa.")
    json_output = Path(__file__).parent / "records_2026_09.json"
    with open(json_output, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"Đã lưu kết quả trích xuất vào: {json_output}")

    # Cập nhật vào sổ Excel
    excel_file = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL_PATH
    print(f"Bắt đầu nạp dữ liệu vào sổ Excel: {excel_file}")
    new_max = update_workbook(records, excel_file)
    print(f"Hoàn thành xuất sắc! Dòng cuối cùng mới: {new_max}")
