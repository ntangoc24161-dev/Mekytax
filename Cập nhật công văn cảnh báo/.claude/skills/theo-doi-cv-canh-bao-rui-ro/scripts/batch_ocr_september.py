"""
batch_ocr_september.py
Thực hiện OCR offline hàng loạt cho 16 công văn (18 tệp PDF) scan tháng 09/2026,
bóc tách dữ liệu và nạp vào sổ theo dõi Excel SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx.
"""

import os
import re
import csv
import json
import sys
import shutil
from pathlib import Path

# Thêm đường dẫn javis-os để import doc2md backends
JAVIS_OS_DIR = r"D:\SECOND BRAIN\javis-os"
if JAVIS_OS_DIR not in sys.path:
    sys.path.insert(0, JAVIS_OS_DIR)

from tools.doc2md.backends import extract_paddle_ocr
from update_so_theo_doi import update_workbook, DEFAULT_EXCEL_PATH

DIR_SCAN = r"D:\HO SO CONG VAN CANH BAO\Cong_van_chua_boc_tach\2026-09"
DIR_ORIG = r"D:\HO SO CONG VAN CANH BAO\2026\2026-09"
CSV_PATH = r"D:\HO SO CONG VAN CANH BAO\DANH_SACH_CONG_VAN_CANH_BAO.csv"

def clean_company_name(s):
    if not s:
        return ""
    s = re.sub(r"^(?:v/v|vv|v/v\s+cảnh\s+báo|cảnh\s+báo|cvr|cbrr|cv|công\s+văn)[\s:.-]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:doanh\s+nghiệp\s+có\s+dấu\s+hiệu\s+rủi\s+ro\s+cao\s+về\s+thuế[^\w]*)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:đối\s+với\s+)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\s:;.-]+(?:mst|mã\s+số\s+thuế).*$", "", s, flags=re.IGNORECASE)
    return s.strip(" ,.:;-\"'\n\r\t")

def run_ocr_and_parse(target_excel=DEFAULT_EXCEL_PATH):
    print("=== BẮT ĐẦU QUY TRÌNH BÓC TÁCH 16 CÔNG VĂN SCAN THÁNG 09/2026 ===")
    
    # 1. Đọc metadata từ CSV bảng kê
    csv_rows = {}
    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, "r", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                fn = row.get("Tên file lưu", "")
                m = re.match(r"^\[(.*?)\]", fn)
                p = m.group(1) if m else fn
                csv_rows.setdefault(p, []).append(row)
                csv_rows[fn] = row

    # 2. Liệt kê các file PDF trong thư mục scan
    pdf_files = [f for f in sorted(os.listdir(DIR_SCAN)) if f.endswith(".pdf")]
    print(f"Tổng số tệp PDF scan cần xử lý: {len(pdf_files)} tệp")

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

    records = []
    seen_msts = set()

    for idx, pdf_name in enumerate(pdf_files, 1):
        pdf_path = os.path.join(DIR_SCAN, pdf_name)
        md_name = pdf_name[:-4] + ".md"
        md_path = os.path.join(DIR_SCAN, md_name)
        md_orig_path = os.path.join(DIR_ORIG, md_name)

        print(f"\n[{idx}/{len(pdf_files)}] Xử lý: {pdf_name}")

        # Kiểm tra xem file md đã có chưa
        md_text = ""
        if os.path.exists(md_path) and os.path.getsize(md_path) > 100:
            print(f" -> Đã có file Markdown: {md_path}")
            with open(md_path, "r", encoding="utf-8", errors="ignore") as fp:
                md_text = fp.read()
        elif os.path.exists(md_orig_path) and os.path.getsize(md_orig_path) > 100:
            print(f" -> Đã có file Markdown tại thư mục gốc: {md_orig_path}")
            with open(md_orig_path, "r", encoding="utf-8", errors="ignore") as fp:
                md_text = fp.read()
            shutil.copy2(md_orig_path, md_path)
        else:
            print(f" -> Đang chạy OCR offline cho {pdf_name}...")
            try:
                pages = extract_paddle_ocr(Path(pdf_path), lang="vi")
                md_chunks = [f"## Trang {p.page_number}\n\n{p.markdown}" for p in pages]
                md_text = "\n\n".join(md_chunks)
                with open(md_path, "w", encoding="utf-8") as fp:
                    fp.write(md_text)
                if os.path.exists(DIR_ORIG):
                    with open(md_orig_path, "w", encoding="utf-8") as fp:
                        fp.write(md_text)
                print(f" -> Hoàn tất OCR {len(pages)} trang.")
            except Exception as e:
                print(f" [LỖI OCR] {pdf_name}: {e}")

        # Trích xuất metadata
        meta = csv_rows.get(pdf_name, {})
        if not meta:
            m = re.match(r"^\[(.*?)\]", pdf_name)
            p = m.group(1) if m else pdf_name
            meta_list = csv_rows.get(p, [{}])
            meta = meta_list[0] if meta_list else {}

        socv = meta.get("Số ký hiệu", "")
        if not socv:
            m_scv = re.match(r"^\[(.*?)\]", pdf_name)
            socv = m_scv.group(1).replace("_", "/") if m_scv else ""

        ngay_cv = meta.get("Ngày ban hành", "09/2026")
        raw_don_vi = meta.get("Đơn vị ban hành", "")
        don_vi_bh = MAP_DON_VI.get(raw_don_vi.strip().upper(), raw_don_vi.strip())
        trich_yeu = meta.get("Trích yếu nội dung", "")

        # Trích xuất MST người nộp thuế
        mst_ban = ""
        mst_matches = re.findall(r"(?:mã số thuế|mst|mã số doanh nghiệp)[:\s]*([0-9]{10}(?:-[0-9]{3})?)", md_text, re.IGNORECASE)
        if mst_matches:
            mst_ban = mst_matches[0]
        else:
            m_mst = re.search(r"\b(36\d{8}|38\d{8}|03\d{8}|[0-9]{10})\b", trich_yeu + " " + pdf_name)
            if m_mst:
                mst_ban = m_mst.group(1)

        # Trích xuất tên người nộp thuế
        ten_ban = ""
        m_nnt = re.search(r"(?:Tên\s+người\s+nộp\s+thuế|Tên\s+doanh\s+nghiệp|Tên\s+đơn\s+vị|Người\s+nộp\s+thuế)[:\s]*([^\n\r,]{5,80})", md_text, re.IGNORECASE)
        if m_nnt:
            cand = clean_company_name(m_nnt.group(1))
            if len(cand) > 5 and not any(w in cand.lower() for w in ["dấu hiệu", "rủi ro", "không hoạt động", "hóa đơn"]):
                ten_ban = cand

        if not ten_ban:
            m_ty = re.search(r"\((?:Cty|Công ty|Doanh nghiệp|DNTN|DNNN)\s+([^)]+)\)", trich_yeu, re.IGNORECASE)
            if m_ty:
                cand = m_ty.group(0).strip("()")
                if not re.match(r"^(?:Công ty|Cty|Doanh nghiệp|DNTN)", cand, re.IGNORECASE):
                    cand = "Công ty " + cand
                ten_ban = cand

        if not ten_ban:
            m_ty2 = re.search(r"(?:CÔNG TY TNHH|CÔNG TY CỔ PHẦN|CÔNG TY CP|CÔNG TY MTV|DOANH NGHIỆP TƯ NHÂN|DNTN|CTY TNHH|CTY CP|Cty\s+|Công ty\s+)[^,;:\n\r()]+", trich_yeu, re.IGNORECASE)
            if m_ty2:
                cand = clean_company_name(m_ty2.group(0))
                if len(cand) > 5:
                    ten_ban = cand

        if not ten_ban:
            m_cty = re.findall(r"(?:CÔNG TY TNHH|CÔNG TY CỔ PHẦN|CÔNG TY CP|DOANH NGHIỆP TƯ NHÂN|DNTN)[^\n\r,.:;]{4,80}", md_text)
            for c in m_cty:
                if "THÀNH PHỐ" not in c and "ĐỒNG NAI" not in c and "VIỆT NAM" not in c and "HẠNH PHÚC" not in c:
                    ten_ban = c.strip()
                    break

        if not ten_ban:
            ten_ban = "Doanh nghiệp theo CV " + socv

        # Trích xuất địa chỉ
        dia_chi = ""
        m_addr = re.search(r"(?:địa chỉ|trụ sở)[:\s]*([^\n\r]{10,150})", md_text, re.IGNORECASE)
        if m_addr:
            dia_chi = m_addr.group(1).strip()
            dia_chi = re.sub(r"^(?:đăng ký kinh doanh|trụ sở chính|trụ sở|đã đăng ký)[:\s]*", "", dia_chi, flags=re.IGNORECASE).strip()

        # Dấu hiệu vi phạm
        reasons = []
        if re.search(r"không hoạt động tại địa chỉ", md_text, re.IGNORECASE):
            reasons.append("Không hoạt động tại địa chỉ đăng ký")
        if re.search(r"chưa nộp tờ khai|không nộp tờ khai", md_text, re.IGNORECASE):
            reasons.append("Chưa nộp tờ khai thuế")
        if re.search(r"chênh lệch.*hóa đơn|chênh lệch.*tờ khai", md_text, re.IGNORECASE):
            reasons.append("Chênh lệch số liệu tờ khai và HĐĐT")
        if re.search(r"cơ quan điều tra|chuyển tin báo", md_text, re.IGNORECASE):
            reasons.append("Chuyển tin báo CQĐT")
        if re.search(r"tạm ngừng kinh doanh", md_text, re.IGNORECASE):
            reasons.append("Tạm ngừng kinh doanh")
        if re.search(r"bãi bỏ", md_text + " " + trich_yeu, re.IGNORECASE):
            reasons.append("Bãi bỏ văn bản cảnh báo trước")
        if re.search(r"gian lận.*hoá đơn|bán lẻ.*hoá đơn", md_text, re.IGNORECASE):
            reasons.append("Nghi vấn gian lận hóa đơn")

        noi_dung = "; ".join(reasons) if reasons else (trich_yeu if trich_yeu else "Cảnh báo doanh nghiệp có rủi ro cao về thuế, hóa đơn")

        # Cờ đếm duy nhất
        mst_dau = 1 if mst_ban and (mst_ban not in seen_msts) else None
        if mst_ban:
            seen_msts.add(mst_ban)

        rec = {
            "so_cv": socv,
            "ngay_cv": ngay_cv,
            "don_vi_bh": don_vi_bh,
            "ten_ban": ten_ban,
            "mst_ban": mst_ban,
            "dia_chi": dia_chi,
            "noi_dung": noi_dung,
            "nhom_nd": "Dấu hiệu rủi ro về thuế",
            "file": os.path.join(DIR_ORIG, pdf_name),
            "nam_cv": 2026,
            "ghi_chu": f"Nguồn: CBRR-ben ban (Scan OCR). Trích yếu: {trich_yeu}.",
            "cv_dau": 1,
            "mst_dau": mst_dau
        }
        records.append(rec)
        print(f" -> Trích xuất thành công: {socv} | {ten_ban} | MST: {mst_ban}")

    print(f"\nTổng số bản ghi trích xuất: {len(records)}")
    out_json = Path(__file__).parent / "records_2026_09_scan.json"
    with open(out_json, "w", encoding="utf-8") as fp:
        json.dump(records, fp, ensure_ascii=False, indent=2)
    print(f"Đã lưu kết quả tại: {out_json}")

    # Ghi vào Excel nếu người dùng yêu cầu
    if target_excel and os.path.exists(target_excel):
        print(f"\nTiến hành nạp vào sổ theo dõi Excel: {target_excel}")
        new_max = update_workbook(records, target_excel)
        print(f"Nạp thành công! Dòng cuối mới: {new_max}")
    return records

if __name__ == "__main__":
    excel_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL_PATH
    run_ocr_and_parse(excel_path)
