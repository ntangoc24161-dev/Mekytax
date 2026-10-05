"""
fast_batch_ocr.py
Thực hiện OCR offline nhanh cho các công văn scan (bỏ qua bảng biểu - chỉ quét trang chính),
trích xuất thông tin doanh nghiệp cảnh báo và nạp vào sổ theo dõi Excel.
"""

import os
import re
import csv
import json
import sys
import io
import time
import shutil
from pathlib import Path
import numpy as np
from PIL import Image
import fitz  # PyMuPDF

# Đường dẫn javis-os & scripts
JAVIS_OS_DIR = r"D:\SECOND BRAIN\javis-os"
SCRIPTS_DIR = r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo\skills\theo-doi-cv-canh-bao-rui-ro\scripts"
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)
if JAVIS_OS_DIR not in sys.path:
    sys.path.insert(0, JAVIS_OS_DIR)

from update_so_theo_doi import update_workbook, DEFAULT_EXCEL_PATH

DIR_SCAN_09 = r"D:\HO SO CONG VAN CANH BAO\Cong_van_chua_boc_tach\2026-09"
DIR_ORIG_09 = r"D:\HO SO CONG VAN CANH BAO\2026\2026-09"
CSV_PATH = r"D:\HO SO CONG VAN CANH BAO\Cong_van_chua_boc_tach\DANH_SACH_CHUA_BOC_TACH.csv"

def clean_company_name(s):
    if not s:
        return ""
    s = re.sub(r"^(?:v/v|vv|v/v\s+cảnh\s+báo|cảnh\s+báo|cvr|cbrr|cv|công\s+văn)[\s:.-]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:doanh\s+nghiệp\s+có\s+dấu\s+hiệu\s+rủi\s+ro\s+cao\s+về\s+thuế[^\w]*)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:đối\s+với\s+)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\s:;.-]+(?:mst|mã\s+số\s+thuế).*$", "", s, flags=re.IGNORECASE)
    return s.strip(" ,.:;-\"'\n\r\t")

def normalize_don_vi(dv_raw):
    if not dv_raw:
        return ""
    dv = dv_raw.strip()
    dv_upper = dv.upper()
    
    # Đồng Nai
    map_dn = {
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
    if dv_upper in map_dn:
        return map_dn[dv_upper]
    
    # TP.HCM và các tỉnh
    dv = re.sub(r"THUẾ CƠ SỞ (\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv = re.sub(r"Thuế CS(\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv = re.sub(r"THUẾ CS(\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv = re.sub(r"THUẾ TỈNH\s+", "Thuế tỉnh ", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TP\.?\s*HCM", "TP. Hồ Chí Minh", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TP\s+HỒ\s+CHÍ\s+MINH", "TP. Hồ Chí Minh", dv, flags=re.IGNORECASE)
    dv = re.sub(r"T\.\s*Quảng Trị|TỈNH QUẢNG TRỊ", "tỉnh Quảng Trị", dv, flags=re.IGNORECASE)
    dv = re.sub(r"T\.\s*Lâm Đồng|TỈNH LÂM ĐỒNG", "tỉnh Lâm Đồng", dv, flags=re.IGNORECASE)
    dv = re.sub(r"T\.\s*Tây Ninh|TỈNH TÂY NINH", "tỉnh Tây Ninh", dv, flags=re.IGNORECASE)
    dv = re.sub(r"T\.\s*Khánh Hòa|TỈNH KHÁNH HÒA", "tỉnh Khánh Hòa", dv, flags=re.IGNORECASE)
    dv = re.sub(r"Tỉnh\s+", "tỉnh ", dv, flags=re.IGNORECASE)
    return dv

def init_ocr():
    from paddleocr import PaddleOCR
    print("[OCR] Khởi tạo mô hình PaddleOCR tiếng Việt (tối ưu tốc độ, bỏ phân loại hướng)...")
    return PaddleOCR(
        lang="vi",
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False
    )

def ocr_page_to_text(ocr, page):
    pix = page.get_pixmap(dpi=100)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    res = ocr.predict(np.array(img))
    if not res or not res[0]:
        return ""
    data = res[0]
    texts = data.get("rec_texts")
    if texts is None:
        texts = data.get("rec_text")
    if texts is None:
        texts = []
    boxes = data.get("rec_boxes")
    if boxes is None:
        boxes = data.get("rec_polys")
    if boxes is None:
        boxes = data.get("dt_polys")
    if boxes is None:
        boxes = []
    if len(texts) == 0:
        return ""
        
    rows = []
    for i, t in enumerate(texts):
        y, x = 0.0, 0.0
        if i < len(boxes):
            pts = list(boxes[i])
            if pts and hasattr(pts[0], "__len__"):
                ys = [float(p[1]) for p in pts]
                xs = [float(p[0]) for p in pts]
                y, x = sum(ys) / len(ys), sum(xs) / len(xs)
        rows.append((y, x, str(t)))
    rows.sort()
    
    lines = []
    buf = []
    last_y = None
    for y, x, t in rows:
        if last_y is None or abs(y - last_y) < 12:
            buf.append((x, t))
            last_y = y if last_y is None else (last_y + y) / 2
        else:
            buf.sort()
            lines.append(" ".join(p[1] for p in buf))
            buf = [(x, t)]
            last_y = y
    if buf:
        buf.sort()
        lines.append(" ".join(p[1] for p in buf))
    return "\n".join(lines)

def run_process_september(target_excel=DEFAULT_EXCEL_PATH):
    print("=== BẮT ĐẦU XỬ LÝ 16 CÔNG VĂN SCAN THÁNG 09/2026 (BỎ QUA BẢNG BIỂU) ===")
    
    # 1. Đọc metadata từ CSV
    csv_rows = {}
    if os.path.exists(CSV_PATH):
        with open(CSV_PATH, "r", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                if row.get("Kỳ") == "2026-09":
                    fn = row.get("Tên tệp PDF", "")
                    csv_rows[fn] = row
                    m = re.match(r"^\[(.*?)\]", fn)
                    if m:
                        csv_rows.setdefault(m.group(1), row)

    pdf_files = [f for f in sorted(os.listdir(DIR_SCAN_09)) if f.endswith(".pdf")]
    print(f"Tổng số tệp PDF scan tháng 9: {len(pdf_files)} tệp")

    ocr = None
    records = []
    seen_msts = set()
    if target_excel and os.path.exists(target_excel):
        try:
            import openpyxl
            wb_ro = openpyxl.load_workbook(target_excel, read_only=True)
            if "THEO_DOI" in wb_ro.sheetnames:
                ws_ro = wb_ro["THEO_DOI"]
                for row in ws_ro.iter_rows(min_row=7, min_col=6, max_col=6, values_only=True):
                    if row[0]:
                        seen_msts.add(str(row[0]).strip())
            wb_ro.close()
            print(f"Đã nạp {len(seen_msts)} MST hiện có từ sổ theo dõi Excel để đối soát duy nhất (MSTDau).")
        except Exception as e:
            print(f"Không thể đọc danh sách MST cũ: {e}")

    for idx, pdf_name in enumerate(pdf_files, 1):
        pdf_path = os.path.join(DIR_SCAN_09, pdf_name)
        md_name = pdf_name[:-4] + ".md"
        md_path = os.path.join(DIR_SCAN_09, md_name)
        md_orig_path = os.path.join(DIR_ORIG_09, md_name)

        print(f"\n[{idx}/{len(pdf_files)}] Xử lý: {pdf_name}")
        md_text = ""
        
        # Nếu đã có file md
        if os.path.exists(md_path) and os.path.getsize(md_path) > 100:
            print(f" -> Đã có sẵn file Markdown: {md_name}")
            with open(md_path, "r", encoding="utf-8", errors="ignore") as fp:
                md_text = fp.read()
        elif os.path.exists(md_orig_path) and os.path.getsize(md_orig_path) > 100:
            print(f" -> Đã có sẵn file Markdown tại thư mục gốc: {md_name}")
            with open(md_orig_path, "r", encoding="utf-8", errors="ignore") as fp:
                md_text = fp.read()
            shutil.copy2(md_orig_path, md_path)

        # Kiểm tra xem md_text đã có MST chưa
        has_mst = False
        if md_text:
            has_mst = bool(re.search(r"(?:mã số thuế|mst|mã số doanh nghiệp|mã s[oố] thu[eế])[:\s]*([0-9]{10}(?:-[0-9]{3})?)", md_text, re.IGNORECASE))
            if not has_mst:
                has_mst = bool(re.search(r"\b(0\d{9}|3\d{9}|\d{10})\b", md_text))

        # Nếu chưa có md hoặc md chưa có MST, tiến hành OCR bổ sung
        if not md_text or not has_mst:
            if ocr is None:
                ocr = init_ocr()
            doc = fitz.open(pdf_path)
            num_pages = len(doc)
            print(f" -> PDF có {num_pages} trang. Tiến hành OCR trang chính (tối đa trang 4)...")

            md_chunks = [md_text] if md_text else []
            start_p = 0
            if md_text:
                existing_pages = re.findall(r"## Trang (\d+)", md_text)
                if existing_pages:
                    start_p = max(int(p) for p in existing_pages)

            for p_idx in range(start_p, min(num_pages, 4)):
                t0 = time.time()
                p_text = ocr_page_to_text(ocr, doc[p_idx])
                elapsed = time.time() - t0
                print(f"    Trang {p_idx + 1}: hoàn tất trong {elapsed:.1f}s")
                md_chunks.append(f"## Trang {p_idx + 1}\n\n{p_text}")

                found = bool(re.search(r"(?:mã số thuế|mst|mã số doanh nghiệp|mã s[oố] thu[eế])[:\s]*([0-9]{10}(?:-[0-9]{3})?)", p_text, re.IGNORECASE))
                if not found:
                    found = bool(re.search(r"\b(0\d{9}|3\d{9}|\d{10})\b", p_text))
                if found:
                    print(f"    (Đã tìm thấy MST tại Trang {p_idx + 1}, dừng quét bảng biểu)")
                    break

            doc.close()
            md_text = "\n\n".join(c for c in md_chunks if c.strip())
            with open(md_path, "w", encoding="utf-8") as fp:
                fp.write(md_text)
            if os.path.exists(DIR_ORIG_09):
                with open(md_orig_path, "w", encoding="utf-8") as fp:
                    fp.write(md_text)

        # Trích xuất metadata
        meta = csv_rows.get(pdf_name, {})
        if not meta:
            m = re.match(r"^\[(.*?)\]", pdf_name)
            p = m.group(1) if m else pdf_name
            meta = csv_rows.get(p, {})

        socv = meta.get("Số công văn", "")
        if not socv:
            m_scv = re.search(r"Số[:\s]*([0-9A-Z/_-]+)", md_text)
            if m_scv:
                socv = m_scv.group(1).strip()
            else:
                m_p = re.match(r"^\[(.*?)\]", pdf_name)
                socv = m_p.group(1).replace("_", "/") if m_p else ""

        ngay_cv = meta.get("Ngày ban hành", "09/2026")
        don_vi_bh = normalize_don_vi(meta.get("Đơn vị ban hành", ""))
        trich_yeu = meta.get("Trích yếu nội dung", "")

        # Trích xuất MST người nộp thuế
        mst_ban = ""
        mst_matches = re.findall(r"(?:mã số thuế|mst|mã số doanh nghiệp|mã s[oố] thu[eế])[:\s]*([0-9]{10}(?:-[0-9]{3})?)", md_text, re.IGNORECASE)
        if mst_matches:
            mst_ban = mst_matches[0]
        else:
            m_mst = re.findall(r"\b(0\d{9}|3\d{9}|[0-9]{10})\b", md_text)
            for m_cand in m_mst:
                if m_cand not in (socv, "20260901", "20260930"):
                    mst_ban = m_cand
                    break

        # Trích xuất tên người nộp thuế
        ten_ban = ""
        m_nnt = re.search(r"(?:Tên\s+ngư[^\s:]*\s+n[^\s:]*\s+thu[^\s:]*|Tên\s+người\s+nộp\s+thuế|Tên\s+doanh\s+nghiệp|Tên\s+đơn\s+vị|Người\s+nộp\s+thuế|NNT)[:\s]*([^\n\r,;]{5,80})", md_text, re.IGNORECASE)
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
            m_cty = re.findall(r"(?:CÔNG TY TNHH|CÔNG TY CỔ PHẦN|CÔNG TY CP|CÔNG TY MTV|DOANH NGHIỆP TƯ NHÂN|DNTN|Công ty TNHH|Công ty CP)[^\n\r,.:;]{4,80}", md_text)
            for c in m_cty:
                if not any(w in c.upper() for w in ["THÀNH PHỐ", "ĐỒNG NAI", "VIỆT NAM", "HẠNH PHÚC", "CỘNG HÒA"]):
                    ten_ban = c.strip()
                    break

        if not ten_ban:
            m_fn = re.search(r"(?:CÔNG TY|CTY|DOANH NGHIỆP|DNTN)[^.]+", pdf_name, re.IGNORECASE)
            if m_fn:
                ten_ban = clean_company_name(m_fn.group(0))

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
            "file": os.path.join(DIR_ORIG_09, pdf_name),
            "nam_cv": 2026,
            "ghi_chu": f"Nguồn: CBRR-ben ban (Scan OCR - Bỏ qua bảng biểu). Trích yếu: {trich_yeu}.",
            "cv_dau": 1,
            "mst_dau": mst_dau
        }
        records.append(rec)
        print(f" -> Trích xuất thành công: {socv} | {ten_ban} | MST: {mst_ban} | ĐV: {don_vi_bh}")

    # Gom nhóm deduplicate theo số công văn để nếu 1 CV có 2 file chỉ lấy 1 dòng đại diện
    dedup_records = []
    seen_socv = set()
    for r in records:
        scv = r["so_cv"]
        if scv not in seen_socv:
            seen_socv.add(scv)
            dedup_records.append(r)
        else:
            # Nếu bản ghi trước chưa có MST mà bản ghi này có thì bổ sung
            for prev in dedup_records:
                if prev["so_cv"] == scv and not prev["mst_ban"] and r["mst_ban"]:
                    prev["mst_ban"] = r["mst_ban"]
                    prev["ten_ban"] = r["ten_ban"]
                    prev["dia_chi"] = r["dia_chi"]

    print(f"\nTổng số bản ghi công văn duy nhất: {len(dedup_records)}")
    out_json = Path(SCRIPTS_DIR) / "records_2026_09_scan.json"
    with open(out_json, "w", encoding="utf-8") as fp:
        json.dump(dedup_records, fp, ensure_ascii=False, indent=2)
    print(f"Đã lưu kết quả JSON tại: {out_json}")

    # Ghi vào Excel
    if target_excel and os.path.exists(target_excel):
        print(f"\nTiến hành nạp vào sổ theo dõi Excel: {target_excel}")
        new_max = update_workbook(dedup_records, target_excel)
        print(f"Nạp thành công! Dòng cuối mới: {new_max}")
    return dedup_records

if __name__ == "__main__":
    excel_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL_PATH
    run_process_september(excel_path)
