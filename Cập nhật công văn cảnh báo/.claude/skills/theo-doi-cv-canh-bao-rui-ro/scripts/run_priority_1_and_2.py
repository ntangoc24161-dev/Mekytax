"""
run_priority_1_and_2.py
Thực hiện tự động toàn bộ 3 giai đoạn:
- Ưu tiên 1: Chạy OCR offline cho 57 công văn scan tháng 08/2026 (66 file PDF), trích xuất thông tin doanh nghiệp bị cảnh báo, bỏ qua hóa đơn, doanh thu, thuế.
- Ưu tiên 2: Quét và bóc tách nhóm công văn số hóa (digital) tháng 07/2026 và 06/2026, bỏ qua hóa đơn, doanh thu, thuế.
- Rà soát bỏ qua các dòng trùng (Số công văn, Tên công ty bán, Tên công ty mua) đối với cả dữ liệu đã có trong Excel và dữ liệu trong các đợt nạp.
- Nạp dữ liệu vào sổ Excel SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx, cập nhật cờ duy nhất (CVDau, MSTDau) và đồng bộ công thức TONG_HOP.
- Hỗ trợ cơ chế checkpoint cache lưu tức thì từng tệp để tự động phục hồi nếu bị tạm dừng đột ngột.
"""

import os
import re
import csv
import json
import sys
import io
import time
import shutil
import concurrent.futures
from datetime import datetime
from pathlib import Path
from PIL import Image
import numpy as np
import fitz  # PyMuPDF
import openpyxl

os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

SCRIPT_DIR = Path(__file__).resolve().parent
CACHE_SCAN_08 = SCRIPT_DIR / "cache_scan_08.json"
CACHE_DIGITAL = SCRIPT_DIR / "cache_digital_07_06.json"
COMPLETION_MARKER = SCRIPT_DIR / "COMPLETED_PRIORITY_1_AND_2.txt"

DEFAULT_EXCEL_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
if not Path(DEFAULT_EXCEL_PATH).is_file():
    DEFAULT_EXCEL_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"

CSV_CHUA_BOC_TACH = r"D:\HO SO CONG VAN CANH BAO\Cong_van_chua_boc_tach\DANH_SACH_CHUA_BOC_TACH.csv"
CSV_TONG = r"D:\HO SO CONG VAN CANH BAO\DANH_SACH_CONG_VAN_CANH_BAO.csv"

DIR_SCAN_08 = r"D:\HO SO CONG VAN CANH BAO\Cong_van_chua_boc_tach\2026-08"
DIR_ORIG_08 = r"D:\HO SO CONG VAN CANH BAO\2026\2026-08"
DIR_ORIG_07 = r"D:\HO SO CONG VAN CANH BAO\2026\2026-07"
DIR_ORIG_06 = r"D:\HO SO CONG VAN CANH BAO\2026\2026-06"

_global_ocr = None
_worker_ocr = None

LOG_FILE = SCRIPT_DIR / "run.log"

def safe_print(msg):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{now_str}] {msg}"
    try:
        print(formatted, flush=True)
    except Exception:
        try:
            print(formatted.encode('ascii', errors='replace').decode('ascii'), flush=True)
        except Exception:
            pass
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(formatted + "\n")
    except Exception:
        pass

def init_worker_ocr():
    global _worker_ocr
    from paddleocr import PaddleOCR
    _worker_ocr = PaddleOCR(
        lang="vi",
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        enable_mkldnn=True,
        cpu_threads=2
    )

def get_ocr():
    global _global_ocr
    if _global_ocr is None:
        safe_print("[OCR] Khởi tạo mô hình PaddleOCR tiếng Việt (tối ưu MKL-DNN, 8 threads)...")
        from paddleocr import PaddleOCR
        _global_ocr = PaddleOCR(
            lang="vi",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            enable_mkldnn=True,
            cpu_threads=8
        )
    return _global_ocr

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

def clean_company_name(s):
    if not s:
        return ""
    s = re.sub(r"^(?:v/v|vv|v/v\s+cảnh\s+báo|cảnh\s+báo|cvr|cbrr|cv|công\s+văn)[\s:.-]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:doanh\s+nghiệp\s+có\s+dấu\s+hiệu\s+rủi\s+ro\s+cao\s+về\s+thuế[^\w]*)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:đối\s+với\s+)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\s:;.-]+(?:mst|mã\s+số\s+thuế).*$", "", s, flags=re.IGNORECASE)
    return s.strip(" ,.:;-\"'\n\r\t")

def normalize_key_str(s):
    if not s:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"[^\w\s]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_don_vi(dv_raw):
    if not dv_raw:
        return ""
    dv = dv_raw.strip()
    dv_upper = dv.upper()
    
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
        "CỤC THUẾ THÀNH PHỐ ĐỒNG NAI": "Thuế thành phố Đồng Nai",
    }
    if dv_upper in map_dn:
        return map_dn[dv_upper]
        
    dv = re.sub(r"THUẾ CƠ SỞ (\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv = re.sub(r"Thuế CS(\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv = re.sub(r"THUẾ CS(\d+)", r"Thuế cơ sở \1", dv, flags=re.IGNORECASE)
    dv = re.sub(r"Thuế cơ sở 28 QLDN", "Thuế cơ sở 28 TP. Hồ Chí Minh", dv, flags=re.IGNORECASE)
    dv = re.sub(r"Thuế cơ sở TP\s*HCM", "Thuế cơ sở TP. Hồ Chí Minh", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TP\.?\s*HCM|TP\s+HỒ\s+CHÍ\s+MINH|TPHCM", "TP. Hồ Chí Minh", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TP\s+HẢI\s+PHÒNG|TP\.?\s*Hải Phòng", "TP. Hải Phòng", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TP\s+CẦN\s+THƠ|TP\.?\s*Cần Thơ", "TP. Cần Thơ", dv, flags=re.IGNORECASE)
    dv = re.sub(r"T\.\s*Quảng Trị|TỈNH QUẢNG TRỊ|tỉnh Quảng Trị", "tỉnh Quảng Trị", dv, flags=re.IGNORECASE)
    dv = re.sub(r"T\.\s*Lâm Đồng|TỈNH LÂM ĐỒNG", "tỉnh Lâm Đồng", dv, flags=re.IGNORECASE)
    dv = re.sub(r"T\.\s*Tây Ninh|TỈNH TÂY NINH", "tỉnh Tây Ninh", dv, flags=re.IGNORECASE)
    dv = re.sub(r"T\.\s*Khánh Hòa|TỈNH KHÁNH HÒA", "tỉnh Khánh Hòa", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TỈNH HƯNG YÊN|Hưng Yên", "tỉnh Hưng Yên", dv, flags=re.IGNORECASE)
    dv = re.sub(r"tỉnh Đăk Lăk|ĐẮK LẮK", "tỉnh Đắk Lắk", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TỈNH AN GIANG", "tỉnh An Giang", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TỈNH BẮC NINH", "tỉnh Bắc Ninh", dv, flags=re.IGNORECASE)
    dv = re.sub(r"TỈNH BÌNH DƯƠNG|Bình Dương", "tỉnh Bình Dương", dv, flags=re.IGNORECASE)
    dv = re.sub(r"THUẾ TỈNH\s+", "Thuế tỉnh ", dv, flags=re.IGNORECASE)
    dv = re.sub(r"\s+", " ", dv).strip()
    return dv

def extract_record_fields(text, meta, pdf_name, default_period="2026-08"):
    socv = meta.get("Số công văn") or meta.get("Số ký hiệu") or ""
    if not socv:
        m_scv = re.search(r"Số[:\s]*([0-9A-Z/_-]+)", text)
        if m_scv:
            socv = m_scv.group(1).strip()
        else:
            m_p = re.match(r"^\[(.*?)\]", pdf_name)
            socv = m_p.group(1).replace("_", "/") if m_p else ""
            
    ngay_cv = meta.get("Ngày ban hành", "")
    if not ngay_cv:
        m_ng = re.search(r"ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})", text, re.IGNORECASE)
        if m_ng:
            d, m, y = m_ng.groups()
            ngay_cv = f"{int(d):02d}/{int(m):02d}/{y}"
        else:
            ngay_cv = default_period

    don_vi_bh = normalize_don_vi(meta.get("Đơn vị ban hành", ""))
    if not don_vi_bh:
        m_dv = re.search(r"(THUẾ\s+[^\n\r]{5,40}|CHI CỤC THUẾ[^\n\r]{5,40}|CỤC THUẾ[^\n\r]{5,40})", text, re.IGNORECASE)
        if m_dv:
            don_vi_bh = normalize_don_vi(m_dv.group(1))

    trich_yeu = meta.get("Trích yếu nội dung", "")
    
    # 1. MST bên bán
    mst_ban = ""
    mst_matches = re.findall(r"(?:mã số thuế|mst|mã số doanh nghiệp|mã s[oố\s]*thu[eế]|mã\s*s\s*thu[eêế])[:\s]*([0-9]{10}(?:-[0-9]{3})?)", text, re.IGNORECASE)
    if mst_matches:
        mst_ban = mst_matches[0]
    else:
        m_fn_mst = re.search(r"_(\d{10})\.", pdf_name)
        if m_fn_mst:
            mst_ban = m_fn_mst.group(1)
        else:
            m_mst = re.findall(r"\b(0\d{9}|3\d{9}|[0-9]{10})\b", text)
            for m_cand in m_mst:
                if m_cand not in (socv, "20260801", "20260831", "20260701", "20260731", "20260601", "20260630"):
                    mst_ban = m_cand
                    break

    # 2. Tên doanh nghiệp bên bán
    ten_ban = ""
    m_nnt = re.search(r"[-–*•\s]*(?:Tên\s+người\s+nộp\s+thuế|Tên\s+doanh\s+nghiệp|Tên\s+đơn\s+vị|Người\s+nộp\s+thuế|NNT)\s*:\s*([^\n\r;]{5,80})", text, re.IGNORECASE)
    if m_nnt:
        cand = clean_company_name(m_nnt.group(1))
        if len(cand) > 5 and not any(w in cand.lower() for w in ["dấu hiệu", "rủi ro", "không hoạt động", "hóa đơn", "kính gửi"]):
            ten_ban = cand

    if not ten_ban and trich_yeu:
        m_ty = re.search(r"\((?:Cty|Công ty|Doanh nghiệp|DNTN|DNNN)\s+([^)]+)\)", trich_yeu, re.IGNORECASE)
        if m_ty:
            cand = m_ty.group(0).strip("()")
            if not re.match(r"^(?:Công ty|Cty|Doanh nghiệp|DNTN)", cand, re.IGNORECASE):
                cand = "Công ty " + cand
            ten_ban = cand

    if not ten_ban:
        m_cty = re.findall(r"(?:CÔNG TY TNHH|CÔNG TY CỔ PHẦN|CÔNG TY CP|CÔNG TY MTV|DOANH NGHIỆP TƯ NHÂN|DNTN|Công ty TNHH|Công ty CP)[^\n\r,.:;]{4,80}", text)
        for c in m_cty:
            if not any(w in c.upper() for w in ["THÀNH PHỐ", "ĐỒNG NAI", "VIỆT NAM", "HẠNH PHÚC", "CỘNG HÒA"]):
                ten_ban = c.strip()
                break

    if not ten_ban:
        m_fn = re.search(r"(?:CÔNG TY|CTY|DOANH NGHIỆP|DNTN|CONG TY)[^._]+", pdf_name, re.IGNORECASE)
        if m_fn:
            ten_ban = clean_company_name(m_fn.group(0))

    if not ten_ban:
        ten_ban = "Doanh nghiệp theo CV " + socv

    # 3. Địa chỉ
    dia_chi = ""
    m_addr = re.search(r"(?:[đd][iị]a\s*ch[iỉí]|trụ\s*sở)[^:\n\r]*[:\s]+([^\n\r]{10,150})", text, re.IGNORECASE)
    if m_addr:
        dia_chi = m_addr.group(1).strip()
        dia_chi = re.sub(r"^(?:đăng ký kinh doanh|trụ sở chính|trụ sở|đã đăng ký)[:\s]*", "", dia_chi, flags=re.IGNORECASE).strip()

    # 4. Dấu hiệu rủi ro
    reasons = []
    if re.search(r"(?:kh[oô]ng\s+ho[aạ]t\s+[đd][oô]ng|khong\s+hoat\s+dong).*?(?:[đd][iị]a\s+ch[iỉí]|[đd]ckd)", text, re.IGNORECASE):
        reasons.append("Không hoạt động tại địa chỉ đăng ký")
    if re.search(r"(?:ch[uư]a|kh[oô]ng)\s+n[oộ]p\s+t[oờ]\s+khai", text, re.IGNORECASE):
        reasons.append("Chưa nộp tờ khai thuế")
    if re.search(r"ch[eê]nh\s+l[eệ]ch.*?(?:h[oó]a\s+[đd][oơ]n|t[oờ]\s+khai)", text, re.IGNORECASE):
        reasons.append("Chênh lệch số liệu tờ khai và HĐĐT")
    if re.search(r"(?:c[oơ]\s+quan\s+[đd]i[eề]u\s+tra|chuy[eể]n\s+tin\s+b[aá]o)", text, re.IGNORECASE):
        reasons.append("Chuyển tin báo CQĐT")
    if re.search(r"t[aạ]m\s+ng[uừ]ng\s+kinh\s+doanh", text, re.IGNORECASE):
        reasons.append("Tạm ngừng kinh doanh")
    if re.search(r"b[aã]i\s+b[oỏ]", text + " " + trich_yeu, re.IGNORECASE):
        reasons.append("Bãi bỏ văn bản cảnh báo trước")
    if re.search(r"gian\s+l[aậ]n.*?(?:h[oó]a\s+[đd][oơ]n|ho[aá]\s+[đd]on)", text, re.IGNORECASE):
        reasons.append("Nghi vấn gian lận hóa đơn")

    noi_dung = "; ".join(reasons) if reasons else (trich_yeu if trich_yeu else "Cảnh báo doanh nghiệp có rủi ro cao về thuế, hóa đơn")

    return {
        "so_cv": socv,
        "ngay_cv": ngay_cv,
        "don_vi_bh": don_vi_bh,
        "ten_ban": ten_ban,
        "mst_ban": mst_ban,
        "dia_chi": dia_chi,
        "noi_dung": noi_dung,
        "ten_mua": None,
        "mst_mua": None,
        "nhom_nd": "Dấu hiệu rủi ro về thuế",
        "file": pdf_name,
        "nam_cv": 2026,
        "trich_yeu": trich_yeu
    }

def process_single_scan_pdf(pdf_name, dir_scan, meta, ocr):
    pdf_path = os.path.join(dir_scan, pdf_name)
    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    text = ""
    for p_idx in range(min(num_pages, 2)):
        page = doc[p_idx]
        # Cắt thông minh 45% phần trên trang chứa thông tin đơn vị, NNT, MST và lý do cảnh báo (bỏ qua bảng kê HĐ phía dưới)
        rect = fitz.Rect(0, 0, page.rect.width, page.rect.height * 0.45)
        pix = page.get_pixmap(dpi=70, clip=rect)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        res = ocr.predict(np.array(img))
        if res and res[0]:
            data = res[0]
            texts = data.get("rec_texts") or data.get("rec_text") or []
            p_text = "\n".join(str(t) for t in texts)
            text += "\n" + p_text
            if re.search(r"(?:mã số thuế|mst|mã số doanh nghiệp|mã s[oố] thu[eế])[:\s]*([0-9]{10}(?:-[0-9]{3})?)", p_text, re.IGNORECASE) or re.search(r"\b(0\d{9}|3\d{9})\b", p_text):
                break
    doc.close()
    rec = extract_record_fields(text, meta, pdf_name, default_period="08/2026")
    rec["file"] = os.path.join(DIR_ORIG_08, pdf_name)
    rec["ghi_chu"] = f"Nguồn: CBRR-ben ban (Scan OCR - Bỏ qua bảng biểu). Trích yếu: {rec.get('trich_yeu', '')}."
    return rec

def _worker_process_scan(args):
    global _worker_ocr
    fn, dir_scan, meta = args
    if _worker_ocr is None:
        init_worker_ocr()
    pdf_path = os.path.join(dir_scan, fn)
    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    text = ""
    for p_idx in range(min(num_pages, 2)):
        page = doc[p_idx]
        # Cắt thông minh 45% phần trên trang chứa thông tin đơn vị, NNT, MST và lý do cảnh báo
        rect = fitz.Rect(0, 0, page.rect.width, page.rect.height * 0.45)
        pix = page.get_pixmap(dpi=70, clip=rect)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        res = _worker_ocr.predict(np.array(img))
        if res and res[0]:
            data = res[0]
            texts = data.get("rec_texts") or data.get("rec_text") or []
            p_text = "\n".join(str(t) for t in texts)
            text += "\n" + p_text
            if re.search(r"(?:mã số thuế|mst|mã số doanh nghiệp|mã s[oố] thu[eế])[:\s]*([0-9]{10}(?:-[0-9]{3})?)", p_text, re.IGNORECASE) or re.search(r"\b(0\d{9}|3\d{9})\b", p_text):
                break
    doc.close()
    rec = extract_record_fields(text, meta, fn, default_period="08/2026")
    rec["file"] = os.path.join(DIR_ORIG_08, fn)
    rec["ghi_chu"] = f"Nguồn: CBRR-ben ban (Scan OCR Song Song - Bỏ qua bảng biểu). Trích yếu: {rec.get('trich_yeu', '')}."
    return fn, rec

def load_json_cache(cache_file):
    if cache_file.exists():
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_json_cache(cache_file, data):
    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        safe_print(f"[CACHE WARN] Không thể lưu cache: {e}")

def run_all(target_excel=DEFAULT_EXCEL_PATH):
    safe_print("="*70)
    safe_print("BẮT ĐẦU QUY TRÌNH XỬ LÝ TỰ ĐỘNG CẢ 3 GIAI ĐOẠN (SONG SONG ĐA TIẾN TRÌNH)")
    safe_print("="*70)
    
    # 1. Đọc dữ liệu hiện có trong Excel
    existing_keys = set()
    seen_msts = set()
    old_max_row = 0
    last_stt = 0
    
    if os.path.exists(target_excel):
        safe_print(f"[EXCEL] Đang nạp dữ liệu hiện có từ: {target_excel}...")
        wb_ro = openpyxl.load_workbook(target_excel, read_only=True, data_only=True)
        ws_ro = wb_ro["THEO_DOI"]
        old_max_row = ws_ro.max_row
        for r, row in enumerate(ws_ro.iter_rows(min_row=7, values_only=True), start=7):
            stt_val = row[0] if len(row) > 0 else None
            if isinstance(stt_val, (int, float)):
                last_stt = max(last_stt, int(stt_val))
            scv = normalize_key_str(row[1]) if len(row) > 1 else ""
            ban = normalize_key_str(row[4]) if len(row) > 4 else ""
            mua = normalize_key_str(row[7]) if len(row) > 7 else ""
            if scv or ban:
                existing_keys.add((scv, ban, mua))
            mst = row[5] if len(row) > 5 else None
            if mst:
                seen_msts.add(str(mst).strip())
        wb_ro.close()
        safe_print(f"[EXCEL] Dòng cuối hiện tại: {old_max_row}, STT cuối: {last_stt}, {len(existing_keys)} bộ khóa duy nhất, {len(seen_msts)} MST.")
    else:
        raise FileNotFoundError(f"Không tìm thấy file Excel tại {target_excel}")

    # Nạp metadata
    meta_chua_boc = {}
    if os.path.exists(CSV_CHUA_BOC_TACH):
        with open(CSV_CHUA_BOC_TACH, "r", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                meta_chua_boc[row.get("Tên tệp PDF", "").strip()] = row
                
    meta_tong = {}
    if os.path.exists(CSV_TONG):
        with open(CSV_TONG, "r", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                meta_tong[row.get("Tên file lưu", "").strip()] = row

    # =========================================================================
    # PHẦN 1: ƯU TIÊN 1 - 57 CÔNG VĂN SCAN THÁNG 08/2026 (66 TỆP PDF)
    # =========================================================================
    safe_print("\n" + "="*70)
    safe_print(">>> GIAI ĐOẠN 1: THỰC HIỆN ƯU TIÊN 1 (57 CÔNG VĂN SCAN THÁNG 08/2026 - SONG SONG)")
    safe_print("="*70)
    
    pdf_files_08 = sorted([f for f in os.listdir(DIR_SCAN_08) if f.endswith(".pdf")])
    safe_print(f"Tổng số tệp PDF scan cần xử lý: {len(pdf_files_08)} tệp")
    
    scan_cache = load_json_cache(CACHE_SCAN_08)
    t0_p1 = time.time()
    
    raw_p1_results = []
    uncached_tasks = []
    for idx, fn in enumerate(pdf_files_08, 1):
        if fn in scan_cache:
            rec = scan_cache[fn]
            safe_print(f"  [{idx}/{len(pdf_files_08)}] [Cache hit] {rec['so_cv']} | {rec['ten_ban']} | MST: {rec['mst_ban']}")
            raw_p1_results.append((fn, rec))
        else:
            meta = meta_chua_boc.get(fn) or meta_tong.get(fn) or {}
            scv_cand = meta.get("Số công văn") or meta.get("Số ký hiệu") or ""
            reused = None
            if scv_cand:
                for cached_fn, cached_rec in scan_cache.items():
                    if cached_rec.get("so_cv") == scv_cand and cached_rec.get("mst_ban"):
                        reused = dict(cached_rec)
                        reused["file"] = os.path.join(DIR_ORIG_08, fn)
                        break
            if reused:
                rec = reused
                scan_cache[fn] = rec
                save_json_cache(CACHE_SCAN_08, scan_cache)
                safe_print(f"  [{idx}/{len(pdf_files_08)}] [Tái sử dụng SCV] {rec['so_cv']} | {rec['ten_ban']} | MST: {rec['mst_ban']}")
                raw_p1_results.append((fn, rec))
            else:
                uncached_tasks.append((idx, fn, DIR_SCAN_08, meta))

    if uncached_tasks:
        max_workers = min(4, max(1, (os.cpu_count() or 4) // 2))
        safe_print(f"\n[SONG SONG] Khởi chạy {max_workers} worker processes OCR song song cho {len(uncached_tasks)} tệp scan...")
        with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers, initializer=init_worker_ocr) as executor:
            futures = {
                executor.submit(_worker_process_scan, (fn, dir_s, m)): (idx, fn)
                for (idx, fn, dir_s, m) in uncached_tasks
            }
            for future in concurrent.futures.as_completed(futures):
                idx, fn = futures[future]
                try:
                    fn_res, rec = future.result()
                    scan_cache[fn_res] = rec
                    save_json_cache(CACHE_SCAN_08, scan_cache)
                    safe_print(f"  [{idx}/{len(pdf_files_08)}] [Song Song OCR OK] {rec['so_cv']} | {rec['ten_ban']} | MST: {rec['mst_ban']}")
                    raw_p1_results.append((fn_res, rec))
                except Exception as e:
                    safe_print(f"  [{idx}/{len(pdf_files_08)}] [LỖI OCR {fn}]: {e}")

    # Rà soát trùng và gom nhóm
    records_p1 = []
    skipped_dup_p1 = 0
    for fn, rec in raw_p1_results:
        key = (normalize_key_str(rec["so_cv"]), normalize_key_str(rec["ten_ban"]), normalize_key_str(rec["ten_mua"]))
        if key in existing_keys:
            skipped_dup_p1 += 1
            continue
        existing_keys.add(key)
        records_p1.append(rec)

    # Deduplicate cùng số công văn trong đợt scan
    dedup_p1 = []
    seen_scv_p1 = {}
    for r in records_p1:
        scv = r["so_cv"]
        if scv not in seen_scv_p1:
            seen_scv_p1[scv] = len(dedup_p1)
            dedup_p1.append(r)
        else:
            prev_idx = seen_scv_p1[scv]
            prev = dedup_p1[prev_idx]
            if not prev["mst_ban"] and r["mst_ban"]:
                prev["mst_ban"] = r["mst_ban"]
                prev["ten_ban"] = r["ten_ban"]
                prev["dia_chi"] = r["dia_chi"]

    p1_time = time.time() - t0_p1
    safe_print(f"\n[KẾT QUẢ GIAI ĐOẠN 1] Bóc tách thành công {len(dedup_p1)} công văn duy nhất (bỏ qua {skipped_dup_p1} dòng trùng) trong {p1_time:.1f}s.")
    with open(SCRIPT_DIR / "records_priority_1.json", "w", encoding="utf-8") as f:
        json.dump(dedup_p1, f, ensure_ascii=False, indent=2)

    # =========================================================================
    # PHẦN 2: ƯU TIÊN 2 - NHÓM CÔNG VĂN SỐ HÓA THÁNG 07/2026 VÀ 06/2026
    # =========================================================================
    safe_print("\n" + "="*70)
    safe_print(">>> GIAI ĐOẠN 2: THỰC HIỆN ƯU TIÊN 2 (CÔNG VĂN SỐ HÓA THÁNG 07 & 06/2026)")
    safe_print("="*70)
    
    t0_p2 = time.time()
    skipped_dup_p2 = 0
    
    def process_digital_dir(dir_path, default_period):
        nonlocal skipped_dup_p2
        files = sorted([f for f in os.listdir(dir_path) if f.endswith(".pdf")])
        safe_print(f"\nQuét thư mục số hóa: {dir_path} ({len(files)} tệp PDF)...")
        month_recs = []
        for idx, pdf_name in enumerate(files, 1):
            pdf_path = os.path.join(dir_path, pdf_name)
            doc = fitz.open(pdf_path)
            first_text = doc[0].get_text() if len(doc) > 0 else ""
            if len(first_text.strip()) < 50:
                doc.close()
                continue
                
            full_text = ""
            for p in doc:
                full_text += p.get_text() + "\n"
                if len(full_text) > 3500:
                    break
            doc.close()
            
            meta = meta_tong.get(pdf_name) or {}
            rec = extract_record_fields(full_text, meta, pdf_name, default_period=default_period)
            rec["file"] = pdf_path
            rec["ghi_chu"] = f"Nguồn: CBRR-ben ban (Digital PDF - Bỏ qua bảng biểu). Trích yếu: {rec.get('trich_yeu', '')}."
            
            # Bỏ qua các văn bản hành chính nội bộ (không có MST và không có tên NNT rủi ro)
            if not rec["mst_ban"] and "Doanh nghiệp theo CV" in rec["ten_ban"]:
                continue
                
            key = (normalize_key_str(rec["so_cv"]), normalize_key_str(rec["ten_ban"]), normalize_key_str(rec["ten_mua"]))
            if key in existing_keys:
                skipped_dup_p2 += 1
                continue
                
            existing_keys.add(key)
            month_recs.append(rec)
            safe_print(f"  [Digital OK] {rec['so_cv']} | {rec['ten_ban']} | MST: {rec['mst_ban']} | ĐV: {rec['don_vi_bh']}")
            
        dedup_month = []
        seen_scv_m = {}
        for r in month_recs:
            scv = r["so_cv"]
            if scv not in seen_scv_m:
                seen_scv_m[scv] = len(dedup_month)
                dedup_month.append(r)
            else:
                prev_idx = seen_scv_m[scv]
                prev = dedup_month[prev_idx]
                if not prev["mst_ban"] and r["mst_ban"]:
                    prev["mst_ban"] = r["mst_ban"]
                    prev["ten_ban"] = r["ten_ban"]
                    prev["dia_chi"] = r["dia_chi"]
        return dedup_month

    recs_07 = process_digital_dir(DIR_ORIG_07, "07/2026")
    recs_06 = process_digital_dir(DIR_ORIG_06, "06/2026")
    records_p2 = recs_07 + recs_06
    p2_time = time.time() - t0_p2
    safe_print(f"\n[KẾT QUẢ GIAI ĐOẠN 2] Bóc tách thành công {len(records_p2)} công văn số hóa (T7: {len(recs_07)}, T6: {len(recs_06)}, bỏ qua {skipped_dup_p2} dòng trùng) trong {p2_time:.1f}s.")
    with open(SCRIPT_DIR / "records_priority_2.json", "w", encoding="utf-8") as f:
        json.dump(records_p2, f, ensure_ascii=False, indent=2)

    # =========================================================================
    # PHẦN 3: NẠP DỮ LIỆU VÀO SỔ THEO DÕI EXCEL & CẬP NHẬT CÔNG THỨC
    # =========================================================================
    all_new_records = dedup_p1 + records_p2
    safe_print("\n" + "="*70)
    safe_print(f">>> GIAI ĐOẠN 3: NẠP {len(all_new_records)} DÒNG DỮ LIỆU MỚI VÀO SỔ EXCEL")
    safe_print("="*70)
    
    if all_new_records:
        # Tạo bản sao lưu .bak
        bak_path = Path(target_excel).with_suffix(".xlsx.bak")
        shutil.copy2(target_excel, bak_path)
        safe_print(f"[OK] Đã tạo bản sao lưu an toàn tại: {bak_path}")

        wb = openpyxl.load_workbook(target_excel)
        ws = wb["THEO_DOI"]
        
        current_row = old_max_row + 1
        current_stt = last_stt + 1
        
        for r_idx, rec in enumerate(all_new_records, 1):
            ws.cell(current_row, 1, current_stt)
            ws.cell(current_row, 2, rec["so_cv"])
            ws.cell(current_row, 3, parse_date(rec["ngay_cv"]))
            ws.cell(current_row, 4, rec["don_vi_bh"])
            ws.cell(current_row, 5, rec["ten_ban"])
            mst = rec["mst_ban"]
            ws.cell(current_row, 6, mst)
            ws.cell(current_row, 7, rec["noi_dung"])
            ws.cell(current_row, 8, rec["ten_mua"])
            ws.cell(current_row, 9, rec["mst_mua"])
            ws.cell(current_row, 10, None)
            ws.cell(current_row, 11, None)
            ws.cell(current_row, 12, None)
            ws.cell(current_row, 13, 0)
            ws.cell(current_row, 14, 0)
            ws.cell(current_row, 15, rec["nhom_nd"])
            ws.cell(current_row, 16, rec["dia_chi"])
            ws.cell(current_row, 17, parse_date(rec.get("ngay_nhan") or rec["ngay_cv"]))
            ws.cell(current_row, 18, None)
            ws.cell(current_row, 19, rec["ghi_chu"])
            ws.cell(current_row, 20, rec["file"])
            ws.cell(current_row, 21, rec["nam_cv"])
            ws.cell(current_row, 22, None)
            ws.cell(current_row, 23, None)
            ws.cell(current_row, 24, 1)
            mst_dau = 1 if mst and (mst not in seen_msts) else None
            if mst:
                seen_msts.add(mst)
            ws.cell(current_row, 25, mst_dau)
            
            current_row += 1
            current_stt += 1
            
        new_max_row = current_row - 1
        safe_print(f"[OK] Đã chèn {len(all_new_records)} dòng mới (từ dòng {old_max_row + 1} đến {new_max_row}, STT từ {last_stt + 1} đến {current_stt - 1})")

        # Đồng bộ công thức sheet TONG_HOP
        if "TONG_HOP" in wb.sheetnames:
            ws_th = wb["TONG_HOP"]
            pattern = re.compile(rf"(THEO_DOI!\$[A-Z]+\$\d+:\$[A-Z]+\$){old_max_row}\b")
            updated_count = 0
            for r in range(1, ws_th.max_row + 1):
                for c in range(1, ws_th.max_column + 1):
                    val = ws_th.cell(r, c).value
                    if isinstance(val, str) and val.startswith("=") and f"${old_max_row}" in val:
                        new_val = pattern.sub(rf"\g<1>{new_max_row}", val)
                        if new_val != val:
                            ws_th.cell(r, c, new_val)
                            updated_count += 1
            safe_print(f"[OK] Đã cập nhật đồng bộ {updated_count} công thức trên sheet TONG_HOP từ dòng {old_max_row} sang {new_max_row}")

        wb.save(target_excel)
        safe_print(f"[OK] Đã lưu thành công file Excel: {target_excel}")
    else:
        safe_print("[THÔNG BÁO] Không có dòng dữ liệu mới nào cần nạp vào Excel.")

    # Đánh dấu hoàn thành toàn bộ 3 giai đoạn
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(COMPLETION_MARKER, "w", encoding="utf-8") as f:
        f.write(f"HOÀN THÀNH TOÀN BỘ 3 GIAI ĐOẠN LÚC: {now_str}\n")
        f.write(f"- Ưu tiên 1: {len(dedup_p1)} công văn scan 08/2026\n")
        f.write(f"- Ưu tiên 2: {len(records_p2)} công văn digital 07 & 06/2026\n")
        f.write(f"- Giai đoạn 3: Nạp tổng cộng {len(all_new_records)} dòng vào Excel\n")
    safe_print(f"\n[HOÀN TẤT] >>> TOÀN BỘ 3 GIAI ĐOẠN ĐÃ ĐƯỢC THỰC HIỆN THÀNH CÔNG <<<")

    return dedup_p1, records_p2

if __name__ == "__main__":
    excel_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EXCEL_PATH
    run_all(excel_path)
