# -*- coding: utf-8 -*-
"""
parse_months_1_to_4.py - Doc va trich xuat toan bo 53 tep trong cac thu muc 2026-01, 2026-02, 2026-03, 2026-04
Luu ket qua vao records_months_1_to_4.json va doi chieu voi So theo doi Excel
"""
import os
import re
import sys
import json
from pathlib import Path
import fitz  # PyMuPDF
import openpyxl

BASE_DIR = Path(r"D:\01. Cong tac kiem tra\HO SO CONG VAN CANH BAO\2026")
EXCEL_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
OUT_JSON = Path(r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo\skills\theo-doi-cv-canh-bao-rui-ro\scripts\records_months_1_to_4.json")

def clean_company_name(s):
    if not s:
        return ""
    s = re.sub(r"^(?:v/v|vv|v/v\s+cảnh\s+báo|cảnh\s+báo|cvr|cbrr|cv|công\s+văn)[\s:.-]*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:doanh\s+nghiệp\s+có\s+dấu\s+hiệu\s+rủi\s+ro\s+cao\s+về\s+thuế[^\w]*)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^(?:đối\s+với\s+)", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\s:;.-]+(?:mst|mã\s+số\s+thuế).*$", "", s, flags=re.IGNORECASE)
    return s.strip(" ,.:;-\"'\n\r\t")

def normalize_mst(mst_raw):
    if not mst_raw:
        return ""
    clean = re.sub(r"[^\d\-]", "", str(mst_raw)).strip()
    return clean

def extract_pdf_info(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    doc.close()
    return full_text

def parse_file_group(group_prefix, files_in_group, month_dir):
    # Determine SoCV, DonVi, NgayCV, MST, Company Name, NoiDung, etc.
    # Group prefix is usually [prefix] or [so_den]_[so_ky_hieu]
    so_cv = ""
    ngay_cv = ""
    don_vi_bh = ""
    ten_ban = ""
    mst_ban = ""
    ten_mua = ""
    mst_mua = ""
    noi_dung = "Cảnh báo rủi ro về thuế và hoá đơn"
    
    # Extract text from all PDFs in the group
    all_texts = []
    main_pdf = ""
    for f in files_in_group:
        p = month_dir / f
        txt = extract_pdf_info(p)
        all_texts.append((f, txt))
        if not main_pdf and not any(x in f.lower() for x in ["bang ke", "bảng kê", "hoa don", "hoá đơn", "danh sach"]):
            main_pdf = f

    if not main_pdf and files_in_group:
        main_pdf = files_in_group[0]

    # Try parsing text from main pdf
    combined_text = "\n".join(t[1] for t in all_texts)
    
    # Extract SoCV
    m_socv = re.search(r"Số\s*:\s*([\d]+/[A-Z0-9\-]+)", combined_text, re.IGNORECASE)
    if m_socv:
        so_cv = m_socv.group(1).strip()
    else:
        # Fallback from prefix
        m_pref = re.search(r"\[([\d]+[_\/][A-Z0-9\-]+)\]", group_prefix)
        if m_pref:
            so_cv = m_pref.group(1).replace("_", "/")
        else:
            m_num = re.search(r"\[(\d+)\]", group_prefix)
            if m_num:
                so_cv = m_num.group(1)

    # Extract NgayCV
    m_ngay = re.search(r"ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(202\d)", combined_text, re.IGNORECASE)
    if m_ngay:
        d, m, y = m_ngay.groups()
        ngay_cv = f"{int(d):02d}/{int(m):02d}/{y}"
    else:
        m_ngay2 = re.search(r"(\d{1,2}/\d{1,2}/202\d)", combined_text)
        if m_ngay2:
            ngay_cv = m_ngay2.group(1)

    # Extract DonViBH
    m_dv = re.search(r"(THUẾ\s+CƠ\s+SỞ\s+\d+|CỤC\s+THUẾ\s+[\w\s]+|THUẾ\s+THÀNH\s+PHỐ\s+ĐỒNG\s+NAI)", combined_text, re.IGNORECASE)
    if m_dv:
        don_vi_bh = m_dv.group(1).strip().title()

    # Extract MST
    msts = re.findall(r"(?:MST|mã\s+số\s+thuế|MÃ\s+SỐ\s+THUẾ)[\s:.]*([0-9]{10}(?:-[0-9]{3})?)", combined_text, re.IGNORECASE)
    if not msts:
        # General 10-13 digit regex
        msts = re.findall(r"\b([0-9]{10}(?:-[0-9]{3})?)\b", combined_text)
    
    # Extract Company Name
    m_cty = re.search(r"(?:Công\s+ty|Cty|DOANH\s+NGHIỆP|DNTN|HTX)\s+([^\n\r,\.]+)", combined_text, re.IGNORECASE)
    if m_cty:
        ten_ban = clean_company_name(m_cty.group(0))

    return {
        "so_cv": so_cv,
        "ngay_cv": ngay_cv,
        "don_vi_bh": don_vi_bh,
        "ten_ban": ten_ban,
        "mst_ban": msts[0] if msts else "",
        "ten_mua": ten_mua,
        "mst_mua": mst_mua,
        "noi_dung": noi_dung,
        "file": main_pdf,
        "all_msts": list(dict.fromkeys(msts)),
        "files": files_in_group,
        "text_preview": combined_text[:300].replace("\n", " ")
    }

def main():
    months = ["2026-01", "2026-02", "2026-03", "2026-04"]
    all_records = []
    
    for m in months:
        m_dir = BASE_DIR / m
        if not m_dir.exists():
            continue
        files = [f for f in os.listdir(m_dir) if f.endswith(".pdf")]
        
        # Group by prefix [xxx]
        groups = {}
        for f in files:
            m_pref = re.match(r"^(\[.*?\])", f)
            pref = m_pref.group(1) if m_pref else f
            groups.setdefault(pref, []).append(f)
            
        print(f"Tháng {m}: {len(files)} tệp -> {len(groups)} nhóm")
        for pref, gfiles in groups.items():
            rec = parse_file_group(pref, gfiles, m_dir)
            rec["month"] = m
            all_records.append(rec)
            
    print(f"\nTổng số nhóm công văn trích xuất từ 4 tháng: {len(all_records)}")
    with open(OUT_JSON, "w", encoding="utf-8") as fp:
        json.dump(all_records, fp, ensure_ascii=False, indent=2)

if __name__ == "__main__":
    main()
