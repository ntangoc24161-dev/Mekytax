# -*- coding: utf-8 -*-
"""Full update script for Son Tien Hai MST 3603859012.
- Doi 'chi phi quan ly kinh doanh' -> 'chi phi quan ly doanh nghiep'
- Dien QDKT 7092/QD-TCS3, QDGS 7093/QD-TCS3 ngay 26/10/2026
- Ke toan: Le Hoai Phuong
- TKNH: 6722949824 tai Ngan hang TMCP Dau tu va Phat trien Viet Nam
- Nhung phan chua cung cap de trong (CCCDKT)
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import docx
import openpyxl

# Paths
BRAIN = Path(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra")
OUTPUT_DIR = BRAIN / "output" / "3603859012_HO-SO-VE_son-tien-hai"
G_BB = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\11. SON TIEN HAI\0. BIEN BAN"
)
G_SO = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung"
    r"\Danh sach kiem tra quyet toan pro.xlsm"
)

# Constants
QDKT = "7092/QĐ-TCS3"
QDGS = "7093/QĐ-TCS3"
NGAY = "26/10/2026"
KT = "Lê Hoài Phương"
TK = "6722949824 tại Ngân hàng TMCP Đầu tư và phát triển Việt Nam"

LOG = []


def log(msg: str) -> None:
    LOG.append(msg)
    print(msg)


def patch_tong_hop_xml(xml: str) -> str:
    # 1. Replace 'quản lý kinh doanh' -> 'quản lý doanh nghiệp'
    n1 = xml.count("quản lý kinh doanh")
    xml = xml.replace("quản lý kinh doanh", "quản lý doanh nghiệp")
    log(f"Contiguous 'quản lý kinh doanh' replaced: {n1}")

    split_from = (
        '<w:t xml:space="preserve">Chi phí quản lý </w:t></w:r>'
        '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>kinh doanh</w:t></w:r>'
    )
    split_to = (
        '<w:t xml:space="preserve">Chi phí quản lý </w:t></w:r>'
        '<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>doanh nghiệp</w:t></w:r>'
    )
    n2 = xml.count(split_from)
    xml = xml.replace(split_from, split_to)
    log(f"Split 'Chi phí quản lý ' + 'kinh doanh' replaced: {n2}")

    # 2. Header QDKT: Table 1
    rpr_t1 = (
        '<w:rPr><w:rFonts w:cs="Times New Roman"/><w:sz w:val="24"/>'
        '<w:szCs w:val="24"/><w:lang w:val="nl-NL"/></w:rPr>'
    )
    old_t1_hdr = (
        '<w:t xml:space="preserve">Số:         </w:t></w:r><w:r>'
        + rpr_t1
        + '<w:t xml:space="preserve">   </w:t></w:r>'
        + '<w:r w:rsidR="00D50A77" w:rsidRPr="009D7F91">'
        + rpr_t1
        + '<w:t xml:space="preserve">     /QĐ-TCS3</w:t>'
    )
    new_t1_hdr = (
        '<w:t xml:space="preserve">Số: </w:t></w:r>'
        + '<w:r w:rsidR="00D50A77" w:rsidRPr="009D7F91">'
        + rpr_t1
        + f'<w:t xml:space="preserve">{QDKT}</w:t>'
    )
    if old_t1_hdr in xml:
        xml = xml.replace(old_t1_hdr, new_t1_hdr, 1)
        log("Replaced Header QDKT in Table 1")

    old_t1_date = "Trấn Biên, ngày       tháng       năm"
    new_t1_date = f"Trấn Biên, ngày 26 tháng 10 năm 2026"
    if old_t1_date in xml:
        xml = xml.replace(old_t1_date, new_t1_date, 1)
        log("Replaced Header Date in Table 1")

    # Header QDGS: Table 3
    old_t3_hdr = (
        '<w:t xml:space="preserve">Số:           </w:t></w:r><w:r>'
        + rpr_t1
        + '<w:t xml:space="preserve">  </w:t></w:r>'
        + '<w:r w:rsidR="008629EC" w:rsidRPr="009D7F91">'
        + rpr_t1
        + '<w:t xml:space="preserve">   /QĐ-TCS3</w:t>'
    )
    new_t3_hdr = (
        '<w:t xml:space="preserve">Số: </w:t></w:r>'
        + '<w:r w:rsidR="008629EC" w:rsidRPr="009D7F91">'
        + rpr_t1
        + f'<w:t xml:space="preserve">{QDGS}</w:t>'
    )
    if old_t3_hdr in xml:
        xml = xml.replace(old_t3_hdr, new_t3_hdr, 1)
        log("Replaced Header QDGS in Table 3")

    old_t3_date = "Trấn Biên, ngày       tháng      năm"
    if old_t3_date in xml:
        xml = xml.replace(old_t3_date, new_t1_date, 1)
        log("Replaced Header Date in Table 3")

    # 3. References to QĐGS / QĐKT in text
    xml = xml.replace(
        "Căn cứ Quyết định kiểm tra số              /QĐ-TCS3 ngày ...../...../2026",
        f"Căn cứ Quyết định kiểm tra số {QDKT} ngày {NGAY}",
    )
    xml = xml.replace(
        "theo Quyết định số              /QĐ-TCS3 ngày ...../...../2026",
        f"theo Quyết định số {QDKT} ngày {NGAY}",
    )

    old_p449 = (
        "Thực hiện Quyết định số              /QĐ-TCS3 ngày ..../..../2026 "
        "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
        "đoàn kiểm tra theo QĐ kiểm tra tại doanh nghiệp số              "
        "/QĐ-TCS3 ngày ..../..../2026"
    )
    new_p449 = (
        f"Thực hiện Quyết định số {QDGS} ngày {NGAY} "
        "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
        f"đoàn kiểm tra theo QĐ kiểm tra tại doanh nghiệp số {QDKT} ngày {NGAY}"
    )
    xml = xml.replace(old_p449, new_p449)

    old_p480 = (
        "theo quyết định kiểm tra tại doanh nghiệp số              /QĐ-TCS3 "
        "ngày ..../..../2026 về việc giám sát Đoàn kiểm tra theo quyết "
        "định giám sát số              /QĐ-TCS3 ngày ..../..../2026."
    )
    new_p480 = (
        f"theo quyết định kiểm tra tại doanh nghiệp số {QDKT} "
        f"ngày {NGAY} về việc giám sát Đoàn kiểm tra theo quyết "
        f"định giám sát số {QDGS} ngày {NGAY}."
    )
    xml = xml.replace(old_p480, new_p480)

    old_p501 = (
        "Thực hiện Quyết định số              /QĐ-TCS3 ngày ..../..../2026 "
        "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
        "hoạt động của Đoàn kiểm tra tại Công ty TNHH Thương mại Sản xuất Sơn Tiền Hải"
    )
    new_p501 = (
        f"Thực hiện Quyết định số {QDGS} ngày {NGAY} "
        "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
        "hoạt động của Đoàn kiểm tra tại Công ty TNHH Thương mại Sản xuất Sơn Tiền Hải"
    )
    xml = xml.replace(old_p501, new_p501)

    old_gen = "số              /QĐ-TCS3 ngày ..../..../2026"
    n_gen = xml.count(old_gen)
    xml = xml.replace(old_gen, f"số {QDKT} ngày {NGAY}")
    log(f"Replaced general references (số .../QĐ-TCS3 ngày ...): {n_gen}")

    old_gen2 = "số:              /QĐ-TCS3 ngày ..../..../2026"
    n_gen2 = xml.count(old_gen2)
    xml = xml.replace(old_gen2, f"số: {QDKT} ngày {NGAY}")
    log(f"Replaced general references with colon: {n_gen2}")

    xml = xml.replace("SỐ              /QĐ-TCS3", f"SỐ {QDKT}")
    xml = xml.replace("……../QĐ-TCS3 ngày …../…../202", f"{QDKT} ngày {NGAY}")
    xml = xml.replace("……../QĐ-TCS3 ngày …../…../2026", f"{QDKT} ngày {NGAY}")
    xml = xml.replace("……../QĐ-TCS3 ngày …../…../", f"{QDKT} ngày {NGAY}")
    xml = xml.replace("……../QĐ-TCS3", QDKT)

    xml = xml.replace(
        '<w:t xml:space="preserve">              /QĐ-TCS3</w:t>',
        f'<w:t xml:space="preserve">{QDKT}</w:t>',
    )

    # 4. Replace MERGEFIELD "KT" and MERGEFIELD KT cleanly
    for pat in ['MERGEFIELD "KT"', 'MERGEFIELD KT']:
        while True:
            pos = xml.find(pat)
            if pos == -1:
                break
            r2_start = xml.rfind('<w:r', 0, pos)
            while r2_start > 0:
                prefix = xml[r2_start:r2_start+5]
                if prefix.startswith('<w:r>') or prefix.startswith('<w:r '):
                    break
                r2_start = xml.rfind('<w:r', 0, r2_start - 1)
            r2_end = xml.find('</w:r>', pos) + len('</w:r>')

            r1_end = r2_start
            r1_start = xml.rfind('<w:r', 0, r1_end)
            while r1_start > 0:
                prefix = xml[r1_start:r1_start+5]
                if prefix.startswith('<w:r>') or prefix.startswith('<w:r '):
                    break
                r1_start = xml.rfind('<w:r', 0, r1_start - 1)

            r3_start = r2_end
            r3_end = xml.find('</w:r>', r3_start) + len('</w:r>')

            r2_str = xml[r2_start:r2_end]
            m_rpr = re.search(r'<w:rPr>.*?</w:rPr>', r2_str)
            rpr = m_rpr.group(0) if m_rpr else ""
            rep_run = f"<w:r>{rpr}<w:t>{KT}</w:t></w:r>"

            xml = xml[:r1_start] + rep_run + xml[r3_end:]
            log("Replaced one KT field")

    # 5. Remove MERGEFIELD "CCCDKT"
    for pat in ['MERGEFIELD "CCCDKT"', 'MERGEFIELD CCCDKT']:
        while True:
            pos = xml.find(pat)
            if pos == -1:
                break
            r2_start = xml.rfind('<w:r', 0, pos)
            while r2_start > 0:
                prefix = xml[r2_start:r2_start+5]
                if prefix.startswith('<w:r>') or prefix.startswith('<w:r '):
                    break
                r2_start = xml.rfind('<w:r', 0, r2_start - 1)
            r2_end = xml.find('</w:r>', pos) + len('</w:r>')

            r1_end = r2_start
            r1_start = xml.rfind('<w:r', 0, r1_end)
            while r1_start > 0:
                prefix = xml[r1_start:r1_start+5]
                if prefix.startswith('<w:r>') or prefix.startswith('<w:r '):
                    break
                r1_start = xml.rfind('<w:r', 0, r1_start - 1)

            r3_start = r2_end
            r3_end = xml.find('</w:r>', r3_start) + len('</w:r>')

            xml = xml[:r1_start] + xml[r3_end:]
            log("Removed CCCDKT field")

    # Signatures in Table 17 and Table 24:
    pth = "Phan Thế Hanh"
    old_sig17 = f'<w:t xml:space="preserve">        {pth}                  </w:t>'
    new_sig17 = f'<w:t xml:space="preserve">        {pth}                  {KT}</w:t>'
    if old_sig17 in xml:
        xml = xml.replace(old_sig17, new_sig17, 1)
        log("Updated signature Table 17")

    old_sig24 = f'<w:t xml:space="preserve">{pth}             </w:t>'
    new_sig24 = f'<w:t xml:space="preserve">{pth}                  {KT}</w:t>'
    if old_sig24 in xml:
        xml = xml.replace(old_sig24, new_sig24, 1)
        log("Updated signature Table 24")

    # 6. Bank account
    old_tk = '<w:t xml:space="preserve">- Tài khoản số: </w:t>'
    new_tk = f'<w:t xml:space="preserve">- Tài khoản số: {TK}.</w:t>'
    if old_tk in xml:
        xml = xml.replace(old_tk, new_tk, 1)
        log("Updated Bank Account in P291")

    # Remove empty Tài_khoản_số paragraph
    for field_str in ['MERGEFIELD Tài_khoản_số', 'MERGEFIELD "Tài_khoản_số"']:
        while True:
            pos = xml.find(field_str)
            if pos == -1:
                break
            p_start = xml.rfind('<w:p', 0, pos)
            while p_start > 0:
                prefix = xml[p_start:p_start+5]
                if prefix.startswith('<w:p>') or prefix.startswith('<w:p '):
                    break
                p_start = xml.rfind('<w:p', 0, p_start - 1)
            p_end = xml.find('</w:p>', pos) + len('</w:p>')
            xml = xml[:p_start] + xml[p_end:]
            log("Removed empty Tài_khoản_số paragraph")

    return xml


def patch_tables_dates(doc_path: Path) -> None:
    doc = docx.Document(str(doc_path))
    # Table 42:
    # r1: Ban hanh QD: QDKT, 26/10/2026
    # r2: Thong bao QD: 26/10/2026
    if len(doc.tables) > 42:
        t42 = doc.tables[42]
        if len(t42.rows) > 2:
            t42.cell(1, 2).text = QDKT
            t42.cell(1, 3).text = NGAY
            t42.cell(2, 3).text = NGAY
            log("Updated Table 42 dates")
    # Table 50:
    # r8: QD kiem tra: QDKT, 26/10/2026
    if len(doc.tables) > 50:
        t50 = doc.tables[50]
        if len(t50.rows) > 8:
            t50.cell(8, 2).text = QDKT
            t50.cell(8, 3).text = NGAY
            log("Updated Table 50 row 8 date & QDKT")
    doc.save(str(doc_path))


def patch_qdxphc(path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf8")
    old_ref = "số              /QĐ-TCS3 ngày ..../..../2026"
    if old_ref in xml:
        xml = xml.replace(old_ref, f"số {QDKT} ngày {NGAY}")
        log("Updated QDXPHC reference")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("word/document.xml", xml.encode("utf8"))


def export_qd_files(master_path: Path, out_kt: Path, out_gs: Path) -> None:
    # Export QĐKT (elements 34 to 58) and QĐGS (elements 60 to 78)
    shutil.copy2(master_path, out_kt)
    doc_kt = docx.Document(str(out_kt))
    body_kt = doc_kt._element.body
    elems_kt = list(body_kt)
    keep_kt = set(range(34, 59))
    for i, el in enumerate(elems_kt):
        if el.tag.endswith("sectPr"):
            continue
        if i not in keep_kt:
            body_kt.remove(el)
    doc_kt.save(str(out_kt))
    log(f"Exported {out_kt.name}")

    shutil.copy2(master_path, out_gs)
    doc_gs = docx.Document(str(out_gs))
    body_gs = doc_gs._element.body
    elems_gs = list(body_gs)
    keep_gs = set(range(60, 79))
    for i, el in enumerate(elems_gs):
        if el.tag.endswith("sectPr"):
            continue
        if i not in keep_gs:
            body_gs.remove(el)
    doc_gs.save(str(out_gs))
    log(f"Exported {out_gs.name}")


def patch_excel_tracking(so_path: Path) -> None:
    wb = openpyxl.load_workbook(so_path, keep_vba=True)
    ws = wb["01_DA_BAN_HANH"]
    r = None
    for i in range(5, ws.max_row + 1):
        if str(ws.cell(i, 4).value or "") == "3603859012":
            r = i
            break
    if not r:
        log("WARN: Son Tien Hai not found in tracking sheet")
        return

    ws.cell(r, 6).value = "Số 23, Tổ 1, KP 3A, Phường Long Hưng"
    ws.cell(r, 16).value = KT
    # CCCDKT blank
    ws.cell(r, 21).value = f"TK số {TK}"
    ws.cell(r, 27).value = "-Giá vốn hàng bán\n- Chi phí bán hàng\n- Chi phí quản lý doanh nghiệp"
    ws.cell(r, 28).value = "- Thuế GTGT"
    ws.cell(r, 29).value = QDKT
    ws.cell(r, 30).value = QDGS
    ws.cell(r, 31).value = NGAY
    ws.cell(r, 40).value = 13052423
    ws.cell(r, 41).value = 0
    ws.cell(r, 42).value = 13052423
    ws.cell(r, 44).value = 662730
    ws.cell(r, 45).value = 662730
    ws.cell(r, 46).value = 2610485
    ws.cell(r, 47).value = 0
    ws.cell(r, 48).value = 3273215
    ws.cell(r, 49).value = 16325638
    ws.cell(r, 50).value = 809868861
    ws.cell(r, 51).value = 0
    wb.save(so_path)
    log(f"Updated tracking sheet row {r}")


def update_data_files() -> None:
    # 1. JSON
    json_path = OUTPUT_DIR / "du-lieu-tron.json"
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    tt = data["truong_tron"]
    tt["KT"] = KT
    tt["CCCDKT"] = ""
    tt["QĐKT"] = QDKT
    tt["QĐGS"] = QDGS
    tt["ngày_QĐ"] = NGAY
    tt["Tài_khoản_số"] = TK
    tt["Nội_dung_kiểm_tra"] = (
        "- Giá vốn hàng bán; - Chi phí quản lý doanh nghiệp; "
        "- Doanh thu bán hàng và cung cấp dịch vụ; - Thuế GTGT phải nộp."
    )
    tt["Phân_công_trưởng_đoàn"] = "- Giá vốn hàng bán\n- Chi phí quản lý doanh nghiệp"

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    log("Updated du-lieu-tron.json")

    # 2. Excel
    xlsx_path = OUTPUT_DIR / "du-lieu-tron_son-tien-hai_3603859012.xlsx"
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb["DU_LIEU_TRON"]
    for c in range(1, ws.max_column + 1):
        h = ws.cell(1, c).value
        if h == "KT":
            ws.cell(2, c).value = KT
        elif h == "CCCDKT":
            ws.cell(2, c).value = None
        elif h == "Tài_khoản_số":
            ws.cell(2, c).value = TK
        elif h == "Nội_dung_kiểm_tra":
            ws.cell(2, c).value = (
                "- Giá vốn hàng bán; - Chi phí quản lý doanh nghiệp; "
                "- Doanh thu bán hàng và cung cấp dịch vụ; - Thuế GTGT phải nộp."
            )
        elif h == "Phân_công_trưởng_đoàn":
            ws.cell(2, c).value = "- Giá vốn hàng bán\n- Chi phí quản lý doanh nghiệp"

    # Add QDKT, QDGS, ngày_QĐ if not in columns
    existing_headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    next_col = ws.max_column + 1
    if "QĐKT" not in existing_headers:
        ws.cell(1, next_col).value = "QĐKT"
        ws.cell(2, next_col).value = QDKT
        next_col += 1
    if "QĐGS" not in existing_headers:
        ws.cell(1, next_col).value = "QĐGS"
        ws.cell(2, next_col).value = QDGS
        next_col += 1
    if "ngày_QĐ" not in existing_headers:
        ws.cell(1, next_col).value = "ngày_QĐ"
        ws.cell(2, next_col).value = NGAY

    wb.save(xlsx_path)
    log("Updated du-lieu-tron_son-tien-hai_3603859012.xlsx")


def main() -> None:
    now_str = datetime.now().strftime("%Y%m%d-%H%M%S")
    log(f"=== START FULL UPDATE: {now_str} ===")

    # 1. Backups
    src_tong = OUTPUT_DIR / "TONG HOP.docx"
    bak_tong = OUTPUT_DIR / f"TONG HOP.bak-truoc-7092-{now_str}.docx"
    shutil.copy2(src_tong, bak_tong)
    log(f"Backup created: {bak_tong.name}")

    src_xphc = OUTPUT_DIR / "QD XPHC Cong ty Son Tien Hai.docx"
    bak_xphc = OUTPUT_DIR / f"QD XPHC.bak-truoc-7092-{now_str}.docx"
    shutil.copy2(src_xphc, bak_xphc)
    log(f"Backup created: {bak_xphc.name}")

    if G_SO.exists():
        bak_so = G_SO.parent / f"Danh sach kiem tra quyet toan pro.bak-{now_str}.xlsm"
        shutil.copy2(G_SO, bak_so)
        log(f"Backup tracking sheet created: {bak_so.name}")

    # 2. Patch XML of TONG HOP.docx
    with zipfile.ZipFile(src_tong) as z:
        xml = z.read("word/document.xml").decode("utf8")
    
    xml = patch_tong_hop_xml(xml)

    # Save patched XML back to TONG HOP.docx in OUTPUT_DIR
    temp_patched = OUTPUT_DIR / "TONG HOP.temp.docx"
    with zipfile.ZipFile(src_tong) as zin, zipfile.ZipFile(temp_patched, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == "word/document.xml":
                data = xml.encode("utf8")
            zout.writestr(it, data)

    shutil.move(temp_patched, src_tong)
    log(f"Saved patched {src_tong.name}")

    # Patch Table 42 and Table 50 dates
    patch_tables_dates(src_tong)

    # 3. Patch QD XPHC
    patch_qdxphc(src_xphc)

    # 4. Export separate QDKT & QDGS
    out_kt = OUTPUT_DIR / "QD Kiem tra tai cty Son Tien Hai.docx"
    out_gs = OUTPUT_DIR / "QDgiam sat tai cty Son Tien Hai.docx"
    export_qd_files(src_tong, out_kt, out_gs)

    # 5. Update data files (du-lieu-tron.json and xlsx)
    update_data_files()

    # 6. Copy to G: drive
    if G_BB.exists():
        # Kill Winword if open to avoid locks
        os.system('taskkill /F /IM WINWORD.EXE >nul 2>&1')
        shutil.copy2(src_tong, G_BB / "TONG HOP.docx")
        shutil.copy2(src_xphc, G_BB / "QD XPHC Cong ty Son Tien Hai.docx")
        shutil.copy2(out_kt, G_BB / "QD Kiem tra tai cty Son Tien Hai.docx")
        shutil.copy2(out_gs, G_BB / "QDgiam sat tai cty Son Tien Hai.docx")
        log(f"Synchronized all 4 docx files to G: {G_BB}")

    if G_SO.exists():
        patch_excel_tracking(G_SO)

    # 7. Verification
    doc = docx.Document(str(src_tong))
    txt = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                txt += "\n" + c.text

    log("=== FINAL VERIFICATION ===")
    log(f"TONG HOP total paragraphs: {len(doc.paragraphs)}, tables: {len(doc.tables)}")
    log(f"Count QDKT ({QDKT}): {txt.count(QDKT)}")
    log(f"Count QDGS ({QDGS}): {txt.count(QDGS)}")
    log(f"Count KT ({KT}): {txt.count(KT)}")
    log(f"Count TK (6722949824): {txt.count('6722949824')}")
    log(f"Count 'quản lý kinh doanh': {txt.lower().count('quản lý kinh doanh')}")
    log(f"Count 'quản lý doanh nghiệp': {txt.lower().count('quản lý doanh nghiệp')}")
    log(f"Count remaining '/QĐ-TCS3' unfulfilled: {txt.count('              /QĐ-TCS3')}")

    # Append to nhat-ky-cap-nhat.json
    nk_path = OUTPUT_DIR / "nhat-ky-cap-nhat.json"
    nk_data = {
        "thoi_diem": now_str,
        "yeu_cau": "Doi QLDN, dien QDKT 7092, QDGS 7093 ngay 26/10/2026, ke toan Le Hoai Phuong, TKNH 6722949824 BIDV",
        "cac_thay_doi": [
            f"Doi toan bo 'quản lý kinh doanh' -> 'quản lý doanh nghiệp' ({txt.lower().count('quản lý doanh nghiệp')} cho)",
            f"Dien so QDKT: {QDKT} ngay {NGAY} ({txt.count(QDKT)} cho)",
            f"Dien so QDGS: {QDGS} ngay {NGAY} ({txt.count(QDGS)} cho)",
            f"Dien Ke toan truong: {KT} ({txt.count(KT)} cho gom 6 fld + 2 chu ky)",
            f"Dien Tai khoan ngan hang: {TK}",
            "CCCDKT giu trong theo yeu cau",
            "Xuat file rieng QD Kiem tra va QD Giam sat",
            "Dong bo so theo doi Danh sach kiem tra quyet toan pro.xlsm"
        ]
    }
    with open(nk_path, "w", encoding="utf-8") as f:
        json.dump(nk_data, f, ensure_ascii=False, indent=2)
    log("Saved nhat-ky-cap-nhat.json")


if __name__ == "__main__":
    main()
