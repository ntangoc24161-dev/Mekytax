# -*- coding: utf-8 -*-
"""Dung lai file 04 tu mau Word (header 3 dong, tblGrid dung cot).
Chi dien noi dung huong 1: phan cong quy che + viec 7 cua An. Khong goi table.columns.
"""
import os, sys, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from cb13_data import NHAN_SU
from cb13_noidung import NHIEM_VU

VAULT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
LAPTOP = (r"H:\Other computers\My Laptop\2. CONG TAC KHAC\4. CONG TAC DANG"
          r"\A PHUONG\2026\06_Tieu_chi_danh_gia\Du_thao_CB13_QuyIII_2026")
MAU = r"D:\SECOND BRAIN\javis-os\server\.staging\04 Quyet dinh phan cong va danh muc san pham qui 3 2026.docx"
OUT_NAME = "04_Quyet_dinh_phan_cong_va_danh_muc_san_pham.docx"


def set_para_text(p, text):
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def set_cell_text(cell, text):
    paras = cell.paragraphs
    if not paras:
        return
    set_para_text(paras[0], text if text is not None else "")
    for p in paras[1:]:
        set_para_text(p, "")


def restore_grid_from_row(table, row_idx=1):
    """Ghi lai w:tblGrid theo tcW cua 1 hang, khong dung table.columns."""
    if row_idx >= len(table.rows):
        row_idx = 0
    widths = []
    seen = set()
    for c in table.rows[row_idx].cells:
        tc = c._tc
        if id(tc) in seen:
            continue
        seen.add(id(tc))
        tcPr = tc.find(qn("w:tcPr"))
        tcW = tcPr.find(qn("w:tcW")) if tcPr is not None else None
        widths.append(tcW.get(qn("w:w")) if tcW is not None else "0")
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tbl.insert(0, grid)
    for child in list(grid):
        grid.remove(child)
    for w in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(w))
        grid.append(col)
    return widths


def patch_04_from_mau(path):
    d = Document(path)
    for p in d.paragraphs:
        t = p.text.strip()
        if t.startswith("- Xét đề nghị của đồng chí Phó Bí thư"):
            set_para_text(p, "- Xét đề nghị của đồng chí Phó Bí thư Chi bộ phụ trách công tác tuyên giáo,")
        if "Trợ lý Bí thư có trách nhiệm" in p.text:
            set_para_text(
                p,
                p.text.replace(
                    "đồng chí Trợ lý Bí thư có trách nhiệm",
                    "đồng chí phụ trách công tác đảng vụ có trách nhiệm",
                ),
            )
        for n in NHAN_SU:
            prefix = "%d. Đồng chí %s" % (n["stt"], n["ten_hoa_thuong"])
            if not t.startswith(prefix):
                continue
            if ":" in t:
                set_para_text(
                    p,
                    "%d. Đồng chí %s - %s: %s"
                    % (n["stt"], n["ten_hoa_thuong"], n["cv_dang"], n["linh_vuc"]),
                )
            else:
                set_para_text(
                    p,
                    "%d. Đồng chí %s - %s" % (n["stt"], n["ten_hoa_thuong"], n["cv_dang"]),
                )

    nv_tables = []
    for t in d.tables:
        hdr = " ".join(c.text for c in t.rows[0].cells)
        if "Nhiệm vụ" in hdr and "Hệ số" in hdr:
            nv_tables.append(t)
            restore_grid_from_row(t, 1)
    if len(nv_tables) != 7:
        raise SystemExit("so bang nhiem vu = %d" % len(nv_tables))

    for n, t in zip(NHAN_SU, nv_tables):
        nvs = NHIEM_VU[n["ten"]]
        body = t.rows[1:]
        data_rows = body[:-1] if body and "TỔNG" in body[-1].cells[1].text else body
        if len(data_rows) < len(nvs):
            raise SystemExit("%s thieu hang: co %d can %d" % (n["ten"], len(data_rows), len(nvs)))
        for i, nv in enumerate(nvs):
            noi_dung, loai, chi_tieu, tien_do, san_pham, phoi_hop, cap_trinh, thoi_han, sl, hs = nv
            diem_qd = round(sl * hs * 100)
            row = data_rows[i]
            set_cell_text(row.cells[0], str(i + 1))
            set_cell_text(row.cells[1], noi_dung)
            set_cell_text(row.cells[2], san_pham)
            set_cell_text(row.cells[3], "%d" % sl)
            set_cell_text(row.cells[4], ("%.1f" % hs).replace(".", ","))
            set_cell_text(row.cells[5], "%d" % diem_qd)
            set_cell_text(row.cells[6], thoi_han)
        tong = sum(round(nv[8] * nv[9] * 100) for nv in nvs)
        if body and "TỔNG" in body[-1].cells[1].text:
            set_cell_text(body[-1].cells[5], "%d" % tong)
    d.save(path)
    print("saved", path)


def drop_empty_ghi_chu(path):
    """Xoa khoi Ghi chu (da lam rong) nam giua bang nhiem vu va bang chu ky."""
    d = Document(path)
    body = d.element.body
    tbls = list(body.findall(qn("w:tbl")))
    if len(tbls) < 3:
        print("KH skip", os.path.basename(path), "tbls", len(tbls))
        return
    start = tbls[1]
    end = tbls[2]
    removed = 0
    el = start.getnext()
    while el is not None and el is not end:
        nxt = el.getnext()
        if el.tag == qn("w:p"):
            body.remove(el)
            removed += 1
        el = nxt
    if removed:
        d.save(path)
    print("KH khoi giua bang removed", removed, os.path.basename(path))


def restore_kh_grid(path):
    d = Document(path)
    t = d.tables[1]
    ws = restore_grid_from_row(t, 1)
    d.save(path)
    print("KH grid", os.path.basename(path), ws)


if __name__ == "__main__":
    src = os.path.join(VAULT, OUT_NAME)
    shutil.copy2(MAU, src)
    print("copied mau ->", src, "size", os.path.getsize(src))
    patch_04_from_mau(src)
    kh_dir = os.path.join(VAULT, "05_Ke_hoach_ca_nhan")
    for fn in sorted(os.listdir(kh_dir)):
        if fn.endswith(".docx"):
            p = os.path.join(kh_dir, fn)
            drop_empty_ghi_chu(p)
            restore_kh_grid(p)
    shutil.copy2(src, os.path.join(LAPTOP, OUT_NAME))
    kh_lt = os.path.join(LAPTOP, "05_Ke_hoach_ca_nhan")
    os.makedirs(kh_lt, exist_ok=True)
    for fn in sorted(os.listdir(kh_dir)):
        if fn.endswith(".docx"):
            shutil.copy2(os.path.join(kh_dir, fn), os.path.join(kh_lt, fn))
    print("copied to laptop")
