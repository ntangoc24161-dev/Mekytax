# -*- coding: utf-8 -*-
"""Tiep tuc tien trinh: xoa dong 524 ngoai 524.xlsx, gop trung, khong lam lai toan so.

Giu nguyen CBRR / HDRR khac / CV cong an khac / F1.
Chi xu ly cum danh sach 524/637 (1770, 1798, 1798.0, BKL, HDRR 'DS 524 DN').
"""
from __future__ import annotations

import re
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lam_sach_so_theo_doi import (  # noqa: E402
    SO_THEO_DOI,
    SYNC_COPIES,
    backup_current,
    build_f0_sheet,
    clean_text,
    is_valid_mst,
    log,
    norm_mst,
    norm_socv,
    sort_records,
    stamp_flags,
    xldate,
)

F524 = SO_THEO_DOI.parent / "INPUT" / "524.xlsx"


def is_524_list_row(socv: str, ghi: str, nd: str) -> bool:
    su = (socv or "").upper().replace(" ", "")
    gl = ((ghi or "") + " | " + (nd or "")).upper()
    src = (ghi or "").upper()
    if "NGUỒN: CBRR" in src or "NGUON: CBRR" in src:
        return False
    if su in {
        "1798",
        "1798.0",
        "1798/TCT-TTKT",
        "1770/CAT-CSKT",
        "05/BKL-CSKT",
        "05/BKL-CSKT(Đ4)",
        "05/BKL-CSKT(D4)",
    }:
        return True
    if "1798/TCT" in su or "1770/CAT" in su or "05/BKL" in su:
        return True
    if "DS 524 DN" in gl or "CV 1798" in gl:
        return True
    if "QUY MÔ ĐỢT: 524" in gl or "QUY MO DOT: 524" in gl:
        return True
    if "QUY MÔ ĐỢT: 637" in gl or "QUY MO DOT: 637" in gl:
        return True
    if "524 CÔNG TY" in gl or "524 CONG TY" in gl:
        return True
    return False


def parse_524_cv(blob: str):
    s = clean_text(blob)
    socv = ""
    m = re.search(r"(\d+[A-Z]?/[A-Za-z0-9Đđ.\-()]+)", s)
    if m:
        socv = m.group(1).rstrip(".,;")
    ngay = None
    m = re.search(r"ngày\s+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", s, re.I)
    if m:
        ngay = xldate(m.group(1))
    donvi = "Công an tỉnh Phú Thọ"
    m = re.search(r"của\s+(.+)$", s, re.I)
    if m:
        donvi = clean_text(m.group(1))
        if donvi.lower().startswith("công an") and "tỉnh" not in donvi.lower() and "phú thọ" in donvi.lower():
            donvi = "Công an tỉnh Phú Thọ"
    return socv, ngay, donvi


def load_524_xlsx():
    wb = load_workbook(F524, read_only=True, data_only=True)
    recs = {}
    ghi_c = Counter()
    for row in wb.active.iter_rows(min_row=2, values_only=True):
        mst = norm_mst(row[1])
        if not is_valid_mst(mst):
            continue
        ten = clean_text(row[2])
        ghi = clean_text(row[3])
        blob = clean_text(row[4])
        socv, ngay, donvi = parse_524_cv(blob)
        ghi_c[ghi] += 1
        if "637" in ghi:
            nd = "Danh sách 637 DN (524 mở rộng) - BKL 05/BKL-CSKT(Đ4)"
            note = (
                "Nguồn: 524.xlsx. Quy mô đợt: 637 Công ty (524 mở rộng). "
                "BKL 05/BKL-CSKT(Đ4) ngày 16/08/2023 của Công an tỉnh Phú Thọ."
            )
        else:
            nd = "Danh sách 524 DN bán rủi ro - CV 1770/CAT-CSKT gắn CV 1798/TCT-TTKT"
            note = (
                "Nguồn: 524.xlsx. Quy mô đợt: 524 Công ty. "
                "CV 1770/CAT-CSKT ngày 29/03/2023 của Công an tỉnh Phú Thọ; "
                "liên kết CV 1798/TCT-TTKT ngày 16/05/2023 của Tổng cục Thuế."
            )
        recs[mst] = {
            "mst": mst,
            "ten": ten,
            "so_cv": norm_socv(socv),
            "ngay_cv": ngay,
            "don_vi_bh": donvi,
            "noi_dung": nd,
            "nhom_nd": "Rủi ro do Công an chuyển",
            "ghi_chu": note,
            "sheet": "F0",
            "nguon": "524.xlsx",
        }
    wb.close()
    return recs, ghi_c


def row_to_rec(row) -> dict:
    return {
        "so_cv": norm_socv(row[1]),
        "ngay_cv": xldate(row[2]) if row[2] is not None else None,
        "don_vi_bh": clean_text(row[3]),
        "ten": clean_text(row[4]),
        "mst": norm_mst(row[5]),
        "noi_dung": clean_text(row[6]),
        "tam_ngung": clean_text(row[7]),
        "nhom_nd": clean_text(row[8]),
        "dia_chi": clean_text(row[9]),
        "ngay_nhan": xldate(row[10]) if row[10] is not None else None,
        "tinh_trang": clean_text(row[11]),
        "ghi_chu": clean_text(row[12]),
        "file": clean_text(row[13]),
        "sheet": "F0",
    }


def absorb_extra(canon: dict, old: dict) -> None:
    for k in ("dia_chi", "tam_ngung", "tinh_trang", "ngay_nhan", "file"):
        if not canon.get(k) and old.get(k) not in (None, "", 0):
            canon[k] = old[k]


def main():
    if not SO_THEO_DOI.is_file():
        raise SystemExit(f"Thieu so theo doi: {SO_THEO_DOI}")
    if not F524.is_file():
        raise SystemExit(f"Thieu 524.xlsx: {F524}")

    canon_map, ghi_c = load_524_xlsx()
    log(f"524.xlsx MST={len(canon_map)} ghi={dict(ghi_c)}")

    log("Doc F0 hien tai (khong doc lai xlsb)...")
    wb_in = load_workbook(SO_THEO_DOI, read_only=True, data_only=True)
    ws = wb_in["F0"]
    keep = []
    dropped_out = []
    dropped_dup = 0
    list_in = 0
    seen_absorb = set()
    total = 0
    for row in ws.iter_rows(min_row=8, values_only=True):
        if not any(row):
            continue
        total += 1
        rec = row_to_rec(row)
        if is_524_list_row(rec["so_cv"], rec["ghi_chu"], rec["noi_dung"]):
            mst = rec["mst"]
            if mst in canon_map:
                absorb_extra(canon_map[mst], rec)
                list_in += 1
                if mst in seen_absorb:
                    dropped_dup += 1
                else:
                    seen_absorb.add(mst)
            else:
                dropped_out.append((rec.get("so_cv"), mst, rec.get("ten"), rec.get("ghi_chu")[:80]))
            continue
        keep.append(rec)
    wb_in.close()

    missing = [m for m in canon_map if m not in seen_absorb]
    f0_recs = keep + list(canon_map.values())

    def _key(r):
        ngay = r.get("ngay_cv") or datetime(1900, 1, 1)
        if not isinstance(ngay, datetime):
            ngay = datetime(1900, 1, 1)
        return (r.get("so_cv") or "", ngay, r.get("mst") or "")

    f0_recs.sort(key=_key)
    f0_recs = stamp_flags(f0_recs)

    log(f"F0 cu={total}")
    log(f"giu (khong phai ds 524)={len(keep)}")
    log(f"dong ds 524 dung MST 524.xlsx={list_in} unique={len(seen_absorb)}")
    log(f"gop trung ds 524={dropped_dup}")
    log(f"xoa ngoai 524.xlsx={len(dropped_out)}")
    for x in dropped_out:
        log(f"  OUT {x}")
    log(f"524.xlsx MST chua co dong ds tren F0 (se nap moi)={len(missing)}")
    log(f"F0 moi={len(f0_recs)} MST={len({r['mst'] for r in f0_recs})}")

    bak = backup_current()
    log(f"Backup {bak}")

    log("Mo workbook de giu F1, thay F0...")
    wb = load_workbook(SO_THEO_DOI)
    if "F0" in wb.sheetnames:
        del wb["F0"]
    build_f0_sheet(wb, f0_recs)
    # F0 phai dung truoc F1
    idx_f0 = wb.sheetnames.index("F0")
    if idx_f0 != 0:
        sheets = wb._sheets
        f0 = sheets.pop(idx_f0)
        sheets.insert(0, f0)

    out_tmp = Path(r"C:\Users\NGOC\AppData\Local\Temp\SoTheoDoi_f0_524_pruned.xlsx")
    wb.save(out_tmp)
    wb.close()
    log(f"Saved tmp {out_tmp} size={out_tmp.stat().st_size}")

    shutil.copy2(out_tmp, SO_THEO_DOI)
    log(f"Copied {SO_THEO_DOI} size={SO_THEO_DOI.stat().st_size}")
    for p in SYNC_COPIES:
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(out_tmp, p)
            log(f"Synced {p}")
        except Exception as e:
            log(f"Skip sync {p}: {e}")

    stats = Path(r"C:\Users\NGOC\AppData\Local\Temp\xoa_524_stats.txt")
    stats.write_text(
        "\n".join(
            [
                f"backup={bak}",
                f"f0_cu={total}",
                f"keep={len(keep)}",
                f"list_in={list_in}",
                f"unique_524={len(seen_absorb)}",
                f"dup_dropped={dropped_dup}",
                f"out_dropped={len(dropped_out)}",
                f"out={dropped_out}",
                f"missing_then_added={len(missing)}",
                f"f0_moi={len(f0_recs)}",
                f"mst={len({r['mst'] for r in f0_recs})}",
                f"socv_1770={sum(1 for r in f0_recs if (r.get('so_cv') or '')=='1770/CAT-CSKT')}",
                f"socv_bkl={sum(1 for r in f0_recs if '05/BKL' in (r.get('so_cv') or ''))}",
                f"socv_1798={sum(1 for r in f0_recs if '1798' in (r.get('so_cv') or ''))}",
            ]
        ),
        encoding="utf-8",
    )
    log("DONE")
    log(stats.read_text(encoding="utf-8"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
