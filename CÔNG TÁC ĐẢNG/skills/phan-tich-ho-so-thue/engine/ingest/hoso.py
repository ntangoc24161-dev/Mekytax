"""Quét thư mục hồ sơ, gộp nhiều năm, khử trùng XML/HĐ."""
from __future__ import annotations

from pathlib import Path

from models import HoaDon, HoSoChuan, ToKhai

from .excel_hd import parse_bang_ke
from .xml_tkhai import file_hash, parse_xml, sniff_xml

SKIP_DIR = {"bien ban", "biên bản"}
SKIP_FILE_PREFIX = ("ptdn01", "ptdn02")


def scan_dir(root: Path) -> HoSoChuan:
    root = Path(root)
    xml_files: list[Path] = []
    xlsx_files: list[Path] = []
    skipped: list[str] = []
    seen_hash: set[str] = set()
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        parts = {x.casefold() for x in p.parts}
        if parts & SKIP_DIR:
            continue
        if p.name.casefold().startswith(SKIP_FILE_PREFIX):
            skipped.append(f"bỏ output tool cũ: {p.name}")
            continue
        suf = "".join(p.suffixes).casefold()
        if ".xml" in suf:
            xml_files.append(p)
        elif p.suffix.casefold() in {".xlsx", ".xlsm"}:
            xlsx_files.append(p)

    to_khai: list[ToKhai] = []
    warnings: list[str] = []
    for p in xml_files:
        if not sniff_xml(p):
            skipped.append(f"XML không phải HSoThueDTu: {p.name}")
            continue
        h = file_hash(p)
        if h in seen_hash:
            skipped.append(f"trùng nội dung: {p.name}")
            continue
        seen_hash.add(h)
        tk = parse_xml(p)
        if tk is None:
            warnings.append(f"parse XML thất bại: {p.name}")
            continue
        to_khai.append(tk)

    to_khai = _pick_latest(to_khai)

    hoa_don: list[HoaDon] = []
    for p in xlsx_files:
        rows = parse_bang_ke(p)
        if not rows:
            skipped.append(f"Excel không phải bảng kê BR/MV: {p.name}")
            continue
        hoa_don.extend(rows)
    hoa_don = _unique_hd(hoa_don)

    mst = ""
    ten = ""
    for tk in to_khai:
        if tk.mst:
            mst = tk.mst
            ten = tk.ten_nnt
            break
    nam = sorted({tk.nam for tk in to_khai if tk.nam} | {
        int(hd.ngay[:4]) for hd in hoa_don if len(hd.ngay) >= 4 and hd.ngay[:4].isdigit()
    })
    return HoSoChuan(
        mst=mst,
        ten_nnt=ten,
        nam=nam,
        to_khai=to_khai,
        hoa_don=hoa_don,
        skipped=skipped,
        warnings=warnings,
    )


def _pick_latest(rows: list[ToKhai]) -> list[ToKhai]:
    best: dict[tuple, ToKhai] = {}
    for tk in rows:
        key = (tk.loai, tk.ky)
        old = best.get(key)
        if old is None:
            best[key] = tk
            continue
        score = (1 if tk.loai_tkhai == "B" else 0, tk.so_lan)
        old_s = (1 if old.loai_tkhai == "B" else 0, old.so_lan)
        if score >= old_s:
            best[key] = tk
    return list(best.values())


def _unique_hd(rows: list[HoaDon]) -> list[HoaDon]:
    seen: set[tuple] = set()
    out: list[HoaDon] = []
    for hd in rows:
        k = (hd.chieu, hd.so, hd.mst_dt, hd.ngay, round(hd.doanh_so, 0), hd.mat_hang[:40])
        if k in seen:
            continue
        seen.add(k)
        out.append(hd)
    return out
