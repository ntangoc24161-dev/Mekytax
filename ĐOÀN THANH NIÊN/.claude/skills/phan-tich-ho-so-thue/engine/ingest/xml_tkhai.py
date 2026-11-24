"""Parse XML HSoThueDTu (kê khai thuế điện tử / HTKK)."""
from __future__ import annotations

import hashlib
import xml.etree.ElementTree as ET
from pathlib import Path

from models import ToKhai

NS_HINT = "kekhaithue.gdt.gov.vn"

MA_LOAI = {"842": "gtgt", "892": "tndn", "953": "tncn", "402": "bctc"}


def _ln(tag: str) -> str:
    return tag.split("}", 1)[-1] if tag else ""


def _text(el: ET.Element | None) -> str:
    if el is None or el.text is None:
        return ""
    return el.text.strip()


def _num(s: str) -> float:
    t = (s or "").replace(",", "").strip()
    if not t or t in {"false", "true"}:
        return 0.0
    try:
        return float(t)
    except ValueError:
        return 0.0


def _find(parent: ET.Element, name: str) -> ET.Element | None:
    for ch in parent.iter():
        if _ln(ch.tag) == name:
            return ch
    return None


def _child(parent: ET.Element, name: str) -> ET.Element | None:
    for ch in list(parent):
        if _ln(ch.tag) == name:
            return ch
    return None


def _ct_map(el: ET.Element | None) -> dict[str, float]:
    out: dict[str, float] = {}
    if el is None:
        return out
    for ch in list(el):
        n = _ln(ch.tag)
        if n.startswith("ct") and n[2:].replace("a", "").replace("b", "").isdigit():
            code = n[2:]
            if list(ch):
                continue
            val = _text(ch)
            if val:
                out[code] = _num(val)
    return out


def sniff_xml(path: Path) -> bool:
    try:
        head = path.read_bytes()[:800]
    except OSError:
        return False
    if head.startswith(b"PK"):
        return False
    try:
        t = head.decode("utf-8", errors="ignore")
    except Exception:
        return False
    return "HSoThueDTu" in t or NS_HINT in t


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def parse_xml(path: Path) -> ToKhai | None:
    try:
        tree = ET.parse(path)
        root = tree.getroot()
    except ET.ParseError:
        return None
    if "HSoThueDTu" not in _ln(root.tag) and NS_HINT not in (root.tag or ""):
        return None
    ma = _text(_find(root, "maTKhai"))
    loai = MA_LOAI.get(ma)
    if not loai:
        return None
    ky_el = _find(root, "KyKKhaiThue")
    ky = _text(_child(ky_el, "kyKKhai")) if ky_el is not None else ""
    nam = _nam_from_ky(ky)
    nnt = _find(root, "NNT")
    tk = ToKhai(
        loai=loai,
        ma_tkhai=ma,
        loai_tkhai=_text(_find(root, "loaiTKhai")) or "C",
        so_lan=int(_num(_text(_find(root, "soLan")))),
        ky=ky,
        nam=nam,
        tu_ngay=_text(_child(ky_el, "kyKKhaiTuNgay")) if ky_el is not None else "",
        den_ngay=_text(_child(ky_el, "kyKKhaiDenNgay")) if ky_el is not None else "",
        mst=_text(_child(nnt, "mst")) if nnt is not None else "",
        ten_nnt=_text(_child(nnt, "tenNNT")) if nnt is not None else "",
        chi_tieu={},
        file=str(path),
    )
    if loai == "gtgt":
        tk.chi_tieu = _ct_map(_find(root, "CTieuTKhaiChinh"))
        # chỉ tiêu lồng (ct23/ct24, ct32/ct33, ct34/ct35)
        for code, val in _deep_ct(root).items():
            tk.chi_tieu.setdefault(code, val)
    elif loai == "tndn":
        pl = _find(root, "PLuc_03_1A_TNDN")
        tk.chi_tieu = _ct_map(pl) if pl is not None else _deep_ct(root)
    elif loai == "tncn":
        tk.chi_tieu = _deep_ct(_find(root, "CTieuTKhaiChinh") or root)
    elif loai == "bctc":
        cuoi = _find(root, "SoCuoiNam")
        dau = _find(root, "SoDauNam")
        tk.bctc_cdkt_cuoi = _ct_map(cuoi)
        tk.bctc_cdkt_dau = _ct_map(dau)
        kq = _find(root, "PL_KQHDSXKD")
        if kq is not None:
            tk.bctc_kqkd = _ct_map(_child(kq, "NamNay"))
        tk.chi_tieu = {**tk.bctc_kqkd}
    return tk


def _deep_ct(el: ET.Element | None) -> dict[str, float]:
    out: dict[str, float] = {}
    if el is None:
        return out
    for ch in el.iter():
        n = _ln(ch.tag)
        if n.startswith("ct") and list(ch) == [] and _text(ch):
            code = n[2:]
            if code[:1].isdigit():
                out.setdefault(code, _num(_text(ch)))
    return out


def _nam_from_ky(ky: str) -> int:
    ky = (ky or "").strip()
    if "/" in ky:
        part = ky.split("/")[-1]
        try:
            return int(part)
        except ValueError:
            return 0
    try:
        return int(ky[:4])
    except ValueError:
        return 0
