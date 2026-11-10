# -*- coding: utf-8 -*-
"""Reparse XML with extra maTKhai codes (01=GTGT, 03=TNDN, 864/394=TNCN kỳ)."""
from __future__ import annotations

import json
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\skills\phan-tich-ho-so-thue\engine")
from ingest.xml_tkhai import _ct_map, _deep_ct, _find, _child, _text, _num, _ln, _nam_from_ky, sniff_xml

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH"
    r"\THANH_THANH_PHUONG_3600840461"
)
CACHE = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\scratch\ttp_3600840461\cache")
YEARS = [2020, 2021, 2022, 2023, 2024, 2025]

MA_LOAI = {
    "842": "gtgt",
    "01": "gtgt",
    "892": "tndn",
    "03": "tndn",
    "953": "tncn",
    "395": "tncn",
    "864": "tncn_ky",
    "394": "tncn_ky",
    "402": "bctc",
    "102": "bc26",
    "106": "tb_hd",
    "107": "tb_hd",
}


def parse_one(path: Path) -> dict | None:
    try:
        tree = ET.parse(path)
        root = tree.getroot()
    except ET.ParseError:
        return {"error": "parse fail"}
    if "HSoThueDTu" not in _ln(root.tag) and "kekhaithue" not in (root.tag or ""):
        return {"error": "not HSoThueDTu"}
    ma = _text(_find(root, "maTKhai"))
    loai = MA_LOAI.get(ma)
    if not loai:
        return {"error": f"unknown ma {ma}", "ma_tkhai": ma}
    ky_el = _find(root, "KyKKhaiThue")
    ky = _text(_child(ky_el, "kyKKhai")) if ky_el is not None else ""
    nam = _nam_from_ky(ky)
    nnt = _find(root, "NNT")
    rec = {
        "loai": loai,
        "ma_tkhai": ma,
        "loai_tkhai": _text(_find(root, "loaiTKhai")) or "C",
        "so_lan": int(_num(_text(_find(root, "soLan")))),
        "ky": ky,
        "nam": nam,
        "tu_ngay": _text(_child(ky_el, "kyKKhaiTuNgay")) if ky_el is not None else "",
        "den_ngay": _text(_child(ky_el, "kyKKhaiDenNgay")) if ky_el is not None else "",
        "mst": _text(_child(nnt, "mst")) if nnt is not None else "",
        "ten": _text(_child(nnt, "tenNNT")) if nnt is not None else "",
        "chi_tieu": {},
        "cdkt_cuoi": {},
        "cdkt_dau": {},
        "kqkd": {},
    }
    if loai == "gtgt":
        rec["chi_tieu"] = _ct_map(_find(root, "CTieuTKhaiChinh"))
        for code, val in _deep_ct(root).items():
            rec["chi_tieu"].setdefault(code, val)
    elif loai == "tndn":
        pl = _find(root, "PLuc_03_1A_TNDN")
        rec["chi_tieu"] = _ct_map(pl) if pl is not None else _deep_ct(root)
        if not rec["chi_tieu"]:
            rec["chi_tieu"] = _deep_ct(root)
    elif loai in {"tncn", "tncn_ky"}:
        rec["chi_tieu"] = _deep_ct(_find(root, "CTieuTKhaiChinh") or root)
    elif loai == "bctc":
        rec["cdkt_cuoi"] = _ct_map(_find(root, "SoCuoiNam"))
        rec["cdkt_dau"] = _ct_map(_find(root, "SoDauNam"))
        kq = _find(root, "PL_KQHDSXKD")
        if kq is not None:
            rec["kqkd"] = _ct_map(_child(kq, "NamNay"))
        rec["chi_tieu"] = {**rec["kqkd"]}
    else:
        rec["chi_tieu"] = _deep_ct(root)
    return rec


def effective(xmls, loai):
    groups = defaultdict(list)
    for x in xmls:
        if x.get("loai") != loai or x.get("error") or x.get("duplicate_of"):
            continue
        groups[str(x.get("ky") or "")].append(x)
    eff = {}
    for ky, items in groups.items():
        items = sorted(
            items,
            key=lambda z: (
                1 if str(z.get("loai_tkhai") or "").upper().startswith("B") else 0,
                int(z.get("so_lan") or 0),
            ),
        )
        best = dict(items[-1])
        best["la_hieu_luc"] = "Có"
        eff[ky] = best
    return eff


def year_sum_rows(rows, chieu, nam, field, tinh_only=False):
    s = 0.0
    n = 0
    for r in rows:
        if r.get("chieu") != chieu or r.get("nam") != nam:
            continue
        if tinh_only and r.get("tinh") != "Có":
            continue
        s += float(r.get(field) or 0)
        n += 1
    return s, n


def main():
    manifest = json.loads((CACHE / "manifest.json").read_text(encoding="utf-8"))
    seen = {}
    out = []
    xmls = [m for m in manifest if m["ext"] == ".xml" and m["goc"] == "CQT"]
    print("xml", len(xmls))
    for m in xmls:
        p = ROOT / m["rel"]
        if m["sha256"] in seen:
            rec = dict(seen[m["sha256"]])
            rec["file"] = m["rel"]
            rec["duplicate_of"] = seen[m["sha256"]]["file"]
            rec["sha256"] = m["sha256"]
            out.append(rec)
            continue
        if not sniff_xml(p):
            rec = {"file": m["rel"], "error": "not HSoThueDTu", "sha256": m["sha256"]}
            seen[m["sha256"]] = rec
            out.append(rec)
            continue
        rec = parse_one(p) or {"error": "parse fail"}
        rec["file"] = m["rel"]
        rec["sha256"] = m["sha256"]
        seen[m["sha256"]] = rec
        out.append(rec)
    (CACHE / "xml.json").write_text(json.dumps(out, ensure_ascii=False, default=str), encoding="utf-8")
    from collections import Counter
    print("loai", Counter(x.get("loai") for x in out if not x.get("error") and not x.get("duplicate_of")))
    print("err", Counter(x.get("error") for x in out if x.get("error")))

    eff_gtgt = effective(out, "gtgt")
    eff_tndn = effective(out, "tndn")
    eff_bctc = effective(out, "bctc")
    eff_tncn = effective(out, "tncn")
    eff_tncn_ky = effective(out, "tncn_ky")
    for name, obj in [
        ("eff_gtgt.json", eff_gtgt),
        ("eff_tndn.json", eff_tndn),
        ("eff_bctc.json", eff_bctc),
        ("eff_tncn.json", eff_tncn),
        ("eff_tncn_ky.json", eff_tncn_ky),
    ]:
        (CACHE / name).write_text(json.dumps(obj, ensure_ascii=False, default=str), encoding="utf-8")
        print(name, len(obj), "keys sample", list(obj.keys())[:8])

    hddt = json.loads((CACHE / "hddt.json").read_text(encoding="utf-8"))
    bk = json.loads((CACHE / "bangke.json").read_text(encoding="utf-8"))
    years = {}
    for y in YEARS:
        hd_ban_gt, hd_ban_n = year_sum_rows(hddt, "BAN", y, "gt", True)
        hd_ban_th, _ = year_sum_rows(hddt, "BAN", y, "thue", True)
        hd_mua_gt, hd_mua_n = year_sum_rows(hddt, "MUA", y, "gt", True)
        hd_mua_th, _ = year_sum_rows(hddt, "MUA", y, "thue", True)
        bk_ban_gt, bk_ban_n = year_sum_rows(bk, "BAN", y, "gt")
        bk_ban_th, _ = year_sum_rows(bk, "BAN", y, "thue")
        bk_mua_gt, bk_mua_n = year_sum_rows(bk, "MUA", y, "gt")
        bk_mua_th, _ = year_sum_rows(bk, "MUA", y, "thue")
        gtgt_34 = gtgt_35 = gtgt_23 = gtgt_24 = gtgt_40 = gtgt_43 = 0.0
        n_ky = 0
        months = []
        for ky, rec in eff_gtgt.items():
            if rec.get("nam") != y:
                continue
            n_ky += 1
            months.append(ky)
            ct = rec.get("chi_tieu") or {}
            gtgt_34 += float(ct.get("34") or 0)
            gtgt_35 += float(ct.get("35") or 0)
            gtgt_23 += float(ct.get("23") or 0)
            gtgt_24 += float(ct.get("24") or 0)
            gtgt_40 += float(ct.get("40") or 0)
            gtgt_43 += float(ct.get("43") or 0)
        tndn = eff_tndn.get(str(y), {})
        bctc = eff_bctc.get(str(y), {})
        years[str(y)] = {
            "hddt_ban_gt": hd_ban_gt, "hddt_ban_th": hd_ban_th, "hddt_ban_n": hd_ban_n,
            "hddt_mua_gt": hd_mua_gt, "hddt_mua_th": hd_mua_th, "hddt_mua_n": hd_mua_n,
            "bk_ban_gt": bk_ban_gt, "bk_ban_th": bk_ban_th, "bk_ban_n": bk_ban_n,
            "bk_mua_gt": bk_mua_gt, "bk_mua_th": bk_mua_th, "bk_mua_n": bk_mua_n,
            "gtgt_nky": n_ky, "gtgt_ky": months,
            "gtgt_34": gtgt_34, "gtgt_35": gtgt_35, "gtgt_23": gtgt_23, "gtgt_24": gtgt_24,
            "gtgt_40": gtgt_40, "gtgt_43": gtgt_43,
            "tndn_ct04": (tndn.get("chi_tieu") or {}).get("04"),
            "tndn_ct12": (tndn.get("chi_tieu") or {}).get("12"),
            "tndn_ct13": (tndn.get("chi_tieu") or {}).get("13"),
            "tndn_ct16": (tndn.get("chi_tieu") or {}).get("16"),
            "tndn_ct18": (tndn.get("chi_tieu") or {}).get("18"),
            "bctc_01": (bctc.get("kqkd") or {}).get("01"),
            "bctc_10": (bctc.get("kqkd") or {}).get("10"),
            "bctc_11": (bctc.get("kqkd") or {}).get("11"),
            "bctc_20": (bctc.get("kqkd") or {}).get("20"),
            "bctc_22": (bctc.get("kqkd") or {}).get("22"),
            "bctc_30": (bctc.get("kqkd") or {}).get("30"),
            "bctc_50": (bctc.get("kqkd") or {}).get("50"),
            "bctc_140": (bctc.get("cdkt_cuoi") or {}).get("140"),
            "bctc_141": (bctc.get("cdkt_cuoi") or {}).get("141"),
            "bctc_152": (bctc.get("cdkt_cuoi") or {}).get("152"),
            "bctc_311": (bctc.get("cdkt_cuoi") or {}).get("311"),
            "bctc_331": (bctc.get("cdkt_cuoi") or {}).get("331"),
            "bctc_411": (bctc.get("cdkt_cuoi") or {}).get("411"),
            "bctc_421": (bctc.get("cdkt_cuoi") or {}).get("421"),
            "che_do": "TT200",
        }
        print(
            y,
            "nky", n_ky,
            "GT34", round(gtgt_34),
            "BK", round(bk_ban_gt),
            "HD", round(hd_ban_gt),
            "BCTC", years[str(y)]["bctc_01"],
            "TNDN04", years[str(y)]["tndn_ct04"],
        )
    summary = json.loads((CACHE / "summary.json").read_text(encoding="utf-8"))
    summary["years"] = years
    summary["eff_gtgt_keys"] = list(eff_gtgt.keys())
    summary["eff_tndn_keys"] = list(eff_tndn.keys())
    summary["eff_bctc_keys"] = list(eff_bctc.keys())
    summary["eff_tncn_keys"] = list(eff_tncn.keys())
    summary["eff_tncn_ky_keys"] = list(eff_tncn_ky.keys())
    (CACHE / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, default=str), encoding="utf-8")
    print("done reparse")


if __name__ == "__main__":
    main()
