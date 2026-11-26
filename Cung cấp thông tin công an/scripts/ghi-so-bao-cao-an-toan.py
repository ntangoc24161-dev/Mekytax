#!/usr/bin/env python3
"""Ghi sổ Bao_cao_cong_van.xlsx an toàn trên Linux/VPS (chốt 20/11/2026).

KHÔNG dùng openpyxl.save cả workbook (mất chart DASHBOARD).
KHÔNG serialize sheet bằng ElementTree.
Chỉ vá XML sheet2/sharedStrings rồi đóng gói lại với ZipInfo Windows.

Lệnh:
  validate  --in FILE.xlsx
  repair    --in FILE.xlsx --out FILE.xlsx
  set-status --in FILE.xlsx --out FILE.xlsx --row N [--row N...] --status "Đang xử lý"
  set-cells  --in FILE.xlsx --out FILE.xlsx --cell O901=Đang xử lý [--cell K903=...]
  set-dates  --in FILE.xlsx --out FILE.xlsx --cell D867=10/11/2026 [--style-row 864]

Sau khi ra file --out: agent tự GOOGLEDRIVE_UPLOAD_UPDATE_FILE đúng fileId sống
1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph (account googledrive_ralph-urbane). Script này
không upload Drive.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

SHEET2 = "xl/worksheets/sheet2.xml"
SST = "xl/sharedStrings.xml"
CHARTS = ("xl/charts/chart1.xml", "xl/charts/chart2.xml")

NS_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
NS_XR2 = 'xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"'
NS_XR3 = 'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"'

WIN_CREATE_SYSTEM = 3
WIN_EXTERNAL_ATTR = 0x81A40000

SST_OPEN_RE = re.compile(rb"<sst\b[^>]*>")
SI_RE = re.compile(r"<si>(.*?)</si>", re.S)
T_RE = re.compile(r"<t(?P<attrs>[^>]*)>(?P<text>.*?)</t>", re.S)


class HealthError(Exception):
    pass


def _md5(path: Path) -> str:
    return hashlib.md5(path.read_bytes()).hexdigest()


def _col_row(ref: str) -> Tuple[str, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref.upper())
    if not m:
        raise ValueError(f"ô không hợp lệ: {ref}")
    return m.group(1), int(m.group(2))


def _si_plain_text(si_xml: str) -> str:
    parts = []
    for m in T_RE.finditer(si_xml):
        parts.append(_xml_unescape(m.group("text")))
    return "".join(parts)


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _xml_unescape(text: str) -> str:
    return (
        text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&amp;", "&")
    )


def _read_workbook(path: Path) -> Tuple[List[str], Dict[str, zipfile.ZipInfo], Dict[str, bytes]]:
    with zipfile.ZipFile(path) as z:
        names = list(z.namelist())
        infos = {i.filename: i for i in z.infolist()}
        data = {n: z.read(n) for n in names}
    return names, infos, data


def _fix_sst_header(sst: bytes, si_count: int) -> bytes:
    m = SST_OPEN_RE.search(sst)
    if not m:
        raise HealthError("thiếu thẻ <sst>")
    old = m.group(0)
    new = re.sub(rb'count="\d+"', f'count="{si_count}"'.encode(), old, count=1)
    new = re.sub(rb'uniqueCount="\d+"', f'uniqueCount="{si_count}"'.encode(), new, count=1)
    if b'count="' not in new:
        # rare: single-quoted attrs
        new = re.sub(rb"count='\d+'", f"count='{si_count}'".encode(), old, count=1)
        new = re.sub(rb"uniqueCount='\d+'", f"uniqueCount='{si_count}'".encode(), new, count=1)
    return sst[: m.start()] + new + sst[m.end() :]


def _ensure_sheet2_namespaces(sheet2: bytes) -> bytes:
    text = sheet2.decode("utf-8")
    m = re.search(r"<worksheet\b[^>]*>", text)
    if not m:
        raise HealthError("thiếu thẻ <worksheet>")
    tag = m.group(0)
    new = tag
    if "xmlns:r=" not in new:
        new = new.replace(
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
            'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' + NS_R,
            1,
        )
    if "xmlns:xr2=" not in new or "xmlns:xr3=" not in new:
        if "xmlns:xr=" in new:
            new = re.sub(
                r'(xmlns:xr="[^"]+")',
                lambda mm: mm.group(1)
                + ("" if "xmlns:xr2=" in new else " " + NS_XR2)
                + ("" if "xmlns:xr3=" in new else " " + NS_XR3),
                new,
                count=1,
            )
        else:
            new = new[:-1] + " " + NS_XR2 + " " + NS_XR3 + ">"
    if new == tag:
        return sheet2
    return (text[: m.start()] + new + text[m.end() :]).encode("utf-8")


def _sst_index_map(sst_text: str) -> Dict[str, int]:
    out: Dict[str, int] = {}
    for i, si in enumerate(SI_RE.findall(sst_text)):
        plain = _si_plain_text(si)
        # first index wins (Excel unique strings)
        out.setdefault(plain, i)
    return out


def _append_sst_string(sst: bytes, text: str) -> Tuple[bytes, int]:
    sst_text = sst.decode("utf-8")
    idx_map = _sst_index_map(sst_text)
    if text in idx_map:
        return sst, idx_map[text]
    # preserve xml:space when leading/trailing whitespace
    if text != text.strip():
        si = f'<si><t xml:space="preserve">{_xml_escape(text)}</t></si>'
    else:
        si = f"<si><t>{_xml_escape(text)}</t></si>"
    if "</sst>" not in sst_text:
        raise HealthError("sharedStrings thiếu </sst>")
    sst_text = sst_text.replace("</sst>", si + "</sst>", 1)
    si_count = sst_text.count("<si>")
    fixed = _fix_sst_header(sst_text.encode("utf-8"), si_count)
    return fixed, si_count - 1


def _set_cell_shared(sheet2: bytes, ref: str, sst_idx: int, default_style: str = "40") -> bytes:
    ref = ref.upper()
    text = sheet2.decode("utf-8")
    pattern = re.compile(
        rf'<c r="{re.escape(ref)}"[^/]*/>|<c r="{re.escape(ref)}"[^>]*>.*?</c>',
        re.S,
    )
    m = pattern.search(text)
    new_cell = f'<c r="{ref}" s="{default_style}" t="s"><v>{sst_idx}</v></c>'
    if m:
        old = m.group(0)
        # keep existing style if present
        sm = re.search(r'\bs="(\d+)"', old)
        if sm:
            new_cell = f'<c r="{ref}" s="{sm.group(1)}" t="s"><v>{sst_idx}</v></c>'
        text = text[: m.start()] + new_cell + text[m.end() :]
        return text.encode("utf-8")

    # Cell missing: insert into its row if the row exists
    _, row = _col_row(ref)
    row_re = re.compile(rf'(<row r="{row}"[^>]*>)(.*?)(</row>)', re.S)
    rm = row_re.search(text)
    if not rm:
        raise HealthError(f"không thấy dòng {row} để ghi ô {ref}")
    open_tag, body, close_tag = rm.group(1), rm.group(2), rm.group(3)
    # insert before </row>, after existing cells
    body2 = body + new_cell
    text = text[: rm.start()] + open_tag + body2 + close_tag + text[rm.end() :]
    return text.encode("utf-8")


def _clone_zipinfo(base: zipfile.ZipInfo, name: str) -> zipfile.ZipInfo:
    info = zipfile.ZipInfo(filename=name, date_time=base.date_time)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.create_system = WIN_CREATE_SYSTEM
    info.create_version = getattr(base, "create_version", 20)
    info.extract_version = getattr(base, "extract_version", 20)
    info.flag_bits = 0
    info.external_attr = WIN_EXTERNAL_ATTR
    info.internal_attr = getattr(base, "internal_attr", 0)
    if getattr(base, "extra", None):
        info.extra = base.extra
    return info


def write_safe_xlsx(
    names: List[str],
    infos: Dict[str, zipfile.ZipInfo],
    data: Dict[str, bytes],
    out_path: Path,
    zipinfo_template: Optional[Dict[str, zipfile.ZipInfo]] = None,
) -> dict:
    """Đóng gói xlsx với ZipInfo Windows; trả báo cáo validate."""
    # Always normalize sheet2 ns + SST counts before pack
    if SHEET2 not in data or SST not in data:
        raise HealthError("thiếu sheet2 hoặc sharedStrings")
    data[SHEET2] = _ensure_sheet2_namespaces(data[SHEET2])
    si = data[SST].count(b"<si>")
    data[SST] = _fix_sst_header(data[SST], si)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zout:
        for name in names:
            base = None
            if zipinfo_template and name in zipinfo_template:
                base = zipinfo_template[name]
            elif name in infos:
                base = infos[name]
            else:
                base = zipfile.ZipInfo(name)
            info = _clone_zipinfo(base, name)
            zout.writestr(info, data[name])
    raw = buf.getvalue()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(raw)
    report = validate_bytes(raw)
    report["path"] = str(out_path)
    report["size"] = len(raw)
    report["md5"] = hashlib.md5(raw).hexdigest()
    return report


def validate_bytes(raw: bytes) -> dict:
    issues: List[str] = []
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        bad = z.testzip()
        if bad is not None:
            issues.append(f"testzip hỏng tại {bad}")
        names = z.namelist()
        for c in CHARTS:
            if c not in names:
                issues.append(f"thiếu {c}")
        s2 = z.read(SHEET2)
        s2_head = s2[:2000].decode("utf-8", "replace")
        for token in ("xmlns:r=", "xmlns:xr2=", "xmlns:xr3="):
            if token not in s2_head:
                issues.append(f"sheet2 thiếu {token}")
        sst = z.read(SST)
        m = re.search(rb'count="(\d+)"[^>]*uniqueCount="(\d+)"', sst) or re.search(
            rb"count='(\d+)'[^>]*uniqueCount='(\d+)'", sst
        )
        si = sst.count(b"<si>")
        if not m:
            issues.append("không đọc được count/uniqueCount SST")
            c = u = None
        else:
            c, u = int(m.group(1)), int(m.group(2))
            if not (c == u == si):
                issues.append(f"SST lệch count={c} unique={u} si={si}")
        info = z.getinfo(SHEET2)
        if info.create_system != WIN_CREATE_SYSTEM:
            issues.append(f"ZipInfo create_system={info.create_system} (cần {WIN_CREATE_SYSTEM})")
        if info.external_attr != WIN_EXTERNAL_ATTR:
            issues.append(
                f"ZipInfo external_attr={hex(info.external_attr)} (cần {hex(WIN_EXTERNAL_ATTR)})"
            )
    ok = not issues
    return {
        "ok": ok,
        "issues": issues,
        "sst_count": c if m else None,
        "sst_unique": u if m else None,
        "sst_si": si,
        "charts": sum(1 for c in CHARTS if c in names),
        "sheet2_create_system": info.create_system,
        "sheet2_external_attr": hex(info.external_attr),
    }


def validate_file(path: Path) -> dict:
    raw = path.read_bytes()
    report = validate_bytes(raw)
    report["path"] = str(path)
    report["size"] = path.stat().st_size
    report["md5"] = _md5(path)
    return report


def cmd_validate(args: argparse.Namespace) -> int:
    report = validate_file(Path(args.input))
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["ok"] else 2


def cmd_repair(args: argparse.Namespace) -> int:
    src = Path(args.input)
    out = Path(args.output or args.input)
    names, infos, data = _read_workbook(src)
    tmpl = None
    if args.template:
        _, tmpl, _ = _read_workbook(Path(args.template))
    report = write_safe_xlsx(names, infos, data, out, zipinfo_template=tmpl)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["ok"] else 2


def _apply_text_cells(
    data: Dict[str, bytes],
    updates: Iterable[Tuple[str, str]],
) -> Dict[str, int]:
    sheet2 = data[SHEET2]
    sst = data[SST]
    applied: Dict[str, int] = {}
    for ref, text in updates:
        ref = ref.upper()
        sst, idx = _append_sst_string(sst, text)
        # style hint: column O thường s="40"
        style = "40" if ref.startswith("O") else "38"
        sheet2 = _set_cell_shared(sheet2, ref, idx, default_style=style)
        applied[ref] = idx
    data[SHEET2] = sheet2
    data[SST] = sst
    return applied


def cmd_set_status(args: argparse.Namespace) -> int:
    src = Path(args.input)
    out = Path(args.output or args.input)
    if not args.row:
        raise SystemExit("cần ít nhất một --row")
    status = args.status
    updates = [(f"O{r}", status) for r in args.row]
    names, infos, data = _read_workbook(src)
    applied = _apply_text_cells(data, updates)
    tmpl = None
    if args.template:
        _, tmpl, _ = _read_workbook(Path(args.template))
    report = write_safe_xlsx(names, infos, data, out, zipinfo_template=tmpl)
    report["applied"] = applied
    report["status"] = status
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["ok"] else 2


def _excel_serial(text: str) -> int:
    """dd/mm/yyyy -> số serial Excel (1900 system)."""
    import datetime as _dt

    m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*", text)
    if not m:
        raise ValueError(f"ngày không hợp lệ (cần dd/mm/yyyy): {text}")
    d = _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    return (d - _dt.date(1899, 12, 30)).days


def _style_of(sheet2: bytes, ref: str) -> Optional[str]:
    text = sheet2.decode("utf-8")
    m = re.search(rf'<c r="{re.escape(ref)}"[^>]*?\bs="(\d+)"', text)
    return m.group(1) if m else None


def _set_cell_number(sheet2: bytes, ref: str, value: int, style: str) -> bytes:
    """Ghi ô số (không sharedString), giữ nguyên các ô khác."""
    ref = ref.upper()
    text = sheet2.decode("utf-8")
    pattern = re.compile(
        rf'<c r="{re.escape(ref)}"[^/]*/>|<c r="{re.escape(ref)}"[^>]*>.*?</c>',
        re.S,
    )
    new_cell = f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    m = pattern.search(text)
    if m:
        text = text[: m.start()] + new_cell + text[m.end() :]
        return text.encode("utf-8")
    _, row = _col_row(ref)
    row_re = re.compile(rf'(<row r="{row}"[^>]*>)(.*?)(</row>)', re.S)
    rm = row_re.search(text)
    if not rm:
        raise HealthError(f"không thấy dòng {row} để ghi ô {ref}")
    text = text[: rm.start()] + rm.group(1) + rm.group(2) + new_cell + rm.group(3) + text[rm.end() :]
    return text.encode("utf-8")


def cmd_set_dates(args: argparse.Namespace) -> int:
    """Ghi ô NGÀY thật (serial Excel). Style lấy từ ô cùng cột ở dòng --style-row (mặc định 864)."""
    src = Path(args.input)
    out = Path(args.output or args.input)
    pairs: List[Tuple[str, str]] = []
    for item in args.cell or []:
        if "=" not in item:
            raise SystemExit(f"--cell cần dạng D867=10/11/2026, nhận được: {item}")
        ref, text = item.split("=", 1)
        pairs.append((ref.strip().upper(), text.strip()))
    if not pairs:
        raise SystemExit("không có ô ngày nào (--cell)")
    names, infos, data = _read_workbook(src)
    sheet2 = data[SHEET2]
    applied: Dict[str, int] = {}
    for ref, text in pairs:
        col, _ = _col_row(ref)
        style = _style_of(sheet2, f"{col}{args.style_row}")
        if not style:
            raise SystemExit(f"không tìm thấy style ngày ở {col}{args.style_row}")
        serial = _excel_serial(text)
        sheet2 = _set_cell_number(sheet2, ref, serial, style)
        applied[ref] = serial
    data[SHEET2] = sheet2
    tmpl = None
    if args.template:
        _, tmpl, _ = _read_workbook(Path(args.template))
    report = write_safe_xlsx(names, infos, data, out, zipinfo_template=tmpl)
    report["applied_serial"] = applied
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["ok"] else 2


def cmd_set_cells(args: argparse.Namespace) -> int:
    src = Path(args.input)
    out = Path(args.output or args.input)
    updates: List[Tuple[str, str]] = []
    for item in args.cell or []:
        if "=" not in item:
            raise SystemExit(f"--cell cần dạng O901=Giá trị, nhận được: {item}")
        ref, text = item.split("=", 1)
        updates.append((ref.strip(), text))
    if args.json_file:
        payload = json.loads(Path(args.json_file).read_text(encoding="utf-8"))
        if not isinstance(payload, list):
            raise SystemExit("--json-file phải là mảng [{ref,text}, ...]")
        for row in payload:
            updates.append((str(row["ref"]), str(row["text"])))
    if not updates:
        raise SystemExit("không có ô nào để ghi (--cell hoặc --json-file)")
    names, infos, data = _read_workbook(src)
    applied = _apply_text_cells(data, updates)
    tmpl = None
    if args.template:
        _, tmpl, _ = _read_workbook(Path(args.template))
    report = write_safe_xlsx(names, infos, data, out, zipinfo_template=tmpl)
    report["applied"] = applied
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["ok"] else 2


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Ghi sổ Bao_cao an toàn (XML + ZipInfo Windows)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_val = sub.add_parser("validate", help="Kiểm SST/namespace/ZipInfo/chart")
    p_val.add_argument("--in", dest="input", required=True)
    p_val.set_defaults(func=cmd_validate)

    p_rep = sub.add_parser("repair", help="Sửa SST/namespace và đóng gói ZipInfo Windows")
    p_rep.add_argument("--in", dest="input", required=True)
    p_rep.add_argument("--out", dest="output", default="")
    p_rep.add_argument("--template", default="", help="xlsx mẫu lấy ZipInfo (tuỳ chọn)")
    p_rep.set_defaults(func=cmd_repair)

    p_st = sub.add_parser("set-status", help="Ghi cột O (Tình trạng) theo dòng")
    p_st.add_argument("--in", dest="input", required=True)
    p_st.add_argument("--out", dest="output", default="")
    p_st.add_argument("--row", type=int, action="append", default=[])
    p_st.add_argument("--status", required=True)
    p_st.add_argument("--template", default="")
    p_st.set_defaults(func=cmd_set_status)

    p_cell = sub.add_parser("set-cells", help="Ghi nhiều ô chữ (sharedStrings)")
    p_cell.add_argument("--in", dest="input", required=True)
    p_cell.add_argument("--out", dest="output", default="")
    p_cell.add_argument("--cell", action="append", default=[])
    p_cell.add_argument("--json-file", default="")
    p_cell.add_argument("--template", default="")
    p_cell.set_defaults(func=cmd_set_cells)

    p_date = sub.add_parser("set-dates", help="Ghi ô NGÀY dạng số Excel (style lấy từ dòng --style-row)")
    p_date.add_argument("--in", dest="input", required=True)
    p_date.add_argument("--out", dest="output", default="")
    p_date.add_argument("--cell", action="append", default=[], help="D867=10/11/2026")
    p_date.add_argument("--style-row", type=int, default=864)
    p_date.add_argument("--template", default="")
    p_date.set_defaults(func=cmd_set_dates)

    return ap


def main(argv: Optional[List[str]] = None) -> int:
    ap = build_parser()
    args = ap.parse_args(argv)
    try:
        return int(args.func(args))
    except HealthError as e:
        print(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False))
        return 2
    except Exception as e:
        print(json.dumps({"ok": False, "error": f"{type(e).__name__}: {e}"}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
