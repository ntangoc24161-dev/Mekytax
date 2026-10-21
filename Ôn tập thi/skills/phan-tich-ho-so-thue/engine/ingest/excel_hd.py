"""Đọc bảng kê hóa đơn Excel (sheet 2.BR / 4.MV)."""
from __future__ import annotations

from datetime import datetime, date
from pathlib import Path

from models import HoaDon

SKIP_SHEET = ("ptdn",)


def _norm(s) -> str:
    return str(s or "").strip().casefold()


def parse_bang_ke(path: Path) -> list[HoaDon]:
    try:
        import openpyxl
    except ImportError:
        return []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return []
    out: list[HoaDon] = []
    for name in wb.sheetnames:
        ln = _norm(name)
        if any(x in ln for x in SKIP_SHEET):
            continue
        chieu = ""
        if "br" in ln or "ban ra" in ln or "bán ra" in ln:
            chieu = "ban"
        elif "mv" in ln or "mua" in ln:
            chieu = "mua"
        else:
            continue
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        out.extend(_sheet_rows(path, name, chieu, rows))
    return out


def _sheet_rows(path: Path, sheet: str, chieu: str, rows: list) -> list[HoaDon]:
    header_i = None
    for i, row in enumerate(rows[:20]):
        joined = " ".join(_norm(c) for c in row if c)
        if "số hoá đơn" in joined or "số hóa đơn" in joined:
            header_i = i
            break
    if header_i is None:
        return []
    # Cột: tìm index theo hàng tiêu đề phụ (Kỳ BCT / Số / Ngày)
    ky_i = so_i = ngay_i = mst_i = ten_i = ds_i = ts_i = thue_i = tong_i = mh_i = None
    scan = rows[header_i : header_i + 4]
    for row in scan:
        for j, c in enumerate(row):
            t = _norm(c)
            if t in {"kỳ bct", "ky bct"} and ky_i is None:
                ky_i = j
            elif ("số hoá đơn" in t or "số hóa đơn" in t) and so_i is None:
                so_i = j
            elif "ngày" in t and "phát hành" in t and ngay_i is None:
                ngay_i = j
            elif "mã số thuế" in t and mst_i is None:
                mst_i = j
            elif ("tên người" in t or "tên người mua" in t or "tên người bán" in t) and ten_i is None:
                ten_i = j
            elif "mặt hàng" in t and mh_i is None:
                mh_i = j
            elif "doanh số" in t and ds_i is None:
                ds_i = j
            elif "thuế suất" in t and ts_i is None:
                ts_i = j
            elif t in {"thuế gtgt"} and thue_i is None:
                thue_i = j
            elif ("tổng tiền" in t or "tổng cộng" in t) and tong_i is None:
                tong_i = j
    # Fallback bảng kê mẫu Sơn Tiến Hải
    if chieu == "ban":
        ky_i = ky_i if ky_i is not None else 2
        so_i = so_i if so_i is not None else 3
        ngay_i = ngay_i if ngay_i is not None else 4
        ten_i = ten_i if ten_i is not None else 6
        mst_i = mst_i if mst_i is not None else 7
        mh_i = mh_i if mh_i is not None else 8
        ds_i = ds_i if ds_i is not None else 9
        ts_i = ts_i if ts_i is not None else 10
        thue_i = thue_i if thue_i is not None else 11
        tong_i = tong_i if tong_i is not None else 12
    else:
        ky_i = ky_i if ky_i is not None else 1
        so_i = so_i if so_i is not None else 2
        ngay_i = ngay_i if ngay_i is not None else 3
        ten_i = ten_i if ten_i is not None else 5
        mst_i = mst_i if mst_i is not None else 6
        mh_i = mh_i if mh_i is not None else 7
        ds_i = ds_i if ds_i is not None else 8
        ts_i = ts_i if ts_i is not None else 9
        thue_i = thue_i if thue_i is not None else 10
        tong_i = tong_i if tong_i is not None else 11

    out: list[HoaDon] = []
    start = header_i + 1
    for ridx, row in enumerate(rows[start:], start=start + 1):
        if not row:
            continue
        so = _cell(row, so_i)
        if not so or _norm(so).startswith("hàng hoá") or _norm(so).startswith("cộng"):
            continue
        so_s = str(so).strip()
        if _norm(so_s) in {"stt", "q1", "q2", "q3", "q4"} or so_s.lower().startswith("q"):
            continue
        ngay = _fmt_ngay(_cell(row, ngay_i))
        if not ngay:
            continue
        try:
            ds = float(_cell(row, ds_i) or 0)
        except (TypeError, ValueError):
            continue
        hd = HoaDon(
            chieu=chieu,
            ky=str(_cell(row, ky_i) or ""),
            so=str(so).strip(),
            ngay=ngay,
            mst_dt=_mst(_cell(row, mst_i)),
            ten_dt=str(_cell(row, ten_i) or "").strip(),
            mat_hang=str(_cell(row, mh_i) or "").strip(),
            doanh_so=ds,
            thue_suat=_rate(_cell(row, ts_i)),
            thue=_float(_cell(row, thue_i)),
            tong=_float(_cell(row, tong_i)),
            file=str(path),
            sheet=sheet,
            row=ridx,
        )
        out.append(hd)
    return out


def _cell(row, i):
    if i is None or i >= len(row):
        return None
    return row[i]


def _float(v) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _rate(v) -> float:
    x = _float(v)
    if x > 1:
        x = x / 100.0
    return x


def _mst(v) -> str:
    s = str(v or "").replace(" ", "").replace(".0", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _fmt_ngay(v) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date().isoformat()
        except ValueError:
            continue
    return s
