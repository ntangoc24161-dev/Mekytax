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
            elif ("số hoá đơn" in t or "số hóa đơn" in t or t in {"số hđ", "so hd", "số ct", "so ct"}) and so_i is None:
                so_i = j
            elif ngay_i is None and (
                "ngày lập" in t
                or ("ngày" in t and "phát hành" in t)
                or "ngày hđ" in t or "ngày ct" in t or "ngày tháng" in t
                or t in {"ngày", "ngay"}
            ):
                ngay_i = j
            elif mst_i is None and ("mã số thuế" in t or "mst" in t):
                # HDDT portal & Bang ke: mua lấy MST người bán; bán lấy MST người mua
                if chieu == "mua" and ("người bán" in t or "người xuất" in t or "bên bán" in t or "ncc" in t):
                    mst_i = j
                elif chieu == "ban" and ("người mua" in t or "người nhận" in t or "bên mua" in t or "khách hàng" in t):
                    mst_i = j
                elif "người bán" not in t and "người mua" not in t and "truyền nhận" not in t and "cung cấp giải pháp" not in t and "giải pháp" not in t and "người xuất" not in t and "người nhận" not in t and "bên bán" not in t and "bên mua" not in t:
                    mst_i = j
            elif ten_i is None and (
                (("tên người bán" in t or "tên người xuất" in t or "tên bên bán" in t or "tên ncc" in t or "tên nhà cung cấp" in t) and chieu == "mua")
                or (("tên người mua" in t or "tên người nhận" in t or "tên bên mua" in t or "tên khách hàng" in t) and chieu == "ban")
                or ("tên người" in t and "bán" not in t and "mua" not in t and "xuất" not in t and "nhận" not in t)
                or ("tên đơn vị" in t or "tên nnt" in t or "tên đối tác" in t or "tên khách hàng" in t or "tên nhà cung cấp" in t)
            ):
                ten_i = j
            elif "mặt hàng" in t and mh_i is None:
                mh_i = j
            elif ds_i is None and ("doanh số" in t or "tổng tiền chưa thuế" in t or "tiền hàng" in t or "doanh thu" in t or "chưa thuế" in t or "giá trị" in t):
                ds_i = j
            elif "thuế suất" in t and ts_i is None:
                ts_i = j
            elif thue_i is None and (t in {"thuế gtgt", "tiền thuế"} or "tổng tiền thuế" in t or ("thuế" in t and "thuế suất" not in t and "chưa thuế" not in t and "mã số thuế" not in t and "mst" not in t)):
                thue_i = j
            elif tong_i is None and (
                "tổng tiền thanh toán" in t
                or (("tổng tiền" in t or "tổng cộng" in t) and "thuế" not in t and "chiết khấu" not in t and "phí" not in t and "chưa" not in t)
            ):
                tong_i = j
    portal = so_i is not None and ngay_i is not None and (ds_i is not None or mst_i is not None)
    # Fallback bảng kê mẫu Sơn Tiến Hải (không đè lên map portal HDDT)
    if not portal and chieu == "ban":
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
    elif not portal:
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
