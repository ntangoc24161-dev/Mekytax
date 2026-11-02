"""Findings -> workbook Excel: mỗi loại thuế một sheet, tô màu theo mức rủi ro."""
from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from models import Finding, HoSoChuan

NAVY = "1B365D"
TEAL = "0F6C8C"
WHITE = "FFFFFF"
BORDER_C = "D5D8DC"
MUTED = "F4F6F7"
TEXT = "1C2833"
HIGH_BADGE = "C0392B"
HIGH_ROW = "FDEDEC"
HIGH_FONT = "922B21"
MED_BADGE = "D68910"
MED_ROW = "FEF9E7"
MED_FONT = "7D6608"
LOW_BADGE = "1E8449"
LOW_ROW = "E8F8F5"
LOW_FONT = "145A32"
SCORE_VERY = "7B241C"

SEV_VI = {"high": "Cao", "medium": "Trung bình", "low": "Thấp"}
SEV_ORDER = {"high": 0, "medium": 1, "low": 2}
TOPIC_SHEETS = (
    ("gtgt", "GTGT", "Thuế GTGT", TEAL),
    ("tndn", "TNDN", "Thuế TNDN", "1A5276"),
    ("tncn", "TNCN", "Thuế TNCN", "6C3483"),
    ("hddt", "Hóa đơn", "Hóa đơn rủi ro", "D35400"),
)
TOPIC_LABEL = {k: label for k, _, label, _ in TOPIC_SHEETS}
DETAIL_HEADERS = (
    "STT",
    "Mức độ",
    "Tiêu chí",
    "Năm",
    "Nội dung bất thường",
    "Căn cứ dữ liệu",
    "Căn cứ tiêu chí",
    "Câu hỏi chất vấn",
)
DETAIL_WIDTHS = (6, 14, 12, 10, 55, 45, 35, 50)
GROUP_HEADERS = ("Tên nhóm", "Số mục", "Cao", "Trung bình", "Thấp")
GROUP_WIDTHS = (55, 12, 12, 14, 12)

_THIN = Border(
    left=Side(style="thin", color=BORDER_C),
    right=Side(style="thin", color=BORDER_C),
    top=Side(style="thin", color=BORDER_C),
    bottom=Side(style="thin", color=BORDER_C),
)
_PATH_TAIL = re.compile(
    r"([^\\/]+\.(?:xlsx|xlsm|xls|xml(?:\.xml)?))(.*)$",
    re.IGNORECASE,
)


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _font(size: int = 11, bold: bool = False, color: str = TEXT, name: str = "Calibri") -> Font:
    return Font(name=name, size=size, bold=bold, color=color)


def _align(h: str = "left", v: str = "center", wrap: bool = True) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _clean(text: str) -> str:
    return (text or "").replace("\u2014", "-")


def _usable(findings: list[Finding]) -> list[Finding]:
    return [f for f in findings if f.severity != "skip"]


def _topic_findings(findings: list[Finding], topic: str) -> list[Finding]:
    rows = [f for f in findings if f.tax_topic == topic and f.id != "m3-score" and f.severity != "skip"]
    rows.sort(key=lambda f: (
        SEV_ORDER.get(f.severity, 9),
        f.nam if f.nam is not None else 10**9,
        f.tieu_chi or "",
        f.title,
    ))
    return rows


def _score_finding(findings: list[Finding]) -> Finding | None:
    return next((f for f in findings if f.id == "m3-score"), None)


def _short_evidence(text: str) -> tuple[str, str | None]:
    raw = _clean(text)
    if not raw:
        return "", None
    m = _PATH_TAIL.search(raw)
    if m:
        short = m.group(1) + m.group(2)
        if short != raw:
            return short, raw
    if "\\" in raw or "/" in raw:
        name = raw.replace("/", "\\").rsplit("\\", 1)[-1]
        if name and name != raw:
            return name, raw
    return raw, None


def _paint(ws: Worksheet, row: int, col: int, value, fill: str, font: Font, align: Alignment,
           border: bool = False, comment: str | None = None) -> None:
    cell = ws.cell(row, col, value)
    cell.fill = _fill(fill)
    cell.font = font
    cell.alignment = align
    if border:
        cell.border = _THIN
    if comment:
        cell.comment = Comment(comment, "PhanTichHoSoThue")


def _merge_paint(ws: Worksheet, r1: int, c1: int, r2: int, c2: int, value, fill: str, font: Font,
                 align: Alignment) -> None:
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.fill = _fill(fill)
            cell.font = font
            cell.alignment = align
    ws.cell(r1, c1).value = value


def _set_widths(ws: Worksheet, widths: tuple[int, ...], start: int = 1) -> None:
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(start + i)].width = w


def _row_height(values: list[str], widths: tuple[int, ...]) -> float:
    lines = 1
    for val, w in zip(values, widths):
        n = max(1, (len(val) + max(w, 8) - 1) // max(int(w * 1.05), 8))
        if n > lines:
            lines = n
    return min(90.0, max(22.0, 16.0 + lines * 13.0))


def _print_setup(ws: Worksheet, header_row: int | None = None) -> None:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.leftMargin = 0.4
    ws.page_setup.rightMargin = 0.4
    ws.page_setup.topMargin = 0.5
    ws.page_setup.bottomMargin = 0.5
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.horizontalCentered = True
    ws.sheet_view.showGridLines = False
    ws.oddHeader.left.text = "&A"
    ws.oddFooter.right.text = "Trang &P / &N"
    if header_row:
        ws.print_title_rows = f"{header_row}:{header_row}"


def _header_row(ws: Worksheet, row: int, headers: tuple[str, ...], fill: str = NAVY) -> None:
    for i, h in enumerate(headers, 1):
        _paint(ws, row, i, h, fill, _font(11, True, WHITE), _align("center", "center"), True)
    ws.row_dimensions[row].height = 24


def _sev_style(sev: str) -> tuple[str, str, str]:
    if sev == "high":
        return HIGH_ROW, HIGH_FONT, HIGH_BADGE
    if sev == "medium":
        return MED_ROW, MED_FONT, MED_BADGE
    return LOW_ROW, LOW_FONT, LOW_BADGE


def _write_detail(ws: Worksheet, start_row: int, rows: list[Finding]) -> int:
    _header_row(ws, start_row, DETAIL_HEADERS)
    r = start_row + 1
    for i, f in enumerate(rows, 1):
        row_fill, row_font, badge = _sev_style(f.severity)
        evidence, full = _short_evidence(f.evidence)
        vals = [
            str(i),
            SEV_VI.get(f.severity, f.severity),
            f.tieu_chi or "",
            "" if f.nam is None else str(f.nam),
            _clean(f.delta),
            evidence,
            _clean(f.legal_basis),
            _clean(f.question_for_nnt),
        ]
        for c, val in enumerate(vals, 1):
            if c == 2:
                _paint(ws, r, c, val, badge, _font(11, True, WHITE), _align("center"), True)
            elif c in (1, 3, 4):
                _paint(ws, r, c, val, row_fill, _font(11, False, row_font), _align("center"), True)
            else:
                comment = full if c == 6 else None
                _paint(ws, r, c, val, row_fill, _font(11, False, row_font), _align("left", "top"), True, comment)
        ws.row_dimensions[r].height = _row_height(vals, DETAIL_WIDTHS)
        r += 1
    last = r - 1
    if rows:
        ws.auto_filter.ref = f"A{start_row}:H{last}"
        ws.freeze_panes = f"A{start_row + 1}"
    return last


def _write_groups(ws: Worksheet, start_row: int, rows: list[Finding]) -> int:
    _paint(ws, start_row, 1, "Nhóm phát hiện", MUTED, _font(12, True, NAVY), _align("left"))
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)
    header = start_row + 1
    _header_row(ws, header, GROUP_HEADERS, TEAL)
    titles = []
    seen: set[str] = set()
    for f in rows:
        if f.title not in seen:
            titles.append(f.title)
            seen.add(f.title)
    counts: dict[str, Counter] = {}
    for f in rows:
        counts.setdefault(f.title, Counter())[f.severity] += 1
    titles.sort(key=lambda t: (-counts[t]["high"], -sum(counts[t].values()), t))
    r = header + 1
    for title in titles:
        c = counts[title]
        n_h, n_m, n_l = c["high"], c["medium"], c["low"]
        total = n_h + n_m + n_l
        fill = HIGH_ROW if n_h else (MED_ROW if n_m else LOW_ROW)
        font_c = HIGH_FONT if n_h else (MED_FONT if n_m else LOW_FONT)
        vals = [_clean(title), str(total), str(n_h), str(n_m), str(n_l)]
        for i, val in enumerate(vals, 1):
            h = "left" if i == 1 else "center"
            this_fill = HIGH_ROW if i == 3 and n_h else fill
            _paint(ws, r, i, int(val) if i > 1 else val, this_fill, _font(11, i > 1, font_c), _align(h), True)
        ws.row_dimensions[r].height = 22
        r += 1
    return r


def _banner_text(hs: HoSoChuan, extra: str = "") -> str:
    nam = ", ".join(str(y) for y in hs.nam) if hs.nam else "(không rõ kỳ)"
    ten = hs.ten_nnt or "(chưa rõ tên NNT)"
    line1 = f"Phân tích hồ sơ kiểm tra thuế  |  {ten}"
    line2 = f"MST {hs.mst or '(không rõ)'}   ·   Kỳ {nam}   ·   {len(hs.to_khai)} tờ khai/BCTC XML   ·   {len(hs.hoa_don)} dòng hóa đơn"
    if extra:
        return f"{line1}\n{line2}\n{extra}"
    return f"{line1}\n{line2}"


def _build_summary(ws: Worksheet, hs: HoSoChuan, findings: list[Finding]) -> None:
    ws.sheet_properties.tabColor = NAVY
    _set_widths(ws, (18, 16, 16, 16, 16, 22, 28, 42))
    usable = [f for f in _usable(findings) if f.id != "m3-score"]
    n_h = sum(1 for f in usable if f.severity == "high")
    n_m = sum(1 for f in usable if f.severity == "medium")
    n_l = sum(1 for f in usable if f.severity == "low")
    _merge_paint(ws, 1, 1, 2, 8, _banner_text(hs), NAVY, _font(16, True, WHITE), _align("left", "center"))
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 22

    score = _score_finding(findings)
    if score:
        title = _clean(score.title)
        band_u = title.upper()
        if "RẤT CAO" in band_u:
            sc_fill, sc_font = SCORE_VERY, WHITE
        elif score.severity == "high":
            sc_fill, sc_font = HIGH_BADGE, WHITE
        elif score.severity == "medium":
            sc_fill, sc_font = MED_BADGE, WHITE
        else:
            sc_fill, sc_font = LOW_BADGE, WHITE
        _merge_paint(
            ws, 4, 1, 4, 8,
            f"{title}   ·   {_clean(score.delta)}",
            sc_fill, _font(13, True, sc_font), _align("left", "center"),
        )
        ws.row_dimensions[4].height = 28
    else:
        _merge_paint(ws, 4, 1, 4, 8, "Chưa có điểm rủi ro tổng hợp", MUTED, _font(12, True, TEXT), _align("left"))

    kpis = (
        (1, 2, "CAO", str(n_h), HIGH_BADGE),
        (3, 4, "TRUNG BÌNH", str(n_m), MED_BADGE),
        (5, 6, "THẤP", str(n_l), LOW_BADGE),
    )
    for c1, c2, label, num, color in kpis:
        _merge_paint(ws, 6, c1, 6, c2, label, color, _font(10, True, WHITE), _align("center"))
        _merge_paint(ws, 7, c1, 7, c2, num, color, _font(22, True, WHITE), _align("center"))
    ws.row_dimensions[6].height = 18
    ws.row_dimensions[7].height = 32
    _merge_paint(
        ws, 6, 7, 7, 8,
        f"Tổng phát hiện\n{len(usable)} mục",
        NAVY, _font(12, True, WHITE), _align("center"),
    )

    _paint(ws, 9, 1, "Phát hiện theo loại thuế", MUTED, _font(12, True, NAVY), _align("left"))
    ws.merge_cells("A9:E9")
    matrix_h = ("Loại thuế", "Cao", "Trung bình", "Thấp", "Tổng")
    _header_row(ws, 10, matrix_h)
    r = 11
    totals = Counter()
    for key, _sheet, label, _tab in TOPIC_SHEETS:
        group = _topic_findings(findings, key)
        c = Counter(f.severity for f in group)
        vals = [label, c["high"], c["medium"], c["low"], len(group)]
        for i, val in enumerate(vals, 1):
            fill = HIGH_ROW if i == 2 and c["high"] else MUTED
            font_c = HIGH_FONT if i == 2 and c["high"] else TEXT
            _paint(ws, r, i, val, fill, _font(11, i == 1 or (i == 2 and c["high"]), font_c),
                   _align("left" if i == 1 else "center"), True)
        totals["high"] += c["high"]
        totals["medium"] += c["medium"]
        totals["low"] += c["low"]
        totals["n"] += len(group)
        r += 1
    sum_vals = ["Tổng", totals["high"], totals["medium"], totals["low"], totals["n"]]
    for i, val in enumerate(sum_vals, 1):
        _paint(ws, r, i, val, NAVY, _font(11, True, WHITE), _align("left" if i == 1 else "center"), True)

    pri_row = r + 2
    _paint(ws, pri_row, 1, "Ưu tiên giải trình", MUTED, _font(12, True, NAVY), _align("left"))
    ws.merge_cells(start_row=pri_row, start_column=1, end_row=pri_row, end_column=8)
    pri_h = ("STT", "Loại thuế", "Tiêu chí", "Năm", "Nội dung bất thường", "Câu hỏi chất vấn")
    _header_row(ws, pri_row + 1, pri_h)
    highs = [f for f in usable if f.severity == "high"]
    if len(highs) < 8:
        highs = highs + [f for f in usable if f.severity == "medium"]
    highs = highs[:8]
    pr = pri_row + 2
    if not highs:
        _merge_paint(ws, pr, 1, pr, 6, "Không có mục ưu tiên.", MUTED, _font(11, False, TEXT), _align("left"))
        pr += 1
    else:
        for i, f in enumerate(highs, 1):
            row_fill, row_font, badge = _sev_style(f.severity)
            vals = [
                str(i),
                TOPIC_LABEL.get(f.tax_topic, f.tax_topic),
                f.tieu_chi or "",
                "" if f.nam is None else str(f.nam),
                _clean(f.delta),
                _clean(f.question_for_nnt),
            ]
            widths = (6, 16, 12, 10, 55, 50)
            for c, val in enumerate(vals, 1):
                if c == 1:
                    _paint(ws, pr, c, val, badge, _font(11, True, WHITE), _align("center"), True)
                elif c in (2, 3, 4):
                    _paint(ws, pr, c, val, row_fill, _font(11, False, row_font), _align("center"), True)
                else:
                    _paint(ws, pr, c, val, row_fill, _font(11, False, row_font), _align("left", "top"), True)
            ws.row_dimensions[pr].height = _row_height(vals, widths)
            pr += 1

    note = pr + 1
    _merge_paint(
        ws, note, 1, note, 8,
        "Màu: đỏ = Cao (ưu tiên chất vấn) · cam = Trung bình · xanh = Thấp.  "
        "Ngưỡng lấy từ sheet TUDIEN tool 4.4. Tiêu chí cần TMS (TC24) không chạy.",
        MUTED, _font(10, False, "5D6D7E"), _align("left"),
    )
    ws.row_dimensions[note].height = 28
    _print_setup(ws, pri_row + 1)
    ws.freeze_panes = "A3"


def _build_topic(ws: Worksheet, hs: HoSoChuan, findings: list[Finding],
                 topic: str, title: str, tab: str) -> None:
    ws.sheet_properties.tabColor = tab
    _set_widths(ws, DETAIL_WIDTHS)
    rows = _topic_findings(findings, topic)
    n_h = sum(1 for f in rows if f.severity == "high")
    n_m = sum(1 for f in rows if f.severity == "medium")
    n_l = sum(1 for f in rows if f.severity == "low")
    extra = f"{title}   ·   {len(rows)} mục  (Cao {n_h} · Trung bình {n_m} · Thấp {n_l})"
    _merge_paint(ws, 1, 1, 2, 8, _banner_text(hs, extra), NAVY, _font(14, True, WHITE), _align("left", "center"))
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 22

    if not rows:
        _merge_paint(
            ws, 4, 1, 5, 8,
            "Không phát hiện mục thuộc chuyên đề này.",
            MUTED, _font(12, False, "5D6D7E"), _align("left", "center"),
        )
        ws.row_dimensions[4].height = 28
        ws.row_dimensions[5].height = 20
        _print_setup(ws)
        ws.freeze_panes = "A3"
        return

    after_groups = _write_groups(ws, 4, rows)
    detail_row = after_groups + 1
    _paint(ws, detail_row, 1, "Chi tiết phát hiện", MUTED, _font(12, True, NAVY), _align("left"))
    ws.merge_cells(start_row=detail_row, start_column=1, end_row=detail_row, end_column=8)
    header = detail_row + 1
    _write_detail(ws, header, rows)
    _print_setup(ws, header)


def _build_phu_luc(ws: Worksheet, hs: HoSoChuan) -> None:
    ws.sheet_properties.tabColor = "7F8C8D"
    _set_widths(ws, (14, 100))
    _merge_paint(
        ws, 1, 1, 2, 2,
        "Phụ lục ingest  |  cảnh báo đọc file và mục đã bỏ qua",
        NAVY, _font(14, True, WHITE), _align("left", "center"),
    )
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 20
    _header_row(ws, 4, ("Loại", "Nội dung"))
    items = [("Cảnh báo", w) for w in hs.warnings] + [("Bỏ", s) for s in hs.skipped]
    if not items:
        _merge_paint(ws, 5, 1, 5, 2, "Không có cảnh báo ingest.", MUTED, _font(11, False, TEXT), _align("left"))
        _print_setup(ws, 4)
        ws.freeze_panes = "A5"
        return
    r = 5
    for kind, text in items:
        fill = MED_ROW if kind == "Cảnh báo" else MUTED
        _paint(ws, r, 1, kind, fill, _font(11, True, TEXT), _align("center"), True)
        _paint(ws, r, 2, _clean(text), fill, _font(11, False, TEXT), _align("left", "top"), True)
        ws.row_dimensions[r].height = _row_height([kind, _clean(text)], (14, 100))
        r += 1
    ws.auto_filter.ref = f"A4:B{r - 1}"
    ws.freeze_panes = "A5"
    _print_setup(ws, 4)


def to_workbook(hs: HoSoChuan, findings: list[Finding]) -> Workbook:
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "Tóm tắt"
    _build_summary(ws0, hs, findings)
    for key, sheet, title, tab in TOPIC_SHEETS:
        ws = wb.create_sheet(sheet)
        _build_topic(ws, hs, findings, key, title, tab)
    _build_phu_luc(wb.create_sheet("Phụ lục"), hs)
    wb.properties.title = f"Phân tích hồ sơ thuế {hs.mst}"
    wb.properties.creator = "PhanTichHoSoThue"
    return wb


def write_xlsx(path: Path, hs: HoSoChuan, findings: list[Finding]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = to_workbook(hs, findings)
    wb.save(path)
    return path
