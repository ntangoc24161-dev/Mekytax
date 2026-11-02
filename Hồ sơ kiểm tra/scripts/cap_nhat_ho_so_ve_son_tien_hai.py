# -*- coding: utf-8 -*-
"""
Cap nhat ho so ve (TONG HOP.docx) vu Son Tien Hai (MST 3603859012),
lay nen phap ly va cach dien tu bo Mo Fa (MST 3603961087) lam phoi.

Chay: python scripts/cap_nhat_ho_so_ve_son_tien_hai.py
"""
from __future__ import annotations

import json
import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path

import docx

sys.stdout.reconfigure(encoding="utf-8")

VU = Path(r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
          r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\11. SON TIEN HAI\0. BIEN BAN")
LIVE = VU / "TONG HOP.docx"
BRAIN = Path(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra")
OUT = BRAIN / "output" / "3603859012_HO-SO-VE_son-tien-hai"

WT = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.S)
WT_SPLIT = re.compile(r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.S)

# 4 noi dung kiem tra - giu dung nhu DI-03 da trinh ky
NOI_DUNG = ("- Giá vốn hàng bán; - Chi phí quản lý kinh doanh; "
            "- Doanh thu bán hàng và cung cấp dịch vụ; - Thuế GTGT phải nộp.")

# Muc III BBKT: nghia vu NSNN. Suy ra tu ket luan 1.2 (da nop theo ke khai)
# va 1.5 (khong con no NSNN): (3)=0, (6)=(4), (7)=(3)+(4)+(5)-(6).
NSNN = {
    "2022": ("0", "0", "0", "0", "0"),
    "2023": ("0", "0", "0", "0", "0"),
    "2024": ("0", "11.591.829", "0", "11.591.829", "0"),
    "2025": ("0", "5.814.545", "10.180.180", "5.814.545", "10.180.180"),
    "2026": ("0", "0", "2.872.243", "0", "2.872.243"),
}

# Nhat ky kiem tra: moc thoi gian theo so JAVIS PRO
NHAT_KY = {
    "Tiến hành kiểm tra": "30/10/2026",
    "Kết thúc kiểm tra": "03/11/2026",
    "Lập biên bản thông báo kết quả KT": "04/11/2026",
    "Báo cáo kết quả kiểm tra": "09/11/2026",
    "Chuyển kết quả KT chờ xử lý": "09/11/2026",
}

RUN_REPL = [
    # --- 1. Nen phap ly, lay theo bo Mo Fa ---
    ("Luật Thanh tra ngày 14 tháng 11 năm 2022",
     "Luật Thanh tra ngày 25 tháng 06 năm 2025"),
    ("13/10/2022", "13/10/2020"),
    ("1377/QĐ-CT ngày 30/06/2026", "1377/QĐ-CT ngày 30/06/2025"),
    # --- 2. Dong bo cach goi loai cuoc kiem tra theo Mo Fa ---
    # (QD va bien ban goi la "kiem tra thue"; ly do giai the giu o To trinh)
    ("về việc kiểm tra quyết toán giải thể tại", "về việc kiểm tra thuế tại"),
    ("về việc kiểm tra quyết toán tại", "về việc kiểm tra thuế tại"),
    ("về việc kiểm tra thuế giải thể tại", "về việc kiểm tra thuế tại"),
    ("Biên bản kiểm tra quyết toán ký ngày", "Biên bản kiểm tra ký ngày"),
    # --- 3. Loi chinh ta, the thuc ---
    ("Thuế cơ sổ", "Thuế cơ sở"),
    ("được lập được lập", "được lập"),
    ("Căn chứ theo Quyết định", "Căn cứ theo Quyết định"),
    ("Quy chế giám, sát hoạt động", "Quy chế giám sát hoạt động"),
    ("kính  xin ý kiến chỉ đạo của Lãnh đạo Cục thuế",
     "kính xin ý kiến chỉ đạo của Lãnh đạo Thuế cơ sở"),
    ("Lãnh đạo Cục;", "- Lãnh đạo Thuế Cơ sở 3;"),
    # --- 4. BB giao nhan ho so: ben nhan khong the trung ben giao ---
    ("Ông/Bà: Nguyễn Thị Ánh Ngọc\t- Chức vụ: Kiểm tra viên",
     "Ông/Bà: Nguyễn Thành Đồng\t- Chức vụ: Kiểm tra viên thuế"),
]

# Nhan checkbox can tick (theo trang thai bo Mo Fa)
TICK_LABELS = [
    "Công bố đúng thời hạn", "Báo cáo người giám sát đúng thời hạn",
    "Không có phản ánh", "Không phát sinh",
]
# Nhung dong co 2 lua chon nam o 2 run rieng: tick run dau, de trong run sau
TICK_FIRST_OF = [
    ("Có", "Không"), ("Đúng", "Không đúng"), ("Tốt", "Không tốt"),
]


def wt_sub(xml: str, old: str, new: str) -> tuple[str, int]:
    """Thay the chi trong noi dung the <w:t> that (khong dinh vao <w:tab/>)."""
    hits = 0

    def fix(m):
        nonlocal hits
        body = m.group(2)
        if old not in body:
            return m.group(0)
        hits += body.count(old)
        return m.group(1) + body.replace(old, new) + m.group(3)

    return WT_SPLIT.sub(fix, xml), hits


def set_para_runs(xml: str, para_text: str, values: list[str]) -> tuple[str, int]:
    """Tim doan van co text ghep dung bang para_text, dat lai noi dung tung run."""
    out, pos, hits = [], 0, 0
    for m in re.finditer(r"<w:p\b.*?</w:p>", xml, re.S):
        block = m.group(0)
        if "".join(WT.findall(block)) != para_text:
            continue
        idx = [0]

        def put(mm):
            i = idx[0]
            idx[0] += 1
            if i >= len(values):
                return mm.group(0)
            return mm.group(1) + values[i] + mm.group(3)

        new_block = WT_SPLIT.sub(put, block)
        out.append(xml[pos:m.start()])
        out.append(new_block)
        pos = m.end()
        hits += 1
    out.append(xml[pos:])
    return "".join(out), hits


def fix_ngay_bao_cao(xml: str) -> tuple[str, int]:
    """Ngay bao cao 23/10 -> 09/11/2026 (theo so JAVIS PRO)."""
    out, pos, hits = [], 0, 0
    pat = re.compile(r'Ngày_báo_cáo"\\@"(d{1,2}|M{1,2})"')
    for m in pat.finditer(xml):
        fmt = m.group(1)
        tail = xml[m.end():m.end() + 1500]
        tm = re.search(r"(<w:t(?:\s[^>]*)?>)(\s*)(\d{1,2})(\s*)(</w:t>)", tail)
        if not tm:
            continue
        new = "09" if fmt.startswith("d") else "11"
        s, e = m.end() + tm.start(), m.end() + tm.end()
        out.append(xml[pos:s])
        out.append(tm.group(1) + tm.group(2) + new + tm.group(4) + tm.group(5))
        pos = e
        hits += 1
    out.append(xml[pos:])
    return "".join(out), hits


def pha_xml(path: Path) -> None:
    ET.fromstring(zipfile.ZipFile(path).read("word/document.xml").decode("utf8"))


def main() -> None:
    assert LIVE.exists(), f"khong thay {LIVE}"
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    bak = VU / f"TONG HOP.bak-truoc-phoi-mofa-{stamp}.docx"
    shutil.copy2(LIVE, bak)
    pha_xml(LIVE)
    print(f"[bak] {bak.name}")

    zin = zipfile.ZipFile(LIVE)
    parts = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    xml = parts["word/document.xml"].decode("utf8")
    log: list[str] = []

    def note(n: int, what: str) -> None:
        log.append(f"{n}x  {what}")

    for old, new in RUN_REPL:
        xml, n = wt_sub(xml, old, new)
        if n:
            note(n, f"{old[:48]!r} -> {new[:48]!r}")

    # checkbox: chuan hoa sang Unicode roi tick theo bo Mo Fa
    xml, n = wt_sub(xml, "\u25a1", "\u2610")
    note(n, "checkbox vuong -> U+2610")
    for lab in TICK_LABELS:
        xml, n = wt_sub(xml, f"{lab} \u2610", f"{lab} \u2611")
        note(n, f"tick {lab!r}")
    for yes, no in TICK_FIRST_OF:
        xml, n = wt_sub(xml, f"{yes} \u2610", f"{yes} \u2611")
        note(n, f"tick {yes!r} (doi lap {no!r})")
    # muc 9 cua Bao cao giam sat: phoi thieu han o checkbox o lua chon thu 3
    xml, n = wt_sub(xml, "Không có thông tin thu thập từ bên ngoài",
                    "Không có thông tin thu thập từ bên ngoài \u2611")
    note(n, "them + tick o checkbox muc 9 lua chon 3")

    # Bao cao ket qua: hai muc cung mang so III -> muc sau thanh IV
    xml, n = WT_SPLIT.subn(
        lambda m: (m.group(1) + "IV. " + m.group(3))
        if m.group(2) == "III. " else m.group(0), xml)
    note(sum(1 for _ in re.finditer(r"<w:t(?:\s[^>]*)?>IV\. </w:t>", xml)),
         "danh so 'III. ' -> 'IV. ' (Kien nghi xu ly)")

    xml, n = fix_ngay_bao_cao(xml)
    note(n, "ngay bao cao -> 09/11/2026")

    # Dieu 1 QDKT va Muc II Bao cao ket qua: dien 4 noi dung that
    xml, n = set_para_runs(xml, "Kiểm tra chấp hành pháp luật thuế", [NOI_DUNG])
    note(n, "Dieu 1 QDKT: dien 4 noi dung kiem tra")
    xml, n = set_para_runs(xml, "Kiểm tra chấp hành pháp luật thuế.",
                           [NOI_DUNG, ""])
    note(n, "Muc II Bao cao ket qua: dien 4 noi dung kiem tra")

    # Tieu de QDKT: bo chu "quyet toan giai the" (giu o To trinh nhu Mo Fa)
    xml, n = set_para_runs(xml, "Về việc kiểm tra quyết toán giải thể ",
                           ["Về việc kiểm tra ", "thuế", "", " "])
    note(n, "tieu de QDKT -> 'Ve viec kiem tra thue'")

    parts["word/document.xml"] = xml.encode("utf8")
    with zipfile.ZipFile(LIVE, "w", zipfile.ZIP_DEFLATED) as zo:
        for name, blob in parts.items():
            zo.writestr(name, blob)
    pha_xml(LIVE)
    print(f"[xml] da ap {len(log)} nhom thay the, XML con hop le")

    # ------- phan bang: Muc III NSNN va Nhat ky -------
    doc = docx.Document(str(LIVE))
    tbl_log: list[str] = []

    def set_cell(cell, value: str) -> None:
        paras = cell.paragraphs
        target = paras[0]
        if not target.runs:
            target.add_run(value)
        else:
            target.runs[0].text = value
            for r in target.runs[1:]:
                r.text = ""
        for p in paras[1:]:
            for r in p.runs:
                r.text = ""

    for t in doc.tables:
        head = " ".join(c.text.strip() for c in t.rows[0].cells)
        if "Số kỳ trước chuyển sang" in head:
            for row in t.rows:
                nam = row.cells[1].text.strip()
                if nam in NSNN and len(row.cells) >= 7:
                    v = NSNN[nam]
                    for off, val in zip((2, 3, 4, 5, 6), v):
                        set_cell(row.cells[off], val)
                    tbl_log.append(f"NSNN {nam}: {' | '.join(v)}")
        if "Số hiệu văn bản" in head:
            for row in t.rows:
                viec = row.cells[1].text.strip()
                if viec in NHAT_KY:
                    set_cell(row.cells[3], NHAT_KY[viec])
                    tbl_log.append(f"Nhat ky '{viec}' -> {NHAT_KY[viec]}")
        if "Nội dung công việc được giao" in head:
            for row in t.rows:
                if row.cells[3].text.strip().startswith("..../..../2026"):
                    set_cell(row.cells[3], "30/10/2026 - 03/11/2026")
                    tbl_log.append("Nhat ky phan cong: 30/10/2026 - 03/11/2026")
        # VB tu choi nhan lai tien nop thua: dien MST theo tung o
        if t.rows and "Mã số thuế" in t.rows[0].cells[0].text:
            cells = [c for r in t.rows for c in r.cells]
            empties = [c for c in cells if not c.text.strip()]
            if len(empties) >= 10:
                for c, d in zip(empties[:10], "3603859012"):
                    set_cell(c, d)
                tbl_log.append("VB tu choi nop thua: dien MST 3603859012")

    doc.save(str(LIVE))
    pha_xml(LIVE)
    print(f"[tbl] da dien {len(tbl_log)} o bang")
    for line in tbl_log:
        print("   ", line)

    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "nhat-ky-cap-nhat.json").write_text(
        json.dumps({"thoi_diem": stamp, "nguon_phoi": "Mo Fa 3603961087",
                    "ban_sao_luu": bak.name,
                    "thay_the_van_ban": log, "dien_bang": tbl_log},
                   ensure_ascii=False, indent=2), encoding="utf8")
    for line in log:
        print("   ", line)


if __name__ == "__main__":
    main()


def vet_con_lai() -> None:
    """Hai cho con lai: ben nhan BB giao nhan ho so va o MST cua VB tu choi."""
    import copy  # noqa: F401  (giu import cuc bo)

    doc = docx.Document(str(LIVE))
    done: list[str] = []

    def set_cell(cell, value: str) -> None:
        p = cell.paragraphs[0]
        if not p.runs:
            p.add_run(value)
        else:
            p.runs[0].text = value
            for r in p.runs[1:]:
                r.text = ""

    # o MST theo tung chu so trong VAN BAN TU CHOI NHAN LAI SO TIEN NOP THUA
    for t in doc.tables:
        if len(t.rows) < 2 or t.rows[1].cells[0].text.strip() != "Mã số thuế:":
            continue
        o = t.rows[1].cells
        if all(not c.text.strip() for c in o[1:11]):
            for cell, d in zip(o[1:11], "3603859012"):
                set_cell(cell, d)
            done.append("VB tu choi nop thua: dien MST 3603859012 theo tung o")

    doc.save(str(LIVE))

    # ben nhan ban giao ho so: phai la thanh vien, khong trung truong doan
    zin = zipfile.ZipFile(LIVE)
    parts = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    xml = parts["word/document.xml"].decode("utf8")

    xml, n = set_para_runs(
        xml, "Ông/Bà: Nguyễn Thị Ánh Ngọc- Chức vụ: Kiểm tra viên",
        ["Ông/", "Bà: Nguyễn Thành Đồng", "- Chức vụ: ", "Kiểm tra", " viên thuế"])
    if n:
        done.append("BB giao nhan ho so: ben nhan -> Nguyen Thanh Dong")
    xml, n = set_para_runs(
        xml, "Nguyễn Thị Ánh NgọcNguyễn Thị Ánh Ngọc ",
        ["Nguyễn Thành Đồng", "Nguyễn Thị Ánh Ngọc", " "])
    if n:
        done.append("BB giao nhan ho so: chu ky ben nhan -> Nguyen Thanh Dong")

    parts["word/document.xml"] = xml.encode("utf8")
    with zipfile.ZipFile(LIVE, "w", zipfile.ZIP_DEFLATED) as zo:
        for name, blob in parts.items():
            zo.writestr(name, blob)
    pha_xml(LIVE)
    for line in done:
        print("   ", line)
    return done
