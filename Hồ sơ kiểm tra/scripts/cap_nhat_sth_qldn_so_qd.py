# -*- coding: utf-8 -*-
"""Son Tien Hai: doi QLDN, dien 7092/7093 ngay 26/10/2026, ke toan, TKNH."""
from __future__ import annotations

import copy
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BRAIN = Path(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra")
WORK = BRAIN / "temp_test" / "sth_inspect"
G_BB = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\11. SON TIEN HAI\0. BIEN BAN"
)
G_SO = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA"
    r"\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung"
    r"\Danh sach kiem tra quyet toan pro.xlsm"
)

TONG_IN = WORK / "TONG HOP.docx"
TONG_OUT = WORK / "TONG HOP.patched.docx"
XPHC_IN = WORK / "QD_XPHC.docx"
XPHC_OUT = WORK / "QD_XPHC.patched.docx"
SO_OUT = WORK / "so_patched.xlsm"

QDKT = "7092/QĐ-TCS3"
QDGS = "7093/QĐ-TCS3"
NGAY = "26/10/2026"
KT = "Lê Hoài Phương"
TK = "6722949824 tại Ngân hàng TMCP Đầu tư và Phát triển Việt Nam"

LOG: list[str] = []


def log(msg: str) -> None:
    LOG.append(msg)
    print(msg)


def load_xml(path: Path) -> str:
    return zipfile.ZipFile(path).read("word/document.xml").decode("utf8")


def save_docx(src: Path, dst: Path, xml: str) -> None:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if it.filename == "word/document.xml":
                data = xml.encode("utf8")
            zout.writestr(it, data)


def rep(xml: str, old: str, new: str, expected: int | None = None, label: str = "") -> str:
    n = xml.count(old)
    if expected is not None and n != expected:
        log(f"WARN {label}: expected {expected} got {n}")
        if n == 0:
            return xml
    xml2 = xml.replace(old, new)
    log(f"OK {label or old[:40]!r} x{n}")
    return xml2


def replace_kt_fields(xml: str) -> str:
    fld = re.compile(
        r"<w:r[^>]*>"
        r"(?:<w:rPr>(?P<rpr>.*?)</w:rPr>)?"
        r'<w:fldChar w:fldCharType="begin"/>'
        r"</w:r>"
        r"<w:r[^>]*>"
        r"(?:<w:rPr>.*?</w:rPr>)?"
        r"<w:instrText[^>]*>\s*MERGEFIELD\s+\"?KT\"?\s*</w:instrText>"
        r"</w:r>"
        r"<w:r[^>]*>"
        r"(?:<w:rPr>.*?</w:rPr>)?"
        r'<w:fldChar w:fldCharType="end"/>'
        r"</w:r>",
        re.S,
    )

    def _kt(m: re.Match) -> str:
        rpr = m.group("rpr")
        rpr_xml = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
        return f"<w:r>{rpr_xml}<w:t>{KT}</w:t></w:r>"

    xml2, n = fld.subn(_kt, xml)
    log(f"OK MERGEFIELD KT -> {KT} x{n}")
    if n != 6:
        log(f"WARN KT field count {n} (ky vong 6)")
    return xml2


def rename_qldn(xml: str) -> str:
    phrase = "quản lý kinh doanh"
    pat = "".join(
        re.escape(c) + r"(?:(?:</w:t>)?(?:<[^>]*>)*?(?:<w:t[^>]*>)?)?"
        for c in phrase
    )

    def _sub(m: re.Match) -> str:
        return m.group(0).replace("kinh doanh", "doanh nghiệp")

    xml2, n = re.subn(pat, _sub, xml)
    log(f"OK quan ly kinh doanh -> quan ly doanh nghiep x{n}")
    left = xml2.count("quản lý kinh doanh")
    if left:
        log(f"WARN still contiguous 'quan ly kinh doanh' x{left}")
    return xml2


def patch_tong(xml: str) -> str:
    rpr = (
        '<w:rPr><w:rFonts w:cs="Times New Roman"/><w:sz w:val="24"/>'
        '<w:szCs w:val="24"/><w:lang w:val="nl-NL"/></w:rPr>'
    )

    # 1. Header QDKT / QDGS
    xml = rep(
        xml,
        (
            '<w:t xml:space="preserve">Số:         </w:t></w:r><w:r>'
            + rpr
            + '<w:t xml:space="preserve">   </w:t></w:r>'
            + '<w:r w:rsidR="00D50A77" w:rsidRPr="009D7F91">'
            + rpr
            + '<w:t xml:space="preserve">     /QĐ-TCS3</w:t>'
        ),
        (
            '<w:t xml:space="preserve">Số: </w:t></w:r>'
            + '<w:r w:rsidR="00D50A77" w:rsidRPr="009D7F91">'
            + rpr
            + f'<w:t xml:space="preserve">{QDKT}</w:t>'
        ),
        1,
        "header QDKT 7092",
    )
    xml = rep(
        xml,
        (
            '<w:t xml:space="preserve">Số:           </w:t></w:r><w:r>'
            + rpr
            + '<w:t xml:space="preserve">  </w:t></w:r>'
            + '<w:r w:rsidR="008629EC" w:rsidRPr="009D7F91">'
            + rpr
            + '<w:t xml:space="preserve">   /QĐ-TCS3</w:t>'
        ),
        (
            '<w:t xml:space="preserve">Số: </w:t></w:r>'
            + '<w:r w:rsidR="008629EC" w:rsidRPr="009D7F91">'
            + rpr
            + f'<w:t xml:space="preserve">{QDGS}</w:t>'
        ),
        1,
        "header QDGS 7093",
    )

    # 2. Cap doi QĐGS + QĐKT
    xml = rep(
        xml,
        "Căn cứ Quyết định kiểm tra số              /QĐ-TCS3 ngày ...../...../2026",
        f"Căn cứ Quyết định kiểm tra số {QDKT} ngày {NGAY}",
        1,
        "QDGS can cu QDKT",
    )
    xml = rep(
        xml,
        "được thành lập theo Quyết định số              /QĐ-TCS3 ngày ...../...../2026",
        f"được thành lập theo Quyết định số {QDKT} ngày {NGAY}",
        1,
        "QDGS Dieu 1",
    )

    kh_old = (
        "Thực hiện Quyết định số              /QĐ-TCS3 ngày ..../..../2026 "
        "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
        "đoàn kiểm tra theo QĐ kiểm tra tại doanh nghiệp số              "
        "/QĐ-TCS3 ngày ..../..../2026"
    )
    kh_new = (
        f"Thực hiện Quyết định số {QDGS} ngày {NGAY} "
        "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
        f"đoàn kiểm tra theo QĐ kiểm tra tại doanh nghiệp số {QDKT} ngày {NGAY}"
    )
    xml = rep(xml, kh_old, kh_new, 1, "ke hoach giam sat")

    bc_old = (
        "Thực hiện Quyết định số              /QĐ-TCS3 ngày ..../..../2026 "
        "của Thuế Cơ sở 3 Thành phố Đồng Nai về việc giám sát "
        "hoạt động của Đoàn kiểm tra"
    )
    bc_new = bc_old.replace(
        "số              /QĐ-TCS3 ngày ..../..../2026",
        f"số {QDGS} ngày {NGAY}",
    )
    xml = rep(xml, bc_old, bc_new, 1, "bao cao giam sat")

    lv_old = (
        "theo quyết định kiểm tra tại doanh nghiệp số              /QĐ-TCS3 "
        "ngày ..../..../2026 về việc giám sát Đoàn kiểm tra theo quyết "
        "định giám sát số              /QĐ-TCS3 ngày ..../..../2026."
    )
    lv_new = (
        f"theo quyết định kiểm tra tại doanh nghiệp số {QDKT} "
        f"ngày {NGAY} về việc giám sát Đoàn kiểm tra theo quyết "
        f"định giám sát số {QDGS} ngày {NGAY}."
    )
    xml = rep(xml, lv_old, lv_new, 1, "bb lam viec giam sat")

    # 3. Con lai: tham chieu QDKT
    g_old = "số              /QĐ-TCS3 ngày ..../..../2026"
    n = xml.count(g_old)
    xml = xml.replace(g_old, f"số {QDKT} ngày {NGAY}")
    log(f"OK tham chieu QDKT chung x{n}")

    g_old2 = "số:              /QĐ-TCS3 ngày ..../..../2026"
    n2 = xml.count(g_old2)
    if n2:
        xml = xml.replace(g_old2, f"số: {QDKT} ngày {NGAY}")
        log(f"OK tham chieu QDKT co hai cham x{n2}")

    xml = rep(xml, "SỐ              /QĐ-TCS3", f"SỐ {QDKT}", 1, "bao cao ket qua SO")

    # Nhat ky / o bang con lai
    xml = rep(
        xml,
        '<w:t xml:space="preserve">              /QĐ-TCS3</w:t>',
        f'<w:t xml:space="preserve">{QDKT}</w:t>',
        None,
        "o bang /QĐ-TCS3 con lai",
    )

    xml = rep(
        xml,
        "……../QĐ-TCS3 ngày …../…../",
        f"{QDKT} ngày 26/10/",
        None,
        "giao QD so+ngay",
    )
    xml = rep(
        xml,
        "<w:t>……../QĐ-TCS3</w:t>",
        f"<w:t>{QDKT}</w:t>",
        None,
        "giao QD so con lai",
    )
    xml = rep(
        xml,
        "<w:t>…../…../202</w:t>",
        "<w:t>26/10/202</w:t>",
        None,
        "giao QD ngay con lai",
    )

    # 4. Ke toan
    xml = replace_kt_fields(xml)
    pth = "Phan Thế Hanh"
    xml = rep(
        xml,
        f'<w:t xml:space="preserve">        {pth}                  </w:t>',
        f'<w:t xml:space="preserve">        {pth}                  {KT}</w:t>',
        1,
        "chu ky BBXN thanh vien",
    )
    xml = rep(
        xml,
        f'<w:t xml:space="preserve">{pth}             </w:t>',
        f'<w:t xml:space="preserve">{pth}            {KT}</w:t>',
        1,
        "chu ky BBXN truong doan",
    )

    # 5. Tai khoan
    xml = rep(
        xml,
        '<w:t xml:space="preserve">- Tài khoản số: </w:t>',
        f'<w:t xml:space="preserve">- Tài khoản số: {TK}.</w:t>',
        1,
        "tai khoan ngan hang",
    )
    m = re.search(
        r"<w:p\b[^>]*>(?:(?!</w:p>).)*?MERGEFIELD Tài_khoản_số(?:(?!</w:p>).)*?</w:p>",
        xml,
        re.S,
    )
    if m:
        xml = xml[: m.start()] + xml[m.end() :]
        log("OK xoa doan mergefield tai khoan rong")
    else:
        log("WARN khong thay doan mergefield tai khoan")

    # 6. Doi ten chi tieu
    xml = rename_qldn(xml)

    # 7. Phan cong: them ban hang, bo doanh thu, dien ngay thuc hien
    old_td = (
        '<w:t>-Chi phí quản lý doanh nghiệp</w:t></w:r>'
        '<w:r w:rsidR="00B03632" w:rsidRPr="00EF7270">'
        '<w:rPr><w:rFonts w:cs="Times New Roman"/><w:sz w:val="24"/>'
        '<w:szCs w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    )
    if xml.count(old_td) == 1:
        pbh = (
            '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
            "<w:rPr>"
            '<w:rFonts w:cs="Times New Roman"/><w:noProof/><w:sz w:val="24"/>'
            '<w:szCs w:val="24"/></w:rPr></w:pPr>'
            "<w:r><w:rPr>"
            '<w:rFonts w:cs="Times New Roman"/><w:noProof/><w:sz w:val="24"/>'
            '<w:szCs w:val="24"/></w:rPr>'
            "<w:t>- Chi phí bán hàng</w:t></w:r></w:p>"
        )
        i = xml.find(old_td)
        pstart = xml.rfind("<w:p ", 0, i)
        xml = xml[:pstart] + pbh + xml[pstart:]
        log("OK phan cong TD: them Chi phi ban hang")
    else:
        log(f"WARN phan cong TD fragment x{xml.count(old_td)}")

    dt = (
        "<w:t>- Doanh thu bán hàng và cung cấp dịch vụ</w:t></w:r></w:p>"
        '<w:p w14:paraId="1953EFD5"'
    )
    if dt in xml:
        xml = xml.replace(dt, '</w:r></w:p><w:p w14:paraId="1953EFD5"', 1)
        log("OK phan cong TV: bo doanh thu")
        empty = (
            '<w:r w:rsidR="00696739" w:rsidRPr="00F34EDF"><w:rPr>'
            '<w:rFonts w:cs="Times New Roman"/>'
            '<w:noProof/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:r>'
        )
        if empty in xml:
            xml = xml.replace(empty, "", 1)
            log("OK don run rong o phan cong TV")
    else:
        n_dt = xml.count("Doanh thu bán hàng và cung cấp dịch vụ")
        log(f"WARN phan cong TV doanh thu fragment missing, con {n_dt} chuoi")

    xml = xml.replace(
        "-Chi phí quản lý doanh nghiệp.",
        "- Chi phí bán hàng\n- Chi phí quản lý doanh nghiệp.",
        1,
    ) if xml.count("-Chi phí quản lý doanh nghiệp.") == 1 else xml

    rg = "<w:t>..../..../2026-..../..../2026</w:t>"
    n = xml.count(rg)
    if n:
        xml = xml.replace(rg, "<w:t>30/10/2026-03/11/2026</w:t>")
        log(f"OK phan cong ngay thuc hien 30/10-03/11 x{n}")

    return xml


def verify_tong(xml: str) -> None:
    text = re.sub(r"</w:p>", "\n", xml)
    text = re.sub(r"<[^>]+>", "", text)
    checks = {
        "7092/QĐ-TCS3": 1,
        "7093/QĐ-TCS3": 1,
        "Lê Hoài Phương": 1,
        "6722949824": 1,
        "quản lý doanh nghiệp": 1,
        "quản lý kinh doanh": 0,
        "/QĐ-XPHC": 1,
    }
    log("--- VERIFY TONG HOP ---")
    log(f"count 7092={text.count('7092/QĐ-TCS3')} 7093={text.count('7093/QĐ-TCS3')}")
    log(f"count KT={text.count(KT)} TK={text.count('6722949824')}")
    log(f"count QLDN={text.count('quản lý doanh nghiệp')} QLKinh={text.count('quản lý kinh doanh')}")
    log(f"count blank so TCS3 con '{text.count('              /QĐ-TCS3')}'")
    log(f"QDXPHC header con trong: {'Số:               /QĐ-XPHC' in text or ' /QĐ-XPHC' in text}")
    if "quản lý kinh doanh" in text:
        log("WARN van con 'quan ly kinh doanh'")
    if KT not in text:
        log("WARN thieu ten ke toan")
    if "6722949824" not in text:
        log("WARN thieu so TK")


def patch_xphc(xml: str) -> str:
    xml = rep(
        xml,
        "số              /QĐ-TCS3 ngày ..../..../2026",
        f"số {QDKT} ngày {NGAY}",
        None,
        "XPHC can cu QDKT",
    )
    return xml


def find_qd_ranges(doc) -> tuple:
    body = list(doc._element.body)
    qdkt_start = qdkt_end = qdgs_start = qdgs_end = None
    for i, el in enumerate(body):
        text = "".join(el.itertext())
        if qdkt_start is None and el.tag.endswith("tbl") and "THUẾ THÀNH PHỐ ĐỒNG NAI" in text:
            nxt = "".join("".join(body[j].itertext()) for j in range(i, min(i + 8, len(body))))
            if "Về việc kiểm tra quyết toán giải thể" in nxt or "Về việc kiểm tra" in nxt:
                if "giám sát" not in nxt.lower()[:400]:
                    qdkt_start = i
        if qdkt_start is not None and qdkt_end is None and el.tag.endswith("tbl") and "Nơi nhận" in text and "Điều 3" in text:
            qdkt_end = i
        if qdkt_end is not None and qdgs_start is None and el.tag.endswith("tbl") and "THUẾ THÀNH PHỐ ĐỒNG NAI" in text:
            nxt = "".join("".join(body[j].itertext()) for j in range(i, min(i + 8, len(body))))
            if "Về việc giám sát" in nxt:
                qdgs_start = i
        if qdgs_start is not None and qdgs_end is None and el.tag.endswith("tbl") and "Nơi nhận" in text and "Điều 4" in text:
            qdgs_end = i
            break
    log(f"QD ranges kt={qdkt_start}-{qdkt_end} gs={qdgs_start}-{qdgs_end}")
    return qdkt_start, qdkt_end, qdgs_start, qdgs_end


def export_qd(master_path: Path, start: int, end: int, out_path: Path) -> None:
    from docx import Document

    shutil.copy2(master_path, out_path)
    doc = Document(str(out_path))
    body = doc._element.body
    elems = list(body)
    keep = set(range(start, end + 1))
    for i, el in enumerate(elems):
        if el.tag.endswith("sectPr"):
            continue
        if i not in keep:
            body.remove(el)
    doc.save(str(out_path))
    log(f"OK export {out_path.name} elems {start}-{end}")


def patch_excel(src: Path, dst: Path) -> None:
    shutil.copy2(src, dst)
    wb = load_workbook(dst, keep_vba=True)
    ws = wb["01_DA_BAN_HANH"]
    r = None
    for i in range(5, ws.max_row + 1):
        if str(ws.cell(i, 4).value or "") == "3603859012":
            r = i
            break
    if r is None:
        log("WARN khong thay hang Son Tien Hai tren so")
        return
    ws.cell(r, 6).value = "Số 23, Tổ 1, KP 3A, Phường Long Hưng"
    ws.cell(r, 16).value = KT
    # CCCD ke toan de trong
    ws.cell(r, 21).value = f"TK số 6722949824 tại Ngân hàng TMCP Đầu tư và Phát triển Việt Nam"
    ws.cell(r, 27).value = "-Giá vốn hàng bán\n- Chi phí bán hàng\n- Chi phí quản lý doanh nghiệp"
    ws.cell(r, 28).value = "- Thuế GTGT phải nộp"
    ws.cell(r, 29).value = QDKT
    ws.cell(r, 30).value = QDGS
    ws.cell(r, 31).value = NGAY
    # so tien da chot tu Bang tinh / QD XPHC
    ws.cell(r, 40).value = 13052423
    ws.cell(r, 41).value = 0
    ws.cell(r, 43).value = 0
    ws.cell(r, 44).value = 662730
    ws.cell(r, 46).value = 2610485
    ws.cell(r, 47).value = 0
    ws.cell(r, 50).value = 809868861
    ws.cell(r, 51).value = 0
    wb.save(dst)
    log(f"OK excel row {r}: KT, TKNH, QDKT/QDGS, tien")


def copy_to_g(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    # kill Word lock
    os.system('taskkill /F /IM WINWORD.EXE >nul 2>&1')
    lock = dst.parent / ("~$" + dst.name[:2] + dst.name[2:])
    # standard Word lock is ~$ + first 2 chars stripped from 8.3, actually ~$ + name[2:]
    for p in dst.parent.glob("~$*.docx"):
        try:
            p.unlink()
        except OSError:
            pass
    shutil.copy2(src, dst)
    if not dst.exists() or dst.stat().st_size < 1000:
        raise SystemExit(f"copy failed {dst}")
    log(f"OK copy {dst.name} {dst.stat().st_size} bytes {datetime.fromtimestamp(dst.stat().st_mtime)}")


def main() -> None:
    WORK.mkdir(parents=True, exist_ok=True)
    shutil.copy2(G_BB / "TONG HOP.docx", TONG_IN)
    shutil.copy2(G_BB / "QD XPHC Cong ty Son Tien Hai.docx", XPHC_IN)
    log(f"copied live TONG HOP {TONG_IN.stat().st_size} {datetime.fromtimestamp(TONG_IN.stat().st_mtime)}")

    xml = load_xml(TONG_IN)
    xml = patch_tong(xml)
    verify_tong(xml)
    save_docx(TONG_IN, TONG_OUT, xml)
    log(f"saved {TONG_OUT} {TONG_OUT.stat().st_size}")

    xxml = load_xml(XPHC_IN)
    xxml = patch_xphc(xxml)
    save_docx(XPHC_IN, XPHC_OUT, xxml)

    from docx import Document

    doc = Document(str(TONG_OUT))
    s1, e1, s2, e2 = find_qd_ranges(doc)
    qdkt_path = WORK / "QD Kiem tra tai cty Son Tien Hai.docx"
    qdgs_path = WORK / "QDgiam sat tai cty Son Tien Hai.docx"
    if None not in (s1, e1, s2, e2):
        export_qd(TONG_OUT, s1, e1, qdkt_path)
        export_qd(TONG_OUT, s2, e2, qdgs_path)
    else:
        log("WARN khong tach duoc QĐ rieng")

    patch_excel(G_SO, SO_OUT)

    copy_to_g(TONG_OUT, G_BB / "TONG HOP.docx")
    copy_to_g(XPHC_OUT, G_BB / "QD XPHC Cong ty Son Tien Hai.docx")
    if qdkt_path.exists():
        copy_to_g(qdkt_path, G_BB / "QD Kiem tra tai cty Son Tien Hai.docx")
    if qdgs_path.exists():
        copy_to_g(qdgs_path, G_BB / "QDgiam sat tai cty Son Tien Hai.docx")
    copy_to_g(SO_OUT, G_SO)

    # doc lai G: de xac nhan
    gxml = load_xml(G_BB / "TONG HOP.docx")
    gtext = re.sub(r"<[^>]+>", "", gxml)
    log("--- READBACK G: TONG HOP ---")
    log(f"7092={gtext.count('7092/QĐ-TCS3')} 7093={gtext.count('7093/QĐ-TCS3')}")
    log(f"KT={gtext.count(KT)} TK={gtext.count('6722949824')} QLKinh={gtext.count('quản lý kinh doanh')} QLDN={gtext.count('quản lý doanh nghiệp')}")


if __name__ == "__main__":
    main()
