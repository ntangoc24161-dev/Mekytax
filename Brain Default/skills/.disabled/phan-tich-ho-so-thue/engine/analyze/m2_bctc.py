"""M2: 28 tiêu chí BCTC / TNDN / TNCN (sheet TUDIEN tool 4.4). Bỏ TC cần TMS."""
from __future__ import annotations

from models import Finding, HoSoChuan, ToKhai

# Ngưỡng ghi trên TUDIEN
UNG_TRUOC = 0.05
HTK_XDCB = 0.30


def run_m2(hs: HoSoChuan) -> list[Finding]:
    out: list[Finding] = []
    bctc_by_nam = {t.nam: t for t in hs.to_khai if t.loai == "bctc"}
    gtgt_by_nam: dict[int, list[ToKhai]] = {}
    tndn_by_nam = {t.nam: t for t in hs.to_khai if t.loai == "tndn"}
    tncn_by_nam = {t.nam: t for t in hs.to_khai if t.loai == "tncn"}
    for t in hs.to_khai:
        if t.loai == "gtgt":
            gtgt_by_nam.setdefault(t.nam, []).append(t)
    years = sorted(set(bctc_by_nam) | set(gtgt_by_nam) | set(tndn_by_nam) | set(tncn_by_nam))
    for nam in years:
        bc = bctc_by_nam.get(nam)
        gt = gtgt_by_nam.get(nam) or []
        tn = tndn_by_nam.get(nam)
        cn = tncn_by_nam.get(nam)
        out.extend(_year(nam, bc, gt, tn, cn))
    return out


def _year(nam: int, bc: ToKhai | None, gt: list[ToKhai], tn: ToKhai | None, cn: ToKhai | None) -> list[Finding]:
    f: list[Finding] = []
    dt_bctc = (bc.bctc_kqkd.get("01") or bc.bctc_kqkd.get("10") or 0) if bc else 0
    dt_gtgt = sum(t.chi_tieu.get("34", 0) for t in gt)
    if bc and gt:
        d = dt_bctc - dt_gtgt
        if abs(d) > 1:
            f.append(_tc("TC1", nam, "gtgt", "high" if abs(d) > 10_000_000 else "medium",
                         "Chênh lệch DT GTGT và DT trên BCTC (KQKD)",
                         f"BCTC [01] {dt_bctc:,.0f} - tổng TK GTGT [34] {dt_gtgt:,.0f} = {d:,.0f}",
                         f"BCTC {bc.file}; các tờ khai GTGT năm {nam} chỉ tiêu [34]",
                         "TUDIEN TC1_N: BCKQHĐKD [01] # tổng 12 tháng/4 quý [34] mẫu 01/GTGT",
                         "Chênh lệch do doanh thu không chịu thuế, xuất sai kỳ, hay kê thiếu?"))
    if bc and dt_bctc:
        ung = bc.bctc_cdkt_cuoi.get("312", 0) + bc.bctc_cdkt_cuoi.get("332", 0)
        if ung > UNG_TRUOC * dt_bctc:
            f.append(_tc("TC2", nam, "tndn", "medium",
                         "Người mua trả tiền trước lớn (>5% DT)",
                         f"[312]+[332]={ung:,.0f} = {ung / dt_bctc:.1%} DT {dt_bctc:,.0f}",
                         f"BCTC CĐKT cuối năm {nam}",
                         "TUDIEN TC2_N ngưỡng 5% DT",
                         "Số dư ứng trước đã xuất HĐ/ghi nhận DT đúng thời điểm chưa?"))
        htk = bc.bctc_cdkt_cuoi.get("141", 0)
        xdcb = bc.bctc_cdkt_cuoi.get("241", 0) + bc.bctc_cdkt_cuoi.get("242", 0)
        if dt_bctc and (htk + xdcb) > HTK_XDCB * dt_bctc:
            f.append(_tc("TC3", nam, "tndn", "high",
                         "HTK + XDCB dở dang > 30% DT (dấu hiệu kê DT sai thời điểm)",
                         f"[141]+[241]+[242]={htk + xdcb:,.0f} = {(htk + xdcb) / dt_bctc:.1%} DT",
                         f"BCTC CĐKT cuối năm {nam} TT200",
                         "TUDIEN TC3_N ngưỡng 30%",
                         "Hàng tồn / XDCB đã nghiệm thu chưa? Vì sao DT thấp hơn HTK?"))
        dthu_chua = bc.bctc_cdkt_cuoi.get("318", 0) + bc.bctc_cdkt_cuoi.get("336", 0)
        if dthu_chua > 0:
            f.append(_tc("TC4", nam, "tndn", "medium",
                         "Doanh thu chưa thực hiện > 0 (có thể thiếu DT TNDN)",
                         f"[318]+[336]={dthu_chua:,.0f}",
                         f"BCTC CĐKT năm {nam}",
                         "TUDIEN TC4_N",
                         "Khoản DT chưa thực hiện đã đủ điều kiện ghi nhận thuế chưa?"))
        gv = bc.bctc_kqkd.get("11", 0)
        dtt = bc.bctc_kqkd.get("10", 0)
        if dtt and gv and dtt - gv < 0:
            f.append(_tc("TC5", nam, "tndn", "high",
                         "Bán hàng dưới giá vốn",
                         f"DT thuần [10] {dtt:,.0f} - GV [11] {gv:,.0f} = {dtt - gv:,.0f}",
                         f"BCTC KQKD năm {nam}",
                         "TUDIEN TC5_N",
                         "Vì sao giá vốn vượt doanh thu thuần? Có xuất dưới giá vốn không?"))
    if bc and gt:
        ct22 = next((t.chi_tieu.get("22", 0) for t in sorted(gt, key=lambda x: x.ky) if t.chi_tieu.get("22") is not None), 0)
        ct43 = next((t.chi_tieu.get("43", 0) for t in sorted(gt, key=lambda x: x.ky, reverse=True)), 0)
        sd_dau = bc.bctc_cdkt_dau.get("152", 0)
        sd_cuoi = bc.bctc_cdkt_cuoi.get("152", 0)
        if abs(ct22 - sd_dau) > 1 or abs(ct43 - sd_cuoi) > 1:
            f.append(_tc("TC9", nam, "gtgt", "medium",
                         "Thuế GTGT khấu trừ trên tờ khai không khớp BCTC [152]",
                         f"TK [22]={ct22:,.0f} vs CĐKT đầu [152]={sd_dau:,.0f}; TK [43]={ct43:,.0f} vs cuối [152]={sd_cuoi:,.0f}",
                         f"GTGT năm {nam} và BCTC CĐKT [152]",
                         "TUDIEN TC9_N TT200",
                         "Chênh lệch do kết chuyển, hoàn thuế, hay hạch toán sai TK 133?"))
        htk = bc.bctc_cdkt_cuoi.get("141", 0)
        tscd = bc.bctc_cdkt_cuoi.get("222", 0)
        bds = bc.bctc_cdkt_cuoi.get("231", 0)
        vat_kh = bc.bctc_cdkt_cuoi.get("152", 0)
        uoc = (htk + tscd + bds) * 0.1
        if uoc and vat_kh > uoc * 1.2:
            f.append(_tc("TC10", nam, "gtgt", "medium",
                         "Thuế GTGT khấu trừ không tương ứng HTK+TSCĐ",
                         f"[152]={vat_kh:,.0f} > 10%×([141]+[222]+[231])={uoc:,.0f}",
                         f"BCTC năm {nam}",
                         "TUDIEN TC10_N",
                         "Phần khấu trừ vượt tài sản hình thành từ đâu?"))
    for t in gt:
        if t.chi_tieu.get("26", 0) > 0 and abs(t.chi_tieu.get("25", 0) - t.chi_tieu.get("24", 0)) < 1:
            f.append(_tc("TC11", nam, "gtgt", "high",
                         "Không phân bổ VAT đầu vào khi có DT không chịu thuế",
                         f"{t.ky}: [26]={t.chi_tieu.get('26', 0):,.0f} mà [25]=[24]={t.chi_tieu.get('24', 0):,.0f}",
                         t.file,
                         "TUDIEN TC11_N",
                         "Đã phân bổ thuế GTGT đầu vào cho DT không chịu thuế chưa?"))
    if tn:
        ln = tn.chi_tieu.get("18", 0)
        if ln < 0:
            f.append(_tc("TC5b", nam, "tndn", "medium",
                         "LN kế toán trước thuế âm (03/TNDN [18])",
                         f"[18]={ln:,.0f}",
                         tn.file,
                         "Mẫu 03/TNDN PL 03-1A chỉ tiêu [18]",
                         "Nguyên nhân lỗ? Chi phí nào bất thường?"))
    if cn is None:
        pass
    # TC24 cần TMS - skip
    f.append(_tc("TC24", nam, "gtgt", "skip",
                 "Xác định sai nghĩa vụ khi kê khai bổ sung (cần TMS)",
                 "Thiếu dữ liệu hạch toán TMS",
                 "Bỏ qua: không truy cập TMS",
                 "TUDIEN TC24_N",
                 "(không chạy)"))
    return [x for x in f if x.severity != "skip"]


def _tc(cid, nam, topic, sev, title, delta, evidence, legal, q) -> Finding:
    return Finding(
        id=cid.lower(), module="m2", tax_topic=topic, severity=sev, title=title,
        delta=delta, evidence=evidence, legal_basis=legal, question_for_nnt=q,
        tieu_chi=cid, nam=nam,
    )
