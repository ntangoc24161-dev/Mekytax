"""M1: rủi ro hóa đơn + chênh lệch HĐ vs tờ khai GTGT (TC27, sheet 3.TTKT-HDDT)."""
from __future__ import annotations

from collections import defaultdict

from models import Finding, HoSoChuan


def run_m1(hs: HoSoChuan) -> list[Finding]:
    out: list[Finding] = []
    out.extend(_anomalies(hs))
    out.extend(_vs_to_khai(hs))
    return out


def _anomalies(hs: HoSoChuan) -> list[Finding]:
    out: list[Finding] = []
    by_key: dict[tuple, list] = defaultdict(list)
    for hd in hs.hoa_don:
        by_key[(hd.chieu, hd.so, hd.mst_dt, hd.ngay, round(hd.doanh_so, 0))].append(hd)
        nam_ok = (not hd.ngay) or (len(hd.ngay) >= 4 and int(hd.ngay[:4]) in (hs.nam or [int(hd.ngay[:4])]))
        if hd.ngay and hs.nam and int(hd.ngay[:4]) not in hs.nam:
            out.append(_f(
                "m1-ngay-ngoai-ky", "medium", "hddt",
                "Hóa đơn ngoài năm hồ sơ",
                f"HĐ {hd.so} ngày {hd.ngay}, năm hồ sơ {hs.nam}",
                f"{hd.file} / {hd.sheet} dòng {hd.row}",
                "Bảng kê HĐ",
                "Hóa đơn này thuộc kỳ nào? Đã kê khai kỳ đó chưa?",
                nam=int(hd.ngay[:4]) if hd.ngay[:4].isdigit() else None,
            ))
        if hd.thue_suat not in (0, 0.05, 0.08, 0.1) and hd.doanh_so:
            out.append(_f(
                "m1-thue-suat", "medium", "hddt",
                "Thuế suất không thuộc 0/5/8/10%",
                f"HĐ {hd.so} suất {hd.thue_suat}",
                f"{hd.file} / {hd.sheet} dòng {hd.row}",
                "Nghị định thuế suất GTGT hiện hành",
                "Căn cứ áp thuế suất này là gì?",
            ))
        if hd.doanh_so < 0 or hd.thue < 0:
            out.append(_f(
                "m1-am", "low", "hddt",
                "Hóa đơn số tiền âm (điều chỉnh/hủy?)",
                f"HĐ {hd.so} doanh số {hd.doanh_so:,.0f}",
                f"{hd.file} / {hd.sheet} dòng {hd.row}",
                "Hóa đơn điều chỉnh/thay thế",
                "Đây là HĐ điều chỉnh hay hủy? Đã lập biên bản điều chỉnh chưa?",
            ))
        _ = nam_ok
    for key, group in by_key.items():
        if len(group) > 1:
            hd = group[0]
            out.append(_f(
                "m1-trung", "high", "hddt",
                "Trùng số hóa đơn cùng MST đối tác",
                f"{len(group)} dòng HĐ {hd.so} MST {hd.mst_dt}",
                "; ".join(f"{x.sheet}:{x.row}" for x in group[:5]),
                "Bảng kê HĐ",
                "Vì sao cùng số HĐ xuất hiện nhiều lần?",
            ))
    return out


def _vs_to_khai(hs: HoSoChuan) -> list[Finding]:
    out: list[Finding] = []
    gtgt = [t for t in hs.to_khai if t.loai == "gtgt"]
    if not gtgt or not hs.hoa_don:
        return out
    hd_ban = defaultdict(lambda: [0.0, 0.0])
    hd_mua = defaultdict(lambda: [0.0, 0.0])
    for hd in hs.hoa_don:
        nam = int(hd.ngay[:4]) if hd.ngay[:4].isdigit() else 0
        if nam < 1990 or nam > 2100:
            continue
        bag = hd_ban if hd.chieu == "ban" else hd_mua
        bag[nam][0] += hd.doanh_so
        bag[nam][1] += hd.thue
    tk_ban = defaultdict(lambda: [0.0, 0.0])
    tk_mua = defaultdict(lambda: [0.0, 0.0])
    for t in gtgt:
        tk_ban[t.nam][0] += t.chi_tieu.get("34", 0)
        tk_ban[t.nam][1] += t.chi_tieu.get("35", 0)
        tk_mua[t.nam][0] += t.chi_tieu.get("23", 0)
        tk_mua[t.nam][1] += t.chi_tieu.get("24", 0)
    years = sorted(set(tk_ban) | set(hd_ban) | set(tk_mua) | set(hd_mua))
    for nam in years:
        if not nam:
            continue
        d_ds = tk_ban[nam][0] - hd_ban[nam][0]
        d_th = tk_ban[nam][1] - hd_ban[nam][1]
        if abs(d_ds) > 1 or abs(d_th) > 1:
            sev = "high" if abs(d_ds) > 10_000_000 else "medium"
            out.append(_f(
                "tc27", sev, "gtgt",
                "Chênh lệch DT tờ khai GTGT và HĐĐT bán ra",
                f"Năm {nam}: TK ct34 {tk_ban[nam][0]:,.0f} - HĐ bán {hd_ban[nam][0]:,.0f} = {d_ds:,.0f} (thuế {d_th:,.0f})",
                "Tờ khai 01/GTGT chỉ tiêu [34]/[35] đối chiếu bảng kê bán ra",
                "TUDIEN TC27_N tool 4.4; sheet 3.TTKT-HDDT",
                "Chênh lệch do hóa đơn điều chỉnh, xuất sai kỳ, hay kê thiếu/thừa?",
                tieu_chi="TC27",
                nam=nam,
            ))
        d_m = tk_mua[nam][0] - hd_mua[nam][0]
        if abs(d_m) > 1:
            sev = "high" if abs(d_m) > 10_000_000 else "medium"
            out.append(_f(
                "m1-mua-tk", sev, "gtgt",
                "Chênh lệch GT mua vào tờ khai GTGT và HĐĐT mua vào",
                f"Năm {nam}: TK ct23 {tk_mua[nam][0]:,.0f} - HĐ mua {hd_mua[nam][0]:,.0f} = {d_m:,.0f}",
                "Tờ khai 01/GTGT [23]/[24] đối chiếu bảng kê mua vào",
                "Sheet 3.TTKT-HDDT (mua vào)",
                "Hóa đơn mua nào chưa kê hoặc kê không có HĐ?",
                nam=nam,
            ))
    return out


def _f(fid, sev, topic, title, delta, evidence, legal, q, tieu_chi="", nam=None) -> Finding:
    return Finding(
        id=fid, module="m1", tax_topic=topic, severity=sev, title=title,
        delta=delta, evidence=evidence, legal_basis=legal, question_for_nnt=q,
        tieu_chi=tieu_chi, nam=nam,
    )
