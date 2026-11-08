# -*- coding: utf-8 -*-
"""Uoc tinh so lieu KPI quy III/2026 cho 7 dong chi chi uy CB13.

Can cu: file mau 'Danh gia PCVP.xlsx' (so Mau 02AB cua PCVP phuong Tran Bien).
Trong file mau, 4 tru A/B/C/D co ty le thuc te/ke hoach lan luot ~91,5% - 86,9%
- 92,5% - 61,1%; KPI tung nhiem vu 71,75% den 89,67%, trung binh 83,01%.
Vi 7 dong chi CB13 duoc de xuat HTTNV (hoan thanh tot nhiem vu: hoan thanh
100% nhiem vu, chat luong tot, KHONG dat tieu chi vuot muc 30% cua HTXSNV)
nen uoc tinh chi dung muc tot va muc dat. Tru D giu thap hon A/B/C dung theo
dang hinh cua file mau.

MOI SO DO FILE NAY SINH RA DEU LA UOC TINH, DUOC TO MAU DO TRONG EXCEL.
"""

DO = "C00000"  # mau chu do cho so lieu Javis uoc tinh

# (A so luong, B chat luong, C tien do, D lanh dao) - ty le thuc te / ke hoach
VUOT = (1.10, 1.05, 1.00, 1.00)   # KPI 103,75%
TOT  = (1.00, 0.92, 0.97, 0.80)   # KPI  92,25%
DAT  = (0.95, 0.88, 0.93, 0.70)   # KPI  86,50%

NHAN = {VUOT: "vượt mức", TOT: "hoàn thành tốt", DAT: "hoàn thành"}

# Mau phan bo goc cho 9 tieu chi Nhom B; xoay theo stt de 7 file khong giong nhau.
# HTTNV: khong dung VUOT (KPI 103,75%) de tranh vo tinh du dieu kien HTXSNV.
_GOC = [TOT, TOT, DAT, TOT, DAT, TOT, TOT, DAT, TOT]


def phan_bo(stt):
    """Tra ve list 9 profile cho mot dong chi (xoay mau goc theo stt)."""
    k = (stt - 1) % 9
    return _GOC[k:] + _GOC[:k]


def co_so(nv):
    """Diem quy doi ke hoach cua 1 tieu chi = so luong x he so x 100 (Phu luc 2A)."""
    if nv is None:
        return 1, 1.0, 100.0
    sl, hs = nv[8], nv[9]
    return sl, hs, round(sl * hs * 100.0, 1)


def o_nhap(nv, prof):
    """8 o nhap cua 1 tieu chi: (A_KH, A_TT, B_KH, B_TT, C_KH, C_TT, D_KH, D_TT)."""
    sl, hs, kh = co_so(nv)
    a, b, c, d = prof
    return (kh, round(kh * a, 1), kh, round(kh * b, 1),
            kh, round(kh * c, 1), kh, round(kh * d, 1))


def kpi(prof):
    return round(sum(prof) / 4.0 * 100, 2)


def so(x):
    """So kieu Viet Nam: dau phay thap phan, bo .0 du."""
    t = ("%.2f" % x).rstrip("0").rstrip(".")
    return t.replace(".", ",")


def ket_qua_text(nv, prof):
    """Cau tom tat ket qua thuc te uoc tinh cho 1 tieu chi."""
    a, b, c, d = prof
    sl, hs, kh = co_so(nv)
    san_pham = nv[4] if nv else "sản phẩm theo lĩnh vực được phân công phụ trách"
    vuot = (" Có khối lượng phát sinh ngoài kế hoạch được bổ sung và quy đổi theo Phụ lục 2A."
            if a > 1 else "")
    return ("Mức %s. Khối lượng kế hoạch: %s sản phẩm, hệ số quy đổi %s "
            "(điểm quy đổi kế hoạch %s). Điểm quy đổi thực tế so với kế hoạch: số lượng %s%%, "
            "chất lượng %s%%, tiến độ %s%%, năng lực lãnh đạo - chỉ đạo - điều hành %s%%; "
            "KPI %s%%.%s Sản phẩm: %s."
            % (NHAN[prof], so(sl), so(hs), so(kh), so(a * 100), so(b * 100),
               so(c * 100), so(d * 100), so(kpi(prof)), vuot, san_pham))


def diem_nhom_b(stt, tieu_chi_b):
    """Diem Nhom B uoc tinh = sum(KPI_i/100 * diem_toi_da_i)."""
    ps = phan_bo(stt)
    return round(sum(kpi(p) / 100.0 * tc[1] for p, tc in zip(ps, tieu_chi_b)), 2)
