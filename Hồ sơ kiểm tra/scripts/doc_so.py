# -*- coding: utf-8 -*-
"""Doc so tien thanh chu tieng Viet (chuan tai chinh).

Dung THAY cho ham VBA DocSo: cac cot chu (53-58) tren so
'Danh sach kiem tra quyet toan pro.xlsm' bi mat dau tieng Viet vi VBE luu
chuoi theo code page ANSI cua he thong. Khi merge ho so, LUON tinh chu bang
ham nay tu con so goc, khong doc tu cot chu tren so.
"""

HANG = [u"không", u"một", u"hai", u"ba", u"bốn",
        u"năm", u"sáu", u"bảy", u"tám", u"chín"]
MUOI = u"mười"      # muoi (10)
MUOI2 = u"mươi"     # muoi (chuc)
LE = u"lẻ"               # le
MOT_LE = u"mốt"          # mot (21, 31...)
LAM = u"lăm"             # lam (15, 25...)
TRAM = u"trăm"
DONG = u"đồng"
DON_VI = [u"", u" nghìn", u" triệu", u" tỷ"]


def _doc_nhom(a, b, c, day_du):
    out = []
    if day_du or a > 0:
        out.append(HANG[a] + u" " + TRAM)
    if b == 0:
        if c > 0 and (day_du or a > 0):
            out.append(LE)
    elif b == 1:
        out.append(MUOI)
    else:
        out.append(HANG[b] + u" " + MUOI2)
    if c == 1:
        out.append(MOT_LE if b > 1 else HANG[1])
    elif c == 5:
        out.append(LAM if b > 0 else HANG[5])
    elif c > 0:
        out.append(HANG[c])
    return u" ".join(x for x in out if x)


def doc_so(n):
    """28410901 -> 'Hai muoi tam trieu bon tram muoi nghin chin tram le mot dong.'"""
    if n is None or str(n).strip() == "":
        return u""
    try:
        n = int(round(float(n)))
    except (TypeError, ValueError):
        return u""
    am = n < 0
    n = abs(n)
    if n == 0:
        return HANG[0].capitalize() + u" " + DONG + u"."
    s = str(n)
    while len(s) % 3:
        s = "0" + s
    nhom = [s[i:i + 3] for i in range(0, len(s), 3)]
    tong = len(nhom)
    phan = []
    for idx, g in enumerate(nhom):
        bac = tong - idx - 1
        if g == "000":
            continue
        a, b, c = int(g[0]), int(g[1]), int(g[2])
        txt = _doc_nhom(a, b, c, idx > 0)
        phan.append((txt + DON_VI[bac % 4]).strip())
    kq = u" ".join(u" ".join(phan).split())
    kq = kq[0].upper() + kq[1:]
    if am:
        kq = u"Âm " + kq[0].lower() + kq[1:]
    return kq + u" " + DONG + u"."


if __name__ == "__main__":
    import io
    out = io.open("doc_so_test.txt", "w", encoding="utf-8")
    for v in [0, 1, 5, 10, 11, 15, 21, 25, 101, 105, 110, 1000, 1005,
              19719634, 4747340, 3943927, 28410901, 1000000000,
              1234567890, 1000000, 1010000, -500000]:
        out.write(u"%s -> %s\n" % (v, doc_so(v)))
    out.close()
