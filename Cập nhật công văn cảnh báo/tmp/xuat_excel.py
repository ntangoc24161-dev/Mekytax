# -*- coding: utf-8 -*-
import openpyxl, collections, re, unicodedata, io, sys
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\03_BAN HANH CANH BAO RUI RO\2026\26.09\CTL\Chi tiet mat hang\HDDT_CHUA_RO_MST_Cong_ty.xlsx'
OUT = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\03_BAN HANH CANH BAO RUI RO\2026\26.09\CTL\Chi tiet mat hang\DOI_CHIEU_MAT_HANG_MUA_BAN.xlsx'
wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)


def load(name):
    out = []
    for r in wb[name].iter_rows(min_row=7, values_only=True):
        if r[3] is None:
            continue
        out.append(dict(mau=r[1], kh=r[2], so=r[3], ngay=r[4],
                        mst_ban=(r[5] or '').strip(), ten_ban=(r[6] or '').strip(),
                        mst_mua=(r[8] or '').strip(), ten_mua=(r[9] or '').strip(),
                        mh=(r[11] or '').strip(), chuathue=r[12] or 0, ts=r[13],
                        thue=r[14] or 0, tong=r[15] or 0, tt=r[17]))
    return out


MV, BR = load('Mua vào'), load('Bán ra')


def nodiac(s):
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s.replace('đ', 'd').replace('Đ', 'D')


def norm(s):
    t = nodiac(s).lower().replace('pc trand', 'pc strand')
    t = re.sub(r'(\d),(\d)', r'\1.\2', t)
    t = re.sub(r'(\d)\.0+(?!\d)', r'\1', t)
    return re.sub(r'[^a-z0-9]', '', t)


def nhom(mh):
    t = nodiac(mh).lower()
    if not mh:
        return 'Z. Khong ghi mat hang'
    if 'thep' in t or 'pc strand' in t:
        return 'A. Thep (day, cuon, DUL PC Strand)'
    if re.search(r'\bvo\b|\bruot\b|lop xe|vanh|yem', t):
        return 'B. Vo / ruot xe (lop)'
    if any(k in t for k in ['lan cat', 'cap lan', 'utp', 'cap mang', 'switch', 'router', 'camera',
                            'may tinh', 'laptop', 'man hinh', 'ups', 'thiet bi', 'dau ghi',
                            'ban phim', 'chuot', 'o cung', 'cpu', 'dien thoai', 'sim', 'modem',
                            'wifi', 'phan mem', 'ban od', 'ghe', 'tu ']):
        return 'C. Thiet bi CNTT / dien tu / noi that'
    if 'thue van phong' in t or 'phi thue' in t or 'cho thue' in t:
        return 'D. Cho thue van phong / mat bang'
    if any(k in t for k in ['xang', 'dau do', 'do 0', 'nhien lieu', 'ron 9']):
        return 'E. Xang dau'
    if any(k in t for k in ['phi dich vu ngan hang', 'phi ngan hang', 'phi chuyen tien', 'phi sms',
                            'phi quan ly', 'phi duy tri', 'internet', 'cuoc', 'phi luu ky',
                            'chung khoan', 'kiem toan', 'luat', 'dich vu']):
        return 'F. Dich vu / tai chinh / vien thong'
    return 'G. Khac'


def agg(rows, kf):
    d = collections.defaultdict(lambda: [0, 0.0, 0.0])
    for r in rows:
        k = kf(r)
        d[k][0] += 1
        d[k][1] += r['chuathue']
        d[k][2] += r['tong']
    return d


out = openpyxl.Workbook()
TNR = 'Times New Roman'
hdr_f = Font(name=TNR, bold=True, size=11, color='FFFFFF')
hdr_fill = PatternFill('solid', fgColor='1F4E79')
ttl_f = Font(name=TNR, bold=True, size=13)
bd = Border(*[Side(style='thin', color='BFBFBF')] * 4)
NUM = '#,##0'
red = PatternFill('solid', fgColor='FFC7CE')
yel = PatternFill('solid', fgColor='FFF2CC')


def sheet(title, headers, rows, widths, numcols, note=None, fills=None):
    ws = out.create_sheet(title)
    ws['A1'] = note or title
    ws['A1'].font = ttl_f
    hr = 3
    for i, h in enumerate(headers, 1):
        c = ws.cell(hr, i, h)
        c.font = hdr_f
        c.fill = hdr_fill
        c.border = bd
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for ri, row in enumerate(rows, hr + 1):
        for ci, v in enumerate(row, 1):
            c = ws.cell(ri, ci, v)
            c.font = Font(name=TNR, size=11)
            c.border = bd
            if ci in numcols:
                c.number_format = NUM
                c.alignment = Alignment(horizontal='right')
        if fills:
            f = fills(row)
            if f:
                for ci in range(1, len(headers) + 1):
                    ws.cell(ri, ci).fill = f
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(hr + 1, 1)
    ws.auto_filter.ref = 'A%d:%s%d' % (hr, get_column_letter(len(headers)), hr + len(rows))
    return ws


# ---- Sheet 1
gm, gb = agg(MV, lambda r: nhom(r['mh'])), agg(BR, lambda r: nhom(r['mh']))
tm = sum(v[1] for v in gm.values())
tb = sum(v[1] for v in gb.values())
r1 = []
for g in sorted(set(gm) | set(gb)):
    m = gm.get(g, [0, 0, 0])
    b = gb.get(g, [0, 0, 0])
    r1.append([g, m[0], m[1], round(m[1] / tm * 100, 2), b[0], b[1], round(b[1] / tb * 100, 2),
               b[1] - m[1], round(b[1] / m[1], 2) if m[1] else None])
r1.append(['TONG CONG', len(MV), tm, 100.0, len(BR), tb, 100.0, tb - tm, round(tb / tm, 2)])
ws = sheet('1. Tong hop nhom hang',
           ['Nhom mat hang', 'Mua: so dong', 'Mua: tien chua thue', 'Mua: %', 'Ban: so dong',
            'Ban: tien chua thue', 'Ban: %', 'Chenh lech (Ban - Mua)', 'Ty le Ban/Mua'],
           r1, [40, 12, 22, 9, 12, 22, 9, 24, 13], {2, 3, 5, 6, 8},
           note='BANG 1. BOC TACH MAT HANG MUA VAO - BAN RA THEO NHOM (dong, chua thue GTGT)')
for c in ws[ws.max_row]:
    c.font = Font(name=TNR, bold=True, size=11)
    c.fill = yel
for rr in range(4, ws.max_row + 1):
    for cc in (4, 7, 9):
        ws.cell(rr, cc).number_format = '#,##0.00'

# ---- Sheet 2
km, kb = agg(MV, lambda r: norm(r['mh'])), agg(BR, lambda r: norm(r['mh']))
nmm, nbb = {}, {}
for r in MV:
    nmm.setdefault(norm(r['mh']), r['mh'])
for r in BR:
    nbb.setdefault(norm(r['mh']), r['mh'])
r2 = []
for k in set(km) | set(kb):
    if not k:
        continue
    m = km.get(k, [0, 0, 0])
    b = kb.get(k, [0, 0, 0])
    nm2 = nbb.get(k) or nmm.get(k)
    pct = None
    if m[0] and b[0]:
        pct = (b[1] - m[1]) / m[1] * 100 if m[1] else None
        if pct is None:
            st = 'Co ca hai phia'
        elif pct < 0:
            st = 'Ban duoi gia mua'
        elif pct < 0.15:
            st = 'Mua - ban gan bang nhau (duoi 0,15%)'
        elif pct > 20:
            st = 'Ban vuot mua tren 20% - thieu dau vao'
        else:
            st = 'Co ca hai phia'
    elif b[0]:
        st = 'CHI CO BAN RA - khong co dau vao'
    else:
        st = 'Chi co mua vao'
    r2.append([nm2, nhom(nm2), m[0], m[1], b[0], b[1], b[1] - m[1],
               round(pct, 3) if pct is not None else None, st])
r2.sort(key=lambda x: -x[5])


def f2(row):
    if row[8].startswith('CHI CO BAN') or row[8].startswith('Ban vuot'):
        return red
    if row[8].startswith('Mua - ban gan'):
        return yel
    return None


ws = sheet('2. Doi chieu mat hang',
           ['Ten mat hang', 'Nhom', 'Mua: dong', 'Mua: tien chua thue', 'Ban: dong',
            'Ban: tien chua thue', 'Chenh lech', 'Bien lai %', 'Danh gia'],
           r2, [62, 34, 11, 22, 11, 22, 20, 11, 40], {3, 4, 5, 6, 7}, fills=f2,
           note='BANG 2. DOI CHIEU TUNG MAT HANG SAU CHUAN HOA TEN (do = can lam ro, vang = bien lai duoi 0,15%)')
for rr in range(4, ws.max_row + 1):
    ws.cell(rr, 8).number_format = '#,##0.000'

# ---- Sheet 3
mvby = collections.defaultdict(list)
for r in MV:
    if r['mh']:
        mvby[norm(r['mh'])].append(r)
for k in mvby:
    mvby[k].sort(key=lambda r: r['ngay'])
pairs = []
for r in BR:
    for m in mvby.get(norm(r['mh']), []):
        d = (r['ngay'] - m['ngay']).days
        if 0 <= d <= 2 and m['chuathue'] > 0 and abs(r['chuathue'] - m['chuathue']) / m['chuathue'] < 0.02:
            pairs.append([m['ngay'], '%s/%s' % (m['kh'], m['so']), m['mst_ban'], m['ten_ban'], m['chuathue'],
                          r['ngay'], '%s/%s' % (r['kh'], r['so']), r['mst_mua'], r['ten_mua'], r['chuathue'],
                          r['chuathue'] - m['chuathue'],
                          round((r['chuathue'] - m['chuathue']) / m['chuathue'] * 100, 4), d, r['mh']])
            break
pairs.sort(key=lambda x: x[0])
ws = sheet('3. Cap mua-ban trung khop',
           ['Ngay mua', 'So HD mua', 'MST NCC', 'Ten NCC', 'Tien mua', 'Ngay ban', 'So HD ban',
            'MST khach', 'Ten khach', 'Tien ban', 'Chenh', 'Bien lai %', 'So ngay', 'Mat hang'],
           pairs, [11, 13, 15, 40, 19, 11, 13, 15, 40, 19, 15, 11, 9, 52], {5, 10, 11},
           note='BANG 3. CAC CAP MUA - BAN TRUNG TIEN (lech duoi 2%) TRONG 0 - 2 NGAY, sap theo ngay mua')
for rr in range(4, ws.max_row + 1):
    ws.cell(rr, 1).number_format = 'dd/mm/yyyy'
    ws.cell(rr, 6).number_format = 'dd/mm/yyyy'
    ws.cell(rr, 12).number_format = '#,##0.0000'

# ---- Sheet 4
mon_m = agg(MV, lambda r: '%04d-%02d' % (r['ngay'].year, r['ngay'].month))
mon_b = agg(BR, lambda r: '%04d-%02d' % (r['ngay'].year, r['ngay'].month))
tm_a = agg([r for r in MV if nhom(r['mh']).startswith('A.')],
           lambda r: '%04d-%02d' % (r['ngay'].year, r['ngay'].month))
tb_a = agg([r for r in BR if nhom(r['mh']).startswith('A.')],
           lambda r: '%04d-%02d' % (r['ngay'].year, r['ngay'].month))
r4 = []
for k in sorted(set(mon_m) | set(mon_b)):
    a = tm_a.get(k, [0, 0, 0])[1]
    b = tb_a.get(k, [0, 0, 0])[1]
    r4.append([k, mon_m.get(k, [0, 0, 0])[0], mon_m.get(k, [0, 0, 0])[1],
               mon_b.get(k, [0, 0, 0])[0], mon_b.get(k, [0, 0, 0])[1],
               a, b, b - a, round(b / a, 2) if a else None])
ws = sheet('4. Dien bien theo thang',
           ['Thang', 'Mua: dong', 'Mua: tien chua thue', 'Ban: dong', 'Ban: tien chua thue',
            'Mua THEP', 'Ban THEP', 'Thep: Ban - Mua', 'Thep: Ban/Mua'],
           r4, [10, 11, 22, 11, 22, 22, 22, 22, 14], {2, 3, 4, 5, 6, 7, 8},
           note='BANG 4. DIEN BIEN MUA VAO - BAN RA THEO THANG (tach rieng nhom thep)')
for rr in range(4, ws.max_row + 1):
    ws.cell(rr, 9).number_format = '#,##0.00'

# ---- Sheet 5
def basem(x):
    return x.split('-')[0]


mm = agg(MV, lambda r: (r['mst_ban'], r['ten_ban']))
bb = agg(BR, lambda r: (r['mst_mua'], r['ten_mua']))
bmap = collections.defaultdict(lambda: [0, 0.0])
for (mst, _), v in bb.items():
    t = bmap[basem(mst)]
    t[0] += v[0]
    t[1] += v[1]
mmap = collections.defaultdict(lambda: [0, 0.0])
for (mst, _), v in mm.items():
    t = mmap[basem(mst)]
    t[0] += v[0]
    t[1] += v[1]
r5 = []
for (mst, ten), v in mm.items():
    o = bmap.get(basem(mst))
    r5.append([mst, ten, 'Nguoi ban (mua vao)', v[0], v[1], o[0] if o else 0, o[1] if o else 0,
               'HAI CHIEU - vua mua vua ban' if o else ''])
for (mst, ten), v in bb.items():
    if basem(mst) in mmap:
        continue
    r5.append([mst, ten, 'Nguoi mua (ban ra)', 0, 0, v[0], v[1], ''])
r5.sort(key=lambda x: -(x[4] + x[6]))
sheet('5. Doi tac',
      ['MST', 'Ten doi tac', 'Vai tro', 'Mua vao: dong', 'Mua vao: tien', 'Ban ra: dong',
       'Ban ra: tien', 'Canh bao'],
      r5, [17, 56, 22, 13, 22, 13, 22, 30], {4, 5, 6, 7},
      note='BANG 5. DOI TAC MUA VAO VA BAN RA (danh dau doi tac hai chieu)',
      fills=lambda row: red if row[7] else None)

del out['Sheet']
out.save(OUT)
print('Da ghi:', OUT)
print('Sheet:', out.sheetnames)
print('Bang 2:', len(r2), '| Bang 3:', len(pairs), '| Bang 5:', len(r5))
