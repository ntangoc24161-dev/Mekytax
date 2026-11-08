# -*- coding: utf-8 -*-
import openpyxl, collections, re, unicodedata, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = r'G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\03_BAN HANH CANH BAO RUI RO\2026\26.09\CTL\Chi tiet mat hang\HDDT_CHUA_RO_MST_Cong_ty.xlsx'
wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)

def load(name):
    out = []
    for r in wb[name].iter_rows(min_row=7, values_only=True):
        if r[3] is None:
            continue
        out.append(dict(stt=r[0], mau=r[1], kh=r[2], so=r[3], ngay=r[4],
                        mst_ban=(r[5] or '').strip(), ten_ban=(r[6] or '').strip(),
                        mst_mua=(r[8] or '').strip(), ten_mua=(r[9] or '').strip(),
                        mh=(r[11] or '').strip(), chuathue=r[12] or 0, ts=r[13],
                        thue=r[14] or 0, tong=r[15] or 0, tt=r[17], src=r[18]))
    return out

MV = load('Mua vào')
BR = load('Bán ra')

def nodiac(s):
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return s.replace('đ','d').replace('Đ','D')

def norm(s):
    """Chuan hoa ten mat hang de so khop: bo dau, lowercase, dong nhat dau thap phan,
    bo khoang trang va dau cau, dong nhat 'trand'->'strand'."""
    t = nodiac(s).lower()
    t = t.replace('thep', 'thep')
    t = t.replace('pc trand', 'pc strand')
    t = re.sub(r'(\d),(\d)', r'\1.\2', t)          # 0,75 -> 0.75
    t = re.sub(r'(\d)\.0+(?!\d)', r'\1', t)        # 6.0 -> 6
    t = re.sub(r'[^a-z0-9]', '', t)
    return t

# ---- phan nhom hang
def nhom(mh):
    t = nodiac(mh).lower()
    if not mh:
        return 'Z. Khong ghi mat hang'
    if 'thep' in t or 'pc strand' in t or 'pc trand' in t:
        return 'A. Thep (day, cuon, DUL PC Strand)'
    if re.search(r'\bvo\b|\bruot\b|lop xe|vanh|yem', t) or re.match(r'^(vo|ruot)\s', t):
        return 'B. Vo / ruot xe (lop)'
    if any(k in t for k in ['lan cat','cap lan','utp','cap mang','day mang','switch','router','camera','may tinh','laptop','man hinh','ups','thiet bi','dau ghi','ban phim','chuot','o cung','ram ','cpu','dien thoai','sim','modem','wifi','phan mem','ban od','ghe','tu ']):
        return 'C. Thiet bi CNTT / dien tu / noi that'
    if 'thue van phong' in t or 'phi thue' in t or 'cho thue' in t:
        return 'D. Cho thue van phong / mat bang'
    if any(k in t for k in ['xang','dau do','do 0','nhien lieu','ron 9']):
        return 'E. Xang dau'
    if any(k in t for k in ['phi dich vu ngan hang','phi ngan hang','lai ','phi chuyen tien','phi sms','phi quan ly','phi duy tri','internet','cuoc','dien thoai co dinh','phi luu ky','chung khoan','kiem toan','luat','dich vu']):
        return 'F. Dich vu / tai chinh / vien thong'
    return 'G. Khac (van phong pham, an uong, tieu hao...)'

def agg_by(rows, keyfn):
    d = collections.defaultdict(lambda: [0, 0.0, 0.0])
    for r in rows:
        k = keyfn(r)
        d[k][0] += 1
        d[k][1] += r['chuathue']
        d[k][2] += r['tong']
    return d

g_mv = agg_by(MV, lambda r: nhom(r['mh']))
g_br = agg_by(BR, lambda r: nhom(r['mh']))

def vn(x, dec=0):
    s = f'{x:,.{dec}f}'
    return s.replace(',', '#').replace('.', ',').replace('#', '.')

print('=' * 104)
print('PHAN 1. TONG QUAN FILE')
print('=' * 104)
for nm, rows in (('MUA VAO', MV), ('BAN RA', BR)):
    ds = [r['ngay'] for r in rows if r['ngay']]
    print(f"{nm}: {len(rows)} dong | {len(set((r['so'], r['kh']) for r in rows))} so HD | "
          f"{min(ds):%d/%m/%Y} - {max(ds):%d/%m/%Y}")
    print(f"   Chua thue: {vn(sum(r['chuathue'] for r in rows))} d | "
          f"Thue GTGT: {vn(sum(r['thue'] for r in rows))} d | "
          f"Tong TT: {vn(sum(r['tong'] for r in rows))} d")
    print(f"   So doi tac: {len(set(r['mst_ban'] or r['mst_mua'] for r in rows))} | "
          f"So ten mat hang khac nhau: {len(set(r['mh'] for r in rows))} | "
          f"Sau chuan hoa: {len(set(norm(r['mh']) for r in rows if r['mh']))}")

print()
print('=' * 104)
print('PHAN 2. BOC TACH THEO NHOM MAT HANG')
print('=' * 104)
print(f"{'Nhom hang':<44}{'Mua: dong':>10}{'Mua: tien chua thue':>24}{'Ban: dong':>10}{'Ban: tien chua thue':>24}")
allg = sorted(set(g_mv) | set(g_br))
tm = sum(v[1] for v in g_mv.values()); tb = sum(v[1] for v in g_br.values())
for g in allg:
    m = g_mv.get(g, [0, 0, 0]); b = g_br.get(g, [0, 0, 0])
    print(f"{g:<44}{m[0]:>10}{vn(m[1]):>24}{b[0]:>10}{vn(b[1]):>24}")
print('-' * 112)
print(f"{'TONG CONG':<44}{len(MV):>10}{vn(tm):>24}{len(BR):>10}{vn(tb):>24}")
print()
print('Ty trong trong tong (%):')
for g in allg:
    m = g_mv.get(g, [0, 0, 0]); b = g_br.get(g, [0, 0, 0])
    print(f"{g:<44}{m[1]/tm*100:>9.2f}%{'':>15}{b[1]/tb*100:>9.2f}%")

print()
print('=' * 104)
print('PHAN 3. DOI CHIEU TUNG MAT HANG (sau chuan hoa ten)')
print('=' * 104)
km = agg_by(MV, lambda r: norm(r['mh']))
kb = agg_by(BR, lambda r: norm(r['mh']))
name_m = {}; name_b = {}
for r in MV: name_m.setdefault(norm(r['mh']), r['mh'])
for r in BR: name_b.setdefault(norm(r['mh']), r['mh'])

both = set(km) & set(kb) - {''}
only_b = set(kb) - set(km) - {''}
only_m = set(km) - set(kb) - {''}
print(f"Mat hang co CA mua lan ban : {len(both):>5} | ban {vn(sum(kb[k][1] for k in both))} d | mua {vn(sum(km[k][1] for k in both))} d")
print(f"Mat hang CHI CO ban ra     : {len(only_b):>5} | ban {vn(sum(kb[k][1] for k in only_b))} d  <= khong thay dau vao")
print(f"Mat hang CHI CO mua vao    : {len(only_m):>5} | mua {vn(sum(km[k][1] for k in only_m))} d  <= ton kho / chi phi")

print()
print('--- 3.1 TOP 30 MAT HANG BAN RA KHONG CO MUA VAO TUONG UNG (theo tien ban) ---')
print(f"{'Tien ban chua thue':>22} {'Dong':>5}  Ten mat hang")
for k in sorted(only_b, key=lambda x: -kb[x][1])[:30]:
    print(f"{vn(kb[k][1]):>22} {kb[k][0]:>5}  {name_b[k][:70]}")
tb_only = sum(kb[k][1] for k in only_b)
print(f"-> Tong {len(only_b)} mat hang chi co ban ra: {vn(tb_only)} d = {tb_only/tb*100:.2f}% doanh thu")

print()
print('--- 3.2 MAT HANG CO CA HAI PHIA: BAN VUOT MUA (top 30 theo chenh lech) ---')
print(f"{'Mua chua thue':>20}{'Ban chua thue':>20}{'Chenh (ban-mua)':>20}  Ten mat hang")
rows_cl = sorted(both, key=lambda k: -(kb[k][1] - km[k][1]))
for k in rows_cl[:30]:
    print(f"{vn(km[k][1]):>20}{vn(kb[k][1]):>20}{vn(kb[k][1]-km[k][1]):>20}  {name_b[k][:52]}")

print()
print('--- 3.3 NHOM A (THEP): can doi mua - ban theo tung quy cach ---')
tha_m = {k: v for k, v in km.items() if k and nhom(name_m[k]).startswith('A.')}
tha_b = {k: v for k, v in kb.items() if k and nhom(name_b[k]).startswith('A.')}
sb = set(tha_b) - set(tha_m)
print(f"Thep: {len(tha_m)} quy cach mua / {len(tha_b)} quy cach ban")
print(f"Thep ban ra KHONG co dau vao cung quy cach: {len(sb)} quy cach, {vn(sum(tha_b[k][1] for k in sb))} d")
print(f"   = {sum(tha_b[k][1] for k in sb)/sum(v[1] for v in tha_b.values())*100:.2f}% doanh thu thep")
print('   Top 20:')
for k in sorted(sb, key=lambda x: -tha_b[x][1])[:20]:
    print(f"   {vn(tha_b[k][1]):>20} {tha_b[k][0]:>3} dong  {name_b[k][:62]}")

print()
print('=' * 104)
print('PHAN 4. BIEN LAI GOP (ban - mua) TREN CAC MAT HANG CO CA HAI PHIA')
print('=' * 104)
buckets = collections.Counter(); bval = collections.Counter()
detail = []
for k in both:
    m, b = km[k][1], kb[k][1]
    if m <= 0: continue
    r = (b - m) / m * 100
    lab = ('<0% (ban duoi gia mua)' if r < 0 else
           '0 - 0,15%' if r < 0.15 else '0,15 - 1%' if r < 1 else
           '1 - 5%' if r < 5 else '5 - 20%' if r < 20 else '>20%')
    buckets[lab] += 1; bval[lab] += b
    detail.append((r, k, m, b))
order = ['<0% (ban duoi gia mua)','0 - 0,15%','0,15 - 1%','1 - 5%','5 - 20%','>20%']
print(f"{'Khoang bien lai':<26}{'So mat hang':>12}{'Tien ban ra':>24}")
for o in order:
    if buckets[o]: print(f"{o:<26}{buckets[o]:>12}{vn(bval[o]):>24}")

print()
print('--- 4.1 Mat hang mua va ban gan nhu BANG NHAU (bien lai < 0,15%) ---')
near = sorted([d for d in detail if abs(d[0]) < 0.15], key=lambda x: -x[3])
print(f"So mat hang: {len(near)} | Tien mua {vn(sum(d[2] for d in near))} | Tien ban {vn(sum(d[3] for d in near))}")
print(f"{'Mua':>20}{'Ban':>20}{'Bien lai':>10}  Ten")
for r, k, m, b in near[:25]:
    print(f"{vn(m):>20}{vn(b):>20}{r:>9.3f}%  {name_b[k][:48]}")

print()
print('=' * 104)
print('PHAN 5. DOI TAC HAI CHIEU (vua la NCC mua vao, vua la khach ban ra)')
print('=' * 104)
mstm = agg_by(MV, lambda r: r['mst_ban'])
mstb = agg_by(BR, lambda r: r['mst_mua'])
nm_m = {r['mst_ban']: r['ten_ban'] for r in MV}
nm_b = {r['mst_mua']: r['ten_mua'] for r in BR}
def base(x): return x.split('-')[0]
mm = {base(k): v for k, v in mstm.items()}
bb = collections.defaultdict(lambda: [0, 0.0, 0.0])
for k, v in mstb.items():
    t = bb[base(k)]; t[0] += v[0]; t[1] += v[1]; t[2] += v[2]
for k in sorted(set(mm) & set(bb), key=lambda x: -(mm[x][1] + bb[x][1])):
    print(f"MST {k}  {str(nm_b.get(k) or nm_m.get(k))[:46]}")
    print(f"    Mua vao: {mm[k][0]:>4} dong / {vn(mm[k][1]):>20} d   |   Ban ra: {bb[k][0]:>4} dong / {vn(bb[k][1]):>20} d")

print()
print('--- 5.1 Cap NCC mua -> khach ban cua nhom VO/RUOT XE ---')
vm = [r for r in MV if nhom(r['mh']).startswith('B.')]
vb = [r for r in BR if nhom(r['mh']).startswith('B.')]
print('Mua tu:')
for k, v in sorted(agg_by(vm, lambda r: (r['mst_ban'], r['ten_ban'])).items(), key=lambda x: -x[1][1]):
    print(f"   {k[0]} {v[0]:>4} dong {vn(v[1]):>18}  {k[1][:44]}")
print('Ban cho:')
for k, v in sorted(agg_by(vb, lambda r: (r['mst_mua'], r['ten_mua'])).items(), key=lambda x: -x[1][1]):
    print(f"   {k[0]} {v[0]:>4} dong {vn(v[1]):>18}  {k[1][:44]}")

print()
print('--- 5.2 Cap NCC mua -> khach ban cua nhom THEP ---')
tm_ = [r for r in MV if nhom(r['mh']).startswith('A.')]
tb_ = [r for r in BR if nhom(r['mh']).startswith('A.')]
print('Mua tu:')
for k, v in sorted(agg_by(tm_, lambda r: (r['mst_ban'], r['ten_ban'])).items(), key=lambda x: -x[1][1]):
    print(f"   {k[0]} {v[0]:>4} dong {vn(v[1]):>20}  {k[1][:44]}")
print('Ban cho:')
for k, v in sorted(agg_by(tb_, lambda r: (r['mst_mua'], r['ten_mua'])).items(), key=lambda x: -x[1][1]):
    print(f"   {k[0]} {v[0]:>4} dong {vn(v[1]):>20}  {k[1][:44]}")

print()
print('=' * 104)
print('PHAN 6. QUAN HE THOI GIAN MUA - BAN (cung mat hang chuan hoa)')
print('=' * 104)
gap = collections.Counter(); gapval = collections.Counter(); ban_truoc = []
mv_by = collections.defaultdict(list)
for r in MV:
    if r['mh']: mv_by[norm(r['mh'])].append(r)
for k in mv_by: mv_by[k].sort(key=lambda r: r['ngay'])
for r in BR:
    k = norm(r['mh'])
    if k not in mv_by: continue
    ds = [(abs((r['ngay'] - m['ngay']).days), (r['ngay'] - m['ngay']).days, m) for m in mv_by[k]]
    ds.sort()
    d0 = ds[0]
    lab = ('Ban TRUOC khi mua' if d0[1] < 0 else 'Cung ngay' if d0[1] == 0 else
           '1 - 3 ngay' if d0[1] <= 3 else '4 - 10 ngay' if d0[1] <= 10 else
           '11 - 30 ngay' if d0[1] <= 30 else '> 30 ngay')
    gap[lab] += 1; gapval[lab] += r['chuathue']
    if d0[1] < 0: ban_truoc.append((r, d0[2], d0[1]))
for o in ['Ban TRUOC khi mua','Cung ngay','1 - 3 ngay','4 - 10 ngay','11 - 30 ngay','> 30 ngay']:
    if gap[o]: print(f"{o:<22}{gap[o]:>6} dong ban ra{vn(gapval[o]):>24} d")
print()
print(f"--- 6.1 Dong ban ra co ngay SOM HON moi lan mua cung mat hang: {len(ban_truoc)} dong, "
      f"{vn(sum(x[0]['chuathue'] for x in ban_truoc))} d")
for r, m, d in sorted(ban_truoc, key=lambda x: -x[0]['chuathue'])[:15]:
    print(f"   Ban {r['ngay']:%d/%m/%Y} HD {r['kh']}/{r['so']} {vn(r['chuathue']):>18} d <- mua som nhat {m['ngay']:%d/%m/%Y} (lech {d} ngay)")
    print(f"      {r['mh'][:80]}")

print()
print('=' * 104)
print('PHAN 7. KIEM TRA CHUOI MUA DI BAN LAI NGAY TRONG NGAY / 1-2 NGAY')
print('=' * 104)
same = []
for r in BR:
    k = norm(r['mh'])
    for m in mv_by.get(k, []):
        d = (r['ngay'] - m['ngay']).days
        if 0 <= d <= 2 and m['chuathue'] > 0 and abs(r['chuathue'] - m['chuathue']) / m['chuathue'] < 0.02:
            same.append((r, m, d, (r['chuathue'] - m['chuathue']) / m['chuathue'] * 100))
            break
print(f"So cap mua-ban trung khop (lech tien < 2%, cach nhau 0-2 ngay): {len(same)}")
print(f"Tong tien ban cua cac cap nay: {vn(sum(x[0]['chuathue'] for x in same))} d")
print(f"{'Ngay mua':<12}{'NCC':<16}{'Ngay ban':<12}{'Khach':<16}{'Tien mua':>18}{'Tien ban':>18}{'Lai %':>8}  Mat hang")
for r, m, d, pct in sorted(same, key=lambda x: -x[0]['chuathue'])[:30]:
    print(f"{m['ngay']:%d/%m/%Y}  {m['mst_ban']:<16}{r['ngay']:%d/%m/%Y}  {r['mst_mua']:<16}"
          f"{vn(m['chuathue']):>18}{vn(r['chuathue']):>18}{pct:>7.3f}%  {r['mh'][:40]}")

print()
print('=' * 104)
print('PHAN 8. DIEN BIEN THEO THANG - NHOM THEP (tien chua thue)')
print('=' * 104)
mon_m = agg_by([r for r in MV if nhom(r['mh']).startswith('A.')], lambda r: f"{r['ngay']:%Y-%m}")
mon_b = agg_by([r for r in BR if nhom(r['mh']).startswith('A.')], lambda r: f"{r['ngay']:%Y-%m}")
print(f"{'Thang':<10}{'Mua thep':>22}{'Ban thep':>22}{'Ban - Mua':>22}{'Ty le ban/mua':>16}")
cm = cb = 0
for k in sorted(set(mon_m) | set(mon_b)):
    m = mon_m.get(k, [0, 0, 0])[1]; b = mon_b.get(k, [0, 0, 0])[1]
    cm += m; cb += b
    rt = f"{b/m:.2f}" if m else ('-' if not b else 'khong co dau vao')
    print(f"{k:<10}{vn(m):>22}{vn(b):>22}{vn(b-m):>22}{rt:>16}")
print('-' * 92)
print(f"{'TONG':<10}{vn(cm):>22}{vn(cb):>22}{vn(cb-cm):>22}{cb/cm:>16.2f}")

print()
print('=' * 104)
print('PHAN 9. TONG HOP CAN DOI & CAC DAU HIEU CAN LAM RO')
print('=' * 104)
print(f"1) Ban ra {vn(tb)} d / Mua vao {vn(tm)} d -> ty le ban/mua = {tb/tm:.2f} lan;")
print(f"   thieu dau vao (ban - mua) = {vn(tb - tm)} d.")
print(f"2) Doanh thu tu mat hang KHONG co bat ky dau vao cung ten: {vn(tb_only)} d ({tb_only/tb*100:.2f}%).")
print(f"3) 33 mat hang mua - ban gan nhu bang nhau (bien lai duoi 0,15%): ban {vn(sum(d[3] for d in near))} d.")
print(f"4) {len(same)} cap mua - ban trung tien (lech <2%) trong vong 0-2 ngay: {vn(sum(x[0]['chuathue'] for x in same))} d doanh thu.")
print(f"5) Nhom chi phi khong he xuat hien o dau ra: xang dau {vn(g_mv['E. Xang dau'][1])} d, "
      f"dich vu-tai chinh {vn(g_mv['F. Dich vu / tai chinh / vien thong'][1])} d, "
      f"khac {vn(g_mv['G. Khac (van phong pham, an uong, tieu hao...)'][1])} d.")
print(f"6) {g_mv['Z. Khong ghi mat hang'][0]} dong mua vao KHONG co ten mat hang: {vn(g_mv['Z. Khong ghi mat hang'][1])} d.")
blank_big = sorted([r for r in MV if not r['mh']], key=lambda r: -r['chuathue'])[:8]
for r in blank_big:
    print(f"   {r['ngay']:%d/%m/%Y} HD {r['kh']}/{r['so']} {r['mst_ban']} {vn(r['chuathue']):>18} d  {r['ten_ban'][:40]}")
print(f"7) Nhom D (cho thue van phong) ban ra {vn(g_br['D. Cho thue van phong / mat bang'][1])} d "
      f"nhung mua vao nhom nay = 0 d.")
gkhac = [r for r in BR if nhom(r['mh']).startswith('G.')]
print(f"8) Nhom G ban ra {len(gkhac)} dong {vn(sum(r['chuathue'] for r in gkhac))} d:")
for r in gkhac:
    print(f"   {r['ngay']:%d/%m/%Y} HD {r['kh']}/{r['so']} {r['mst_mua']} {vn(r['chuathue']):>18} d  {r['mh'][:55]}")
