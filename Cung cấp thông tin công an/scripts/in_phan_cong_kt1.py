"""In phan cong KT1: doc NHAP_HO_SO, chon dong To KT=KT1 + tinh trang 'Mới',
gom theo (can bo, cong van), dong dau trang 1 cong van ghi can bo/MST/ten cty/ngay phan cong,
moi can bo mot PDF, tong so trang chan, sau do ghep PDF tong hop.
Dung: python in_phan_cong_kt1.py [--apply]   (--apply: doi 'Mới' -> 'Đang xử lý' bang Excel COM)
"""
import sys, re, datetime, collections, unicodedata
from pathlib import Path
import openpyxl, pymupdf

XLSX = Path(r'I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx')
ROOT = XLSX.parent / '1. CUNG CAP CONG AN NAM 2026'
OUT = Path(r'D:\SECOND BRAIN\javis-os\brains\Cung cấp thông tin công an\exports\in-phan-cong-kt1')
FONT = r'C:\Windows\Fonts\arialbd.ttf'


ONLY_ROWS = {int(a.split('=')[1]) for a in sys.argv if a.startswith('--row=')}  # vd --reprint --row=816 --row=817
REPRINT = '--reprint' in sys.argv  # in lai cac dong da 'Đang xử lý'


def norm_name(s):
    return ' '.join(str(s).split()).upper() if s else None


def find_cv_pdf(folder_no, so_cv):
    num = re.match(r'\s*[A-Za-z]?(\d+)', str(so_cv or ''))
    cands = [d for d in ROOT.iterdir() if d.is_dir() and d.name.startswith(f'{folder_no}.')] if folder_no else []
    for d in cands:
        pdfs = sorted(d.glob('*.pdf'), key=lambda p: -p.stat().st_size)
        pdfs = [p for p in pdfs if num and num.group(1) in p.name and 'TRIEN KHAI' not in p.name.upper()] or pdfs
        if pdfs:
            return pdfs[0]
    return None


def main(apply=False):
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb['NHAP_HO_SO']
    cv = {}
    jobs, missing = [], []
    for n, r in enumerate(ws.iter_rows(min_row=5, values_only=True), 5):
        if r[2]:
            cv = dict(so=r[2], ngay=r[3], folder=r[1], row=n)
        if ONLY_ROWS and n not in ONLY_ROWS:
            continue
        if str(r[14]).strip() != ('Đang xử lý' if REPRINT else 'Mới'):
            continue
        to = str(r[12]).strip().upper() if r[12] else ''
        if to == 'KT1' or (not to and (r[9] or r[2])):
            if to != 'KT1' or not r[13]:
                missing.append((n, cv.get('so'), r[9], to or '(chưa có tổ)', 'thiếu tổ' if to != 'KT1' else 'thiếu cán bộ'))
                continue
            jobs.append(dict(row=n, cb=norm_name(r[13]), mst=r[9], ten=r[10], ngay_pc=r[16], cv=dict(cv)))
    OUT.mkdir(parents=True, exist_ok=True)
    by_cb = collections.defaultdict(list)
    for j in jobs:
        by_cb[j['cb']].append(j)
    done_rows, report, outs = [], [], []
    for cb, items in by_cb.items():
        doc = pymupdf.open()
        for j in items:
            pdf = find_cv_pdf(j['cv'].get('folder'), j['cv'].get('so'))
            if not pdf:
                missing.append((j['row'], j['cv'].get('so'), j['mst'], 'KT1', 'không tìm thấy PDF công văn'))
                continue
            src = pymupdf.open(pdf)
            start = len(doc)
            doc.insert_pdf(src)
            pg = doc[start]
            nd = j['ngay_pc']
            nd = nd.strftime('%d/%m/%Y') if isinstance(nd, datetime.datetime) else (nd or 'chưa có')
            line = f"CB: {cb} | MST: {j['mst']} | {j['ten']} | Ngày PC: {nd} | CV {j['cv']['so']}"
            k = pg.rect.width / 595.0
            fs = 11 * k
            avail = pg.rect.width - 30 * k
            fnt = pymupdf.Font(fontfile=FONT)
            tl = fnt.text_length(line, fontsize=fs)
            if tl > avail:
                fs = fs * avail / tl * 0.98  # co chu de vua 1 dong
            pg.draw_rect(pymupdf.Rect(0, 0, pg.rect.width, fs * 2.2), color=None, fill=(1, 1, 1))
            pg.insert_text((15 * k, fs * 1.5), line, fontname='arialbd', fontfile=FONT, fontsize=fs, color=(0, 0, 0))
            done_rows.append(j['row'])
        if len(doc) % 2:
            doc.new_page(width=doc[0].rect.width, height=doc[0].rect.height)
        if len(doc):
            f = OUT / f"{unicodedata.normalize('NFKD', cb).encode('ascii', 'ignore').decode().replace(' ', '_')}.pdf"
            doc.save(f, garbage=3, deflate=True)
            outs.append(f)
            report.append((cb, len(items), len(doc)))
    total = pymupdf.open()
    for f in outs:
        total.insert_pdf(pymupdf.open(f))
    if len(total):
        import shutil
        tn = f'Phan cong xu ly KT1 tu dong {min(done_rows)} den dong {max(done_rows)} ngay in {datetime.date.today():%d-%m-%Y}.pdf'
        total.save(OUT / tn, garbage=3, deflate=True)
        DRIVE = Path(r'I:\My Drive\KIEM SOAT HSO CUNG CAP\IN PHAN CONG')
        DRIVE.mkdir(parents=True, exist_ok=True)
        shutil.copy2(OUT / tn, DRIVE / tn)
        print('SAVED', DRIVE / tn)
    print('REPORT', report)
    print('MISSING', missing)
    print('ROWS', done_rows)
    if apply and done_rows:
        import win32com.client
        xl = win32com.client.DispatchEx('Excel.Application')
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.AskToUpdateLinks = False
        try:
            w = xl.Workbooks.Open(str(XLSX), UpdateLinks=0, ReadOnly=False)
            s = w.Worksheets('NHAP_HO_SO')
            for n in done_rows:
                s.Cells(n, 15).Value = 'Đang xử lý'
            w.Save()
            w.Close(SaveChanges=True)
            print('STATUS UPDATED', len(done_rows))
        finally:
            xl.Quit()


if __name__ == '__main__':
    main('--apply' in sys.argv)




