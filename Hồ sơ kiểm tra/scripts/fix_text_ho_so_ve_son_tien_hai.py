#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Pass 2: sua cac truong con sot tu mau DNTN Thuy Hong + dien so tien
vao Ket luan BBKT, Bao cao ket qua va Quyet dinh XPHC cua vu Son Tien Hai."""
import shutil
import sys
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

SRC = '/brains/Hồ sơ kiểm tra/exports/son-tien-hai-3603859012/TONG HOP.docx'

DIACHI = 'Số 23, Tổ 1, KP 3A, Phường Long Hưng, Thành phố Đồng Nai'
NGANH = 'Bán buôn vật liệu, thiết bị lắp đặt khác trong xây dựng'

# so tien chot tu Bang tinh (sheet DTCHEP + TINH PHAT + CHAM NOP (C T))
TRUY_THU = '13.052.423'
TRUY_THU_CHU = 'Mười ba triệu không trăm năm mươi hai nghìn bốn trăm hai mươi ba đồng'
PHAT = '2.610.485'
PHAT_CHU = 'Hai triệu sáu trăm mười nghìn bốn trăm tám mươi lăm đồng'
CHAM_NOP = '387.865'
TONG_PHAT = '2.998.350'
TONG_CONG = '16.050.773'
TONG_CONG_CHU = 'Mười sáu triệu không trăm năm mươi nghìn bảy trăm bảy mươi ba đồng'
GIAM_LO = '809.868.861'

# (old, new) - thu tu quan trong: chuoi dai dat truoc
SUBS = [
    # dia chi
    ('Số 23, Tổ 1, KP 3A, Tỉnh Đồng Nai, Thành phố Đồng Nai', DIACHI),
    ('Số 23, Tổ 1, KP 3A, Tỉnh Đồng Nai, Đồng Nai', DIACHI),
    ('Số 23, Tổ 1, KP 3A, Tỉnh Đồng Nai', DIACHI),
    # CCCD nguoi dai dien (dang la CCCD cua Thuy Hong)
    ('075163005918 cấp ngày: 21.09.2021', '034083000616 cấp ngày: 15.01.2023'),
    ('075163005918', '034083000616'),
    # nganh nghe
    ('Bán lẻ hàng hóa khác mới trong các cửa hàng chuyên doanh...', NGANH + '.'),
    ('Bán lẻ hàng hóa khác mới trong các cửa hàng chuyên doanh', NGANH),
    # loai hinh / hach toan
    ('- Loại hình doanh nghiệp: Doanh nghiệp tư nhân',
     '- Loại hình doanh nghiệp: Công ty trách nhiệm hữu hạn'),
    # thoi han kiem tra theo QDKT: 20 ngay
    ('- Thời gian tiến hành kiểm tra: 10 ngày',
     '- Thời gian tiến hành kiểm tra: 20 ngày'),
    # ngay cap GCN DKKD
    ('đăng ký kinh doanh số 3603859012 ngày ..../..../........',
     'đăng ký kinh doanh số 3603859012 ngày 09/03/2022'),
    ('đăng ký doanh nghiệp số: 3603859012 Ngày cấp:  ..../..../........',
     'đăng ký doanh nghiệp số: 3603859012 Ngày cấp: 09/03/2022'),
    # don vi thanh vien
    ('- Các đơn vị thành viên (nếu có): Các xí nghiệp, chi nhánh, cửa hàng '
     '(Ghi rõ nơi đóng trụ sở hình thức hạch toán).',
     '- Các đơn vị thành viên: CHI NHÁNH CÔNG TY TNHH THƯƠNG MẠI SẢN XUẤT SƠN TIỀN HẢI, '
     'mã số thuế: 3603859012-001.'),
    # hoa don: DN thanh lap 2022, dung HDDT theo ND 123/2020, khong phai TT 39/2014
    ('- Về việc xử lý hóa đơn: Do Công Ty sử dụng hóa đơn theo Thông tư số '
     '39/2014/TT-BTC ngày 31/3/2014 của Bộ Tài chính và đã thực hiện hủy hóa đơn '
     'theo quy định.',
     '- Về việc xử lý hóa đơn: Công ty sử dụng hóa đơn điện tử theo quy định tại Nghị '
     'định số 123/2020/NĐ-CP ngày 19/10/2020 của Chính phủ (được sửa đổi, bổ sung tại '
     'Nghị định số 70/2025/NĐ-CP ngày 20/3/2025) và đã thực hiện ...... theo quy định.'),
    # ngay header Bao cao (con sot ngay cua Thuy Hong)
    ('Trấn Biên, ngày 23   tháng  10 năm 2026',
     'Trấn Biên, ngày       tháng       năm 2026'),

    # ---- BBKT muc IV: hinh thuc xu phat + bien phap khac phuc
    ('- Phạt tiền với số tiền: .......400.000 đồng (Bằng chữ Một triệu năm trăm bảy '
     'mươi bảy nghìn năm trăm mười ba đồng), trong đó:',
     f'- Phạt tiền với số tiền: {PHAT} đồng (Bằng chữ: {PHAT_CHU}), trong đó:'),
    ('+ Phạt về hành vi vi phạm thủ tục thuế:  đồng;',
     '+ Phạt về hành vi vi phạm thủ tục thuế: 0 đồng;'),
    ('+ Phạt về hành vi khai sai dẫn đến thiếu số tiền thuế phải nộp (theo tỷ lệ 20% '
     'trên số thuế truy thu): .......400.000 đồng;',
     '+ Phạt về hành vi khai sai dẫn đến thiếu số tiền thuế phải nộp (theo tỷ lệ 20% '
     f'trên số thuế truy thu): {PHAT} đồng;'),
    ('-  Truy thu thuế: 39.937.049.637.700.000 đồng (Bằng chữ: Ba triệu chín trăm '
     'chín mươi ba nghìn bảy trăm lẻ năm đồng), trong đó:',
     f'-  Truy thu thuế: {TRUY_THU} đồng (Bằng chữ: {TRUY_THU_CHU}), trong đó:'),
    ('+ Thuế TNDN (TM 1052), số tiền: 2.662.470 đồng.',
     f'+ Thuế TNDN (TM 1052), số tiền: {TRUY_THU} đồng.'),
    ('+ Thuế GTGT (TM 1701), số tiền: ...... đồng;',
     '+ Thuế GTGT (TM 1701), số tiền: 0 đồng;'),
    ('Giảm lỗ: 6.404.948 đồng;', f'Giảm lỗ: {GIAM_LO} đồng;'),
    ('Giảm khấu trừ:  đồng;', 'Giảm khấu trừ: 0 đồng;'),

    # ---- Bao cao ket qua kiem tra
    ('- Kiến nghị xử lý truy thu: 39.937.049.637.700.000\t\t\t\t\t đồng, trong đó:',
     f'- Kiến nghị xử lý truy thu: {TRUY_THU} đồng, trong đó:'),
    ('+ Thuế TNDN:  \t2.662.470 đồng.', f'+ Thuế TNDN: {TRUY_THU} đồng.'),
    ('+ Thuế GTGT: \t...... đồng.', '+ Thuế GTGT: 0 đồng.'),
    ('- Kiến nghị xử phạt: Tổng cộng số tiền phạt: 15.775.134.589.284.000 đồng, trong đó:',
     f'- Kiến nghị xử phạt: Tổng cộng số tiền phạt: {TONG_PHAT} đồng, trong đó:'),
    ('với số tiền:  77.877.246.617.437.000 đồng.', f'với số tiền: {CHAM_NOP} đồng.'),
    ('của Chính phủ, với số tiền: .......400.000 đồng.',
     f'của Chính phủ, với số tiền: {PHAT} đồng.'),
    ('Nghị định số 125/2020/NĐ-CP ngày 19/10/2020 của Chính phủ và Nghị định '
     '310/2025/NĐ-CP ngày 02/12/2025 của Chính phủ:  đồng;',
     'Nghị định số 125/2020/NĐ-CP ngày 19/10/2020 của Chính phủ và Nghị định '
     '310/2025/NĐ-CP ngày 02/12/2025 của Chính phủ: 0 đồng;'),

    # ---- Quyet dinh xu phat VPHC
    ('a) Hình thức xử phạt hành chính: phạt tiền, mức phạt:\t .......400.000 đồng.',
     f'a) Hình thức xử phạt hành chính: phạt tiền, mức phạt: {PHAT} đồng.'),
    ('+ Phạt kê khai sai (TM 4254):\t .......400.000 đồng.',
     f'+ Phạt kê khai sai (TM 4254): {PHAT} đồng.'),
    ('+ Phạt về hành vi vi phạm thủ tục thuế (TM 4254):  đồng;',
     '+ Phạt về hành vi vi phạm thủ tục thuế (TM 4254): 0 đồng;'),
    ('+ Thuế TNDN (TM 1052), số tiền:  \t2.662.470 đồng;',
     f'+ Thuế TNDN (TM 1052), số tiền: {TRUY_THU} đồng;'),
    ('+ Thuế GTGT (TM 1701), số tiền:  ...... \tđồng;',
     '+ Thuế GTGT (TM 1701), số tiền: 0 đồng;'),
    ('+ Tiền chậm nộp thuế TNDN (TM 4918), số tiền:  45.448.362.900.000.000 đồng;',
     f'+ Tiền chậm nộp thuế TNDN (TM 4918), số tiền: {CHAM_NOP} đồng;'),
    ('+ Tiền chậm nộp thuế GTGT (TM 4931), số tiền: ...... đồng;',
     '+ Tiền chậm nộp thuế GTGT (TM 4931), số tiền: 0 đồng;'),
    ('Tổng số tiền truy thu, phạt: 55.712.184.226.984.000 đồng (bằng chữ: Năm triệu '
     'năm trăm bảy mươi mốt nghìn hai trăm mười tám đồng).',
     f'Tổng số tiền truy thu, phạt: {TONG_CONG} đồng (bằng chữ: {TONG_CONG_CHU}).'),
    # loi dan chieu NÐ 125 trong To trinh QD xu phat
    ('Nghị định số 125/2022/NĐ-CP ngày 19/10/2022 của Chính phủ',
     'Nghị định số 125/2020/NĐ-CP ngày 19/10/2020 của Chính phủ'),
    ('Nơi cấp Sở Tài chính Thành phố Đồng Nai cấp',
     'Nơi cấp: Sở Tài chính tỉnh Đồng Nai'),
]


def rewrite(p, old, new):
    """Thay old -> new trong paragraph, uu tien thay trong 1 run de giu dinh dang."""
    for r in p.runs:
        if old in r.text:
            r.text = r.text.replace(old, new)
            return True
    full = p.text
    if old not in full:
        return False
    runs = p.runs
    if not runs:
        return False
    runs[0].text = full.replace(old, new)
    for r in runs[1:]:
        r._element.getparent().remove(r._element)
    return True


def iter_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                yield p


def main():
    shutil.copy2(SRC, SRC + '.bak-truoc-fix-text')
    doc = Document(SRC)
    counts = {old: 0 for old, _ in SUBS}
    for p in list(iter_paragraphs(doc)):
        for old, new in SUBS:
            while old in p.text:
                if not rewrite(p, old, new):
                    break
                counts[old] += 1
    doc.save(SRC)
    for old, n in counts.items():
        flag = 'OK ' if n else '!! '
        print(f'{flag}{n}x  {old[:70]}')


if __name__ == '__main__':
    sys.exit(main())
