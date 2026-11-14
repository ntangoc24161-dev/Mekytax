# -*- coding: utf-8 -*-
"""Dien bien ban tu Bang tinh (sheet DTCHEP, LOAI LUONG, TINH PHAT) cho Tam Hoang Phat (MST 3603721832)."""

import os
import copy
import shutil
import docx
import datetime
import sys
import win32com.client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
STAMP = datetime.datetime.now().strftime("%Y%m%d-%H%M")

BRAIN_EXPORT_DIR = r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\exports\tam-hoang-phat-3603721832"
BRAIN_TONG_HOP = os.path.join(BRAIN_EXPORT_DIR, "TONG HOP.docx")
BRAIN_TONG_HOP_NAMED = os.path.join(BRAIN_EXPORT_DIR, "TONG HOP - Tam Hoang Phat.docx")

G_DIR = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\10. TAM HOANG PHAT\0. BIEN BAN"
G_TONG_HOP = os.path.join(G_DIR, "TONG HOP.docx")
G_BAK_SRC = os.path.join(G_DIR, "TONG HOP.bak-truoc-cap-nhat-20261109-172629.docx")
G_BAK_NEW = os.path.join(G_DIR, "TONG HOP.bak-truoc-dien-dtchep-%s.docx" % STAMP)

def replace_slice(parent, start_el, end_el, new_elements):
    current = start_el
    to_remove = []
    while current is not None:
        to_remove.append(current)
        if current == end_el:
            break
        current = current.getnext()
    
    for new_el in new_elements:
        start_el.addprevious(new_el)
    
    for el in to_remove:
        parent.remove(el)

def main():
    print("1. Nạp tài liệu nguồn...")
    doc_old = docx.Document(G_BAK_SRC)
    doc_cur = docx.Document(BRAIN_TONG_HOP)

    old_body = doc_old.element.body
    cur_body = doc_cur.element.body

    # Lấy các khối đã chuẩn hóa 3 năm (2023-2025)
    # Block 1 (BBXN TV): 248..257
    block1 = [copy.deepcopy(old_body[i]) for i in range(248, 258)]

    # Block 2 (BBXN TD): 286..300
    block2 = [copy.deepcopy(old_body[i]) for i in range(286, 301)]

    # Block 3 (BBKT): 340..357
    block3 = [copy.deepcopy(old_body[i]) for i in range(340, 358)]

    # Block 4 (MUC III): 359
    block4 = copy.deepcopy(old_body[359])

    # Cập nhật chuẩn MST trên bảng kê lương theo tờ khai QTT TNCN (memory fact)
    # Table 2024: item 9 trong block2
    t_2024 = docx.table.Table(block2[9], doc_cur)
    t_2024.rows[1].cells[2].paragraphs[0].text = '3100211137'

    # Table 2025: item 14 trong block2
    t_2025 = docx.table.Table(block2[14], doc_cur)
    t_2025.rows[1].cells[2].paragraphs[0].text = '044090000152'

    print("2. Đổ số liệu vào tài liệu...")
    children = list(cur_body)

    # 1. Thay thế Block 4 (Mục III Nghĩa vụ NSNN)
    muc_iii_p = None
    for el in children:
        if el.tag.endswith('p') and 'III. TÌNH HÌNH THỰC HIỆN NGHĨA VỤ' in (el.text or ''):
            muc_iii_p = el
            break

    muc_iii_tbl = muc_iii_p.getnext()
    assert muc_iii_tbl.tag.endswith('tbl')
    muc_iii_tbl.addprevious(block4)
    cur_body.remove(muc_iii_tbl)
    print("  -> Đã cập nhật Mục III (7x7) cho 3 năm 2023-2025.")

    # 2. Thay thế Block 3 (BBKT)
    bbkt_intro_p = None
    for el in children:
        if el.tag.endswith('p') and 'II. NỘI DUNG KIỂM TRA TẠI TRỤ SỞ CỦA NGƯỜI NỘP THUẾ:' in (el.text or ''):
            bbkt_intro_p = el
            break

    start_b3 = bbkt_intro_p.getnext()
    end_b3 = muc_iii_p.getprevious()
    replace_slice(cur_body, start_b3, end_b3, block3)
    print("  -> Đã cập nhật BBKT cho 3 năm 2023-2025.")

    children = list(cur_body)

    # 3. Thay thế Block 2 (BBXN Trưởng đoàn)
    bbxn_td_intro = None
    found_tv = False
    for el in children:
        if el.tag.endswith('p') and 'II. Kết quả kiểm tra:' in (el.text or ''):
            if not found_tv:
                found_tv = True
            else:
                bbxn_td_intro = el
                break

    start_b2 = bbxn_td_intro.getnext()
    end_b2 = None
    curr = start_b2
    while curr is not None:
        nxt = curr.getnext()
        if nxt is not None and nxt.tag.endswith('p') and 'Biên bản này được lập' in (nxt.text or ''):
            end_b2 = curr
            break
        curr = nxt

    replace_slice(cur_body, start_b2, end_b2, block2)
    print("  -> Đã cập nhật BBXN Trưởng đoàn kèm bảng kê lương 3 năm 2023-2025.")

    children = list(cur_body)

    # 4. Thay thế Block 1 (BBXN Thành viên)
    bbxn_tv_intro = None
    for el in children:
        if el.tag.endswith('p') and 'II. Kết quả kiểm tra:' in (el.text or ''):
            bbxn_tv_intro = el
            break

    start_b1 = bbxn_tv_intro.getnext()
    end_b1 = None
    curr = start_b1
    while curr is not None:
        nxt = curr.getnext()
        if nxt is not None and nxt.tag.endswith('p') and 'Biên bản này được lập' in (nxt.text or ''):
            end_b1 = curr
            break
        curr = nxt

    replace_slice(cur_body, start_b1, end_b1, block1)
    print("  -> Đã cập nhật BBXN Thành viên cho 3 năm 2023-2025.")

    # 5. Khắc phục 2 vị trí sót Số QĐKT
    # Báo cáo kết quả kiểm tra (T39 dòng 1 cột 2) và Danh mục hồ sơ (T47 dòng 8 cột 2)
    for t in doc_cur.tables:
        if len(t.rows) >= 9 and 'Số TT' in t.rows[0].cells[0].text and 'Nội dung công việc' in t.rows[0].cells[1].text:
            if 'Ban hành quyết định' in t.rows[1].cells[1].text:
                t.rows[1].cells[2].paragraphs[0].text = '6899/QĐ-TCS3'
                print("  -> Đã chuẩn hóa Số hiệu QĐKT tại Báo cáo kết quả kiểm tra.")
        if len(t.rows) >= 9 and 'Trích yếu văn bản' in t.rows[0].cells[1].text:
            for r in t.rows:
                if 'Quyết định kiểm tra' in r.cells[1].text:
                    r.cells[2].paragraphs[0].text = '6899/QĐ-TCS3'
                    print("  -> Đã chuẩn hóa Số hiệu QĐKT tại Danh mục hồ sơ gửi pháp lý.")

    print("3. Lưu tài liệu hoàn chỉnh...")
    doc_cur.save(BRAIN_TONG_HOP)
    doc_cur.save(BRAIN_TONG_HOP_NAMED)
    print(f"  -> Lưu Brain: {BRAIN_TONG_HOP}")

    if os.path.exists(G_DIR):
        if os.path.exists(G_TONG_HOP):
            shutil.copyfile(G_TONG_HOP, G_BAK_NEW)
            print(f"  -> Đã tạo sao lưu ổ G: {G_BAK_NEW}")
        doc_cur.save(G_TONG_HOP)
        print(f"  -> Lưu ổ G: {G_TONG_HOP}")

    print("4. Kiểm tra mở bằng Microsoft Word COM...")
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        d = word.Documents.Open(os.path.abspath(BRAIN_TONG_HOP))
        page_count = d.ComputeStatistics(2)
        print(f"  -> Mở Microsoft Word thành công: {page_count} trang.")
        d.Close(False)
    finally:
        word.Quit()

    # Dọn dẹp file test
    test_out = os.path.join(BRAIN_EXPORT_DIR, "test_out.docx")
    if os.path.exists(test_out):
        os.remove(test_out)

if __name__ == "__main__":
    main()
