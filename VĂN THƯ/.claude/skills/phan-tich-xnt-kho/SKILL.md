---
name: Phân tích XNT kho
description: "Đối chiếu XNT kho 152/155/156 với BCTC HTK, xuất kho với giá vốn 632, giải tỏa TC3."
group: Tài chính
updated: '2026-10-15'
---
# Phân tích XNT kho

Đối chiếu sổ Nhập-Xuất-Tồn với BCTC hàng tồn kho (TC3) và xuất kho với giá vốn 632. DN thương mại dùng 156; DN sản xuất cộng 152+155 (+156 nếu có).

Skill trục: `phan-tich-ho-so-thue` playbook Bước 6. Giá thành / định mức: `phan-tich-gia-thanh-dinh-muc`. Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "phân tích XNT", "sổ kho 152 155 156", "đối chiếu tồn kho BCTC", "TC3", "NXT".
- Có sổ XNT hoặc NXT trong `2.2. Kho - Chi phi - Gia thanh/` và XML/BCTC.

## Chuẩn bị

1. Sổ tổng hợp / chi tiết XNT: 152 (NVL), 155 (thành phẩm), 156 (hàng hóa) theo năm.
2. BCTC XML: DNN đọc HTK ở **ct140**; báo cáo đầy đủ dùng chỉ tiêu **[141]** = 152+155+156.
3. Sổ 632 hoặc kết quả `doi-chieu-511-632` để khớp xuất kho vs giá vốn.

## Quy trình

1. Gom tồn CK từng TK từng năm từ sổ XNT (không lấy dòng Tổng cộng).
2. DN thương mại: lấy 156. Sản xuất: 152+155, cộng 156 nếu có hàng hóa.
3. Đối chiếu tồn CK vs BCTC:
   - Đầy đủ: [141] = 152+155+156, khớp 100% hoặc ghi lệch kèm nguyên nhân.
   - DNN: dùng **ct140**, không kết luận HTK=0 chỉ vì ct141=0.
4. Map năm BCTC theo doanh thu nếu nhãn `kyKKhai` lệch.
5. Xuất kho vs GV 632 (và [02] BCTC nếu đã chạy giá thành).
6. Tỷ lệ HTK / Doanh thu; ngưỡng TC3 trên TUDIEN: HTK + XDCB dở dang > **30%** DT (`skills/phan-tich-ho-so-thue/references/tieu-chi-28.md`).
7. Sheet Master `Tong_hop_XNT_va_Kho`: bảng theo năm + nhận xét TC3 (giải tỏa / xác nhận rủi ro). Thiếu sổ: ghi thiếu nguồn, không để 0 giả.

File: `Phan_tich_XNT_va_Gia_thanh_<MST>_<KY>.xlsx` hoặc cập nhật thẳng Master; lưu `cau-thanh/`.

## Định dạng số tiền

Hiển thị chuẩn VN: `.` hàng nghìn, `,` phần lẻ. Excel `#,##0` / `#,##0.00`.

## Bẫy

- Chỉ lấy 152 rồi bảo lệch [141].
- Kết luận HTK=0 trên BCTC DNN vì đọc nhầm ct141.
- Chốt TC3 trước khi xuất kho 156 khớp 632 (DN thương mại).

## Kiểm chứng

- Tổng tồn CK đúng tập TK theo loại hình DN.
- Lệch BCTC có số và nguyên nhân, không im lặng.
- Sheet Master có nhận xét TC3.
- Bước tiếp: `phan-tich-gia-thanh-dinh-muc`.
