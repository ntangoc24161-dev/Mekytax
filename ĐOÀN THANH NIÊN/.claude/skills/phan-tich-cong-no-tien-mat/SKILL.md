---
name: Phân tích công nợ và tiền mặt
description: "Quét chi tiền mặt TK 111, cụm HĐ cùng ngày, công nợ 131/331 và sao kê ngân hàng 112."
group: Tài chính
updated: '2026-10-15'
---
# Phân tích công nợ và tiền mặt

Quét phiếu chi TK 111 vượt ngưỡng, cụm nhiều HĐ cùng ngày cùng NCC (CD1 BT7), công nợ 131/331 (TC20), sao kê TK 112. Gắn BT7 từ `lam-sach-cd1-mst-hoa-don`.

Skill trục: `phan-tich-ho-so-thue` playbook Bước 8. Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "quét tiền mặt 111", "cụm hóa đơn cùng ngày", "sao kê ngân hàng", "công nợ 131 331", "UNC", "ngưỡng 20 triệu 5 triệu".
- Có sổ 111 / 131 / 331 / 112 hoặc sao kê trong `2.3` / `2.1`.

## Chuẩn bị

1. Sổ chi tiết 1111, 1121, 131, 331 (Excel).
2. Sao kê NH: ưu tiên Excel; PDF scan thì ghi hiện trạng, không kết luận thiếu UNC.
3. Bảng kê / HĐĐT mua vào đã có `MST_chuan` để SUMIFS cụm ngày+NCC.
4. Ngưỡng: > **20.000.000** đến 06/2025; ≥ **5.000.000** từ 07/2025 (Luật GTGT 48/2024). Theo kỳ phát sinh HĐ.

## Quy trình

1. **TK 111:** liệt kê phiếu chi ≥ ngưỡng. Phân loại: nộp 112, trả vay/mượn 336/3388, lương 334, mua hàng.
2. **BT7 cụm cùng ngày:** SUMIFS doanh số+thuế theo ngày + MST/NCC; flag nhóm vượt ngưỡng dù từng HĐ nhỏ. Bắt buộc chứng từ NH theo TT 219/2013 và Luật 48/2024.
3. **TK 112 / sao kê:** đối chiếu số dư / phát sinh lớn; nếu chỉ PDF thì kiến nghị Excel + chọn mẫu UNC gốc (HĐ ≥ ngưỡng, NCC F0).
4. **TK 131:** phải thu, người mua trả trước (131 Có) vs ghi nhận DT (TC2 nếu > 5% DT).
5. **TK 331:** nợ NCC lớn (TC20); thanh toán NH với NCC lớn / F0.
6. Master sheet `Cong_no_va_Ngan_hang`: số thật + Phần 4 báo cáo thiếu dữ liệu và kiến nghị đoàn. Không 0 giả.

File: `Phan_tich_Tien_mat_va_Ngan_hang_<MST>_<KY>.xlsx` tại `cau-thanh/`.

## Định dạng số tiền

Hiển thị chuẩn VN: `.` hàng nghìn, `,` phần lẻ. Excel `#,##0` / `#,##0.00`.

## Bẫy

- Chỉ soi từng dòng HĐ, bỏ cụm cùng ngày.
- Kết luận không có chứng từ NH khi NNT mới gửi PDF.
- Ngưỡng theo ngày chạy file thay vì kỳ phát sinh HĐ.
- Ô diễn giải bắt đầu bằng `=`.

## Kiểm chứng

- Danh sách chi 111 có ngưỡng đúng giai đoạn.
- Cụm BT7 có số nhóm / số dòng flag.
- Sao kê PDF được ghi là PDF, không bịa thiếu UNC.
- Sheet Master có kiến nghị.
- Bước tiếp: Master 8 sheet + 28 TC trong `phan-tich-ho-so-thue`.
