---
name: Tổng hợp TNCN và tra cứu đa nguồn
description: "Tổng hợp QTT TNCN, tờ khai quý, danh sách NLĐ và tra cứu đa nguồn; đối chiếu TC22/TC23."
group: Tài chính
updated: '2026-10-15'
---
> **Hợp đồng đầu ra v1.4:** module nội bộ. Kết quả ghi vào sheet đích của sáu tệp KQ (xem skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md). Không giao tệp chuyên đề Tong_hop_*/Doi_chieu_* / Master 8 sheet như sản phẩm.

# Tổng hợp TNCN và tra cứu đa nguồn

Bóc XML quyết toán TNCN (mẫu 953 / 05/QTT), tờ khai khấu trừ quý (05/KK), phụ lục NLĐ; đối chiếu quý vs năm (TC22) và chi phí lương BCTC (TC23); nếu có file tra cứu đa nguồn thì rà 1 nguồn / nhiều nguồn và trùng CCCD.

Skill trục: `phan-tich-ho-so-thue` playbook Bước 4. Chính sách Luật 109/2025 nằm ở skill `thue-tncn` (shared-skills), **không** thay skill này. Playbook lệnh: `playbook-lenh.md`.

Đầu vào NLĐ thô: `tong-hop-dau-vao-hddt-bang-ke-nld`.

## Khi nào dùng

- "tổng hợp TNCN", "QTT TNCN", "đa nguồn", "trùng CCCD", "TC22 TC23", "05/KK vs quyết toán năm".
- Đã có XML TNCN trong `1.1. To khai XML/`.

## Chuẩn bị

1. XML QTT TNCN các năm trong kỳ; tốt hơn đủ 05/KK quý.
2. Phụ lục 05-1/BK-TNCN hoặc 05-2.
3. (Tuỳ chọn) file tra cứu TNCN đa nguồn tại `3.3. Tra cuu TNCN da nguon & Nhan su/`.
4. (Tuỳ chọn) bảng lương / TK 334, 622, 641, 642 để TC23.
5. Bảng kê NLĐ đã gộp (mẫu 02) nếu có.

## Quy trình

1. Đọc XML QTT từng năm: số NLĐ, tổng TNCT, số thuế phải nộp / đã khấu trừ / nộp thừa thiếu.
2. Ghép tờ khai quý 05/KK vs quyết toán năm (TC22). Liệt kê quý thiếu XML, không bịa số.
3. Xuất / cập nhật danh sách NLĐ theo năm vào `1.4. Tong hop danh sach NLD theo QTT TNCN/` và `cau-thanh/`.
4. Nguồn 05-2, công ty không trích 10%: đánh dấu đề xuất kiểm tra cam kết.
5. Nếu có file đa nguồn ở `3.3`: map CCCD/MST cá nhân; kết luận 1 nguồn hay ≥ 2 nguồn; thuế khấu trừ nơi khác; **trùng CCCD** (1 số gắn nhiều họ tên).
6. TC23: dòng tiền / chi phí lương BCTC vs TNCT trên QTT. Nhớ 622 trên giá thành thường chỉ công nhân trực tiếp; QTT gồm cả 641/642.
7. Ghi sheet `KQ2_TNCN_LD2` của KQ2.1: bảng số + khối nhận xét module (kê khai/khấu trừ/đa nguồn/trùng CCCD, ngoại lệ 3.4). Thiếu đa nguồn: vẫn lập QTT, ghi `thiếu nguồn 3.3`. Mức rủi ro/điểm/đề cương vẫn chỉ ở KQ3.

File: `Tong_hop_TNCN_va_Da_nguon_<MST>_<KY>.xlsx` tại `cau-thanh/`.

## Định dạng số tiền

Hiển thị chuẩn VN: `.` hàng nghìn, `,` phần lẻ. Excel `#,##0` / `#,##0.00`. Cấm ô text bắt đầu bằng `=`.

## Bẫy

- Coi 622 giá thành = toàn bộ quỹ lương QTT.
- Kết luận 100% 1 nguồn khi chưa có file `3.3`.
- Bỏ qua trùng CCCD khác họ tên.
- Điền 0 giả cho quý thiếu XML.

## Kiểm chứng

- Đủ năm có XML QTT; quý thiếu được liệt kê.
- Sheet Master có nhận xét, không chỉ bảng số.
- Trùng CCCD (nếu có file đa nguồn) được nêu tên/năm.
- Bước tiếp: Master 28 TC trong `phan-tich-ho-so-thue`.
