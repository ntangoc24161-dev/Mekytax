---
name: Phân tích giá thành và định mức
description: "Bóc giá thành 621/622/627 hoặc cơ cấu GV thương mại, khớp 632 với BCTC [02], giải tỏa TC28."
group: Tài chính
updated: '2026-10-15'
---
> **Hợp đồng đầu ra v1.4:** module nội bộ. Kết quả ghi vào sheet đích của sáu tệp KQ (xem skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md). Không giao tệp chuyên đề Tong_hop_*/Doi_chieu_* / Master 8 sheet như sản phẩm.

# Phân tích giá thành và định mức

Giải tỏa TC28: DN sản xuất bóc 621/622/627, bảng giá thành, định mức; DN thương mại lập cơ cấu Doanh thu - GV 632 - lãi gộp - chi phí BH/QL, không bắt buộc 621/622/627.

Skill trục: `phan-tich-ho-so-thue` playbook Bước 7. Kho: `phan-tich-xnt-kho`. Doanh thu-GV từng HĐ: `doi-chieu-511-632`. Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "giá thành", "định mức NVL", "621 622 627", "TK 154", "TC28", "cơ cấu giá vốn thương mại".
- Có sổ 154 / bảng Z SP (sản xuất) hoặc sổ 632 + TNDN/BCTC (thương mại).

## Chuẩn bị

**Sản xuất:** sổ 154 / 621 / 622 / 627, bảng tính giá thành, định mức tiêu hao đã ban hành, sổ 632, BCTC [02].

**Thương mại:** sổ 632, kết quả 511-632 nếu có, chỉ tiêu DT/GV/CP BH/QL trên TNDN hoặc BCTC.

## Quy trình

### Sản xuất

1. Bóc cơ cấu 621 (NVL), 622 (nhân công TT), 627 (SXC) theo năm / mã SP.
2. Khớp giá vốn xuất bán TK 632 với BCTC [02] (100% hoặc ghi lệch).
3. Đối chiếu định mức kỹ thuật vs NVL xuất dùng thực tế (từ XNT 152).
4. 622 giá thành ≠ quỹ lương QTT (còn 641/642).

### Thương mại

1. Sheet `Gia_thanh_va_Dinh_muc`: DT - GV 632 - lãi gộp - CP bán hàng / QLDN (từ TNDN).
2. Không bắt buộc 621/622/627.
3. Xuất kho 156 khớp 632 trước khi chốt TC3/TC28 (`phan-tich-xnt-kho`).
4. Dùng tỷ lệ lãi gộp / DT bán dưới giá từ `doi-chieu-511-632` nếu đã chạy.

Cả hai: ghi sheet `KQ2_Gia_thanh_va_Dinh_muc` của KQ2.2 kèm khối nhận xét module biên LN / TC28 / cơ cấu GV (ngoại lệ 3.4). Thiếu nguồn: viết hiện trạng, không 0 giả. Mức rủi ro/điểm vẫn chỉ ở KQ3.

TC28 trên TUDIEN: phân tích ngang/dọc, giá vốn tăng nhanh hơn DT (`tieu-chi-28.md`).

## Định dạng số tiền

Hiển thị chuẩn VN: `.` hàng nghìn, `,` phần lẻ. Excel `#,##0` / `#,##0.00`. Tỷ lệ `12,34%`.

## Bẫy

- Ép DN thương mại phải có 621/622/627.
- Coi 622 = toàn bộ lương QTT.
- Chốt TC28 khi 632 chưa khớp [02] / xuất kho.

## Kiểm chứng

- 632 vs [02] có số khớp hoặc lệch có nguyên nhân.
- Sheet Master có kết luận, không chỉ bảng.
- Bước tiếp: `doi-chieu-511-632` (nếu chưa) rồi Master 28 TC.
