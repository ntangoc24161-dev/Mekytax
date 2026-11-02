---
type: decision
created: 2026-10-23
updated: 2026-10-23
---

## Quy trình phân tích hồ sơ kiểm tra thuế v1.4 - Hoàn thành

**Kỳ:** 2026-10-21 đến 2026-10-23  
**Hồ sơ:** T&T Biên Hòa (MST 3603883505), kỳ kiểm tra 2022-2025  
**Phương pháp:** Quy trình 3 giai đoạn v1.4 vòng 2b (Prompt P7)  
**Tổng file kết quả:** 6 tệp (KQ1.1, KQ1.2, KQ1.3, KQ2.1, KQ2.2, KQ3)

---

## GĐ1 - KIỂM KÊ & ĐÁNH GIÁ HỒ SƠ

**Công việc:**
- Quét hồ sơ từ thư mục CQT (1. DỮ LIỆU THO NGÀNH THUẾ) và NNT (2. HỒ SỢ DN CUNG CẤP)
- Chuẩn hóa dữ liệu vào 14 sheet IN00–IN14 (Manifest, HĐDT, bảng kê, sổ kế toán, kho, lãi vay, v.v.)
- Xác định mức hồ sơ theo số file: **Mức D** (CQT đủ 11 file, NNT thiếu sao kê trích xuất + GDLK)
- Kỳ kiểm tra: **2022-2025** (4 năm)

**Tiến độ:**
- 11 file Đủ (HĐDT, bảng kê, tờ khai, BCTC, NLĐ, TMS, sổ, phụ lục, v.v.)
- 1 file Mỏng: IN09 (Kho, giá thành) - chỉ 6 dòng
- 1 file Chưa nộp: IN12 (Phụ lục GDLK)
- 1 file Chưa trích: IN11 (Sao kê ngân hàng = 5 PDF scan, 28 trang, 0 text)

**File sinh ra:** KQ1.1 (CQT), KQ1.2 (NNT), KQ1.3 (Đánh giá)

---

## GĐ2 - PHÂN TÍCH DỮ LIỆU

### Phần A: Đối chiếu GTGT-TNDN-BCTC

**Đối chiếu GTGT (4 năm, 4 kỳ/năm = 16 kỳ tờ khai):**
- Đối chiếu bảng kê mua vào ↔ tờ khai GTGT [23]/[24] ↔ HĐDT
- 164/169 dòng hóa đơn có số liệu; 5 dòng chưa trích
- Cờ máy: Khớp 100% theo dòng HĐ (cùng số tiền, cùng ngày)

**Đối chiếu NCC mua vào (139 nhà cung cấp, tổng 57,71 tỷ đ):**
- Quét danh sách NCC từ bảng kê + HĐDT
- Đối chiếu với 3 nguồn: Sổ CVCBRR, Công văn Công an 2026, Danh bạ ĐKKD
- **Phát hiện: 1 NCC F0** (NNM Group, MST 0317409384, CV 6174) + 24 NCC trạng thái bất thường

**Đối chiếu 511-632 (287 cặp doanh thu-giá vốn):**
- Khớp 100% theo HĐ, ngày, mặt hàng
- Tính lãi gộp = 3,36% (58,89 tỷ / 56,91 tỷ)
- Phát hiện: 1 HĐ bán dưới giá (HĐ #6, 23/01/2024, lỗ 745.000 đ)

### Phần B: Phân tích EBITDA & khả năng trích lãi vay

**Dữ liệu lấy từ IN04_BCTC_TNDN (4 năm đầy đủ):**

| Năm | KQKD 50 (LN BDKD) | CDKT 141 (Khấu hao) | TNDN ct16 (Lãi vay) | EBITDA Gross | EBITDA 30% |
|---|---:|---:|---:|---:|---:|
| 2022 | -3.389.745 | 2.931.342 | 0 | -458.403 | -137.521 |
| 2023 | 132.907.178 | 20.489.064 | 0 | 153.396.242 | 46.018.873 |
| 2024 | 222.517.254 | 25.175.213 | 0 | 247.692.467 | 74.307.740 |
| 2025 | 155.551.382 | 21.789.832 | 0 | 177.341.214 | 53.202.364 |
| **Tổng** | **507.586.069** | **70.385.451** | **0** | **577.971.520** | **173.391.456** |

**Kết luận:** Công ty **không vay nợ** (chi phí lãi vay = 0 đ). Không có rủi ro K9 (lãi vay không hợp lệ).

**File sinh ra:** KQ2.1 (với sheet KQ2_EBITDA_Phan_tich), KQ2.2

---

## GĐ3 - KẾT LUẬN & ĐÁNH GIÁ

**80 quy tắc rủi ro:**
- 67 quy tắc **Đã chạy**: 2 Cao, 3 Trung bình, 75 Thấp
- 9 quy tắc **Chưa đủ DL**: Chờ IN11 (sao kê), IN12 (GDLK), K13 (số liệu ngành)
- 3 quy tắc **Không áp dụng**: Công ty thương mại thuần túy

**Phát hiện rủi ro:**
1. **Mức Cao (2 quy tắc, 2 chỉ tiêu):**
   - HD01 (Điểm 3): Mua hàng NCC F0 - NNM Group (2 HĐ, 2.442.000 đ)
   - TN22 (Điểm 3): Giá vốn tăng 98,53%, lãi gộp = 1,47%

2. **Mức Trung bình (3 quy tắc):**
   - DT01 (Điểm 2): HĐ bán dưới giá (745.000 đ)
   - TN03 (Điểm 2): Lãi gộp mỏng
   - HD02 (Điểm 2): 24 NCC trạng thái bất thường

3. **Mức Thấp: 75 quy tắc** (phân tích chi tiết ở KQ3)

**28 tiêu chí rủi ro:**
- Đầy đủ 28/28 (TC1-TC28)
- Cao 2 (TC26, TC28), Trung bình 5, Thấp/rõ 21

**Chuỗi bằng chứng:**
- 5 mắt xích: HĐ → HĐKT → TT → Kho → SD
- Danh sách chứng từ: Hóa đơn gốc, sổ chi tiết, phiếu chi, sao kê, sổ kho

**File sinh ra:** KQ3 (2 sheet: KQ3 + Danh_gia_28_tieu_chi)

---

## **Vì sao / Áp dụng**

**Vì sao cần:** Phân tích toàn diện từ CQT đến DN cho 4 năm, dùng quy trình 3 giai đoạn v1.4 để:
1. Xác định hồ sơ đủ hay thiếu → Quyết định chạy GĐ nào
2. Phát hiện lệch/rủi ro từ dữ liệu thô (GĐ2) → Chuyển thành kết luận (GĐ3)
3. Chuẩn bị bằng chứng để chất vấn tại trụ sở

**Áp dụng lần sau:**
- Dùng **6 tệp mẫu chuẩn v1.4** làm khung, chỉ điền dữ liệu theo IN00-IN14
- Theo **quy trình 3 giai đoạn**: GĐ1 → kiểm kê → xác định mức → GĐ2 (nếu đủ) → GĐ3 (nếu có sổ)
- **Chú ý IN04:** Dữ liệu 2022-2024 nằm trong dòng riêng (không phải cột F), cần quét từ dòng 3 đến hết
- **Chú ý EBITDA:** Lấy KQKD 50, CDKT 141, TNDN ct16 từ 4 năm → Tính EBITDA Gross → Áp dụng 30%
