# Giảm thuế TNDN, TNCN & GTGT (2020-2026)

**Group:** Tài chính  
**Description:** Xác định điều kiện áp dụng, tính mức giảm, kê khai 3 mẫu và xử lý vượt ngưỡng theo ND 114/2020, ND 92/2021, ND 15/2022, ND 398/2026

## Khi nào dùng

- "DN doanh thu bao nhiêu được giảm thuế?"
- "Cách tính số tiền giảm TNDN năm 2026?"
- "Nộp mẫu nào khi doanh thu vượt 10 tỷ?"
- "DN 2020 lỗ có được miễn tiền chậm nộp?"
- "Giảm GTGT 2022 áp dụng dịch vụ nào?"
- "Hộ kinh doanh Q3-Q4/2021 chịu tác động Covid được miễn gì?"

## Chuẩn bị

1. **Dữ liệu:** Doanh thu thực tế, loại hình DN, thời kỳ kinh doanh, số thuế phải nộp từ tờ khai
2. **Tài liệu:** 4 Nghị định (ND 114/2020, ND 92/2021, ND 15/2022, ND 398/2026) và 3 phụ lục mẫu kê khai
3. **Công cụ:** Excel/Sheets để tính % giảm, bảng danh mục GTGT (Phụ lục I-III ND 15/2022)

## Quy trình chi tiết

### **Bước 1: XÁC ĐỊNH KỲ ÁP DỤNG & NGƯỠNG DOANH THU**

| Kỳ | Nghị định | Đối tượng | Ngưỡng DT | Điều kiện thêm | Loại thuế |
|:---|:---|:---|:---:|:---|:---|
| 2020 | ND 114/2020 | DN tất cả loại | ≤ 200 tỷ | (không) | TNDN |
| 2021 | ND 92/2021 | DN tất cả loại | ≤ 200 tỷ | DT 2021 < DT 2019 | TNDN |
| 2021 Q3-Q4 | ND 92/2021 | Hộ kinh doanh/cá nhân | N/A | Địa bàn Covid (UBND tỉnh công bố) | TNCN+GTGT+TTĐB+TN+BVMT |
| 2020-2021 | ND 92/2021 | DN lỗ 2020 | N/A | Lỗ năm 2020 | Tiền chậm nộp |
| 2022 | ND 15/2022 | DN tất cả loại | N/A | (không) | GTGT (20% giảm) |
| 2026-2027 | ND 398/2026 | DN, HCX, ĐVSK, hộ KD | ≤ 10 tỷ | (không) | TNCN + TNDN |

---

### **Bước 2: KIỂM TRA ĐIỀU KIỆN ĐỦ**

#### **A. Kiểm doanh thu (ND 114/2020, 92/2021, 398/2026)**

```
Nếu kỳ tính thuế là năm dương lịch:
  DT_năm = Tổng toàn bộ bán hàng + cung cấp DV + trợ giá, phụ thu, phụ trội
  (KHÔNG bao gồm: giảm trừ DT, DT hoạt động tài chính, thu nhập khác)

Nếu DN mới thành lập hoạt động < 12 tháng:
  DT_chuẩn = DT_thực_tế ÷ số_tháng_hoạt_động × 12
  (nếu hoạt động trong tháng → tính đủ tháng)

Nếu DN chia/tách/sáp nhập:
  - ND 114/2020, 92/2021: Cộng thời gian kỳ cuối < 03 tháng vào năm 2021
  - ND 398/2026: DN hình thành từ chia/tách sau ngày NQ có hiệu lực:
    Nếu tổng DT các DN sau chia/tách > 10 tỷ → KHÔNG được giảm
```

**Kết luận:** 
- ✅ Đủ điều kiện → Sang Bước 3 (Tính mức giảm)
- ❌ Vượt ngưỡng hoặc không đủ điều kiện → Không được giảm, nộp thuế bình thường

---

#### **B. Kiểm loại hình (ND 398/2026 áp dụng cho)**

✅ Được giảm:
- Doanh nghiệp (công ty TNHH, công ty cổ phần, doanh nghiệp Nhà nước)
- Hợp tác xã, liên hiệp HCX
- Đơn vị sự nghiệp
- Tổ chức khác (NKT, tổ chức tôn giáo, hội, hiệp hội...) có hoạt động sản xuất kinh doanh
- Hộ kinh doanh, cá nhân kinh doanh (đặc biệt nếu địa bàn Covid)

---

#### **C. Kiểm địa bàn (ND 92/2021 Q3-Q4/2021 — GTGT)**

**Chỉ áp dụng nếu:**
- Hộ KD/cá nhân KD
- **Tại các huyện/thị xã** nằm trong **Danh sách của UBND tỉnh** công bố năm 2021
- Danh sách ghi rõ các địa bàn "dừng hoặc ngừng hoạt động sản xuất kinh doanh" do phong tỏa/cách ly Covid

📌 **Lưu ý:** Nếu không có danh sách từ UBND tỉnh → Xem trên cổng thông tin thuế địa phương hoặc liên hệ Chi cục Thuế

---

### **Bước 3: TÍNH SỐ TIỀN GIẢM**

#### **ND 114/2020 & ND 92/2021 (Giảm TNDN)**

```
Công thức:
  Số_thuế_được_giảm = Số_thuế_TNDN_phải_nộp_kỳ × 30%

Trong đó:
  Số_thuế_TNDN_phải_nộp = Thu_nhập_chịu_thuế × Thuế_suất (20%)
  
Lưu ý: 
  - Tính TRÊN TOÀN BỘ thu nhập (gồm cả cộng lỏ năm trước, lợi nhuận chuyển nhượng BĐS)
  - TRỪ ĐI số thuế đang được ưu đãi (nếu có)
  - KHÔNG tính trên lỗ (giảm thuế chỉ áp dụng cho năm có lợi nhuận)

Ví dụ 2020:
  DN A: Thu nhập = 50 tỷ đ, Thuế suất = 20%
  Thuế phải nộp = 50 tỷ × 20% = 10 tỷ đ
  Giảm = 10 tỷ × 30% = 3 tỷ đ
  → Nộp = 10 tỷ - 3 tỷ = 7 tỷ đ
```

**Cách tính tạm nộp (ND 114/2020 & 92/2021):**

```
Tạm_nộp_quý = Thuế_dự_kiến_phải_nộp_quý × 70%
(Doanh nghiệp tự xác định dựa trên DT dự kiến)

Quyết toán cuối năm:
  - Nếu Thuế_thực_tế < Thuế_tạm_nộp → Được hoàn trả hoặc bù trừ
  - Nếu Thuế_thực_tế > Thuế_tạm_nộp → Nộp bổ sung (có tính tiền chậm nộp)
```

---

#### **ND 92/2021 (Miễn toàn bộ HK, Q3-Q4/2021)**

```
Công thức:
  Số_thuế_được_miễn = Tổng_thuế_phát_sinh_Q3_Q4_2021

Bao gồm miễn 5 loại thuế:
  1. Thuế TNCN từ hoạt động sản xuất kinh doanh
  2. Thuế GTGT
  3. Thuế TTĐB
  4. Thuế tài nguyên
  5. Thuế BVMT
  
Ngoại lệ KHÔNG miễn:
  - Phần mềm, dịch vụ nội dung số (giải trí, trò chơi, phim, ảnh, nhạc số, quảng cáo số)

Cách xác định số thuế:
  - Nếu cơ quan thuế ra Thông báo nộp tiền → Lấy số trên Thông báo
  - Nếu KHÔNG có Thông báo → Tự xác định theo Tờ khai thuế của hộ KD
  - Nếu khai theo kỳ → Xác định tương ứng DT thực tế phát sinh Q3-Q4
  - Nếu không xác định được DT chính xác → Dùng DT bình quân tháng theo giá trị HĐ
```

---

#### **ND 15/2022 (Giảm GTGT 20%)**

```
Công thức theo phương pháp:

A. Phương pháp khấu trừ (cơ sở bình thường):
   Thuế_GTGT_đầu_ra = Giá_hàng_bán_chưa_có_thuế × 8%  (từ 10%)
   
B. Phương pháp tỷ lệ % (hộ KD, DN nhỏ):
   Tỷ_lệ_tính = Tỷ_lệ_gốc × 80%
   Ví dụ: Nếu tỷ lệ gốc 10% → Áp dụng 8%

Danh mục giảm: Phụ lục I (180+ mục hàng hóa/dịch vụ)
Ngoại lệ KHÔNG giảm:
  - Viễn thông, tài chính, ngân hàng, chứng khoán, bảo hiểm
  - Bất động sản, kim loại sắt, than, dầu tinh chế, hóa chất
  - Hàng hóa TTĐB, công nghệ thông tin
```

---

#### **ND 398/2026 (Giảm TNCN + TNDN 30%)**

```
Công thức:
  Số_thuế_được_giảm_2026_2027 = Số_thuế_phải_nộp_kỳ × 30%

Áp dụng:
  - TNCN: Cá nhân kinh doanh
  - TNDN: Doanh nghiệp, HCX, ĐVSK, tổ chức khác

Xử lý khi vượt 10 tỷ đ:
  - Tính từ kỳ phát sinh DT lũy kế > 10 tỷ
  - Phải nộp bổ sung số thuế đã giảm từ đầu năm
  - KHÔNG tính tiền chậm nộp đến hạn nộp của kỳ phát sinh DT > 10 tỷ
  
Ví dụ:
  HK A: 2026 Q1-Q2 DT = 8 tỷ (được giảm)
        Q3 DT = 3 tỷ → Lũy kế = 11 tỷ (vượt)
  → Từ Q3 không được giảm nữa
  → Cuối năm nộp bổ sung thuế đã giảm Q1-Q2 (không có tiền chậm nộp)
```

---

### **Bước 4: KÊ KHAI & NỘP**

#### **ND 114/2020 — 1 mẫu**

```
Tờ khai: Mẫu tờ khai TNDN thông thường (không có mẫu riêng)
+ Phụ lục: Danh sách chi tiết số thuế được giảm theo từng lần nộp hoặc quý
Nộp cùng với: Hồ sơ khai thuế định kỳ (quý hoặc năm)
Thời điểm: Tạm nộp hàng quý (70% × thuế phát sinh)
```

---

#### **ND 92/2021 — 5 mẫu**

**Đối với DN giảm TNDN:**
```
Mẫu 1: "Phụ lục Thuế TNDN được giảm theo NQ 406/NQ-UBTVQH15"
  - Thể hiện: Tổng TNDN phải nộp vs Số TNDN được giảm (quý, 6 tháng hoặc năm)
  - Nộp cùng: Tờ khai TNDN định kỳ
```

**Đối với HK/cá nhân KD miễn thuế Q3-Q4:**
```
Mẫu 2: "Danh sách HK, cá nhân KD được miễn thuế" (Mẫu 01-1/DS-MTHK)
  - Lập bởi: Cơ quan thuế quản lý HK tại địa bàn
  
Mẫu 3: "Quyết định miễn thuế" (Mẫu 01/MTHK)
  - Ký bởi: Chi cục trưởng → Gửi từng HK

Mẫu 4: "Thông báo miễn thuế" (Mẫu 01/TBSMT-CNKD)
  - Thể hiện: Số tiền miễn từng loại thuế
  
Mẫu 5: "Bản xác định số thuế được miễn" (Mẫu 01-1/PL-CNKD)
  - Dùng khi: HK KHÔNG có Thông báo nộp tiền từ cơ quan thuế
  - Tự xác định: Số thuế phát sinh Q3-Q4 → Miễn toàn bộ
  - Nộp cùng: Tờ khai thuế định kỳ
```

**Đối với DN lỗ 2020 xin miễn tiền chậm nộp:**
```
Mẫu 6: "Công văn đề nghị miễn tiền chậm nộp" (Mẫu 01/MTCN)
  - Gửi: Chi cục Thuế quản lý
  - Nêu rõ: Số lỗ năm 2020 (từ BCTC đã kiểm toán hoặc Tờ khai)
  - Kèm: Biên bản kiểm toán hoặc quyết định/kết luận thanh tra (nếu có)

Mẫu 7: "Thông báo không chấp thuận miễn" (Mẫu 02/MTCN) — nếu bị từ chối

Mẫu 8: "Quyết định miễn tiền chậm nộp" (Mẫu 03/MTCN) — nếu được duyệt
  - Ghi rõ: Số tiền chậm nộp từng loại (thuế, tiền sử dụng/thuê đất)
  - Hiệu lực: Từ ngày ký
```

**Giảm GTGT 2021 Q4:**
```
Mẫu 9: "Phụ lục giảm GTGT" (không chỉ định mẫu, nhưng phải kê khai)
  - Hàng hóa/dịch vụ: Vận tải, lưu trú, ăn uống, xuất bản, du lịch (Danh mục 200+ mục)
  - Cách ghi: Trên hóa đơn = Mức thuế suất theo QĐ × 70%
  - Nộp cùng: Tờ khai GTGT của kỳ
```

---

#### **ND 15/2022 — 1 mẫu**

```
Mẫu 1: "Phụ lục giảm GTGT theo NQ 43/2022/QH15" (Mẫu 01, Phụ lục IV)
  - Danh sách: Hàng hóa/dịch vụ được giảm (Phụ lục I-III ND 15/2022)
  - Ghi đầy đủ: 
    * Tên hàng hóa/dịch vụ
    * Giá chưa có thuế
    * Thuế suất gốc & thuế suất giảm
    * Số tiền giảm
  - Nộp cùng: Tờ khai GTGT tháng/quý
  
Lưu ý hoá đơn:
  - Phương pháp khấu trừ: Ghi "8%" + tiền thuế và tổng tiền
  - Phương pháp tỷ lệ %: Ghi "Đã giảm... đồng (20% mức tỷ lệ) theo NQ 43/2022"
```

---

#### **ND 398/2026 — 3 mẫu**

**Mẫu 1: Cá nhân kinh doanh tự kê khai (Mẫu 01)**
```
Tờ khai: Mẫu 01 "Giảm TNCN theo NQ 43/2026/QH16"
  - Cột: [Địa điểm KD] [Hoạt động] [Nhóm ngành] [Thuế TNCN phải nộp] [Thuế TNCN được giảm]
  - Nộp cùng: Tờ khai TNCN của kỳ (Q/6M/Năm)
  - Lần đầu: Cùng với kài lần đầu năm 2026
  - Bổ sung: Nếu cần điều chỉnh
```

**Mẫu 2: Tổ chức khấu trừ/khai thay (Mẫu 02)**
```
Tờ khai: Mẫu 02 "Giảm TNCN cho cá nhân KD"
  - Cột: [Tên HK/cá nhân] [MST] [Nhóm ngành] [Thuế TNCN phải nộp] [Được giảm]
  - Nộp cùng: Tờ khai TNCN của tổ chức (tháng/quý/lần phát sinh)
```

**Mẫu 3: Điều chỉnh khi vượt 10 tỷ (Mẫu 03)**
```
Tờ khai: Mẫu 03 "Điều chỉnh thuế phải nộp & thuế được giảm"
  - Cột: [Hoạt động] [Địa điểm] [Nhóm ngành] [Thuế đã kê khai giảm] [Phải nộp do vượt]
  - Khi: DT lũy kế vượt 10 tỷ (nhập khoảng nào vượt)
  - Nộp: Cùng với tờ khai của kỳ phát sinh DT > 10 tỷ
  - Ghi: "Kỳ nộp từ Q3 năm 2026 (lũy kế 11 tỷ)" hoặc tương tự
```

---

### **Bước 5: XỬ LÝ CÁC TRƯỜNG HỢP ĐẶC BIỆT**

#### **A. Khi phát hiện không đủ điều kiện (ND 114/2020, 92/2021, 398/2026)**

```
Qua thanh tra, kiểm tra, kiểm toán:
  ✓ DN không thuộc đối tượng được giảm HOẶC
  ✓ Doanh thu > ngưỡng quy định
  
Xử lý:
  1. DN phải nộp bổ sung: Số thuế phải nộp - Số thuế đã nộp
  2. Phạt vi phạm: Tính trên số thuế thiếu
  3. Tiền chậm nộp: Từ ngày hết hạn nộp thuế đến ngày nộp bổ sung
  
Trừ ND 398/2026:
  - Nộp bổ sung KHÔNG tính tiền chậm nộp
```

---

#### **B. Khi doanh thu thực tế vượt ngưỡng giữa năm (ND 398/2026)**

```
Tình huống: HK A doanh thu Q1+Q2 = 8 tỷ (được giảm)
           Q3 = 3 tỷ → Lũy kế = 11 tỷ (vượt 10 tỷ)

Xử lý:
  1. Từ Q3 trở đi: KHÔNG được giảm nữa
  2. Q3 kê khai: 100% thuế (không giảm)
  3. Cuối năm (khi quyết toán):
     - Nộp bổ sung: Thuế được giảm Q1+Q2 (do vượt ngưỡng)
     - KHÔNG tính tiền chậm nộp đến hạn nộp tờ khai Q3
     - Ghi Mẫu 03 "Điều chỉnh"

Lưu ý: Không bị xử phạt thêm nếu kê khai đúng & nộp bổ sung đúng hạn
```

---

#### **C. Khi khai bổ sung hoặc thanh tra phát hiện tăng thuế (ND 398/2026)**

```
Trường hợp: DN khai bổ sung hoặc thanh tra làm tăng TNDN
            nhưng vẫn đáp ứng điều kiện ≤ 10 tỷ DT

Xử lý:
  - Số thuế tăng thêm: Cũng được giảm 30%
  - Kê khai: Mẫu 03 (Điều chỉnh)
  - Nộp bổ sung: Số thuế tăng × 70% (vì đã giảm)

Ví dụ:
  DN khai bổ sung, thuế tăng thêm 100 triệu
  Nộp bổ sung = 100 triệu × 70% = 70 triệu
```

---

#### **D. Khi nộp thừa hoặc cơ quan thuế phát hiện giảm (ND 398/2026)**

```
Trường hợp: Khai bổ sung hoặc thanh tra làm GIẢM TNDN

Xử lý:
  - Điều chỉnh giảm: Số thuế giảm được giảm 30% (nếu đủ điều kiện)
  - Xử lý tiền nộp thừa: 
    * Bù trừ với nợ khác (quý/năm tiếp theo)
    * Hoàn trả
    * Bù trừ + hoàn trả (nếu nợ không đủ)
```

---

### **Bước 6: DANH MỤC & THAM KHẢO**

#### **ND 114/2020 (Phụ lục)**
- 1 mẫu: Phụ lục TNDN được giảm (ghi tổng chỉ tiêu)

#### **ND 92/2021 (Phụ lục II)**
- Mẫu 01-1/DS-MTHK: Danh sách HK miễn thuế (cơ quan thuế lập)
- Mẫu 01/MTHK: Quyết định miễn thuế
- Mẫu 01/TBSMT-CNKD: Thông báo miễn thuế (gửi HK)
- Mẫu 01-1/PL-CNKD: Bản xác định miễn (HK tự lập khi không có TB nộp)
- Mẫu 01/MTCN: Công văn đề nghị miễn tiền chậm nộp
- Mẫu 02/MTCN: Thông báo từ chối miễn tiền chậm nộp
- Mẫu 03/MTCN: Quyết định miễn tiền chậm nộp
- **Phụ lục I:** Danh mục 200+ hàng hóa/dịch vụ giảm GTGT 2021 (vận tải, lưu trú, ăn uống, xuất bản, du lịch...)

#### **ND 15/2022 (Phụ lục I, III, IV)**
- **Phụ lục I:** Danh mục ngoại lệ KHÔNG giảm GTGT 2022 (viễn thông, tài chính, BĐS, hóa chất...)
- **Phụ lục II:** Danh mục TTĐB KHÔNG giảm
- **Phụ lục III:** Danh mục công nghệ thông tin KHÔNG giảm
- **Phụ lục IV:** Mẫu 01 "Phụ lục giảm GTGT" (kê khai hàng hóa/dịch vụ được giảm)

#### **ND 398/2026 (Phụ lục)**
- Mẫu 01: Cá nhân KD tự kê khai
- Mẫu 02: Tổ chức khấu trừ/khai thay
- Mẫu 03: Điều chỉnh khi vượt 10 tỷ

---

## Bẫy thường gặp

### ❌ Lỗi 1: Tính giảm trên lỗ
```
SAI:  DN lỗ 100 triệu → Giảm thuế = 100 triệu × 30% = 30 triệu
ĐÚNG: DN lỗ → KHÔNG được giảm (giảm chỉ áp dụng khi lợi nhuận)
      Nếu sau khi bù lỗ kỳ trước còn lợi nhuận → Tính giảm trên phần lợi nhuận đó
```

### ❌ Lỗi 2: Quên điều kiện "2021 giảm so với 2019"
```
SAI:  DN 2021 DT 150 tỷ ≤ 200 tỷ → Được giảm
ĐÚNG: DN 2021 DT 150 tỷ NHƯNG 2019 DT 200 tỷ (2021 < 2019 không thỏa)
      → KHÔNG được giảm (ND 92/2021 có điều kiện so sánh)
      (Trừ DN mới thành lập, chia, tách, sáp nhập 2020-2021)
```

### ❌ Lỗi 3: Nhầm danh mục GTGT 2022
```
SAI:  Công ty viễn thông bán SIM card (GTGT 10%) → Giảm còn 8%
ĐÚNG: Viễn thông nằm ngoại lệ (Phụ lục I) → KHÔNG giảm (vẫn 10%)
```

### ❌ Lỗi 4: Kê khai sai mẫu khi vượt 10 tỷ
```
SAI:  HK doanh thu vượt 10 tỷ năm 2026 → Cứ nộp Mẫu 01 như bình thường
ĐÚNG: Từ kỳ vượt → Nộp Mẫu 03 "Điều chỉnh" (bóc tách phần vượt)
      Ghi rõ: Tháng/quý nào vượt, số thuế nộp bổ sung
```

### ❌ Lỗi 5: Quên không tính tiền chậm nộp (ND 398/2026)
```
SAI:  HK vượt 10 tỷ năm 2026 → Nộp bổ sung + tiền chậm nộp
ĐÚNG: Nộp bổ sung (không có tiền chậm nộp) theo Mẫu 03
      Chỉ có tiền chậm nộp nếu nộp muộn quá hạn
```

---

## Ví dụ minh họa

### Ví dụ 1: DN năm 2020 (ND 114/2020)

```
DN B: 2020 DT = 150 tỷ đ (< 200 tỷ ✓)
      Thu nhập = 30 tỷ đ
      Thuế suất TNDN = 20%
      
Tính:
  Thuế phải nộp = 30 tỷ × 20% = 6 tỷ đ
  Giảm 30% = 6 tỷ × 30% = 1,8 tỷ đ
  Nộp thực tế = 6 tỷ - 1,8 tỷ = 4,2 tỷ đ
  
Tạm nộp quý:
  - Tạm nộp = 70% × thuế quý dự kiến (tự xác định)
  - Giả sử Q1-Q3 nộp 70% từng quý
  - Cuối năm quyết toán: Nếu thuế thực < tạm nộp → Hoàn trả
```

### Ví dụ 2: HK năm 2021 (ND 92/2021)

```
HK C: Q3-Q4/2021, địa bàn Bắc Ninh (có tác động Covid)
      GTGT phát sinh T7-T12/2021 = 500 triệu đ
      TNCN phát sinh = 300 triệu đ
      
Được miễn toàn bộ:
  - TNCN 300 triệu: Miễn ✓
  - GTGT 500 triệu: Miễn ✓ (trừ phần mềm/dịch vụ số)
  
Xử lý:
  - Cơ quan thuế ra Quyết định miễn (Mẫu 01/MTHK)
  - Gửi Thông báo (Mẫu 01/TBSMT-CNKD): Miễn 800 triệu đ
  - HK không cần nộp nộp thêm T7-T12
```

### Ví dụ 3: DN năm 2026 vượt 10 tỷ (ND 398/2026)

```
DN D: 2026 (DT ≤ 10 tỷ)
      - Q1: DT = 2 tỷ (được giảm 30%)
      - Q2: DT = 3 tỷ (được giảm 30%)
      - Q3: DT = 4 tỷ (lũy kế = 9 tỷ, được giảm 30%)
      - Q4: DT = 2 tỷ (lũy kế = 11 tỷ, VƯỢT 10 tỷ)
      
Cách xử lý:
  - Q1-Q3: Kê khai Mẫu 01 (giảm 30% mỗi quý)
  - Q4: Kê khai Mẫu 01 (KHÔNG giảm, vì đã vượt)
  - Cuối năm quyết toán:
    * Tính tổng thuế phải nộp năm = Tổng thu nhập × 20%
    * Nộp bổ sung: (Thuế đã giảm Q1-Q3) - (Nộp thực tế Q1-Q3)
    * Kê khai Mẫu 03 "Điều chỉnh"
    * KHÔNG tính tiền chậm nộp
```

---

## Công cụ hỗ trợ

### File Excel mẫu

**Cột tính toán:**
```
| Kỳ | DT_quý | DT_lũy_kế | Thuế_phải_nộp | Mức_giảm | Thuế_giảm | Nộp_thực_tế |
```

### Checklist trước khi nộp

- [ ] Xác nhận DT thực tế (so với HĐ, sổ cái)
- [ ] Kiểm DN/HK có đủ điều kiện (DT, loại hình, thời kỳ)
- [ ] Tính chính xác số tiền giảm
- [ ] Chọn mẫu kê khai đúng (01, 02, 03...)
- [ ] Ghi chữ số, ghi chữ, ký xác nhận rõ ràng
- [ ] Nộp đúng hạn (cùng với tờ khai thuế)
- [ ] Giữ bản sao cho hồ sơ

---

## Khi nào liên hệ cơ quan thuế

- ❓ Không chắc địa bàn có tác động Covid (ND 92/2021)
- ❓ Cần làm rõ "giảm so với năm 2019" (ND 92/2021)
- ❓ Không biết danh mục GTGT giảm 2022 (ND 15/2022)
- ❓ DN chia/tách, cần xác định tổng DT (ND 398/2026)
- ❓ Kê khai sai, cần sửa bổ sung (tất cả ND)

---

**Cập nhật:** Hiệu lực áp dụng từ 29/10/2026 (ND 398/2026)  
**Quản lý:** Cục Thuế Nhập Cảnh, Chi cục Thuế Địa phương
