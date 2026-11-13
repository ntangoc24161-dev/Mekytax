# Đối chiếu tỷ suất sinh lợi ngành (Kiểm tra TNDN bất thường)

**Group:** Tài chính  
**Description:** So sánh tỷ suất lợi nhuận thực tế DN với chuẩn ngành theo Đồng Nai (2017-2023); phát hiện TNDN bất thường & ghi nhận rủi ro

## Khi nào dùng

- "DN này TNDN bao nhiêu, có bất thường so với ngành không?"
- "Tỷ suất sinh lợi chuẩn của ngành xây dựng, may mặc, linh kiện điện tử là bao nhiêu?"
- "Kiểm tra xem DN khai thiếu TNDN, có dấu hiệu trốn thuế không?"
- "DN hoạt động 3 năm chưa tới, so sánh với chuẩn ngành năm nào?"
- "Ngành bán lẻ, bán buôn chuẩn tỷ suất bao nhiêu?"

## Chuẩn bị

1. **Thông tin DN:** Mã ngành, năm kiểm tra, TNDN khai, doanh thu thực
2. **Bảng tỷ suất chuẩn:** Danh sách 36 ngành Đồng Nai (2017-2023, từ Cục Thuế)
3. **Công cụ:** Excel/Sheets, máy tính

---

## Quy trình

### **Bước 1: XÁC ĐỊNH NGÀNH & NĂM CỤ THỂ**

**Ngành của DN:**
- Xác định từ **Giấy phép kinh doanh** hoặc **Hóa đơn**
- Xác định **mã ngành VSIC** (Danh mục ngành kinh tế)
- So sánh với **Danh sách 36 nhóm ngành** trong bảng chuẩn

**Năm kiểm tra:**
- Sử dụng tỷ suất chuẩn của **năm được kiểm tra**
- Nếu DN mới thành lập (< 1 năm) → Có thể không áp dụng chuẩn ngành

**Bảng 36 nhóm ngành chuẩn (Đồng Nai):**

| STT | Nhóm ngành | 2017 | 2018 | 2019 | 2020 | 2021 | 2022 | 2023 |
|:---|:---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| 1 | Nông lâm nghiệp | | | | 0.0125 | 0.0173 | 0.0212 | 1.54% |
| 2 | Khai khoáng | | | | 0.0176 | 0.0168 | 0.0176 | 2.59% |
| 3 | SX thực phẩm | | | | 0.0288 | 0.0291 | 0.0394 | 1.72% |
| 4 | SX thức ăn gia súc | | | | 0.0285 | 0.013 | 0.0285 | 3.17% |
| 5 | Dệt | | | | 0.0216 | 0.0334 | 0.0345 | 1.16% |
| 6 | May mặc | | | | 0.0249 | 0.0231 | 0.0134 | 1.24% |
| 7 | SX giày dép, vali, túi | | | | 0.016 | 0.0161 | 0.0165 | 1.89% |
| 8 | SX gỗ | | | | 0.025 | 0.0203 | 0.0208 | 0.84% |
| 9 | SX giấy | | | | 0.0208 | 0.0222 | 0.0227 | 2.77% |
| 10 | In ấn | | | | 0.021 | 0.0278 | 0.0304 | 0.99% |
| 11 | SX hóa chất | | | | 0.0455 | 0.03 | 0.0455 | 3.38% |
| 12 | SX cao su, plastic | | | | 0.0281 | 0.0156 | 0.0281 | 2.65% |
| 13 | SX bê tông, xi măng, gốm sứ | | | | 0.0221 | 0.0162 | 0.0221 | 1.12% |
| 14 | SX kim loại | | | | 0.0149 | 0.027 | 0.0278 | 0.99% |
| 15 | SX SP bằng kim loại | | | | 0.0125 | 0.0116 | 0.0125 | 1.20% |
| 16 | SX linh kiện điện tử | | | | 0.0367 | 0.0305 | 0.0408 | **4.47%** |
| 17 | SX thiết bị điện | | | | 0.0275 | 0.0216 | 0.0302 | 2.24% |
| 18 | SX máy móc, phụ tùng | | | | 0.0296 | 0.0241 | 0.0296 | 2.43% |
| 19 | SX giường, tủ, bàn, ghế | | | | 0.0271 | 0.0235 | 0.0122 | 1.34% |
| 20 | Công nghiệp chế biến khác | | | | 0.0189 | 0.0149 | 0.0265 | 2.98% |
| 21 | Nước, điện, rác thải | | | | 0.017 | 0.0193 | 0.0196 | 1.57% |
| 22 | Xây dựng, lắp đặt | | | | 0.0116 | 0.01 | 0.0116 | 1.26% |
| 23 | Bán xe | | | | 0.0056 | 0.0051 | 0.0064 | 0.59% |
| 24 | Bán buôn khác | | | | 0.0073 | 0.0069 | 0.0074 | 0.94% |
| 25 | Bán buôn máy móc, thiết bị | | | | 0.0109 | 0.0086 | 0.011 | 1.07% |
| 26 | Bán buôn đồ dùng gia đình | | | | 0.0089 | 0.0067 | 0.0089 | 0.78% |
| 27 | **Bán buôn thực phẩm** | | | | 0.0035 | 0.0032 | 0.0038 | **0.23%** |
| 28 | Bán buôn kim loại | | | | 0.0055 | 0.005 | 0.0055 | 0.51% |
| 29 | Bán buôn nhiên liệu | | | | 0.0059 | 0.0087 | 0.0117 | 0.71% |
| 30 | Bán buôn nông lâm sản | | | | 0.0038 | 0.0028 | 0.0038 | **0.23%** |
| 31 | Bán buôn vật liệu xây dựng | | | | 0.007 | 0.0053 | 0.007 | 0.71% |
| 32 | Bán lẻ | | | | 0.0086 | 0.0094 | 0.0125 | 0.66% |
| 33 | Vận tải | | | | 0.0323 | 0.0255 | 0.0323 | 1.60% |
| 34 | Kho bãi | | | | 0.0325 | 0.0303 | 0.0325 | 1.18% |
| 35 | Dịch vụ | | | | 0.031 | 0.024 | 0.031 | 1.14% |
| 36 | Kinh doanh bất động sản | | | | 0.0494 | 0.0344 | 0.0494 | 3.54% |

**Dữ liệu 2017-2019 trong bảng chưa đầy đủ → Sử dụng năm gần nhất có dữ liệu**

---

### **Bước 2: TÍNH TỶ SUẤT SINH LỢI THỰC TẾ CỦA DN**

**Công thức:**

```
Tỷ suất sinh lợi thực tế = TNDN thực tế ÷ Doanh thu thực tế
(hoặc Lợi nhuận ÷ Doanh thu)

Tỷ suất sinh lợi % = (TNDN ÷ DT) × 100%
```

**Lấy số liệu từ:**
- Tờ khai TNDN (Mẫu 01/TNDN)
- Báo cáo tài chính hàng năm
- Sổ sách kế toán (nếu chưa nộp quyết toán)

**Ví dụ:**
```
DN A năm 2023:
- Doanh thu: 50 tỷ đ
- TNDN: 1 tỷ đ (lợi nhuận sau thuế)
- Tỷ suất sinh lợi = 1 ÷ 50 = 0.02 = 2%

DN B năm 2023:
- Doanh thu: 100 tỷ đ
- Lợi nhuận: 2 tỷ đ
- Tỷ suất sinh lợi = 2 ÷ 100 = 0.02 = 2%
```

**📌 Lưu ý:**
- **TNDN thực tế** = Lợi nhuận chịu thuế (sau khoán lỗ, khấu hao, chi phí hợp lý)
- Không dùng lợi nhuận ròng trước thuế (đã trừ 10% surtax)
- Nếu DN khai lỗ → Tỷ suất = 0% (không so sánh được)

---

### **Bước 3: SO SÁNH TỶLỆ THỰC TẾ VỚI CHUẨN NGÀNH**

**Công thức đánh giá:**

```
Sai lệch % = (Tỷ suất thực - Tỷ suất chuẩn) ÷ Tỷ suất chuẩn × 100%

Hoặc đơn giản: So sánh tỷ lệ trực tiếp
```

**3 trường hợp:**

#### **A. Tỷ suất thực tế ≈ Tỷ suất chuẩn (sai lệch < ±20%)**

```
✅ Bình thường, hợp lý với ngành

Ví dụ:
- DN ngành SX linh kiện điện tử, chuẩn 4.47%
- Tỷ suất thực: 3.8% ~ 5.2% → Bình thường
- Sai lệch: -15% đến +16% → Chấp nhận
```

#### **B. Tỷ suất thực tế LỰA CHỌN CAO hơn chuẩn (sai lệch > +20%)**

```
⚠️ CÓ THỂ bất thường (Dấu hiệu tích cực):
  - DN có hiệu quả kinh doanh cao
  - Sản phẩm, dịch vụ khác biệt, giá cao
  - Tiết kiệm chi phí tốt
  - Không nhất thiết là bất thường → Điều tra sâu hơn

Hành động:
  1. Kiểm tra xem DN có khai thiếu chi phí không
  2. Kiểm tra cơ cấu doanh thu (có thể high-margin products)
  3. Kiểm tra chi phí quản lý, bán hàng
  4. Có thể chấp nhận nếu có giải thích hợp lý
```

#### **C. Tỷ suất thực tế THẤP hơn chuẩn (sai lệch < -20%)**

```
⚠️ NGUY CÓ BẤT THƯỜNG CAO:
  - DN khai thiếu doanh thu (bỏ sót hóa đơn)
  - DN khai thừa chi phí (chi phí giả, không hợp lý)
  - DN hoạt động kém hiệu quả
  - Dấu hiệu TRỐN THUẾ (giảm thiểu TNDN)

Hành động:
  1. Kiểm tra doanh thu:
     - Đối chiếu với số lượng, giá trị HĐ thực tế
     - Kiểm tra sổ chi tiết, quỹ bán hàng
     - Đối chiếu với tờ khai GTGT (doanh thu GTGT)
  2. Kiểm tra chi phí:
     - Duyệt chi phí COGS, quản lý
     - Kiểm tra chi phí lương, năng lượng, vận tải
     - Kiểm tra khấu hao (đúng tiêu chuẩn không)
  3. Đối chiếu với nguyên vật liệu:
     - Nếu DN sản xuất: Kiểm tra hao mòn, định mức
  4. Ghi nhận rủi ro ⚠️ cần điều tra sâu
```

---

### **Bước 4: PHÂN TÍCH NGUYÊN NHÂN SAIỆCH**

**Nếu thấp hơn chuẩn (< -20%):**

#### **A. Kiểm Tra Doanh Thu**

```
1. Đối chiếu hóa đơn và sổ sách
   ✓ Tổng HĐ phát hành = Doanh thu khai?
   ✓ HĐ hủy, chiết khấu, trả hàng đã trừ?
   ✓ Doanh thu GTGT = Doanh thu TNDN?

2. Đối chiếu tờ khai GTGT + TNDN
   ✓ Doanh thu đầu vào (GTGT Giá bán) = TNDN?

3. Kiểm tra sổ chi tiết:
   ✓ Sổ nhật ký, sổ chi tiết khách hàng
   ✓ Quỹ bán hàng (giá trị bán, số lượng)

4. Phỏng vấn DN:
   ✓ Có hoạt động bán hàng khác không khai?
   ✓ Có khách hàng ngoài sổ sách?
   ✓ Có bán qua người trung gian, đại lý ngoài sổ?

NẾU PHÁT HIỆN:
  → Bổ sung doanh thu, tính bổ sung TNDN + phạt + lãi
```

#### **B. Kiểm Tra Chi Phí**

```
1. COGS (Giá vốn hàng bán):
   ✓ Định mức hao mòn từ từng lô hàng?
   ✓ Tính FIFO, LIFO hay bình quân?
   ✓ Có hàng tồn kho bất thường?
   ✓ Có hàng nhập hóa đơn giả?

2. Lương, thưởng:
   ✓ Lương thực tế vs khai:
     - Kiểm tra hợp đồng lao động
     - Kiểm tra sổ thanh toán lương
     - So sánh với mức bình quân ngành
   ✓ Có chi trả ngoài sổ (tiền mặt) không?

3. Năng lượng (điện, xăng, nước):
   ✓ Hóa đơn điện, nước, năng lượng
   ✓ Công suất sản xuất vs lượng năng lượng sử dụng

4. Vận tải:
   ✓ Giá vận chuyển thực tế so với HĐ
   ✓ Phù hợp với khối lượng, khoảng cách?

NẾUPHÁT HIỆN CHI PHÍ KHÔNG HỢP LÝ:
  → Loại bỏ chi phí giả, tính bổ sung TNDN + phạt
```

#### **C. Kiểm Tra Khấu Hao**

```
✓ Thời gian trích khấu hao có phù hợp với khung?
✓ Phương pháp khấu hao đúng (đường thẳng, số dư giảm dần)?
✓ Có vận dụng khấu hao nhanh hợp lệ?
✓ Có nâng cấp TSCĐ mà chưa tính lại khấu hao?

NẾUU SAI:
  → Xác định lại khấu hao đúng quy định
```

#### **D. Đối Chiếu Với Dữ Liệu Ngoài**

```
✓ Tờ khai GTGT: Doanh thu GTGT (hàng chịu thuế + chưa chịu + miễn)
✓ Sổ CVCBRR: Kiểm tra tính chính xác NCC, mua vào
✓ Bảng kê thương nhân: Doanh thu bán ra thực tế
✓ Sóhệ ngân hàng: Chuyển khoản vs doanh thu khai
✓ Kiểm tra thanh tra trước đó: Có sai sót nào từ năm trước?
```

---

### **Bước 5: GIAO VỤ & GHI NHẬN RỦI RO**

**Nếu phát hiện sai lệch bất thường (< -20% so với chuẩn):**

#### **A. Ghi Nhận Rủi Ro**

```
Mức độ rủi ro:
  🔴 CAO: Sai lệch < -50%
  🟠 TRUNG BÌNH: Sai lệch -20% đến -50%
  🟡 THẤP: Sai lệch -10% đến -20%

Loại rủi ro:
  1. Doanh thu bị khai thiếu (bỏ sót HĐ, bán ngoài sổ)
  2. Chi phí khai thừa (chi phí giả, không hợp lý)
  3. Khấu hao sai tiêu chuẩn
  4. Hao mòn vật tư không phù hợp với sản lượng
  5. Tính chất bất thường (dấu hiệu trốn thuế)
```

#### **B. Tính Toán Bổ Sung**

```
Công thức:
  TNDN bổ sung = (DT thực × Tỷ suất chuẩn) - TNDN khai
  Phạt: TNDN bổ sung × 10% (ngoài thuế TNDN)
  Lãi: TNDN bổ sung × Lãi suất (theo định chỉ số)

Ví dụ:
  DN khai: DT 100 tỷ, TNDN 1 tỷ (1%)
  Chuẩn ngành: 3%
  TNDN bổ sung = 100 × 3% - 1 = 2 tỷ
  Thuế bổ sung = 2 × 20% = 400 triệu
  Phạt = 400 triệu × 10% = 40 triệu
  Lãi = 400 triệu × (Lãi suất theo năm)
```

---

## Các trường hợp không áp dụng chuẩn ngành

### ❌ DN mới thành lập (< 1 năm hoạt động)
```
Khó áp dụng chuẩn → Tạm không so sánh
Theo dõi năm tiếp theo
```

### ❌ DN có thay đổi cấu trúc (chia, tách, sáp nhập, cổ phần hóa)
```
Doanh thu / TNDN có thể bị gián đoạn
Không so sánh trực tiếp với chuẩn
Cần xem xét riêng từng giai đoạn
```

### ❌ DN bị ảnh hưởng sự kiện đặc biệt (thiên tai, dịch bệnh, khủng hoảng)
```
Năm 2020-2021 (COVID-19): Chuẩn có biến động
Năm 2023 hậu COVID: Chuẩn trở lại bình thường
Cần so sánh với năm tương tự
```

### ❌ DN hoạt động với sản phẩm/thị trường đặc biệt
```
- Sản phẩm cao cấp, cao cổ, tỷ suất cao hơn chuẩn
- Sản phẩm bán quyền, bán lẻ, tỷ suất thấp hơn chuẩn
- Cần xem xét cơ cấu doanh thu
```

### ❌ DN có hoạt động xuất nhập khẩu (tính theo ngoại tệ)
```
Chuẩn là giá VND → Ngoại tệ biến động
Cần điều chỉnh theo tỷ giá năm khai
```

---

## Ví dụ minh họa

### Ví dụ 1: SX Linh kiện điện tử - Tỷ suất cao hơn chuẩn

```
DN A: Sản xuất linh kiện điện tử
- Năm kiểm tra: 2023
- Chuẩn ngành: 4.47%
- Doanh thu: 500 triệu đ
- TNDN khai: 30 triệu đ
- Tỷ suất thực: 30 ÷ 500 = 6%

Phân tích:
  Sai lệch = (6% - 4.47%) ÷ 4.47% = +34% (CAO HƠN)
  
  ✅ Không bất thường, khả năng:
    - Sản phẩm cao cấp (margin cao)
    - Chiều dài chi phí thấp (sản xuất hiệu quả)
    - Kiểm tra: Chi phí COGS có phù hợp không?
    
Kết luận: Chấp nhận, không cần bổ sung
```

### Ví dụ 2: Bán buôn thực phẩm - Tỷ suất thấp hơn chuẩn

```
DN B: Bán buôn thực phẩm
- Năm kiểm tra: 2023
- Chuẩn ngành: 0.23%
- Doanh thu: 200 tỷ đ (khai)
- TNDN khai: 0 đ (khai hòa vốn)
- Tỷ suất thực: 0%

Phân tích:
  Sai lệch = (0% - 0.23%) ÷ 0.23% = -100% (THẤP RẤT NHIỀU)
  🔴 RỦI RO CAO → Cần điều tra sâu
  
Hành động:
  1. Kiểm tra doanh thu:
     - HĐ phát hành có bằng 200 tỷ?
     - Tờ khai GTGT cho bao nhiêu doanh thu?
     - Sổ bán hàng có trùng hợp?
  2. Kiểm tra chi phí:
     - COGS (hàng nhập) bằng doanh thu?
     - Có chi phí vận tải, kho bãi, lương?
  3. Nếu phát hiện doanh thu bị khai thiếu:
     - Bổ sung: DN X × 0.23% - 0 = DN X × 0.23%
  4. Nếu phát hiện chi phí thừa:
     - Loại bỏ chi phí giả, tính lại TNDN

Ví dụ: Phát hiện doanh thu thực 250 tỷ (khai 200 tỷ)
  → TNDN bổ sung = 250 × 0.23% - 0 = 575 triệu
  → Phạt + lãi
```

### Ví dụ 3: May mặc - Tỷ suất gần chuẩn

```
DN C: Sản xuất may mặc
- Năm kiểm tra: 2023
- Chuẩn ngành: 1.24%
- Doanh thu: 500 tỷ đ
- TNDN khai: 6.5 tỷ đ
- Tỷ suất thực: 6.5 ÷ 500 = 1.3%

Phân tích:
  Sai lệch = (1.3% - 1.24%) ÷ 1.24% = +4.8% (nhỏ)
  ✅ Bình thường, nằm trong ±20%
  
Kết luận: Chấp nhận, không cần điều tra thêm
```

### Ví dụ 4: SX hóa chất - Kiểm tra chi phí bất thường

```
DN D: Sản xuất hóa chất
- Năm kiểm tra: 2023
- Chuẩn ngành: 3.38%
- Doanh thu: 1000 tỷ đ
- TNDN khai: 15 tỷ đ (1.5%)
- Tỷ suất thực: 15 ÷ 1000 = 1.5%

Phân tích:
  Sai lệch = (1.5% - 3.38%) ÷ 3.38% = -56% (THẤP NHIỀU)
  🔴 RỦI RO CAO → Điều tra sâu

Kiểm tra chi phí:
  - Lương nhân viên: 300 tỷ (khai) → 30% doanh thu
  - Lương bình quân ngành: ~15-20% doanh thu
  
Phát hiện:
  ❌ Chi phí lương khai thừa 100 tỷ
  ❌ Khấu hao khai 50 tỷ (phải kiểm tra TSCĐ)
  ❌ Chi phí vận tải 100 tỷ (phải kiểm tra HĐ)

Hành động:
  → Loại bỏ chi phí không hợp lý
  → Tính lại TNDN
  → TNDN bổ sung = 1000 × 3.38% - 15 = 18.8 tỷ
  → Thuế + phạt + lãi
```

---

## Bẫy thường gặp

### ❌ Lỗi 1: Sử dụng chuẩn sai năm

```
SAI:  DN kiểm tra năm 2023, chuẩn năm 2021
ĐÚNG: Sử dụng chuẩn năm 2023 (năm được kiểm tra)
      Nếu năm đó chưa có chuẩn → Dùng năm gần nhất có dữ liệu
```

### ❌ Lỗi 2: So sánh sai doanh thu

```
SAI:  Doanh thu GTGT = Doanh thu TNDN (bao gồm cả chiết khấu)
ĐÚNG: Doanh thu TNDN = DT thực bán (sau trừ chiết khấu, trả hàng, miễn thuế)
      Không bao gồm doanh thu tài chính
```

### ❌ Lỗi 3: Không xem xét ngành phụ

```
SAI:  DN vừa sản xuất vừa bán lẻ, dùng chuẩn ngành sản xuất (3.38%)
ĐÚNG: Xác định ngành chính (>50% doanh thu)
      Nếu 50-50: Tính bình quân 2 chuẩn
```

### ❌ Lỗi 4: Bắt tất cả DN phải đúng chuẩn

```
SAI:  DN tỷ suất 2% > 1.24% chuẩn → Bắt dùng chuẩn
      DN tỷ suất 0.5% < 1.24% chuẩn → Tính bổ sung

ĐÚNG: Chuẩn là tham chiếu, không bắt buộc
      Sai lệch > ±20% → Điều tra nguyên nhân
      Nếu giải thích hợp lý → Chấp nhận
```

### ❌ Lỗi 5: Không phân biệt lợi nhuận trước/sau thuế

```
SAI:  TNDN = Lợi nhuận trước thuế (chưa trừ 10% surtax)
ĐÚNG: TNDN = Lợi nhuận chịu thuế TNDN (đã trừ chi phí hợp lý)
      (Nếu dùng lợi nhuận sau thuế → Chia 0.9 để convert ngược)
```

---

## Công cụ hỗ trợ

### Excel Template

**Bảng đối chiếu tỷ suất:**
```
| STT | Công ty | Ngành | Năm | DT (tỷ) | TNDN khai (tỷ) | Tỷ suất % | Chuẩn % | Sai lệch % | Rủi ro | Hành động |
```

### Checklist

- [ ] Xác định chính xác ngành của DN (mã VSIC)
- [ ] Chọn năm chuẩn phù hợp
- [ ] Tính tỷ suất thực tế (TNDN ÷ DT)
- [ ] So sánh với chuẩn, tính sai lệch %
- [ ] Nếu sai lệch ±20%:
  - [ ] Kiểm tra doanh thu (HĐ, tờ khai GTGT)
  - [ ] Kiểm tra chi phí (COGS, lương, khấu hao)
  - [ ] Kiểm tra dữ liệu ngoài (sổ CVCBRR, sxnt)
- [ ] Ghi nhận rủi ro + hành động

---

**Cập nhật:** Chuẩn tỷ suất sinh lợi ngành 2017-2023 từ Cục Thuế Tỉnh Đồng Nai  
**Quản lý:** Cục Thuế Nhập Cảnh, Chi cục Thuế Địa phương
