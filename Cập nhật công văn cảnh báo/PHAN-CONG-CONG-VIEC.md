# Phân công công việc - Phòng cảnh báo rủi ro

> **Mô hình phòng làm việc:** Trưởng phòng (người dùng) → Javis Core (điều phối) → Agents thi công + 1 QC gác cổng

> **ĐIỂM KẾT THÚC CÔNG VIỆC (cập nhật 09/11/2026):** một lô hồ sơ coi là XONG khi **nạp sổ F0/F1 và sắp xếp lưu trữ đã qua QC đạt**. Đối chiếu NCC và soạn thảo công văn **chỉ thực hiện khi Trưởng phòng yêu cầu rõ bằng một lệnh riêng**; không yêu cầu thì không chạy và không tính là việc còn dở.

---

## Cấu trúc tổ chức

```
┌─────────────────────────────────┐
│  Trưởng phòng cảnh báo rủi ro   │  (Người dùng)
│  - Giao lệnh xử lý              │
│  - Phê duyệt QC                 │
│  - Nhận báo cáo kết quả         │
└──────────────┬──────────────────┘
               │
               ↓
┌─────────────────────────────────┐
│  Javis Core "Trưởng phòng cảnh báo" │  (Trợ lý điều phối)
│  - Kích hoạt workflow           │
│  - Báo cáo tiến độ              │
│  - Chuyển tiếp báo cáo QC       │
└──────────────┬──────────────────┘
               │
      ┌────────┼────────┐
      ↓        ↓        ↓
   [7 Agents] [QC] [Support]
```

---

## Chi tiết vai trò từng bộ phận

### 1. TRƯỞNG PHÒNG CẢNH BÁO (Người dùng)

**Trách nhiệm:**
- Giao lệnh: "Cập nhật CV từ ngày X" hoặc "Xử lý file vừa gửi lên Zalo"
- Phê duyệt hoặc từ chối kết quả QC
- Yêu cầu sửa lại nếu QC phát hiện lỗi
- Nhận báo cáo tổng hợp hàng ngày

**Quyền hạn:**
- Quyết định cuối cùng về việc release kết quả
- Yêu cầu recheck hoặc sửa lại từ bất kỳ agent nào
- Điều chỉnh quy trình nếu cần

---

### 2. JAVIS CORE - "Trưởng phòng cảnh báo" (Trợ lý điều phối)

**Trách nhiệm:**
- **Nhận lệnh** từ Trưởng phòng
- **Kích hoạt workflow** đúng tính chất công việc
- **Chuyển tiếp báo cáo QC** cho Trưởng phòng
- **Báo cáo tiến độ hàng ngày:**
  - Số CV xử lý được
  - Số lỗi QC phát hiện
  - Số đã phê duyệt release
  - Đang chực ở bước nào

**Deliverable:**
```
📊 BÁNG CÁO TIẾN ĐỘ [HÔM NAY]

✅ Hoàn thành: 
  - CV 123/2026 (UBND Đồng Nai) - Release ✅
  - CV 456/2026 (Công an tỉnh) - Chờ phê duyệt ⏳

❌ Lỗi QC cần sửa:
  - CV 789/2026: MST sai (agent Cập nhật sổ đang sửa)
  
🔄 Đang chạy:
  - 5 CV trong bước OCR (agent Số hoá)
  - 3 CV trong bước đối chiếu (agent Đối chiếu NCC)
```

---

### 3. 7 AGENTS (Thi công chuyên môn)

#### 🔵 Agent 1: Tra cứu iOffice
- **Nhập:** Lệnh từ Trưởng phòng (hoặc Zalo/Drive)
- **Xuất:** Danh sách file + đường dẫn
- **Không làm:** OCR, bóc tách, nạp sổ
- **Phê duyệt:** QC (nếu yêu cầu)

#### 🔵 Agent 2: Số hoá OCR
- **Nhập:** File PDF/ảnh từ agent 1
- **Xuất:** File `.md` sạch (sau QC ✅)
- **Không làm:** Bóc tách, nạp sổ
- **Phê duyệt:** QC → Trưởng phòng

#### 🔵 Agent 3: Cập nhật Sổ (2 nhiệm vụ)
- **Nhập:** File `.md` (từ agent 2)
- **Xuất (nhiệm vụ 1):** 
  - Sheet F0/F1 cập nhật (sau QC ✅)
  - Báo số dòng nạp + số trùng
- **Xuất (nhiệm vụ 2 - Sắp xếp):**
  - File công văn di chuyển đến thư mục chuẩn
  - Tên file chuẩn (Năm_SoCV_CoQuan_SoCongTy)
- **Không làm:** Đối chiếu, soạn công văn
- **Phê duyệt:** QC → Trưởng phòng

#### 🟡 Agent 4: Đối chiếu NCC - CHỈ CHẠY KHI ĐƯỢC YÊU CẦU
- **Khởi động:** lệnh riêng của Trưởng phòng, ví dụ "đối chiếu NCC của MST ... kỳ ...". Không nằm trong dây chuyền trục.
- **Nhập:** Số CV từ agent 3 (từ sổ F0/F1) + Bảng kê/HĐĐT
- **Xuất:** 
  - Phần 1, Bảng 2.1, 2.2, 2.3, Phần 3 (sau QC ✅)
  - Ghi rõ nguồn (Bảng kê / HĐĐT / Chung)
- **Không làm:** Bóc tách, nạp sổ, soạn công văn
- **Phê duyệt:** QC → Trưởng phòng

#### 🟡 Agent 5: Soạn thảo - CHỈ CHẠY KHI ĐƯỢC YÊU CẦU
- **Khởi động:** lệnh riêng của Trưởng phòng, ví dụ "soạn công văn trả lời CV 7268". Không nằm trong dây chuyền trục.
- **Nhập:** Dữ liệu từ agent 3 + 4 + luật pháp
- **Xuất:**
  - `TRA LOI CV [SoCV]_[CoQuan]_du thao.docx`
  - `TO TRINH [SoCV].docx`
  - Trích dẫn Điều/khoản/Luật đầy đủ
- **Không làm:** Bịa số liệu, gọi pháp chế không ghi chỗ trống
- **Phê duyệt:** QC → Trưởng phòng

#### 🔵 Agent 6: Javis Core (phục vụ)
- **Nhập:** Yêu cầu thêm từ Trưởng phòng
- **Xuất:** Hỗ trợ, tìm tài liệu, giải thích luật, sửa công văn nhỏ
- **Không làm:** Bóc tách, nạp sổ (giao cho agent chuyên)

---

### 4. QC DỮ LIỆU (Kiểm soát chất lượng)

**Trách nhiệm:**
- Gác cổng dây chuyền trục: **Thu thập / OCR / Bóc tách nạp sổ / Sắp xếp lưu trữ**. QC bước Sắp xếp đạt là **chốt kết thúc công việc**.
- Gác cổng hai khâu Đối chiếu và Soạn thảo **chỉ khi hai khâu đó được yêu cầu chạy**.
- **Báo cáo kết quả cho Trưởng phòng:**
  - ✅ **PASS** → Agent này hoàn thành, tiếp bước tiếp
  - ⚠️ **CẢNH BÁO** → Dữ liệu trùng nhưng có thể duyệt (Trưởng phòng quyết)
  - ❌ **LỖI** → Yêu cầu agent sửa lại + recheck

**Deliverable:**
```
🔍 KẾT QUẢ QC - CV 123/2026

📋 Giai đoạn OCR: ✅ PASS
   ✓ Ký tự rác: OK
   ✓ Bảng biểu: OK
   ✓ Cấu trúc heading: OK

📋 Giai đoạn Bóc tách: ❌ LỖI
   ✗ MST F0: "0314123789 " (khoảng trắng)
   ✗ Ngày CV: "15-10-2026" (sai định dạng)
   → YÊU CẦU AGENT "Cập nhật Sổ" SỬA LẠI

📋 Giai đoạn Sắp xếp lưu trữ: ⏳ CHỜ BƯỚC TRƯỚC SỬA
Đối chiếu và Soạn thảo: KHÔNG ĐƯA VÀO LÔ NÀY (chưa được yêu cầu)

✍️ GỬI CHO TRƯỞNG PHÒNG: Phê duyệt/từ chối/yêu cầu sửa thêm
```

---

## Luồng báo cáo (Reporting Flow)

```
┌──────────────┐
│ Agent (làm)  │
└──────┬───────┘
       │ báo cáo kết quả
       ↓
┌──────────────┐
│ QC (kiểm tra)│
└──────┬───────┘
       │ báo cáo: ✅/⚠️/❌
       ↓
┌──────────────────────┐
│ Trưởng phòng (quyết) │  ← PHẢI ĐỌC BÁNG CÁO
│ - Duyệt ✅            │
│ - Từ chối ❌          │
│ - Yêu cầu sửa thêm   │
└──────┬───────────────┘
       │
       ↓
   [Tiếp bước tiếp hoặc yêu cầu sửa]
```

**Quy tắc:**
- QC **không quyết định** được release (chỉ báo cáo)
- **Trưởng phòng** quyết định cuối cùng
- Agent không được skip QC
- Lỗi QC phải sửa trước khi báo "hoàn thành"

---

## Ví dụ: Xử lý 1 lô công văn từ đầu đến cuối

```
BƯỚC 1: TRƯỞNG PHÒNG GIAO LỆNH
└─ "Xử lý hồ sơ tại G:\...\01_HO SO CONG AN CHUYEN DEN\2026\HO SO MOI"

BƯỚC 2: JAVIS CORE "Trưởng phòng cảnh báo" KÍCH HOẠT DÂY CHUYỀN TRỤC
└─ wf:day-chuyen-xu-ly-cv-canh-bao, 4 bước, mỗi bước một cổng QC

BƯỚC 3: AGENTS CHẠY, QC GÁC TỪNG CỔNG
├─ Agent Tra cứu      -> bảng tệp + "TONG: n tep"
├─ QC: pass true      -> tiếp
├─ Agent Số hoá       -> bảng đường dẫn .md + số ký tự
├─ QC: pass true      -> tiếp
├─ Agent Nạp sổ       -> F0 truoc/sau, F1 truoc/sau, da nap, bo trung
├─ QC: pass false (MST sai) -> agent sửa, recheck
├─ QC: pass true      -> tiếp
├─ Agent Sắp xếp      -> từng dòng "đường dẫn cũ -> đường dẫn mới"
└─ QC: pass true      -> CHỐT KẾT THÚC

BƯỚC 4: BÁO CÁO TRƯỞNG PHÒNG RỒI DỪNG
└─ Số dòng nạp F0/F1, số bỏ trùng, danh sách tệp đã sắp xếp.
   KHÔNG tự chạy đối chiếu. KHÔNG tự soạn công văn. HẾT VIỆC.

BƯỚC 5 (TUỲ CHỌN, CHỈ KHI TRƯỞNG PHÒNG YÊU CẦU)
└─ "đối chiếu NCC của MST ..." hoặc "soạn công văn trả lời CV ..."
   -> wf:tra-loi-ho-so-cong-an, 2 bước, cũng có QC gác cổng
```

**Trượt cổng QC hết số lần thử lại thì quy trình dừng thật**, không đẩy dữ liệu rỗng xuống bước sau.

---

## Ghi chú về báo cáo hàng ngày

**Javis Core gửi báo cáo cho Trưởng phòng mỗi lúc:**
1. ✅ Công văn hoàn thành (QC ✅, chờ phê duyệt)
2. ⚠️ Lỗi QC phát hiện (chờ agent sửa)
3. 📊 End-of-day summary (tổng số xử lý, lỗi, phê duyệt)

**Format báo cáo:**
```
🕐 [Thời gian] 
📋 Công việc: [Số CV / tên công ty]
📊 Trạng thái: [✅ DONE / ⏳ DOING / ❌ ERROR]
👤 Agent: [Tên agent chịu trách]
🔍 QC: [Kết quả nếu có]
⏭️ Bước tiếp: [Nếu chưa xong]
```

