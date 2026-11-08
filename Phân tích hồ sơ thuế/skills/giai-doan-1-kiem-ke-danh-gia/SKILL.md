---
name: Giai đoạn 1 - Kiểm kê & Đánh giá hồ sơ thuế
description: "Kiểm kê manifest IN00-IN13, xác định mức hồ sơ A-E, rà soát khả năng chạy 80 quy tắc, xếp danh mục đề nghị DN bổ sung (KQ1)."
group: Tài chính
updated: '2026-10-15'
---

# Giai đoạn 1 – Kiểm kê hồ sơ & Đánh giá mức độ dữ liệu

Giai đoạn đầu tiên và bắt buộc đối với mọi hồ sơ kiểm tra thuế khi tiếp nhận. Quét toàn bộ thư mục hồ sơ trên đĩa (kể cả file ngoài cấu trúc chuẩn), lập bảng kiểm kê manifest (IN00), đánh giá độ bao phủ dữ liệu (IN01–IN13), xác định mức hồ sơ từ A đến E, rà soát khả năng chạy của 80 quy tắc rủi ro và lập danh mục đề nghị Người nộp thuế (NNT) bổ sung hồ sơ theo thứ tự ưu tiên mở khóa quy tắc.

> Áp dụng nghiêm ngặt **Khối Q** (Quy tắc chung): Không xóa file nguồn gốc; không tính toán rủi ro hoặc kết luận vi phạm ở bước này; không để số 0 giả khi thiếu nguồn.

---

## Khi nào kích hoạt

- Bắt đầu tiếp nhận một thư mục hồ sơ doanh nghiệp mới hoặc khi NNT vừa nộp bổ sung tài liệu.
- Các câu lệnh/yêu cầu: "kiểm kê hồ sơ", "đánh giá mức độ hồ sơ", "kiểm tra độ đủ dữ liệu", "chạy giai đoạn 1", "lệnh P1", "xác định mức A-E".
- Hồ sơ ở bất kỳ trạng thái dữ liệu nào (từ chỉ có vài file XML/PDF rời rạc đến đầy đủ sổ sách).

---

## Phân loại 5 mức hồ sơ (A – E)

| Mức | Dữ liệu thực tế có trong thư mục | Khả năng chạy phân tích |
|---|---|---|
| **A** | Chỉ có dữ liệu cơ quan thuế: HĐĐT, bảng kê, tờ khai GTGT, QTT TNDN, BCTC nộp thuế, TNCN, cảnh báo rủi ro. | Chạy Giai đoạn 2 (57 quy tắc dữ liệu CQT). |
| **B** | Dữ liệu mức A cộng thêm BCTC, tờ khai do doanh nghiệp nộp/gửi riêng. | Chạy Giai đoạn 2 + đối chiếu bản DN vs bản CQT (phát hiện DQ). |
| **C** | Dữ liệu mức B cộng thêm Sổ kế toán nội bộ (IN08: Nhật ký chung, Sổ Cái) và Sổ 511/632 (IN13). | Chạy Giai đoạn 3 một phần (511–632, công nợ, GT06, TN17). |
| **D** | Dữ liệu mức C cộng thêm Sổ Nhập–Xuất–Tồn kho (TK 152/155/156) và Bảng giá thành/định mức (IN09). | Chạy Giai đoạn 3 thêm nhóm Kho & Giá thành (KH01–KH07). |
| **E** | Dữ liệu mức D cộng thêm Sổ chi tiết tiền mặt (IN10: TK 111), Sao kê ngân hàng và UNC (IN11: TK 112). | Chạy Giai đoạn 3 toàn diện (đủ 80 quy tắc). |

---

## Đầu vào

### Đầu vào bắt buộc:
- Đường dẫn thư mục hồ sơ doanh nghiệp trên đĩa: `<FOLDER>`.
- Thông tin định danh: Mã số thuế `<MST>`, Tên doanh nghiệp `<TEN>`, Niên độ kiểm tra `<KY>`.

### Quy định phân loại nguồn gốc thư mục (Quyết định K1):
- Nhãn CQT hoặc DN suy từ thư mục cấp 1:
  - `1. DU LIEU THO NGANH THUE` = Dữ liệu cơ quan thuế (1.1 = tờ khai XML; 1.3 = Nghĩa vụ nộp thuế & Thông báo TMS;...).
  - `2. HO SO NNT CUNG CAP` = Dữ liệu công ty cung cấp (2.5 = bảng kê mua vào bán ra;...).
  - `3. THONG TIN TRA CUU` (3.3 = file tra cứu đa nguồn).
  - `4. KET QUA PHAN TICH` = Nơi chứa kết quả.

### Các file quét và phân loại vào 15 mẫu nhập (IN00–IN14):
1. **IN00_Manifest**: Mọi file đầu vào (tên, ngày nhận, mã băm SHA-256, định dạng đọc, số dòng).
2. **IN01_HDDT**: Hóa đơn điện tử portal CQT (chiều MUA và BAN).
3. **IN02_BangKe**: Bảng kê mua vào / bán ra do NNT kê khai.
4. **IN03_ToKhai_GTGT**: Tờ khai 01/GTGT chính thức và bổ sung (XML).
5. **IN04_BCTC_TNDN**: BCTC (CĐKT, KQKD, LCTT) và tờ khai QTT 03/TNDN (XML).
6. **IN05_NLD_TNCN**: Danh sách người lao động trích từ phụ lục 05-1/BK-TNCN và 05-2/BK-TNCN.
7. **IN06_QTT_TNCN**: Tờ khai QTT 02/QTT-TNCN hoặc 05/QTT-TNCN và các kỳ tạm tính quý.
8. **IN07_NCC_CanhBao**: Danh mục NCC thuộc diện CVCBRR (F0, F1), Công an, ĐKKD bất thường.
9. **IN08_So_KeToan**: Sổ Nhật ký chung, Sổ Cái các tài khoản (PS Nợ 133, 111, 112, 131, 331).
10. **IN09_Kho_GiaThanh**: Sổ XNT kho 152, 155, 156, Bảng tính giá thành 154/621/622/627, định mức.
11. **IN10_Chi_tien_mat**: Sổ chi tiết tiền mặt TK 1111 (phiếu chi).
12. **IN11_Sao_ke_NH**: File sao kê ngân hàng TK 1121 (dữ liệu dòng tiền và ủy nhiệm chi).
13. **IN12_GDLK**: Phụ lục giao dịch liên kết theo Nghị định 132/2020/NĐ-CP.
14. **IN13_So_511_632**: Sổ chi tiết doanh thu 511 và giá vốn 632 theo dòng bút toán.
15. **IN14_TMS_Nghia_vu**: Lịch sử tuân thủ từ hệ thống TMS (nợ thuế, chậm nộp, kết luận TTKT - thư mục 1.3).

---

## Quy trình thực hiện

1. **Quét thư mục toàn diện:** Đọc toàn bộ cây thư mục `<FOLDER>`, tính mã băm SHA-256 cho từng file để bảo toàn tính toàn vẹn và truy vết.
2. **Xếp loại nguồn vào IN00–IN13:** Phân loại rõ `Nguồn gốc` (CQT hay DN), `Định dạng` (`Cau_truc`, `PDF-da-trich`, `PDF-chua-trich`, `Anh-chua-doc`). Nếu có sao kê PDF scan chưa OCR thì gắn cờ `PDF-chua-trich`.
3. **Đánh giá độ đủ từng nguồn:** Đếm số dòng thực, kiểm tra số dòng thiếu MST hoặc Ký hiệu HĐ, kiểm tra có tờ khai bổ sung không. Gán nhãn: `Đủ`, `Mỏng` (dưới 12 dòng), `Thiếu`, `PDF chưa trích`, hoặc `Chưa nhận`.
4. **Rà soát 80 quy tắc rủi ro:** Chiếu theo cột "Dữ liệu cần" của sheet `04_Quy_tac_rui_ro` để xác định từng quy tắc: `Chạy được ngay`, `Chạy một phần`, hay `Bị chặn`. Nêu rõ quy tắc bị chặn vì thiếu nguồn IN nào.
5. **Kết luận mức hồ sơ (A–E) & Chỉ định giai đoạn tiếp theo:**
   - Nếu ở mức A hoặc B: Chuyển tiếp chạy **Giai đoạn 2** (`giai-doan-2-phan-tich-du-lieu-co-quan-thue`).
   - Nếu đã đạt mức C, D hoặc E: Có thể chuyển tiếp sang **Giai đoạn 3** (`giai-doan-3-hoan-chinh-ho-so`).
6. **Lập danh mục đề nghị bổ sung:** Gom nhóm các tài liệu DN chưa cung cấp, sắp xếp theo thứ tự giải phóng được nhiều quy tắc nhất:
   - Ưu tiên 1: Sổ kho và giá thành IN09 (mở khóa 10 quy tắc).
   - Ưu tiên 2: Sao kê ngân hàng và UNC IN11 (mở khóa 5 quy tắc).
   - Ưu tiên 3: Sổ kế toán tổng hợp IN08 (mở khóa 4 quy tắc).
   - Ưu tiên 4: Sổ chi tiết 511/632 IN13 (mở khóa 4 quy tắc).
   - Ưu tiên 5: Sổ chi tiết tiền mặt IN10 (mở khóa 2 quy tắc).

---

## Cổng dừng (Stop gates)

- **Dừng ngay** nếu đường dẫn thư mục `<FOLDER>` không tồn tại hoặc không thể đọc được.
- **Dừng ngay** nếu không xác định được mã số thuế `<MST>` và kỳ kiểm tra `<KY>`.
- **Cấm tự ý phân tích số liệu hoặc tính chênh lệch rủi ro** trong giai đoạn 1.
- **Cấm xóa hoặc sửa file nguồn gốc** của người nộp thuế (chỉ cho phép ghi nhận và lưu trữ).

---

## Hợp đồng đầu ra

Xuất file Excel kết quả **chỉ chứa các giá trị tĩnh**:
`KQ1_<MST>_<KY>_v<n>.xlsx` (mẫu chuẩn tại `KQ1_MAU_Kiem_ke_danh_gia.xlsx`) gồm 3 sheet:
1. **`KQ1_Kiem_ke`**: Bảng kiểm kê chi tiết mọi nguồn nhận được (mã nguồn, định dạng, nguồn gốc, số dòng, mã SHA-256, đánh giá độ đủ).
2. **`KQ1_Kha_nang_phan_tich`**: Danh mục 80 quy tắc với khả năng chạy tương ứng theo hiện trạng hồ sơ.
3. **`KQ1_De_nghi_bo_sung`**: Danh mục tài liệu đề nghị NNT bổ sung xếp theo số lượng quy tắc được mở khóa.

---

## Câu lệnh thực thi mẫu

### Lệnh đầy đủ (P1):
```text
Giai đoạn 1 – Kiểm kê hồ sơ <FOLDER>, MST <MST>, kỳ <KY>. Áp dụng khối Q.
1. Quét toàn bộ thư mục, kể cả file nằm ngoài cấu trúc 4 bước. Xếp mỗi file vào một nguồn IN00–IN13. File không xếp được thì liệt kê riêng, không bỏ qua.
2. Mỗi nguồn ghi: nguồn gốc (CQT/DN), định dạng (XML/Excel/PDF/ảnh), kỳ phủ, số dòng, số dòng thiếu MST hoặc ký hiệu HĐ, có bản bổ sung không, mã băm SHA-256. Kết luận: Đủ / Mỏng / Thiếu / PDF chưa trích / Chưa nhận.
3. Với 80 quy tắc, nêu quy tắc nào chạy được ngay, chạy một phần, bị chặn; bị chặn vì nguồn nào.
4. Kết luận mức hồ sơ A–E và giai đoạn chạy tiếp (2 hoặc 3).
5. Lập danh mục đề nghị bổ sung cho công ty, xếp theo số quy tắc được mở khóa.
Chưa phân tích rủi ro ở bước này.
Đầu ra: KQ1_<MST>_<KY>_v1.xlsx gồm KQ1_Kiem_ke, KQ1_Kha_nang_phan_tich, KQ1_De_nghi_bo_sung.
```

### Lệnh siêu ngắn:
```text
Kiểm kê GĐ1 hồ sơ <FOLDER> MST <MST> kỳ <KY> xuất KQ1
```
