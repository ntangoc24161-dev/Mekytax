---
name: phieu-chuyen-noi-bo-giai-the
description: Lập Phiếu chuyển nội bộ hồ sơ giải thể đề nghị điều chỉnh nộp thừa về 0 để đóng mã số thuế doanh nghiệp theo đúng mẫu quy định.
group: Tài chính
tags: [giai-the, phieu-chuyen, nop-thua, dong-mst]
---

# Lập Phiếu chuyển nội bộ hồ sơ giải thể

Kỹ năng tự động hóa lập Phiếu chuyển thông tin/hồ sơ nội bộ gửi Tổ Quản lý hỗ trợ doanh nghiệp/quản lý nợ để điều chỉnh số thuế nộp thừa về 0, hoàn tất điều kiện đóng mã số thuế (MST) sau quyết toán giải thể.

## 1. Nguồn mẫu chuẩn (Template)
- **Đường dẫn phôi:** `skills/phieu-chuyen-noi-bo-giai-the/templates/PHIEU CHUYEN NOI BO-GIAI THE.docx`
- Giữ nguyên toàn bộ định dạng bảng, font chữ Times New Roman và cấu trúc thẩm quyền theo Nghị định 30/2020/NĐ-CP.

## 2. Dữ liệu đầu vào cần thiết
- **Thông tin NNT:** Tên người nộp thuế, Mã số thuế, Địa chỉ trụ sở.
- **Căn cứ pháp lý & hồ sơ:**
  - Biên bản kiểm tra quyết toán giải thể (Số, ngày lập).
  - Quyết định xử phạt vi phạm hành chính (Số QĐ, ngày ban hành, số tiền xử phạt/truy thu nếu có).
  - Văn bản cam kết từ chối nhận lại tiền nộp thừa của doanh nghiệp (ngày lập).
- **Chi tiết các khoản nộp thừa:**
  - Tiểu mục hạch toán (Ví dụ: 4254 - Lệ phí môn bài, 4944 - Tiền chậm nộp khác...).
  - Số tiền nộp thừa đề nghị tất toán về 0 (đơn vị: đồng).

## 3. Quy trình thực hiện
1. **Kiểm tra điều kiện:**
   - Đảm bảo doanh nghiệp đã có văn bản từ chối nhận lại tiền nộp thừa (văn bản số 20).
   - Đã ban hành hoặc lập xong Quyết định xử phạt VPHC / Biên bản kiểm tra.
2. **Điền phôi `PHIEU CHUYEN NOI BO-GIAI THE.docx`:**
   - Cập nhật số hiệu văn bản, ngày tháng năm phát hành phiếu chuyển.
   - Điền thông tin doanh nghiệp, số quyết định xử phạt, biên bản kiểm tra.
   - Bảng kê chi tiết các khoản nộp thừa theo từng tiểu mục và số tiền tương ứng.
   - Nơi nhận: Tổ Quản lý hỗ trợ doanh nghiệp / Quản lý nợ, lưu VT, Tổ KT.
3. **Thẩm định & Lưu trữ:**
   - Trợ lý **Giám sát hồ sơ** đối chiếu khớp đúng số tiền nộp thừa với TMS/Bảng tính.
   - Xuất file Word hoàn chỉnh vào thư mục hồ sơ vụ kiểm tra.
