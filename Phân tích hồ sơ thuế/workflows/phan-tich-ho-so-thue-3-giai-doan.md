---
type: workflow
name: Phân tích hồ sơ thuế 3 giai đoạn
slug: phan-tich-ho-so-thue-3-giai-doan
status: active
group: Tài chính
description: "Chuỗi Lệnh 1-3 quy trình 1.4: KQ1.1/1.2/1.3 rồi KQ2.1/2.2 rồi KQ3; module ghi sheet đích."
steps:
  - agent: thu-ho-so-kiem-ke
    task: "Lệnh 1. Đọc {{input}}: folder tuyệt đối, MST, tên DN, kỳ kiểm tra, loại hình nếu có. Tổ chức cây 4 bước nếu chưa chuẩn. Sao mẫu KQ1.1 KQ1.2 KQ1.3 rồi điền IN00-IN05/IN08-IN14 (IN07 không ở KQ1.1; IN05 có Loai_dong). Xuất KQ1.1, KQ1.2, KQ1.3 v1. Kết luận mức A-E. Không phân tích rủi ro, không xóa nguồn."
  - agent: truong-phong-phan-tich-ho-so
    task: "Đọc KQ1.3 từ {{prev}}. Công bố mức A-E. Mức không chặn GĐ2/GĐ3. Ghi rõ IN còn thiếu. Lệnh tiếp: Lệnh 2 rồi Lệnh 3. Không tự kiểm kê lại."
  - agent: chuyen-vien-du-lieu-dau-vao
    task: "Gộp HĐĐT + bảng kê + NLĐ + CD1 BT5-BT9 theo folder trong {{input}}. Ghi KQ1.1!IN01,IN05 và KQ1.2!IN02. Một phiên bản mới của tệp đích. Không xuất mẫu 02 hay Doi_chieu_BT5_BT9 để giao. Thiếu 1 nguồn: ô trống, không điền 0."
  - agent: chuyen-vien-gtgt-tndn
    task: "Đối chiếu bảng kê - tờ khai GTGT - QTT TNDN - HĐĐT. Ghi KQ2.1!KQ2_Doi_chieu_GTGT. Chỉ số, chênh lệch, cờ máy. Kỳ trống bỏ dòng. Không xuất Tong_hop_doi_chieu_GTGT_TNDN. Một phiên bản mới của KQ2.1."
  - agent: chuyen-vien-ncc-hoa-don
    task: "Đối chiếu 100% NCC với CVCBRR F0/F1, Công an, danh bạ. Ghi KQ2.1!IN07_NCC_CanhBao và KQ2_NCC_mua_vao. F0/F1 là nhóm F; mức Cao chỉ ở KQ3. Không xuất Tong_hop_NCC. Một phiên bản mới của KQ2.1. Quy tắc khác CHƯA_CHẠY."
  - agent: chuyen-vien-tncn
    task: "Tổng hợp QTT TNCN, NLĐ, đa nguồn nếu có 3.3. Ghi KQ2.1!KQ2_TNCN_LD2. Thiếu 3.3: ô trống. Không xuất Tong_hop_TNCN. Một phiên bản mới của KQ2.1."
  - agent: phan-tich-vien-cqt
    task: "Lệnh 2. Lắp KQ2.1 và KQ2.2 từ mẫu sau KQ1 và các sheet vừa ghi ({{prev}}). Ô công ty trống nếu chưa có IN02/IN08-IN13. Không ghi mức rủi ro/nhận xét/tổng tiền vào KQ2. Chưa có KQ1 thì dừng, yêu cầu Lệnh 1."
  - agent: chuyen-vien-doanh-thu-gia-von
    task: "Đối chiếu sổ 511 vs 632 (K4) nếu có IN13. Ghi KQ2.2!KQ2_511_632. Không có sổ: để trống, KHÔNG_ĐỦ_DỮ_LIỆU ở KQ3, không suy từng dòng từ tờ khai. Một phiên bản mới của KQ2.2."
  - agent: chuyen-vien-kho-gia-thanh
    task: "XNT 152/155/156 vs BCTC (K6) và giá thành nếu có IN09. Ghi KQ2.2!KQ2_Kho_Gia_thanh và KQ2_Gia_thanh_va_Dinh_muc. Thiếu sổ: ô trống, không kết luận xuất khống. Một phiên bản mới của KQ2.2."
  - agent: chuyen-vien-cong-no-tien-mat
    task: "Quét 111, cụm cùng ngày, 131/331, sao kê 112. Ngưỡng K5 >= 20tr đến 30/06/2025 và >= 5tr từ 01/07/2025. P23=10%. Ghi KQ2.2!KQ2_Cong_no_Tien_mat. PDF chưa OCR: PDF - chưa kết luận. Một phiên bản mới của KQ2.2."
  - agent: phan-tich-vien-so-sach-dn
    task: "Lệnh 3 luôn chạy sau Lệnh 2, kể cả mức A/B. Sao KQ3, 80 quy tắc, 44 Trực tiếp cột B=Trực tiếp. Thiếu nguồn: KHÔNG_ĐỦ_DỮ_LIỆU / Chưa kết luận, cột AA. Chống trùng T1-T10. Đề cương cột AB-AD AH; chuỗi bằng chứng AI AJ. Không tách sheet KQ3 cũ. Không ghi đè bản cũ."
  - agent: truong-phong-phan-tich-ho-so
    task: "Chốt: đường dẫn sáu tệp, tổng cột W theo sắc thuế, 5-10 trọng tâm từ KQ3, hồ sơ còn thiếu. Master 8 sheet / 28 TC chỉ khi cán bộ yêu cầu, sinh từ KQ3 (K3). Từ chối sheet PL và workbook trọn gói. Font Times New Roman. Số tiền dấu chấm nghìn, phẩy lẻ."
updated: 2026-10-19
---
Chuỗi Lệnh 1 rồi Lệnh 2 rồi Lệnh 3 cho một doanh nghiệp, theo bộ mẫu v1.4 (sáu tệp KQ1.1, KQ1.2, KQ1.3, KQ2.1, KQ2.2, KQ3).

Đầu vào bắt buộc trong {{input}}: folder tuyệt đối của DN, MST, kỳ kiểm tra. Tên DN và loại hình nếu có thì ghi luôn.

Mức A-E không chặn GĐ2/GĐ3. Thiếu nguồn: ô trống, không điền 0; quy tắc ghi TREO hoặc KHÔNG_ĐỦ_DỮ_LIỆU.

Playbook: `skills/phan-tich-ho-so-thue/playbook-3-giai-doan.md`.

Workflow đang `status: active`. Chạy từng DN khi có folder tuyệt đối, MST và kỳ. Không tự chạy hàng loạt 13 DN.
