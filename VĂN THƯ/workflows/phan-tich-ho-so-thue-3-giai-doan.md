---
type: workflow
name: Phân tích hồ sơ thuế 3 giai đoạn
slug: phan-tich-ho-so-thue-3-giai-doan
status: active
group: Tài chính
description: Chuỗi KQ1-KQ2-KQ3 theo mức A-E, bộ mẫu v1.2, 11 trợ lý phòng phân tích.
steps:
  - agent: thu-ho-so-kiem-ke
    task: "P1. Đọc {{input}}: folder tuyệt đối, MST, tên DN, kỳ kiểm tra. Tổ chức cây 4 bước nếu chưa chuẩn. Quét IN00-IN14, checklist A4 duy nhất tại 4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx. Xuất KQ1_<MST>_<KY>_v<n>.xlsx (3 sheet). Kết luận mức A-E. Không phân tích rủi ro, không xóa nguồn."
  - agent: truong-phong-phan-tich-ho-so
    task: "Đọc KQ1 từ {{prev}}. Công bố mức A-E và lệnh tiếp: A/B thì P2 (KQ2); C/D/E thì vẫn chạy phần CQT để giữ Chắc rồi P3 (KQ3). Ghi rõ IN nào còn thiếu. Không tự kiểm kê lại."
  - agent: chuyen-vien-du-lieu-dau-vao
    task: "Gộp HĐĐT + bảng kê + NLĐ mẫu 02, CD1 BT5-BT9 theo folder trong {{input}} và mức/IN từ {{prev}}. Xuất 2 file trong cau-thanh/. Thiếu 1 nguồn thì sheet còn lại vẫn chạy, Tom_tat ghi Thiếu/Mỏng. Không trộn nguồn."
  - agent: chuyen-vien-gtgt-tndn
    task: "Đối chiếu bảng kê - tờ khai GTGT - QTT TNDN - HĐĐT (2 sheet CBRR) theo {{input}} và {{prev}}. Xuất Tong_hop_doi_chieu_GTGT_TNDN_<MST>.xlsx. Kỳ trống thì bỏ dòng, không điền 0 giả."
  - agent: chuyen-vien-ncc-hoa-don
    task: "Đối chiếu 100% NCC với CVCBRR F0/F1, Công an, danh bạ. F0 và F1 đều Cao (K2). HD07 TREO nếu chưa có tham chiếu ngành (K13). Xuất Tong_hop_NCC_mua_vao_CV_trang_thai.xlsx. Không rút gọn."
  - agent: chuyen-vien-tncn
    task: "Tổng hợp QTT TNCN, NLĐ 05-1/05-2, đa nguồn TC22/TC23 nếu có 3.3. Xuất Tong_hop_TNCN_va_Da_nguon_<MST>_<KY>.xlsx. Thiếu 3.3 thì vẫn lập QTT, ghi thiếu tra cứu đa nguồn."
  - agent: phan-tich-vien-cqt
    task: "P2. Chạy 57 quy tắc GĐ2 trên IN01-IN07/12/14 sau KQ1 và các file cau-thanh vừa có ({{prev}}). Xuất KQ2_<MST>_<KY>_v<n>.xlsx đủ 7 sheet. 23 quy tắc sổ ghi KHÔNG_ĐỦ_DỮ_LIỆU. Chưa có KQ1 thì dừng, yêu cầu P1."
  - agent: chuyen-vien-doanh-thu-gia-von
    task: "Đối chiếu sổ 511 vs 632 (K4) nếu KQ1 có IN13. Xuất Doi_chieu_511_632_<MST>_<KY>.xlsx. Không có sổ: ghi KHÔNG_ĐỦ_DỮ_LIỆU, không suy từng dòng từ tờ khai."
  - agent: chuyen-vien-kho-gia-thanh
    task: "XNT 152/155/156 vs BCTC (K6 [141]/ct140) và giá thành/định mức nếu có IN09. Thiếu sổ kho: ghi thiếu, đánh giá sơ bộ TC3 từ BCTC, không kết luận xuất khống. Cập nhật sheet Master Tong_hop_XNT_va_Kho và Gia_thanh_va_Dinh_muc."
  - agent: chuyen-vien-cong-no-tien-mat
    task: "Quét 111, cụm cùng ngày, 131/331, sao kê 112. Ngưỡng K5 dùng >= 20tr đến 30/06/2025 và >= 5tr từ 01/07/2025. PDF chưa OCR: PDF - chưa kết luận (K11). Thiếu IN10/IN11: ghi KHÔNG_ĐỦ_DỮ_LIỆU. Xuất Phan_tich_Tien_mat_va_Ngan_hang_<MST>_<KY>.xlsx."
  - agent: phan-tich-vien-so-sach-dn
    task: "P3 chỉ khi mức C/D/E. Chạy 23 quy tắc sổ, chuỗi 5 mắt xích, nâng sơ bộ GĐ2 thành chắc hoặc giải tỏa. Xuất KQ3_<MST>_<KY>_v<n>.xlsx. Mức A/B: không chạy, ghi bị chặn, trả lại trưởng phòng. Không ghi đè bản KQ cũ."
  - agent: truong-phong-phan-tich-ho-so
    task: "Chốt Master 8 sheet (28 tiêu chí, 7 cột; thiếu mẫu thì chỉ sheet 08 theo K3) và đề cương kiểm tra, cột Nhánh xử lý để trống. Tóm tắt 1 trang: mức hồ sơ, đường dẫn KQ1/KQ2/KQ3, 5-10 trọng tâm, hồ sơ còn thiếu. File tại 4. KET QUA PHAN TICH/. Font Times New Roman. Số tiền dấu chấm nghìn, phẩy lẻ."
updated: 2026-10-16
---
Chuỗi KQ1 rồi KQ2 rồi KQ3 cho một doanh nghiệp, theo bộ mẫu v1.2 (`Bo_bieu_mau_chuan_phan_tich_ho_so_thue_v2.xlsx`, 80 quy tắc, IN00-IN14, K1-K13).

Đầu vào bắt buộc trong {{input}}: folder tuyệt đối của DN, MST, kỳ kiểm tra. Tên DN và loại hình (thương mại/sản xuất) nếu có thì ghi luôn.

Nhánh A-E không có if/else trong YAML, nên mỗi bước tự đọc mức từ KQ1:

- Mức A/B: chạy đủ P2, bước sổ DN ghi KHÔNG_ĐỦ_DỮ_LIỆU hoặc bị chặn, không xuất KQ3.
- Mức C/D/E: chạy P2 để giữ phần Chắc từ CQT, rồi P3 xuất KQ3.
- Thiếu nguồn: không điền số 0 giả; ghi TREO hoặc KHÔNG_ĐỦ_DỮ_LIỆU.

Playbook: `skills/phan-tich-ho-so-thue/playbook-3-giai-doan.md`.

Workflow đang `status: active`. Chạy từng DN khi có folder tuyệt đối, MST và kỳ. Không tự chạy hàng loạt 13 DN.
