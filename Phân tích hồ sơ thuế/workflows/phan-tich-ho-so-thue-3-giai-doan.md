---
type: workflow
name: Phân tích hồ sơ thuế 3 giai đoạn
slug: phan-tich-ho-so-thue-3-giai-doan
status: active
group: Tài chính
description: "Chuỗi Lệnh 1-3 quy trình 1.4 vòng 2: KQ1 rồi KQ2.1/KQ2.2 rồi KQ3 + phụ lục GĐ3; phòng 6 vai."
steps:
  - agent: thu-ho-so-kiem-ke
    task: "Lệnh 1. Đọc {{input}}: folder tuyệt đối, MST, tên DN, kỳ kiểm tra, loại hình nếu có. Tổ chức cây 4 bước nếu chưa chuẩn. Sao mẫu KQ1.1 KQ1.2 KQ1.3 rồi điền IN00-IN05/IN08-IN14 (IN07 không ở KQ1.1; IN05 có Loai_dong; IN02 đủ mua+bán). Xuất KQ1.1, KQ1.2, KQ1.3 v1. Kết luận mức A-E. Không phân tích rủi ro, không xóa nguồn."
  - agent: truong-phong-phan-tich-ho-so
    task: "Đọc KQ1.3 từ {{prev}}. Công bố mức A-E. Mức không chặn GĐ2/GĐ3. Ghi rõ IN còn thiếu. Lệnh tiếp: chuẩn hóa dữ liệu đầu vào rồi Lệnh 2 (KQ2.1 + KQ2.2) rồi Lệnh 3. Không tự kiểm kê lại."
  - agent: chuyen-vien-du-lieu-dau-vao
    task: "Gộp HĐĐT + bảng kê + NLĐ + CD1 BT5-BT9 theo folder trong {{input}}. Ghi KQ1.1!IN01,IN05 và KQ1.2!IN02 (đủ mua vào và bán ra). Một phiên bản mới của tệp đích. Không xuất mẫu 02 hay Doi_chieu_BT5_BT9 để giao. Thiếu 1 nguồn: ô trống, không điền 0."
  - agent: phan-tich-vien-cqt
    task: "Lệnh 2 phần KQ2.1. Sao mẫu KQ2.1. Lập IN07; điền GTGT, NCC 100% (giữ HD04), TNCN, EBITDA; Tong_hop_co bằng công thức. Được phép nhận xét module 3.4. Không ghi mức rủi ro/điểm/đề cương toàn hồ sơ. Ô thiếu nguồn để trống. Một phiên bản mới của KQ2.1. Chưa có KQ1 thì dừng."
  - agent: phan-tich-vien-so-cty
    task: "Lệnh 2 phần KQ2.2. Sao mẫu KQ2.2. Điền KQ2_511_632 (tổng hợp năm + đánh giá tuân thủ) và KQ2_511_632_Chi_tiet (bán dưới giá); kho; giá thành; công nợ tiền mặt (K5, P23). Nhận xét module 3.4. Không có sổ: ô trống, không suy từ tờ khai. Tong_hop_co bằng công thức. Một phiên bản mới của KQ2.2."
  - agent: phan-tich-vien-so-sach-dn
    task: "Lệnh 3 luôn chạy sau Lệnh 2, kể cả mức A/B. Sao KQ3: sheet KQ3 80 quy tắc + sheet Danh_gia_28_tieu_chi_rui_ro. Thiếu nguồn: KHÔNG_ĐỦ_DỮ_LIỆU / Chưa kết luận, cột AA. Chống trùng T1-T10. Xuất thêm KQ3_Danh_gia_Dau_hieu_Bat_thuong và KQ3_NNT_Giai_trinh_Chi_tiet (10 sheet có số liệu thật). Áp Q14-Q20. Không ghi đè bản cũ."
  - agent: truong-phong-phan-tich-ho-so
    task: "Chốt: đường dẫn KQ1.1-KQ3 và 2 file phụ GĐ3, tổng cột W theo sắc thuế, 5-10 trọng tâm từ KQ3, hồ sơ còn thiếu. Sheet 28 tiêu chí nằm trong KQ3 là chuẩn. Từ chối sheet PL, workbook trọn gói, mẫu 01-07, Master 8 sheet tách ngoài. Font Times New Roman. Số tiền dấu chấm nghìn, phẩy lẻ."
updated: 2026-10-23
---
Chuỗi Lệnh 1 rồi Lệnh 2 rồi Lệnh 3 cho một doanh nghiệp, theo bộ mẫu v1.4 vòng 2.

Phòng làm việc gọn **6 vai**: Thủ hồ sơ GĐ1, Dữ liệu đầu vào, Phân tích viên KQ2.1 CQT, Phân tích viên KQ2.2 sổ công ty, Phân tích viên kết luận GĐ3, Trưởng phòng.

Đầu vào bắt buộc trong {{input}}: folder tuyệt đối của DN, MST, kỳ kiểm tra. Tên DN và loại hình nếu có thì ghi luôn.

Mức A-E không chặn GĐ2/GĐ3. Thiếu nguồn: ô trống, không điền 0; quy tắc ghi TREO hoặc KHÔNG_ĐỦ_DỮ_LIỆU. KQ2 được nhận xét module 3.4; kết luận mức/điểm chỉ ở KQ3.

Playbook: `skills/phan-tich-ho-so-thue/playbook-3-giai-doan.md`.

Workflow đang `status: active`. Chạy từng DN khi có folder tuyệt đối, MST và kỳ. Không tự chạy hàng loạt.
