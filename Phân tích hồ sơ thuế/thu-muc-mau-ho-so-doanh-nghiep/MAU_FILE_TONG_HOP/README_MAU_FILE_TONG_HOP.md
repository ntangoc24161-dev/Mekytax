# Bộ mẫu phân tích hồ sơ thuế (v1.4, P7 ngày 19/10/2026)

Sáu tệp giao theo hồ sơ. Sao chép rồi điền. Không dựng lại cấu trúc, không đổi tên sheet/cột.

Quy trình: `skills/phan-tich-ho-so-thue/references/Quy_trinh_3_giai_doan_va_mau_bieu_chot_v1.4.md`.
Prompt: `skills/phan-tich-ho-so-thue/references/Prompt_P7_AI_chi_dung_bo_quy_trinh_1.4.md`.

## 1. Sáu tệp kết quả (giao)

| Giai đoạn | Tệp mẫu | Sheet |
|---|---|---|
| GĐ1 | `KQ1.1_Kiem_ke_CQT.xlsx` | IN00, IN01, IN03, IN04, IN05 (Loai_dong), IN14 |
| GĐ1 | `KQ1.2_Kiem_ke_ho_so_NNT.xlsx` | IN00, IN02, IN08-IN13 |
| GĐ1 | `KQ1.3_Danh_gia_ho_so.xlsx` | KQ1_Kiem_ke, KQ1_Kha_nang_phan_tich, KQ1_De_nghi_bo_sung |
| GĐ2 | `KQ2.1_Phan_tich_CQT.xlsx` | KQ2_Tong_hop_co, KQ2_Doi_chieu_GTGT, KQ2_NCC_mua_vao, KQ2_TNCN_LD2, IN07_NCC_CanhBao |
| GĐ2 | `KQ2.2_Phan_tich_ho_so_cty.xlsx` | KQ2_Tong_hop_co, KQ2_511_632, KQ2_Kho_Gia_thanh, KQ2_Gia_thanh_va_Dinh_muc, KQ2_Cong_no_Tien_mat |
| GĐ3 | `KQ3_Ket_luan_danh_gia.xlsx` | một sheet KQ3, 39 cột A-AM |

KQ2 chỉ số, chênh lệch, tỷ lệ, cờ máy. Kết luận chỉ ở KQ3.

## 2. Tra cứu nội bộ (không giao)

`Bo_bieu_mau_chuan_phan_tich_ho_so_thue.xlsx` (bản 1.4): sheet 04 (80 quy tắc), 05 (tham số, P23 = 10%), 06 (mã BCTC), 08 (28 tiêu chí).

## 3. Mẫu cũ (chỉ nguồn nhập, T8)

Thư mục `cu-v1.2/`: `KQ1_MAU_Kiem_ke_danh_gia.xlsx`, `KQ2_MAU_Phan_tich_CQT.xlsx`, `KQ3_MAU_Hoan_chinh_ho_so.xlsx`, `Bo_bieu_mau_chuan_phan_tich_ho_so_thue_v2.xlsx`, bộ 01-07. Đọc một lần, ánh xạ vào IN, ghi IN00. Không tính lại, không xuất theo mẫu cũ.

## 4. Nguyên tắc

1. Times New Roman mọi ô. Giữ định dạng có điều kiện (đỏ lệch/vượt/F0/F1, cam chưa đủ nguồn, xanh khớp).
2. Ô công ty trống khi chưa có, không điền 0 (T10).
3. Không xuất sheet PL, workbook trọn gói, tệp chuyên đề, Master 8 sheet (trừ khi cán bộ yêu cầu, sinh từ KQ3).
