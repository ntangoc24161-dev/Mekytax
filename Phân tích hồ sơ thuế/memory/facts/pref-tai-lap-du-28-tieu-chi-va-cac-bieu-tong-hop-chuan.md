---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-05
updated: 2026-10-05
---
Người dùng yêu cầu tuyệt đối không rút gọn sheet 28 tiêu chí rủi ro hay các sheet tổng hợp chuyên đề trong Master Excel phân tích hồ sơ doanh nghiệp:

1. Sheet `Danh_gia_28_tieu_chi_rui_ro` phải lập đủ **28/28 tiêu chí (TC1 -> TC28)** chuẩn **7 cột** (`ID`, `Tên tiêu chí`, `Phân loại`, `Ngưỡng rủi ro`, `Kết quả phân tích trên hồ sơ DN`, `Mức độ rủi ro`, `Kiến nghị kiểm tra`), có phân hóa rủi ro và kiến nghị chi tiết như bản chuẩn Vạn Phát Thành, không dùng bản rút gọn của tool v1.
2. Các sheet chuyên đề nội bộ (`Tong_hop_GTGT_TNDN`, `Tong_hop_NCC_CVCBRR`, `Tong_hop_TNCN_va_Da_nguon`, `Tong_hop_XNT_va_Kho`, `Gia_thanh_va_Dinh_muc`, `Doi_chieu_511_632`, `Cong_no_va_Ngan_hang`) phải bóc tách dữ liệu thực tế từ XML/sổ chi tiết, không để biểu mẫu sơ bộ giá trị 0.
3. Master Excel đủ **8 sheet** chuẩn; đồng bộ `4. KET QUA PHAN TICH/`, `cau-thanh/`, `7. Sao luu tong hop/`, `ket-qua/`, `sao-luu/`.
4. Áp dụng cho T&T Biên Hòa và **tất cả doanh nghiệp kiểm tra sau này**.

**Vì sao:** Bản rút gọn làm mất bức tranh rủi ro và không đủ để lập đề cương kiểm tra tại trụ sở.
**Áp dụng:** Mỗi lần lập/cập nhật Master Excel phải đối chiếu cấu trúc với bản Vạn Phát Thành trước khi chốt.

### Bài học kỹ thuật bắt buộc khi tái lập
- BCTC mẫu DNN: hàng tồn kho thường ở **ct140**, không phải ct141 (ct141 có thể = 0 dù DN có tồn kho).
- Một số file BCTC XML gắn nhãn `kyKKhai` lệch 1 năm so với niên độ doanh thu; phải **map theo doanh thu** khớp GTGT/TNDN, không tin nhãn năm thô.
- DN thương mại: sheet `Gia_thanh_va_Dinh_muc` chuyển sang phân tích GV 632 + chi phí BH/QL (không bắt buộc 621/622/627/154).
- Xuất kho sổ vật tư TK 156 phải khớp GV 632 trước khi kết luận TC28/TC3.
