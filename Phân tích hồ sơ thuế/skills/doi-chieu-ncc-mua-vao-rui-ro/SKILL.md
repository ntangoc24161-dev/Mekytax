---
name: Đối chiếu NCC mua vào với CVCBRR và danh bạ
description: Đối chiếu NCC bảng kê mua vào với sổ CVCBRR (F0, F1), sổ Công an và danh bạ NNT gom 1 sheet kèm chi tiết hóa đơn.
group: Tài chính
origin: javis-learned
status: active
created: 2026-09-30
updated: '2026-10-29'
learned_updated: '2026-10-29'
---
## Khi nào dùng

- Đối chiếu nhà cung cấp / bên bán trên bảng kê **mua vào** với Sổ CVCBRR (F0, F1), Sổ Công an (`Bao_cao_cong_van.xlsx`) và Danh bạ NNT.
- Kiểm tra công ty đang xét có nằm trong báo cáo công văn (F0, F1, sổ Công an) hay không.
- Xuất file Excel chuẩn 2 sheet: `Tong_hop_doi_chieu` + `Chi_tiet_hoa_don`.
- So khớp với bản chuẩn Vạn Phát Thành khi nghi file bị rút gọn.
- User báo danh bạ đã cập nhật và yêu cầu chạy lại.
- User hỏi vì sao một NCC bị xếp F0/Công an: mở lại file kết quả, dẫn đúng dòng, cột nguồn và số công văn.

## Chuẩn bị

1. **Bảng kê mua vào**: `Bang_ke_mua_vao_ban_ra_tong_hop_*.xlsx` (sheet `mua vao` đã gộp, đã lọc dòng Tổng cộng), hoặc file **IN02** (`IN02_BangKe_*.xlsx`) nếu hồ sơ đã tổng hợp theo mẫu IN02. Với IN02 chỉ lấy dòng có `Tinh_vao_tong` = Có.
2. **Sổ CVCBRR**: `SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx` (F0 bên bán, F1 bên mua).
3. **Sổ Công an**: `Bao_cao_cong_van.xlsx`.
4. **Danh bạ NNT** (thư mục 3.2): trạng thái ĐKKD, ngày đóng MST và ngày tạm ngừng.
5. Nếu thư mục 3.1 của hồ sơ trống, dùng sổ CVCBRR và sổ Công an dùng chung trên Drive của user, rồi ghi rõ nguồn trong báo cáo.

## Cách chạy / Quy trình xuất

File Excel gồm 2 sheet, font **Times New Roman**:

### 1. Sheet `Tong_hop_doi_chieu`
- **Phần 1:** Đối chiếu NNT đang xét với F0, F1 và Sổ Công an, tra theo cả MST và tên; ghi kết luận an toàn/có cảnh báo. Nếu danh bạ không có MST của chính NNT, ghi "cần tra riêng trạng thái ĐKKD".
- **Phần 2:**
  - **Bảng 2.1 (F0):** NCC dính công văn cảnh báo (khoảng 18 cột: Số dòng BK, Năm phát sinh, Từ/Đến ngày HĐ, Mặt hàng chính, Tình trạng F0, Số CV, Ngày CV, Đơn vị ban hành, Nội dung cảnh báo, Trạng thái ĐKKD...).
  - **Bảng 2.2:** Liệt kê **100% NCC** trên bảng kê mua vào (bình thường + bất thường), có dòng **Tổng cộng toàn bộ**; đánh dấu nhóm trạng thái ĐKKD bất thường. **Cấm** chỉ giữ nhóm bất thường.
  - Phân nhóm NCC: A dính cảnh báo; B ĐKKD bất thường; C là bên mua trong F1; D không có MST (cá nhân); E **chưa có dữ liệu danh bạ**; G đang hoạt động, không cảnh báo.
  - Nhóm A phải ghi **nguồn cụ thể từng NCC** ("F0 - sổ CVCBRR" hoặc "Công an - sổ Công an", kèm số và ngày công văn). Nhãn nhóm không được gộp kiểu "F0/Công an"; nếu cần tách thành A1 - F0 và A2 - Công an.
- **Phần 3 (bắt buộc):** (1) Đánh giá rủi ro NNT, (2) Đánh giá rủi ro NCC theo nhóm, (3) Đề xuất trọng tâm kiểm tra chứng từ, (4) Hạn chế dữ liệu (độ phủ danh bạ, nguồn sổ đã dùng).

### 2. Sheet `Chi_tiet_hoa_don`
- Toàn bộ dòng HĐ mua vào trong kỳ, sắp **`Ngay_HD` tăng dần**.
- Đủ khoảng **20 cột**, gồm nhóm chứng từ kế toán (`Loại chứng từ`, `Số chứng từ`, `Ngày chứng từ`), nhóm F0 (`Cảnh báo F0`, `Số CV F0`, `Đơn vị ban hành`, `Nội dung cảnh báo`) và cột `Trạng thái hoạt động ĐKKD` riêng (không gộp nhầm vào ghi chú số tiền).
- Đầu vào là IN02 hoặc bảng kê thì không có sổ chứng từ. Khi đó để trống nhóm chứng từ và ghi rõ lý do trong Phần 3, không bịa số chứng từ.

### 3. Khi user yêu cầu chạy lại với danh bạ mới
- Trước khi chạy, so file danh bạ trên Drive với bản đã dùng (thời gian sửa, md5). Giống hệt thì báo user là chưa có bản mới, đừng chạy lại ra kết quả cũ.
- Có bản mới thì chạy lại, ghi đè file kết quả và báo bảng so sánh từng nhóm A-G trước/sau, kèm độ phủ danh bạ mới (số MST tra được).

## Bẫy

- **Bẫy rút gọn:** chỉ xuất NCC bất thường / cắt Phần 3 / thiếu cột chứng từ-F0 -> mất NCC chủ lực (ví dụ T&T từng thiếu 108 NCC ~99,5% doanh số).
- Nhầm F1 (bên mua) với F0 (bên bán). Một NCC vừa có ĐKKD bất thường vừa nằm trong F1 (vd Điện Lực Trảng Bom ở hồ sơ PHL) phải nêu cả hai.
- **Gộp F0 với Công an trong một nhãn:** nhãn "A - Có công văn cảnh báo (F0/Công an)" khiến người đọc hiểu NCC có công văn F0 trong khi thực tế chỉ có trong sổ Công an (vd Toàn Việt 3601946862 ở hồ sơ PHL, CV 2808/YC-CSKT ngày 03/04/2026 của Cơ quan CSĐT Công an tỉnh Đồng Nai). Khi trả lời user, nói đúng nguồn từng NCC; nếu chưa đối chiếu hết sổ CVCBRR thì nói rõ là chưa loại trừ F0.
- Khớp sót MST có/không gạch nhánh (chuẩn hóa 10 và 13 số).
- **Danh bạ không phủ hết NCC:** NCC không tra được thì xếp nhóm E "chưa có dữ liệu", **không** coi là đang hoạt động. Ví dụ PHL: danh bạ cũ chỉ có 86 MST, 166/198 NCC (104,76 tỷ) chưa rõ trạng thái; danh bạ mới phủ 193/197 thì nhóm E còn 4 NCC.
- **Không dùng cột trạng thái trong file tổng hợp đối tác của NNT** (vd `Tong_hop_danh_sach_doi_tac...` ghi "Đang hoạt động (Khớp danh bạ)") vì có thể mâu thuẫn với danh bạ thật. Luôn lấy trạng thái từ danh bạ gốc.
- **HĐ lập sau ngày đóng/ngừng:** so `Ngay_HD` với ngày đóng trong danh bạ và liệt kê riêng các HĐ lập sau ngày đó.
- **Ngày đóng MST khác ngày tạm ngừng:** danh bạ có thể cho cả ngày đóng MST lẫn ngày bắt đầu tạm ngừng. Đừng gộp vào một cột "Ngày đóng/ngừng" rồi chỉ kiểm theo ngày đóng: HĐ lập trong thời gian tạm ngừng sẽ bị sót. Tách 2 cột và kiểm cả hai, hoặc ghi rõ trong Phần 3 là chưa kiểm phần tạm ngừng.
- **Ngày đóng đổi giữa các bản danh bạ:** ví dụ Điện Lực Trảng Bom từ 25/06/2025 sang 17/09/2025 làm số HĐ sau ngày đóng từ 2 về 0. Luôn ghi bản danh bạ đã dùng.
- **NCC cá nhân không MST:** sổ Công an tra theo MST nên không khớp tự động; phải dò thêm theo tên và ghi "cần xác minh".
- Sort `Chi_tiet` theo MST thay vì `Ngay_HD`.
- Ô text bắt đầu bằng `=` gây lỗi Removed Records Formula.

## Kiểm chứng

- Tổng doanh số mọi NCC trên Bảng 2.2 = tổng bảng kê/IN02 mua vào sau lọc.
- Tổng số NCC các nhóm A-G = tổng NCC trên Bảng 2.2.
- Có đủ Phần 1 + 2.1 + 2.2 (100% NCC) + Phần 3 (kèm hạn chế dữ liệu).
- Mỗi NCC nhóm A có cột nguồn ghi rõ F0 hay Công an, kèm số/ngày CV; không còn nhãn gộp "F0/Công an".
- `Chi_tiet_hoa_don` đủ dòng HĐ, đủ cột chuẩn, font Times New Roman.
- Phần kiểm HĐ sau ngày đóng ghi rõ đã kiểm ngày đóng MST, ngày tạm ngừng hay cả hai.

## Lịch sử (tự học)
- [2026-10-05] User chỉ ra file Tong_hop_NCC của T&T thiếu thông tin so với Vạn Phát Thành; skill cũ mô tả Bảng 2.2 như chỉ liệt kê NCC bất thường, dễ dẫn tới bản rút gọn. Cập nhật bắt buộc 100% NCC, Phần 3 và đủ cột Chi_tiet.
- [2026-10-22] Hồ sơ Phát Hoàng Long: đầu vào là IN02, danh bạ chỉ phủ 86 MST, file tổng hợp đối tác 3.2 ghi trạng thái mâu thuẫn danh bạ, có NCC cá nhân không MST trong sổ Công an, có HĐ lập sau ngày đóng. Bổ sung nhóm E, bẫy độ phủ danh bạ và cách xử lý IN02.
- [2026-10-22] Khi chạy cho Phát Hoàng Long (22/10/2026), skill cũ có 4 chỗ thiếu. (1) Chỉ nêu đầu vào là Bang_ke tổng hợp, chưa nêu IN02, mà IN02 không có cột chứng từ. (2) Không có nhóm 'chưa có dữ liệu danh bạ', trong khi danh bạ chỉ phủ 86 MST nên 166 NCC (104,76 tỷ) dễ bị xếp nhầm vào nhóm bình thường. (3) Kh
- [2026-10-22] Chạy lại PHL với danh bạ mới: cột "Ngày đóng/ngừng" gộp ngày đóng MST với ngày tạm ngừng nên phép kiểm chỉ tính ngày đóng. Lần yêu cầu chạy lại đầu tiên, danh bạ trên Drive giống hệt bản cũ (cùng md5). Bổ sung bước kiểm bản danh bạ trước khi chạy lại và bẫy tách ngày tạm ngừng.
- [2026-10-22] Khi chạy lại cho Phát Hoàng Long ngày 22/10/2026 lộ ra 2 chỗ thiếu. (1) Cột ngày đóng/ngừng gộp ngày đóng MST với ngày tạm ngừng, và phép kiểm 'HĐ sau ngày đóng' chỉ tính ngày đóng, nên HĐ lập trong thời gian tạm ngừng bị sót. Javis phải tự ghi đây là hạn chế. (2) Lần đầu user bảo chạy lại, danh bạ
- [2026-10-29] User hỏi vì sao Toàn Việt (hồ sơ PHL) bị xác định có công văn F0. Thực tế nguồn là sổ Công an; nhãn nhóm A "F0/Công an" gây hiểu nhầm. Bổ sung quy tắc ghi nguồn từng NCC nhóm A, tách nhãn F0 và Công an, kèm bước kiểm chứng.
- [2026-10-29] Ngày 29/10/2026 user hỏi vì sao Toàn Việt 3601946862 trong hồ sơ Phát Hoàng Long bị xác định có công văn F0. Javis mở lại file thì thấy nguồn là sổ Công an (CV 2808/YC-CSKT), nhưng nhãn nhóm A 'A - Có công văn cảnh báo (F0/Công an)' gộp hai nguồn nên gây hiểu nhầm. Skill cũ chỉ ghi 'A dính cảnh báo
