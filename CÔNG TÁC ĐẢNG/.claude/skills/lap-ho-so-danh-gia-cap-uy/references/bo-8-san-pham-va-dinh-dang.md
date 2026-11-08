# Bộ 8 sản phẩm CV 4988 và định dạng chuẩn CB13

Mẫu sống: `bao-cao/danh-gia-quy-III-2026-CB13/`. Kỳ sau copy cấu trúc, đổi kỳ trong tên file và nội dung. Không FitToPages. Không gọi `table.columns`.

## 8 sản phẩm nộp (21 file in)

| Mã | Sản phẩm | File | Khi in |
|---|---|---|---|
| A1 | Tờ trình đề xuất xếp loại | `01_To_trinh_de_xuat_xep_loai.docx` | Cả file |
| A2 | Danh sách trích ngang | `02_Danh_sach_trich_ngang.xlsx` | Sheet `Trich ngang`, ngang |
| A3 | Biểu tổng hợp đề xuất | `03_Bang_tong_hop_de_xuat.xlsx` | Chỉ sheet `Tong hop de xuat` |
| A4 | Quyết định phân công + danh mục | `04_Quyet_dinh_phan_cong_va_danh_muc_san_pham.docx` | Cả phụ lục |
| A5 | Kế hoạch công tác quý cá nhân | `05_Ke_hoach_ca_nhan/KH_<ky>_<stt>_<Ten>.docx` | 7 file |
| A6 | Bản tự đánh giá 30 + 70 | `06_Ban_tu_danh_gia/TDG_<ky>_<stt>_<Ten>.xlsx` | Chỉ sheet `Khung đánh giá` |
| B1 | Báo cáo kết quả | `07_Bao_cao_ket_qua_<ky>.docx` | Cả file |
| B2 | Biểu thống kê cấp ủy viên | `08_Thong_ke_cap_uy_vien.xlsx` | Sheet `Thong ke`, ngang |

## Không in, không nộp

- `00_CHECKLIST_HO_SO.md`, `_scripts/`
- Sheet `Tinh diem KPI` trong file 03
- Sheet `Tinh KPI 2A-2B` trong 7 file TDG
- Không để ghi chú hướng dẫn (LƯU Ý, "thay bằng minh chứng", "Javis ước tính") trong 8 sản phẩm nộp.

## Định dạng bắt buộc (chốt 03/11/2026)

Chung: Times New Roman (kể cả `w:eastAsia`). Header Word 3 dòng, bảng 3 hàng × 2 cột, copy từ mẫu, không gọi `table.columns`. Trái: Đảng bộ / CHI BỘ 13 / `Số:`. Phải: ĐẢNG CỘNG SẢN VIỆT NAM / gạch dưới / **Trấn Biên**, ngày ..... tháng ..... năm. Không ghi Đồng Nai ở dòng ngày. `wdPaperA4 = 7`. openpyxl A4 = 9. Văn bản Chi bộ 13 ban hành. Nơi nhận: Ban Thường vụ Đảng ủy Thuế thành phố Đồng Nai. Ước tính màu `C00000`. KPI: 8 ô nhập A/B/C/D KH+TT, `AVERAGE` ra KPI%, `IFERROR` ra điểm đạt.

Ba lớp in, không trộn:

1. **01 / 04 / 07** (Word đứng): NĐ 30, lề 20/20/30/15 mm, nội dung cỡ 14, tên cơ quan header cỡ 13. Script `_scripts/ap_nd30_font14_va_excel_fit.py`.
2. **05 KH** (Word ngang, 7 file): mẫu `KH_QIII-2026_1_Le-Huu-Nghia.docx`. Lề 720 twips (1,27 cm) bốn phía, bảng cỡ 11 pt, tblW 15538, Số `-ĐU/KH`, bỏ dòng `-----` và dòng kèm Quyết định. Script `_scripts/ap_dinh_dang_mau_kh_tdg.py`. Word COM kỳ III đo 2 trang.
3. **06 TDG** (Excel ngang, 7 file): mẫu `TDG_QIII-2026_1_Le-Huu-Nghia.xlsx`. `scale=63`, `fitToHeight=0`, không `fitToWidth`. Lề 0,2/0,2/0,25/0,25 inch, zoom 70, F4 Trấn Biên, gạch ngang dưới tên chi bộ (`xmlns:r` bắt buộc). Chỉ in sheet `Khung đánh giá` vùng `A1:I61`.
4. **02 / 03 / 08** (Excel): `fitToWidth=1`, không copy scale 63% từ TDG.

## Phân công chuẩn (quy chế 05-QĐ/CB ngày 10/2/2026 + file phân công chi ủy)

Map sang Chi bộ 13 hiện nay (không dùng Bí thư Chi bộ 1-4):

1. Lê Hữu Nghĩa - Bí thư, phụ trách chung.
2. Trần Ngọc Luận - Phó Bí thư, tuyên giáo; điều hành khi Bí thư vắng; chỉ đạo dân vận, thi đua, CNTT, tài chính Đảng, thông tin báo cáo (theo quy chế PBT).
3. Nguyễn Khắc Trung - Chi ủy viên, Tổ trưởng Tổ đảng 1, kiểm tra giám sát; theo dõi sinh hoạt tổ đảng.
4. Nguyễn Đăng Phương - Chi ủy viên, Tổ trưởng Tổ đảng 2, đảng vụ, hồ sơ đảng viên, phát triển đảng viên.
5. Lê Thị Lan - Chi ủy viên, tài chính Đảng, đảng phí, tài sản; kiêm biên bản họp Chi ủy.
6. Phạm Đồng Xuân - Chi ủy viên, dân vận và Đoàn Thanh niên.
7. Nguyễn Thị Lệ An - Chi ủy viên, công tác Đảng, theo dõi quần chúng ưu tú, tạo nguồn. Phụ lục An có 7 nhiệm vụ (thêm: cụ thể hóa, triển khai nghị quyết theo lĩnh vực; sl 3, hs 1,0, 300 điểm).

## Mức xếp loại

- HTTNV (mặc định): hoàn thành 100% nhiệm vụ, không đặt tiêu chí vượt mức. Ước tính `_GOC` chỉ TOT 92,25% và DAT 86,50%.
- HTXSNV: chỉ khi user chốt. Cần 3/9 tiêu chí vượt mức (KPI 103,75%). Biểu 08: cột G ghi số HTXSNV.

## Kỳ sau làm gì

1. Copy thư mục bộ mẫu thành `danh-gia-<ky>-CB13`.
2. Đổi kỳ trong tên file, header, tiêu đề, thời hạn.
3. Vá nội dung 7 đồng chí; giữ định dạng bằng script trên, lấy KH_1 / TDG_1 của bộ vừa copy làm mẫu.
4. Dual save vault + thư mục Laptop kỳ mới.
5. User điền minh chứng, số/ngày văn bản. Ô đỏ chỉ là ước tính.
