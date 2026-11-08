---
name: "Lập hồ sơ đánh giá xếp loại cấp ủy"
description: "Lập hồ sơ đánh giá xếp loại cấp ủy CB13 theo CV 4988: 8 sản phẩm, HTTNV/HTXSNV, mẫu in 05/06, dual save."
group: "Công tác Đảng"
origin: javis-learned
status: active
created: 2026-11-03
updated: 2026-11-03
---
## Khi nào dùng
- Lập hoặc nhân bản hồ sơ đánh giá xếp loại cấp ủy Chi bộ 13 theo Công văn 4988 (quý, 6 tháng, năm).
- User gửi quy chế, phân công chi ủy, rồi bảo chỉnh bộ đang có.
- Đổi mức đề xuất (HTTNV / HTXSNV), đổi kỳ, đổi nơi nộp, hoặc "chỉ sửa nội dung, giữ định dạng".
- User gửi file mẫu KH/TDG rồi bảo áp định dạng cho các file còn lại trong thư mục 05 và 06.
- Ví dụ: "lập hồ sơ quý IV", "dùng bộ quý III làm mẫu kỳ sau", "7 đồng chí hoàn thành tốt nhiệm vụ".

Việc lập file quản lý hồ sơ đảng viên (4 sheet) thuộc skill `lap-file-quan-ly-ho-so-dang-vien`. Soạn VB hành chính TCS3 thuộc skill `soan-thao-van-ban`.

## Chuẩn bị
- Bộ mẫu sống (copy cấu trúc, đổi kỳ): `bao-cao/danh-gia-quy-III-2026-CB13/`. Chi tiết 8 sản phẩm: `references/bo-8-san-pham-va-dinh-dang.md`.
- Mẫu in KH: `05_Ke_hoach_ca_nhan/KH_QIII-2026_1_Le-Huu-Nghia.docx`.
- Mẫu in TDG: `06_Ban_tu_danh_gia/TDG_QIII-2026_1_Le-Huu-Nghia.xlsx`.
- Script: `_scripts/` trong thư mục bộ mẫu. Kỳ sau clone định dạng 05/06 bằng `_scripts/ap_dinh_dang_mau_kh_tdg.py` (đổi đường dẫn mẫu cho khớp kỳ mới).
- Bản Laptop: `H:\Other computers\My Laptop\2. CONG TAC KHAC\4. CONG TAC DANG\A PHUONG\2026\06_Tieu_chi_danh_gia\Du_thao_CB13_QuyIII_2026\`. Kỳ mới tạo thư mục song song, không ghi đè quý cũ.
- Khung 30 + 70: file phụ lục cấp ủy trực thuộc trong `06_Tieu_chi_danh_gia`. KPI 4 trụ: `-Phụ lục 2A&2B.docx`. Số liệu ước tính tham khảo `Danh gia PCVP.xlsx`. Không bịa số thực tế của kỳ.
- Đọc quy chế / phân công nếu user vừa gửi. File `.doc` đọc bằng Word COM. Mặc định chỉ đọc, không ingest.
- Đọc `memory/facts/` về chi ủy, nơi nộp, mức xếp loại đã chốt.

## Cách chạy
Kỳ mới: copy cả thư mục bộ mẫu, đổi kỳ trong tên file và nội dung, giữ định dạng. Có bộ cũ rồi thì vá nội dung, không sinh lại từ `gen_docx.py` / `gen_xlsx.py`.

Mặc định **HTTNV**. HTXSNV chỉ khi user nói rõ. Ước tính tô đỏ `C00000`, giữ công thức sống.

## Quy trình
1. Chốt 4 tham số: kỳ đánh giá, nơi nộp, 7 đồng chí, mức đề xuất. Thiếu mức thì HTTNV.
2. Lấy phân công từ quy chế 05-QĐ/CB và file phân công chi ủy. Map chức danh cũ (Đảng ủy bộ phận, Bí thư Chi bộ 1-4) sang Chi bộ 13 / tổ đảng. Không chép chức danh cơ cấu cũ.
3. Cập nhật `_scripts/cb13_data.py`, `cb13_noidung.py`, `cb13_uoc_tinh.py` theo kỳ mới.
4. **Chỉ sửa nội dung, giữ định dạng:** `patch_noidung_httnv.py`. File 04 lấy khung từ mẫu Word (header 3 dòng, không gọi `table.columns`). 01/04/07: `_scripts/ap_nd30_font14_va_excel_fit.py`. 05/06: clone từ KH_1 / TDG_1 bằng `_scripts/ap_dinh_dang_mau_kh_tdg.py`.
5. Đủ 8 sản phẩm CV 4988, 7 file KH + 7 file TDG. Chi bộ 13 ban hành. Nơi nộp: Ban Thường vụ Đảng ủy Thuế thành phố Đồng Nai.
6. HTTNV: chỉ mức tốt 92,25% và đạt 86,50%, không dùng mức vượt 103,75%. HTXSNV: đủ 3/9 tiêu chí vượt mức. An có 7 nhiệm vụ (thêm cụ thể hóa NQ, sl 3, hs 1,0, 300 điểm).
7. Ghi song song vault và Laptop. Không ingest file đính kèm trừ khi user yêu cầu.
8. Trả lời: mức xếp loại, 7 đồng chí (lĩnh vực + điểm ước tính), file đã sửa, việc user phải điền minh chứng.

## Bẫy
- Không regenerate từ template khi user dặn giữ định dạng. Không gọi `table.columns` (dàn đều cột header). `wdPaperA4 = 7` (9 là A5). openpyxl A4 = 9.
- Không `FitToPages`: co chữ xuống 5-7 pt. KH cỡ bảng 11 pt, lề 1,27 cm là do mẫu user chốt, không phải để nhét 1 trang.
- Excel 06: clone `scale=63%`, không `fitToWidth=1`. Khi gắn gạch header phải có `xmlns:r` trên worksheet và file `.rels`, nếu không file hỏng XML.
- 02/03/08 vẫn `fitToWidth=1`. Đừng áp scale 63% cho cả pack Excel.
- Sheet `Tinh diem KPI` và `Tinh KPI 2A-2B` theo dõi nội bộ, không in, không nộp. Bỏ ghi chú hướng dẫn Javis khỏi bản in.
- Biểu 08: cột F = số HTTNV, cột G = số HTXSNV nằm trong HTTNV. Đề xuất 7 HTTNV thì F = 1/1/5, G = 0/0/0.
- Không bịa số liệu KPI thực tế. Ô đỏ là ước tính; HTTNV thì A thực tế không lớn hơn A kế hoạch.
- Không dùng ký tự em dash.

## Kiểm chứng
- 7/7 KH, 7/7 TDG, tờ trình, trích ngang, tổng hợp, quyết định, báo cáo, thống kê cùng mức xếp loại và cùng phân công.
- Excel: công thức KPI / tổng điểm còn sống; số đỏ A thực tế không lớn hơn A kế hoạch khi mức HTTNV.
- Word COM: 01/04/07 giấy 7 A4 đứng, lề 20/20/30/15 mm, chữ 14. KH: giấy 7 A4 ngang, lề 1,27 cm, bảng 11 pt, Số `-ĐU/KH`.
- TDG: `scale=63`, lề 0,2/0,2/0,25/0,25 inch, F4 Trấn Biên, hết ghi chú Javis cột I.
- Copy Laptop xong mới dẫn đường. Đếm file trong thư mục lúc trả lời.

## Lịch sử (tự học)
- [2026-11-03] User gửi Quy chế 05-QĐ/CB ngày 10/2/2026 và phân công chi ủy, chốt 7 đồng chí HTTNV, chỉ sửa nội dung. Chuẩn định dạng = bộ `bao-cao/danh-gia-quy-III-2026-CB13`.
- [2026-11-03] Hướng 1: cộng việc 7 của An (cụ thể hóa NQ) vào bản 04, giữ lĩnh vực QCƯT. Tách 7 file KH và 7 file TDG. Bỏ ghi chú hướng dẫn Javis.
- [2026-11-03] Vá python-docx làm dàn cột file 04. Dựng lại từ mẫu Word, header 3 dòng, Số 07-QĐ/CB, Trấn Biên.
- [2026-11-03] User chốt header 3 dòng, rồi trả cỡ 14 và lề NĐ 30 cho 01/04/07, bỏ FitToPages. 02/03/08 fit hết cột 1 trang in.
- [2026-11-03] User gửi mẫu KH_1 và TDG_1: KH lề 720 twips, bảng 11pt, Số `-ĐU/KH`; TDG scale 63%, lề hẹp, Trấn Biên. User yêu cầu lưu thành skill để kỳ sau áp tiếp.
