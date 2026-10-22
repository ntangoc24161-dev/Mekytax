---
type: workflow
name: Dây chuyền xử lý công văn cảnh báo
slug: day-chuyen-xu-ly-cv-canh-bao
status: active
group: Tài chính
description: Tải iOffice, OCR, nạp sổ F0/F1, đối chiếu NCC rồi soạn văn bản cảnh báo.
steps:
  - agent: tra-cuu-tai-cv-ioffice
    task: "Việc 1/3 thu thập. Từ {{input}}: tra VNPT iOffice Đồng Nai theo trích yếu, số CV hoặc khoảng ngày; tải tệp đính kèm về D:\\HO SO CONG VAN CANH BAO (theo năm-tháng); giải nén zip. Báo danh sách Số CV, trích yếu, đường dẫn tệp. Không OCR, không nạp sổ."
    max_retries: 1
  - agent: so-hoa-ocr-cong-van
    task: "Việc 2/3 thu thập. Từ {{prev}} và {{input}}: chuyển PDF hoặc ảnh scan thành .md bằng pdf-sang-md (OCR offline nếu scan). Tách nhiều công văn trong một tệp. Gắn cờ PDF-da-trich hoặc PDF-chua-trich. Bàn giao đường dẫn .md. Không nạp sổ."
    max_retries: 1
  - agent: cập-nhật-công-văn-cảnh-báo-rủi-ro
    task: "Việc 3/3 thu thập. Từ {{prev}}: bóc tách trường chuẩn, lọc trùng bộ ba Số CV + tên bên bán + tên bên mua, nạp dòng mới vào F0 hoặc F1 của SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx, sao lưu trước khi ghi. Báo số nạp F0/F1 và số bỏ trùng."
    max_retries: 1
  - agent: doi-chieu-rui-ro-ncc
    task: "Nhận {{prev}}. Đối chiếu NCC mua vào (Bảng kê và HĐĐT nếu có trong {{input}}) với F0/F1, sổ Công an và danh bạ NNT. Xuất đủ Phần 1, Bảng 2.1, Bảng 2.2 (100% NCC), Bảng 2.3 theo kỳ, Phần 3. Ghi nguồn Bảng kê/HĐĐT/Chung."
    max_retries: 1
  - agent: soan-thao-van-ban
    task: "Dùng {{input}} và {{prev}} soạn văn bản đúng NĐ 30 và ký hiệu TCS3. Vụ Công an: TRA LOI CV <Số CV>_<CQ>_du thao.docx và to trinh <Số CV>.docx. Cần dẫn chiếu điều khoản thì ghi chỗ trống và chuyển brain Cung cấp thông tin công an, trợ lý phap-che-chinh-sach-thue. Không bịa số liệu."
    max_retries: 1
updated: "2026-10-16"
---

Quy trình trục của phòng công văn cảnh báo: **tải iOffice** rồi **OCR**, **nạp sổ F0/F1**, **đối chiếu NCC**, cuối cùng **soạn thảo**.

Pháp chế không nằm trong chuỗi này. Căn cứ luật, nghị định, mức phạt gọi trợ lý `phap-che-chinh-sach-thue` ở brain **Cung cấp thông tin công an**.

Giao từng việc 1/2/3 khi chỉ cần một khâu (ví dụ chỉ OCR lô scan, hoặc chỉ nạp sổ từ file `.md` có sẵn). Không đưa Phân tích hồ sơ thuế hay Ôn thi công chức vào quy trình này.
