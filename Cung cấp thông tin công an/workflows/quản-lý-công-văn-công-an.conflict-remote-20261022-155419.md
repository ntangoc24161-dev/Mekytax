---
type: workflow
name: Quản lý công văn công an
slug: quản-lý-công-văn-công-an
status: active
group: Tài chính
description: Tiếp nhận-phân tổ-ghi sổ rồi dự thảo công văn và tờ trình cho DN thuộc KT1-KT2-KT3.
steps:
  - agent: quản-lý-hồ-sơ-công-an
    task: "Tiếp nhận công văn từ CV MOI, iOffice hoặc Zalo. Bóc tách file gộp, trích xuất .md tại D:\\HSO CONG AN FILE MD, tra DB NNT.xls (cột Tên cơ quan thuế) và masothue.com. Cửa kiểm CQ quản lý trước: chỉ phân KT1-KT3 và gán CBQL KT1 khi thuộc TCS3 (Vĩnh Cửu=KT2); CQ khác ghi tên CQ vào ô Tổ KT, không đưa vào gói Soạn thảo. Cập nhật Bao_cao_cong_van.xlsx (NHAP_HO_SO), lưu PDF theo STT năm 2026. Lập gói bàn giao DN KT1+KT2+KT3. Không soạn Word. Đầu vào: {{input}}"
    max_retries: 1
  - agent: soan-thao-tra-loi-cong-an
    task: "Từ gói bàn giao bước trước ({{prev}}) và yêu cầu gốc ({{input}}), dự thảo đủ DN thuộc KT1+KT2+KT3: file TRA LOI CV <Số CV>_<CQ>_du thao.docx và to trinh <Số CV>.docx, lưu vào thư mục hồ sơ trên Drive. Đúng hướng kính gửi (Thuế TP nếu có DON-KTr1). Số liệu live từ DB NNT.xls; chỗ thiếu tô đỏ và nhãn [CHƯA XÁC ĐỊNH]. Không bịa nợ thuế/VPHC."
    verify_agent: soan-thao-tra-loi-cong-an
    max_retries: 1
updated: "2026-10-16"
---

Quy trình trục của phòng cung cấp hồ sơ Công an: **Điều phối** xong mới **Soạn thảo**.

Hai trợ lý hỗ trợ không nằm trong chuỗi mặc định, gọi khi cần:
- `khai-thac-du-lieu-ho-so` - Excel 24 cột, đối chiếu NCC/công nợ, hoặc soi sâu GĐ1-3 khi user yêu cầu.
- `phap-che-chinh-sach-thue` - tra điều khoản khi soạn thảo vướng căn cứ.

Không đưa ôn thi / khảo thí vào quy trình này.
