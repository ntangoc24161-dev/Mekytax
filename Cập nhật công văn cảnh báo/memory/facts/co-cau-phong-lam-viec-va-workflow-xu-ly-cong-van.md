---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-16
updated: 2026-10-16
---
Người dùng duyệt phòng làm việc công văn cảnh báo rồi điều chỉnh lần 2 (2026-10-16):

**Trợ lý tại brain Cập nhật công văn cảnh báo:**
- Tra cứu và tải công văn iOffice (`tra-cuu-tai-cv-ioffice.md`) - việc 1/3 thu thập.
- Số hóa OCR công văn (`so-hoa-ocr-cong-van.md`) - việc 2/3 thu thập.
- Cập nhật công văn cảnh báo rủi ro (`cập-nhật-công-văn-cảnh-báo-rủi-ro.md`) - việc 3/3, chỉ bóc tách nạp F0/F1.
- Đối chiếu rủi ro NCC (`doi-chieu-rui-ro-ncc.md`).
- Soạn thảo văn bản (`soan-thao-van-ban.md`).

**Workflow:** `workflows/day-chuyen-xu-ly-cv-canh-bao.md` gồm 5 bước (status: active): iOffice → OCR → nạp sổ F0/F1 → đối chiếu NCC → soạn thảo. Pháp chế gọi riêng ở brain Cung cấp thông tin công an.

**Chuyển sang brain Cung cấp thông tin công an:** trợ lý Pháp chế (`phap-che-chinh-sach-thue`) giữ 8 skill `thue-gtgt`, `thue-tndn`, `thue-tncn`, `hoa-don-dien-tu`, `xu-phat-vphc-thue-hoa-don`, `quy-trinh-kiem-tra-thue`, `quy-che-phoi-hop-kiem-tra-thue`, `luat-phong-chong-tham-nhung`.

**Loại trừ:** Phân tích hồ sơ thuế, Ôn thi công chức.

**Vì sao:** Một trợ lý thu thập ôm iOffice + OCR + nạp sổ thì khó giao từng khâu. Pháp chế dùng chung cho phúc đáp Công an nên đặt ở brain đó.

**Áp dụng:** Giao từng việc 1/2/3 khi chỉ cần một khâu. Cần dẫn chiếu luật thì đổi sang brain Cung cấp thông tin công an.
