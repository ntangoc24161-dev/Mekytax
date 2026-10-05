---
type: reference
provenance: assistant
origin: javis-learned
created: 2026-09-24
updated: 2026-09-24
---
Tiến trình thu thập và giải đề kíp P3-39917 kết nối HTTP trực tiếp đến taxedu.gdt.gov.vn, giải mã gói tin cục bộ qua `decrypt_dcv2.js`, đối soát với ngân hàng 1.247 câu và solver quy tắc `tax_solver_engine.py`. Quá trình này không gọi API LLM trả phí nên không bị giới hạn bởi hạn mức token AI, có thể chạy đủ 40 đề mà không phát sinh chi phí.
