---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-19
updated: 2026-10-19
---
Trước khi nhập liệu và xuất hồ sơ đi, trong thư mục công ty (ví dụ `1. TD\<Tên Cty>\`) phải tạo thư mục `0. BIEN BAN\`. File `Bang tinh.xlsx` (18 sheet: KQKDSXKD, DTCHEP, KQKD, CĐKT, GTGT...) được khởi tạo bằng cách nhân bản từ file mẫu chung tại `1. TD\BIEN BAN\Bang tinh.xlsx`, sau đó mới chạy các bước nhập dữ liệu và bóc tách.

**Ai làm:** Trợ lý khởi tạo hồ sơ (`khoi-tao-ho-so`). Không giao cho agent Lập hồ sơ kiểm tra (agent đó chỉ xuất Word/app).

**Vì sao:** Bước này mở đường cho `nhap-bang-tinh-tu-bao-cao-hdkd` và `boc-tach-bang-tinh-phan-tich`, thuộc chuỗi khởi tạo dữ liệu trước soạn thảo.

**Áp dụng:** Mỗi vụ mới: tạo `0. BIEN BAN` → copy `Bang tinh.xlsx` từ mẫu chung → mới nhập số liệu / xuất hồ sơ đi. Đã ghi vào mục Chuẩn bị của skill `lap-ho-so-kiem-tra` và bước 2 của agent `khoi-tao-ho-so`.
