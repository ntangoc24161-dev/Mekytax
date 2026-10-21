---
type: decision
provenance: user
origin: javis-learned
created: 2026-10-19
updated: 2026-10-19
---
**Quyết định (2026-10-19):** Sổ theo dõi kiểm tra quyết toán (JAVIS PRO) chỉ giữ **1 file `.xlsm`** duy nhất (có macro VBA tự chuyển dòng giữa 2 sheet trạng thái), xoá file `.xlsx` không macro cùng các bản sao lưu kèm theo, đổi tên chuẩn thành `Danh sach kiem tra quyet toan pro.xlsm`.

**Vì sao:** File `.xlsx` trước đó chỉ tạo để dự phòng mở trên máy/phần mềm chặn macro, nhưng gây rủi ro 2 bản dữ liệu lệch nhau; mã VBA trong `.xlsm` không ảnh hưởng tới Word Mail Merge khi đọc trực tiếp dữ liệu từ file này.

**Áp dụng:** Khi thao tác với sổ theo dõi kiểm tra, chỉ dùng duy nhất file `Danh sach kiem tra quyet toan pro.xlsm` tại `G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\`; không tạo lại bản `.xlsx` song song nữa.
