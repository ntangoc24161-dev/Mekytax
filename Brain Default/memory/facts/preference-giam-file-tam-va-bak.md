---
type: preference
provenance: user
origin: javis-learned
created: 2026-10-07
updated: 2026-10-07
---

Người dùng yêu cầu kiểm soát chặt chẽ và không để phát sinh các file rác, file `.bak`, file `.json` tạm hoặc file `.md` trung gian rải rác trong vault và trên máy tính cá nhân.

**Vì sao:**
Quá trình chạy lệnh, đọc tài liệu PDF/OCR, backup và tạo script dễ sinh ra các file `.bak`, `.json`, `.tmp`, `.lock` hoặc script một lần nằm vương vãi (đặc biệt tại thư mục gốc vault), làm rối mắt và sai lệch cấu trúc dữ liệu người dùng.

**Áp dụng cho tất cả phiên làm việc:**
1. **Không tạo file nháp / script tại thư mục gốc vault:** Mọi script tạm thời, dữ liệu trung gian phải nằm trong `D:\tmp` hoặc thư mục scratch/con của skill, không ghi vào root `Brain Default`.
2. **Tự động dọn sạch sau khi hoàn thành:** Mọi file markdown trung gian (convert để đọc OCR), file txt/json nháp, lock file sau khi xử lý xong phải xóa bỏ ngay trong cùng phiên làm việc.
3. **Quản lý file Session / Cache đúng chỗ:** Các file cookie, session (như `ioffice_session.json`) phải đặt bên trong thư mục skill tương ứng hoặc `.javis/`, không để ở root.
4. **Không tạo `.bak` bừa bãi:** Không tự ý tạo các bản sao lưu `.bak` cạnh file gốc; nếu cần lưu trữ lịch sử thì gom vào thư mục chuyên biệt `sao-luu/` hoặc `D:\tmp\backup\`.
5. **Dọn dẹp rác hệ thống định kỳ:** Chủ động xóa file rác Office (`~$*`, `~WRL*.tmp`), file lock rỗng, file scratch thử nghiệm.

