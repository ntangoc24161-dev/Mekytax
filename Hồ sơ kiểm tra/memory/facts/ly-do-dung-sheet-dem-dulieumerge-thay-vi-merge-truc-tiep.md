---
superseded_by: yeu-cau-ket-noi-merge-truc-tiep-01-da-ban-hanh
type: decision
provenance: assistant
origin: javis-learned
created: 2026-11-13
updated: 2026-11-13
source: chat 2026-11-13 (ce012a55)
---
Không nên merge trực tiếp từ sheet theo dõi `01_DA_BAN_HANH` vì 3 lý do kỹ thuật:
1. Thiếu các cột số tiền bằng chữ (như `Tong_truy_thu_chu`, `Tong_tien_phat_chu`...) cần cho biên bản và quyết định xử phạt.
2. Header của sheet theo dõi có nhiều tầng gộp ô (merged cells) và có dấu tiếng Việt, không phẳng như yêu cầu của Word Mail Merge.
3. Sheet theo dõi chứa nhiều doanh nghiệp nên phải lọc thủ công, trong khi sheet đệm `DuLieuMerge` tự khoá đúng 1 dòng vụ việc qua công thức liên kết và hàm `=DocSo(...)`.


## Lịch sử
- [2026-11-13] Bị thay thế bởi [[yeu-cau-ket-noi-merge-truc-tiep-01-da-ban-hanh]] (user đổi thông tin).
