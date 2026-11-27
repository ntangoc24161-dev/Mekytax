---
type: wiki
title: Playbook lệnh - Loại lương chủ DNTN và chủ TNHH MTV
updated: 2026-11-27
---

# Playbook lệnh - Loại lương chủ DNTN / TNHH MTV

Skill: `loai-luong-chu-dntn-tnhh-mtv`. Bước 4b sau TNCN trong `playbook-tat-ca-skill.md`.

Thay `<FOLDER>`, `<MST>`, `<TEN>`, `<KY>`.

**Cần:** loại hình DNTN hoặc TNHH MTV cá nhân; XML 05-1 (IN05); tốt hơn có TMS `1.3` và bảng lương `2.4`.

**Lệnh:**
```
Xét loại chi phí lương chủ DNTN / chủ TNHH MTV kiêm GĐ
cho <TEN> MST <MST> tại <FOLDER> kỳ <KY>
theo skill loai-luong-chu-dntn-tnhh-mtv.
1) Xác định loại hình. Không phải DNTN/TNHH MTV cá nhân thì ghi KHÔNG_ÁP_DỤNG.
2) Tách TMS DIRNA (GĐ) và NAME_FIRS (chủ) từng năm, kể cả ngày đổi giữa năm.
3) Ghép 05-1 (và bảng lương nếu có) theo người + năm. Xếp A1/A2 hoặc B1/B2/B3.
4) Chỉ cộng vào khoản loại các dòng A1/B1. A2/B2 thiếu HĐLĐ ghi rủi ro chứng từ riêng.
5) Ghi KQ2.1!KQ2_TNCN_LD2 (số + nhận xét module). Kết luận chỉ ở KQ3.
6) Nếu có Bang tinh: điền sheet LOẠI LƯƠNG. Chưa trừ cột Số kiểm tra 642/641.
Không loại nhân viên khác. Không điền 0 giả. Font Times New Roman.
```

**Ra:** phiên bản mới của KQ2.1 (khối loại lương trên `KQ2_TNCN_LD2`). Word/Excel dự thảo tại `0. BIEN BAN` chỉ khi lệnh yêu cầu lập biên bản.

**Thiếu TMS:** vẫn xếp theo 05-1 + GCN hiện hành; ghi thiếu lịch sử đổi chủ/GĐ.

**Thiếu 05-1:** `KHÔNG_ĐỦ_DỮ_LIỆU`, không bịa số lương.

**Câu siêu ngắn:**
```
Xét loại lương chủ DNTN/TNHH MTV <FOLDER>
```
