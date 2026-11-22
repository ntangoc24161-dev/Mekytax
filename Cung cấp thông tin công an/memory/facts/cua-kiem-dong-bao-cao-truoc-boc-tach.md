---
type: decision
provenance: user
origin: javis-learned
created: 2026-11-20
updated: 2026-11-20
source: chat 2026-11-20 (mwnu8fh5tv6n99vy)
---
Trước khi bóc tách/ghi sổ `Bao_cao_cong_van.xlsx`:

1. Kiểm tra / nghi file đang mở.
2. **Nhắc qua bot CA** `@Kiem_soat_dieu_phoi_bot`.
3. **Tự đóng + chạy tiếp, không chờ anh trả lời chat:** gọi `scripts/cua-kiem-dong-bao-cao.py` (best-effort `javis_app_close` Excel trên máy đang chạy Javis; trên VPS Docker thường không tắt được Excel Windows của anh).
4. Chờ ngắn Drive sync (mặc định 60s) rồi **tự bóc tách/ghi đè** đúng 1 file sống Drive `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`.

**Vì sao:** Anh chốt 20/11/2026: sau khi nhắc bot CA, VPS phải tự đóng (mức có thể) và làm tiếp quy trình, không dừng chờ xác nhận.

**Giới hạn thật:** VPS Docker không điều khiển máy Windows của anh. "Tự đóng" trên VPS = nhắc + best-effort close local + không chờ chat + ghi đè cloud. Nếu Excel vẫn mở trên máy anh, vẫn nên đóng tay để tránh conflict Drive Desktop.

**Áp dụng:** workflow `quản-lý-công-văn-công-an` / `boc-tach-phan-cong-cong-van`; agent `khai-thac-du-lieu-ho-so`; skill `theo-doi-cung-cap-ca`; cron 12h.
