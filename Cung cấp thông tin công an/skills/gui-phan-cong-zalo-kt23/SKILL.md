---
name: gui-phan-cong-zalo-kt23
description: "Gửi Zalo phân công KT2/KT3: tin ngắn (tên DN, số CV, ngày, hạn), xem trước rồi gửi, cập nhật Đang xử lý."
group: Tài chính
---

# Gửi phân công Zalo KT2/KT3

## Khi nào dùng
- Vừa phân công (hoặc người dùng lệnh gửi) công văn thuộc Tổ **KT2** hoặc **KT3**.
- Lệnh kiểu: "gửi Zalo KT2", "gửi phân công cho đầu mối KT3", "bàn giao KT2/KT3 qua Zalo".
- Không dùng cho KT1 (KT1 đi đường in PDF `in-phan-cong-kt1`).

## Chuẩn bị
- Zalo Agent MCP đã đấu, bật, mức **Toàn quyền**; gói `javis.zalo` đã cài (có `zalo_send_image`).
- Đầu mối đã chốt (memory `facts/dau-moi-zalo-phan-cong-kt2-kt3.md`):

| Tổ | Người nhận | threadId | threadType |
|---|---|---|---|
| KT2 | CCT- Nga- TT | `973578993300096841` | 0 |
| KT3 | CCT- Tieu Diem | `5588572486369282885` | 0 |

- File PDF công văn phải nằm **trong brain** (vd đã copy vào `exports/gui-phan-cong-zalo/`). Tool chỉ gửi file trong vault.
- Sổ theo dõi: `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\Bao_cao_cong_van.xlsx` (sheet `NHAP_HO_SO`, cột O = Tình trạng).

## Mẫu tin Zalo (BẮT BUỘC)
Tin chữ **chỉ** gồm 4 mục, mỗi mục một dòng, không thêm gì khác:

```
<Tên công ty>
CV <Số CV> ngày <dd/mm/yyyy>
Hạn thực hiện: trước <dd/mm/yyyy>
```

Nếu thiếu hạn trên sổ: ghi `Hạn thực hiện: [CHƯA XÁC ĐỊNH]`.

**Cấm** đưa vào tin chữ: folder, tình trạng, MST, cơ quan ban hành, số đến, dòng "phân công Tổ KT2/KT3", câu "gửi kèm ảnh/PDF/scan", caption dài.

Nhiều DN trong cùng một CV cùng tổ: mỗi DN một khối 3 dòng như trên, cách nhau một dòng trống; hoặc gửi lần lượt từng DN nếu tin quá dài.

## Quy trình
1. Xác định Tổ KT (`KT2` hoặc `KT3`). Một công văn có cả hai tổ: xử lý **tách** từng đầu mối, mỗi bên chỉ DN thuộc tổ đó.
2. Lấy đúng `threadId` / `threadType` trong bảng trên. Không đoán tên gần giống.
3. Copy PDF vào `exports/gui-phan-cong-zalo/`. Nếu `zalo_send_file` (PDF) lỗi vì thiếu kết nối realtime, raster hóa PDF sang JPG (lô 2 trang/lần) để gửi ảnh; **không** viết câu thông báo scan vào tin chữ.
4. **Xem trước qua Telegram bot CA (BẮT BUỘC, chốt 20/11/2026):** gửi bản nháp tin chữ + file PDF (hoặc JPG) + tên đầu mối KT2/KT3 về bot **`@Kiem_soat_dieu_phoi_bot`** (Trợ lý quản lý công văn CA), chat chủ `8799384849`, bằng script:
   `python "/brains/Cung cấp thông tin công an/scripts/gui-telegram-bot-ca.py" --text "<nháp tin + Tổ + đầu mối>" --file "<đường tuyệt đối PDF/JPG>" --caption "[XEM TRƯỚC] ..."`.
   **Không** dùng `POST /telegram/send-file` (đó là bot Javis chính). Có thể nhắc ngắn trên khung chat web rằng đã gửi Telegram bot CA. **Dừng lại.** Chỉ gửi Zalo đầu mối khi người dùng xác nhận rõ (vd "gửi", "chốt gửi", "ok gửi", "OK").
5. Sau khi được chốt: gửi tin chữ (`zalo_send_message` hoặc CLI `msg send`), rồi gửi file/ảnh (`zalo_send_image` / CLI `msg send-image`). Caption ảnh chỉ ghi ngắn `CV <Số CV>` nếu tool bắt buộc có caption.
6. **Cập nhật sổ ngay sau khi gửi thành công:** cột O (Tình trạng) từ `Mới` sang `Đang xử lý` trên `Bao_cao_cong_van.xlsx`.
   - Trước khi ghi: chạy `python scripts/cua-kiem-dong-bao-cao.py`.
   - **Linux/VPS (BẮT BUỘC):** account `googledrive_ralph-urbane`, fileId sống `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`.
     1. Tải sổ về file tạm.
     2. `python scripts/ghi-so-bao-cao-an-toan.py set-status --in <tạm> --out <tạm> --row <dòng> [--row ...] --status "Đang xử lý"`
     3. `python scripts/ghi-so-bao-cao-an-toan.py validate --in <tạm>` → phải `ok: true`.
     4. `GOOGLEDRIVE_UPLOAD_UPDATE_FILE` đúng fileId sống (không tạo bản dated).
     5. Tải lại → `validate` lần nữa.
   - Máy Windows có Excel: Excel COM (`win32com`), giống `in-phan-cong-kt1 --apply`.
   - Cấm: `openpyxl.save` cả workbook; ZipInfo Linux mới; ElementTree serialize sheet; nhầm account `googledrive_attach-speos`.
7. Báo lại chat: đã gửi cho ai, threadId, số CV, msgId (nếu có), đã/ chưa cập nhật sổ.

## Bẫy
- Nhiều bạn Zalo tên gần giống (vd `CCT- Thuy Nga- KK`): chỉ gửi đúng threadId đã chốt.
- File nằm ngoài vault (Drive `I:\...`) thì phải copy vào brain trước khi gọi tool gửi.
- Chưa có Zalo hoặc quyền dưới Toàn quyền: dừng và báo thiếu, không giả gửi.
- Không gửi khi chưa phân Tổ KT hoặc hồ sơ còn thiếu thông tin bắt buộc trên sổ.
- **Không gửi thẳng đầu mối** khi chưa có xác nhận xem trước từ người dùng.
- Không đổi trạng thái sổ trước khi gửi thành công.

## Kiểm chứng
- `zalo_search_threads` với đúng tên đầu mối trả đúng 1 DM khớp threadId trong bảng.
- Tin chữ khớp mẫu 4 mục (không có folder / tình trạng / câu gửi scan).
- Tool gửi trả thành công; cột O đã là `Đang xử lý` (hoặc đã báo rõ lý do chưa ghi được sổ sống).

## Lịch sử (tự học)
- [2026-11-18] Chốt mẫu tin ngắn: tên DN, số CV + ngày, hạn; bỏ folder/tình trạng/câu gửi scan/dòng phân tổ.
- [2026-11-18] Bắt buộc xem trước rồi mới gửi đầu mối; sau gửi thành công cập nhật sổ sang `Đang xử lý`.
- [2026-11-18] PDF realtime `zalo_send_file` dễ lỗi trên phiên MCP: fallback JPG qua CLI `msg send-image`, gửi lô 2 trang.
- [2026-11-18 15:27] Kênh xem trước chốt là **Telegram** (send-file + caption nháp); khung chat web chỉ báo đã gửi xem trước. Người dùng chọn "OK quy trình lần sau", không gửi lại CV đã bàn giao.
- [2026-11-20] Chuyển kênh xem trước sang bot CA `@Kiem_soat_dieu_phoi_bot` qua `scripts/gui-telegram-bot-ca.py`; thôi dùng `/telegram/send-file` cho việc CA.
- [2026-11-20] Cửa kiểm đóng Bao_cao + chờ Drive sync trước khi cập nhật sổ.
- [2026-11-20] VPS bắt buộc `scripts/ghi-so-bao-cao-an-toan.py` trước khi `UPLOAD_UPDATE`.
