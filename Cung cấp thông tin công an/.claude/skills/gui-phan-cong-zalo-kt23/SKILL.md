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
4. **Xem trước qua Telegram (BẮT BUỘC, chốt 18/11/2026):** gửi bản nháp tin chữ + file PDF (hoặc JPG) + tên đầu mối KT2/KT3 về **Telegram** của chủ bằng `POST http://127.0.0.1:7777/telegram/send-file` (caption ghi rõ `[XEM TRƯỚC]`, Tổ, đầu mối, và nguyên văn tin Zalo sẽ gửi). Có thể nhắc ngắn trên khung chat web rằng đã gửi Telegram. **Dừng lại.** Chỉ gửi Zalo đầu mối khi người dùng xác nhận rõ (vd "gửi", "chốt gửi", "ok gửi", "OK").
5. Sau khi được chốt: gửi tin chữ (`zalo_send_message` hoặc CLI `msg send`), rồi gửi file/ảnh (`zalo_send_image` / CLI `msg send-image`). Caption ảnh chỉ ghi ngắn `CV <Số CV>` nếu tool bắt buộc có caption.
6. **Cập nhật sổ ngay sau khi gửi thành công:** cột O (Tình trạng) từ `Mới` sang `Đang xử lý` trên `Bao_cao_cong_van.xlsx`.
   - **Đường chính trên Linux/Javis:** Google Drive Composio account `googledrive_ralph-urbane` (`cungcapthongtin.tcs3tdn@gmail.com`), fileId sổ sống (vd `1OWkt_HFG4Ew-i5Gf9FpXD71gPGsXmXph`). Tải về → vá ô O (ưu tiên sửa XML trong xlsx để giữ chart DASHBOARD) → `GOOGLEDRIVE_UPLOAD_UPDATE_FILE` → tải lại kiểm tra mở được và còn chart.
   - Máy Windows có Excel: cũng có thể dùng Excel COM (`win32com`), giống `in-phan-cong-kt1 --apply`.
   - Không `load_workbook` + `save` openpyxl đè cả file (dễ hỏng biểu đồ). Không nhầm account mặc định `googledrive_attach-speos` (mail khác).
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
