---
name: Bàn giao liên phòng
description: "Checklist bàn giao việc giữa các brain/phòng: đường dẫn, MST, brain đích, agent/workflow, đầu ra."
group: Vận hành
---

# Bàn giao liên phòng

## Khi nào dùng

- Văn thư vừa phân việc và cần chuyển sang brain chuyên.
- Chat đang ở sai brain (Default / VĂN THƯ) nhưng việc thuộc Công an, Cảnh báo, Phân tích, Hồ sơ kiểm tra, Ôn tập.
- Tạo Kanban / nhắn người dùng bước chạy tiếp sau khi ghi sổ.
- Ví dụ: "bàn giao hồ sơ CTL sang cảnh báo", "việc này chạy ở brain nào", "viết gói bàn giao sang phân tích".

## Chuẩn bị

- Đọc map: `wiki/Ban-do-lien-ket-phong-brain.md` (hoặc `skills/ban-giao-lien-phong/references/ban-do.md`).
- Có đường dẫn hồ sơ tuyệt đối nếu việc đụng file ngoài vault.
- Biết MST / tên NNT / kỳ nếu là việc thuế.

## Cách chạy

Điền đủ khối bàn giao dưới đây (thiếu ô nào thì hỏi tối đa 1-3 câu, hoặc ghi rõ "chưa có"):

```text
## BÀN GIAO LIÊN PHÒNG
- Việc: <1 câu>
- Brain đích: <tên brain đúng>
- Agent: <slug> (hoặc Workflow: <slug>, status active/off)
- Đường dẫn hồ sơ: <absolute path hoặc relative trong brain>
- MST / NNT / kỳ: <nếu có>
- Đầu vào đã có: <file/sheet/sổ>
- Đầu ra cần: <Word/Excel/sổ/cột Kết quả báo lại>
- Hạn: <ngày>
- Báo về: sổ Văn thư số đến <n> / chat hiện tại
- Cấm lẫn: <vd không chạy phân tích 80 quy tắc trên brain Hồ sơ kiểm tra>
```

Sau khi điền:

1. Ghi tóm tắt vào sổ Văn thư ô `Kết quả báo lại` hoặc tin nhắn cho người dùng.
2. **Hỏi xác nhận trước khi coi là bàn giao xong** (bắt buộc khi brain đích ≠ brain đang chat):
   - Nói rõ: "Việc này nên chạy ở **brain X**, trợ lý/workflow **Y**. Anh xác nhận để em hướng dẫn đổi?"
   - Dashboard: kèm khối `JAVIS_ASK` tối đa 4 nút, ví dụ Đồng ý đổi / Ở lại hub chỉ ghi sổ / Đổi sang brain khác / Để sau.
   - Telegram/Zalo: liệt kê số 1-4 (không có nút).
3. Chỉ sau khi chủ xác nhận: nêu thao tác tiếp.
   - Web: **đổi brain X → mở Cộng sự → chọn agent/workflow Y**.
   - Telegram/Zalo: gõ `/brain X` rồi gọi đúng agent/workflow (kiểm bằng `/status`).
4. Không giả vờ đã gọi agent ở brain khác. Không nói "đã chuyển brain". Không tạo task nền trừ khi người dùng yêu cầu chạy nền và đã chọn đúng brain đích.

## Quy trình chọn brain (rút gọn)

| Dấu hiệu trong việc | Brain đích |
|---|---|
| Công an, cung cấp hồ sơ, Bao_cao_cong_van | Cung cấp thông tin công an |
| Cảnh báo rủi ro, F0/F1, đối chiếu mặt hàng mua bán | Cập nhật công văn cảnh báo |
| KQ1/KQ2/KQ3, 3 giai đoạn, Master 28 tiêu chí | Phân tích hồ sơ thuế |
| DI-01..07, hồ sơ đi/về, biên bản kiểm tra | Hồ sơ kiểm tra |
| Vào sổ, thể thức NĐ 30, phát hành VB đi | VĂN THƯ |
| Đề thi, cheat sheet, giải đề cuốn chiếu | Ôn tập thi |

## Bẫy

- Agent cùng slug có thể tồn tại ở **hai chỗ** (VĂN THƯ mirror + brain chuyên). **Nguồn chân lý = brain chuyên**; ưu tiên đổi brain chuyên khi cần memory/facts của vụ (vd CTL). Chi tiết danh sách mirror: `wiki/Backlog-van-hanh-lien-phong.md`.
- Workflow `off` vẫn ghi trong bàn giao nhưng phải nói "bật tay hoặc chạy từng agent".
- Orch Kanban đang off hầu hết brain: đừng hứa "tự chạy nền" nếu chưa bật AI tự vận hành đúng brain đích (chỉ bật tạm có chủ đích - xem backlog mục 4).
- Bỏ qua bước hỏi xác nhận rồi bảo "đã bàn giao" là sai quy trình.

## Kiểm chứng

- Khối bàn giao có đủ Brain đích + Agent/Workflow + Đường dẫn + Đầu ra.
- Người đọc biết thao tác tiếp theo trong 1 câu, không cần đoán.
