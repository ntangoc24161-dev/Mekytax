---
type: wiki
title: Playbook lệnh - Đối chiếu mặt hàng mua vào bán ra
updated: 2026-11-03
---

# Playbook `doi-chieu-mat-hang-mua-ban`

Catalog: `skills/phan-tich-ho-so-thue/playbook-tat-ca-skill.md` (skill cạnh). Không nằm trên đường ống bước 0-10.

## Cần

File HĐĐT sheet mua vào + bán ra của một NNT. Có cột SL / đơn giá thì so vật lý; không có thì chỉ so tiền và ghi phủ chưa đủ.

## Lệnh

```
Đối chiếu mặt hàng mua vào với bán ra theo skill doi-chieu-mat-hang-mua-ban
cho <FOLDER> MST <MST> kỳ <KY>.
Lọc HĐ hủy / ghi chú, cắt kỳ chồng, so tiền + SL + đơn giá.
Xuất DOI_CHIEU_MAT_HANG_MUA_BAN.xlsx đủ 6 sheet:
0. Tom tat ket luan  1. Tong hop nhom hang  2. Doi chieu mat hang
3. Cap mua-ban trung khop  4. Dien bien theo thang  5. Doi tac (100% + Tổng cộng).
Đo biên theo đồng/ĐVT trên cặp cùng SL 0-2 ngày, không kết luận thiếu hàng khi SL tháng ≈ 1,000.
Không xóa nguồn. Không nhầm F0/F1 hay 511-632.
```

**Siêu ngắn:** `Đối chiếu mặt hàng HĐĐT mua-bán <FOLDER>`

Word khi được hỏi: `BAO_CAO_DOI_CHIEU_MAT_HANG_MUA_BAN.docx` cùng thư mục.

## File ra (hỏi riêng một DN)

`DOI_CHIEU_MAT_HANG_MUA_BAN.xlsx` (+ Word nếu hỏi). Font Times New Roman, số `#,##0` / `#,##0.00`.

## Trong pipeline 1.4

Không giao file `DOI_CHIEU_*` như sản phẩm bộ sáu tệp KQ. Không tự dựng sheet KQ2. Chứng cứ đưa vào nhận xét `KQ2_Kho_Gia_thanh` hoặc dấu hiệu KQ3 khi cán bộ hỏi.

## Thiếu thì

Thiếu SL: chạy tiền, ghi phủ chưa đủ. Thiếu một phía mua hoặc bán: chạy phía còn lại, không giả 0. Không nhầm với `doi-chieu-ncc-mua-vao-rui-ro`, `doi-chieu-511-632`, `phan-tich-xnt-kho`.
