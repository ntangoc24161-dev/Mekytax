---
name: Tổ chức hồ sơ kiểm tra thuế
description: "Tổ chức thư mục hồ sơ kiểm tra thuế 4 bước, phân loại file, checklist đủ/thiếu A4, gom 1 nơi."
group: Tài chính
updated: '2026-10-14'
---
> **Hợp đồng đầu ra v1.4:** module nội bộ. Kết quả ghi vào sheet đích của sáu tệp KQ (xem skills/phan-tich-ho-so-thue/playbook-lenh-theo-buoc.md). Không giao tệp chuyên đề Tong_hop_*/Doi_chieu_* / Master 8 sheet như sản phẩm.

# Tổ chức hồ sơ kiểm tra thuế

Lớp **khung hồ sơ**: dựng thư mục 4 bước `<TEN>_<MST>`, phân loại file vào đúng chỗ, kiểm kê đủ/thiếu, xuất **1 checklist A4**, dọn nguồn cũ. **Không** chạy đối chiếu số liệu, 28 tiêu chí, Master 8 sheet.

Cây thư mục và map loại file: `references/cay-thu-muc-4-buoc.md`. Playbook lệnh: `playbook-lenh.md`.

## Khi nào dùng

- "sắp xếp hồ sơ", "phân loại thư mục", "tổ chức folder 4 bước", "checklist đủ thiếu", "gom 1 nơi", "xóa file gốc cũ", "kiểm kê hồ sơ", `thu-muc-mau-ho-so-doanh-nghiep`.
- User đưa đường dẫn kế hoạch / thư mục DN lộn xộn, cần chuẩn hóa trước khi phân tích.
- Nhiều DN trong 1 thư mục kế hoạch: mỗi DN 1 folder + 1 checklist; file tiến độ đoàn (mẫu 07) ở **gốc kế hoạch**.

Không dùng khi:

- Đã có folder chuẩn, cần đối chiếu XML/HĐĐT/28 tiêu chí / Master → `phan-tich-ho-so-thue`.
- Checklist nghiệp vụ mở hồ sơ (BCTC, GDLK, nhánh thiếu thuế/trốn thuế) → `mo-ho-so-ttkt`.
- Xuất bộ Word DI-01..DI-07 / app hồ sơ đi → `ho-so-kiem-tra-ke-hoach` hoặc `lap-ho-so-kiem-tra`.

## Chuẩn bị

1. Đường dẫn tuyệt đối thư mục nguồn (ổ G / máy). Thiếu thì hỏi 1 câu rồi dừng.
2. MST + tên ngắn ASCII (vd `TAM_HOANG_PHAT_3603721832`). Suy được từ folder/file thì làm luôn.
3. Mẫu: `thu-muc-mau-ho-so-doanh-nghiep/` (brain này).
4. Checklist mẫu: `thu-muc-mau-ho-so-doanh-nghiep/MAU_FILE_TONG_HOP/01_CHECKLIST_TONG_HOP_HO_SO_MAU.xlsx`.
5. (Tuỳ chọn, nhiều DN) mẫu 07: `07_TONG_HOP_TIEN_DO_DOAN_KIEM_TRA_MAU.xlsx` ở gốc kế hoạch, **không** kẹp vào 1 DN.

## Quy trình

### 1. Dựng thư mục chuẩn

Sao chép cây 4 khối từ mẫu, đổi tên `<TEN_DOANH_NGHIEP>_<MST>`. Đủ 4 khối `1.` `2.` `3.` `4.` và các thư mục con trong `references/cay-thu-muc-4-buoc.md`. Không tự bịa tên thư mục khác.

### 2. Phân loại và di chuyển file

Map theo **loại tài liệu**, không theo tên folder cũ:

| Loại | Đích |
|---|---|
| XML tờ khai GTGT/TNDN/TNCN/BCTC/KHBS | `1.1. To khai XML/` |
| HĐĐT portal CQ thuế (`HDDT.xlsx`) | `1.2. Hoa don dien tu/` |
| PTDN, ETAX, thông báo nợ | `1.3. Nghia vu nop thue & Thong bao/` |
| Danh sách NLĐ theo QTT TNCN | `1.4. Tong hop danh sach NLD theo QTT TNCN/` |
| NKC, Sổ Cái, sổ 111/112/131/331/3388/341 | `2.1. So ke toan tong hop va chi tiet/` |
| XNT 152/155/156, 154, giá thành, định mức | `2.2. Kho - Chi phi - Gia thanh/` |
| Sao kê NH Excel/PDF, UNC | `2.3. Ngan hang & Dong tien/` |
| Lương, chấm công, HĐLĐ, khấu hao 214, 242 | `2.4. Lao dong - Tien luong - TSCD/` |
| Bảng kê mua/bán **do công ty cung cấp** | `2.5. Bang ke mua vao ban ra/` |
| Sổ CVCBRR F0/F1, Sổ Công an | `3.1. Canh bao rui ro CVCBRR & Cong an/` |
| Danh bạ đối tác / trạng thái ĐKKD | `3.2. Danh ba trang thai NNT/` |
| Tra cứu TNCN đa nguồn | `3.3. Tra cuu TNCN da nguon & Nhan su/` |
| QĐ kiểm tra, BB, giải trình | `3.4. Giai trinh & Ho so bo sung/` |

HĐĐT CQ thuế **không** trộn chung folder với bảng kê NNT. Đường dẫn Windows dài: ghi bằng tiền tố `\\?\`.

### 3. Kiểm kê và xuất 1 checklist

Quét file thật từng thư mục con. Ghi **Đạt / Thiếu / Chưa có** kèm **số file** và **nội dung cụ thể còn thiếu** (vd "Sao kê Vietcombank Excel + PDF 2023-2025"), không ghi chung chung "thiếu hồ sơ Bước 2".

Xuất **duy nhất** `4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx`:

- In 1 trang A4 đứng, Fit to 1 page, Times New Roman, Quốc hiệu/Tiêu ngữ, thông tin DN, hộp kết luận, ô ký 2 cấp.
- Sheet danh sách file thực tế (tên, dung lượng, vị trí).
- Xóa `CHECKLIST_BUOC_1` … `CHECKLIST_BUOC_4` và mọi `.md` checklist rải ở thư mục con.

Script `thu-muc-mau-ho-so-doanh-nghiep/kiem_tra_tinh_day_du_ho_so.py` **có thể lệch tên cũ** (`1.2. Hoa don dien tu & Bang ke`). Ưu tiên quét theo cây hiện tại; đừng tin script mù.

### 4. Gom 1 nơi, xóa nguồn cũ

Sau khi đối soát số lượng + dung lượng:

- Xóa thư mục/file gốc đã chuyển hết (vd `To khai\<MST>`, folder tên cũ `1. SON TIEN HAI`).
- Mỗi file chỉ còn **1 chỗ** trong `<TEN>_<MST>/`.
- Chưa chuyển xong thì **không** xóa nguồn.

### 5. Nhiều DN: file tiến độ đoàn

Cập nhật `TONG_HOP_TIEN_DO_...xlsx` ở gốc kế hoạch: mỗi DN 1 dòng, cột hồ sơ NNT còn thiếu vs hồ sơ CQT cần trích. **Không** thay checklist 1 DN.

### 6. Dừng đúng ranh giới

Báo cáo: folder chuẩn, số file, mục Đạt/Thiếu, phiếu yêu cầu NNT, việc CQT cần kết xuất. **Không** gộp bảng kê, không đối chiếu GTGT/TNDN, không chạy 28 tiêu chí. Bước tiếp theo: `tong-hop-dau-vao-hddt-bang-ke-nld` rồi skill trục `phan-tich-ho-so-thue` (playbook từ 0b).

## Lệnh mẫu

```
Tổ chức hồ sơ 4 bước tại <FOLDER_KE_HOACH hoặc FOLDER_DN>
MST <MST>, tên <TEN>, kỳ <KY>.
Tạo thư mục <TEN>_<MST>, phân loại file, xuất 1 checklist A4
tại 4. KET QUA PHAN TICH/CHECKLIST_TONG_HOP_HO_SO.xlsx.
Xóa nguồn cũ sau khi đối soát. Không phân tích số liệu.
```

## Bẫy

- Nhầm skill này với `phan-tich-ho-so-thue` (tính số) hoặc `mo-ho-so-ttkt` (checklist nghiệp vụ).
- Giữ checklist từng bước trong `1.` `2.` `3.` `4.`.
- Để song song folder cũ và folder chuẩn.
- Nhét bảng kê NNT vào `1.2` (HĐĐT portal).
- Coi mẫu 07 (tiến độ đoàn) = mẫu 01 (checklist 1 DN).
- Ô text Excel bắt đầu bằng `=` (lỗi Removed Records Formula).
- Xóa nguồn khi copy chưa đủ hoặc file đang bị Excel khóa.

## Kiểm chứng

- Tồn tại `<TEN>_<MST>` đủ 4 khối và thư mục con chuẩn.
- Chỉ còn 1 file `CHECKLIST_TONG_HOP_HO_SO.xlsx` trong `4. KET QUA PHAN TICH/`.
- Cột giải thích nêu đúng tài liệu thiếu, chia NNT cung cấp / CQT trích xuất.
- Nguồn cũ đã xóa (hoặc ghi rõ chưa xóa vì còn file).
- Chưa tạo/ghi đè Master phân tích trừ khi user ra lệnh skill kia.

## Lịch sử (tự học)

- [2026-10-14] Tách lớp tổ chức thư mục + checklist khỏi `phan-tich-ho-so-thue` theo yêu cầu user.
