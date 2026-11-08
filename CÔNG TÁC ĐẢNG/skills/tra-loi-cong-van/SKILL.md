---
name: Trả lời công văn Công an
description: Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho cơ quan Công an theo quy chuẩn thể thức và dữ liệu quản lý thuế.
group: Tài chính
updated: '2026-09-25'
learned_updated: '2026-09-25'
---

# Trả lời công văn Công an

Quy trình nghiệp vụ và chuẩn thể thức hành chính để lập hồ sơ dự thảo công văn phúc đáp và tờ trình lãnh đạo phê duyệt cung cấp thông tin, tài liệu người nộp thuế theo yêu cầu của cơ quan Công an (CSĐT, CSKT, CSHS) và các cơ quan chức năng.

## Khi nào dùng
- Khi tiếp nhận yêu cầu soạn thảo văn bản trả lời công văn của cơ quan Công an, Viện kiểm sát, Tòa án hoặc cơ quan điều tra các cấp.
- Cần lập dự thảo Công văn trả lời (`TRA LOI CV <Số CV>_<CQ>_du thao.docx`) và Tờ trình lãnh đạo (`to trinh <Số CV>.docx`).
- Cần tra cứu, tổng hợp và đối chiếu dữ liệu thuế lịch sử của các doanh nghiệp theo từng giai đoạn yêu cầu từ cơ sở dữ liệu `I:\My Drive\DB NNT.xls`.
- Cần căn chỉnh thể thức văn bản hành chính theo đúng quy chuẩn kỹ thuật (phông chữ, cỡ chữ, bảng lề in, giãn dòng, khoảng cách chữ ký).

## Chuẩn bị & Dữ liệu nguồn
- **Thư mục hồ sơ gốc:** `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\<STT>. CV <SO_CV>...`
- **File Markdown trích xuất công văn:** `D:\HSO CONG AN FILE MD\<SO_CV>...md`
- **File Danh bạ Người nộp thuế (DB NNT):** `I:\My Drive\DB NNT.xls` (sheet `NNT`)
  - Chứa đầy đủ thông tin định danh: Mã số thuế, Tên doanh nghiệp, Địa chỉ ĐKKD (phường/xã, quận/huyện), Người đại diện / Giám đốc, Số GCN ĐKKD / Ngày cấp, Ngày bắt đầu kinh doanh.
  - Tổ kiểm tra quản lý (KT1, KT2, KT3), Cán bộ quản lý (CBQL).
  - Trạng thái đăng ký thuế tổ chức (*Đang hoạt động*, *NNT không hoạt động tại địa chỉ đã đăng ký*, *NNT ngừng HĐ nhưng chưa hoàn thành thủ tục chấm dứt hiệu lực*...).
  - Lý do và ngày đóng trạng thái / Số và ngày thông báo đóng/khóa MST (cột `Lý do đóng MST tổ chức`, `Ngày đóng trạng thái tổ chức`).
- **File mẫu chuẩn:**
  - Dự thảo công văn: `TRA LOI CV <Số CV>_<Tên CQ>_du thao.docx`
  - Tờ trình lãnh đạo: `to trinh <Số CV>.docx`
- **Cơ sở dữ liệu thuế bổ sung (kê khai & hóa đơn):**
  - Tờ khai thuế GTGT, TNDN, TNCN, Báo cáo tài chính (BCTC).
  - Thông báo phát hành/sử dụng HĐĐT, báo cáo tình hình sử dụng hóa đơn giấy, bảng kê HĐĐT theo các kỳ/năm yêu cầu.
  - Tình hình nợ thuế, tiền chậm nộp, vi phạm hành chính, kết luận thanh tra/kiểm tra.
- **Nhân sự phê duyệt ký chuẩn (Tổ Kiểm tra 1):**
  - Cán bộ đề xuất: **Nguyễn Thị Ánh Ngọc** (hoặc cán bộ thụ lý hồ sơ).
  - Lãnh đạo Tổ KTR1: Tổ trưởng **Nguyễn Trường Giang**.
  - Lãnh đạo phê duyệt ký: Phó Trưởng thuế cơ sở **Phạm Đồng Xuân** (hoặc Lãnh đạo phụ trách).

## Quy trình thực hiện

### Bước 1: Trích xuất và xác định phạm vi yêu cầu
1. Đọc file trích xuất markdown tại `D:\HSO CONG AN FILE MD` hoặc công văn trong thư mục hồ sơ gốc.
2. **Xác định nội dung yêu cầu cụ thể:** Phân tích kỹ nội dung công văn của cơ quan Công an để nắm rõ các tài liệu, số liệu và thông tin được yêu cầu (hồ sơ pháp lý, tình hình kê khai thuế GTGT/TNDN/TNCN, BCTC, hóa đơn, nợ thuế, vi phạm...). Nội dung công văn trả lời phải căn cứ và bám sát trực tiếp vào các nội dung yêu cầu này.
3. Lập danh sách doanh nghiệp cần xác minh: Tên công ty, Mã số thuế, Địa chỉ ĐKKD.
4. Xác định phạm vi thời gian yêu cầu cung cấp (ví dụ: giai đoạn từ năm 2020 đến 2025) và danh mục tài liệu cần trích xuất.

### Bước 2: Thu thập dữ liệu từ DB NNT (`I:\My Drive\DB NNT.xls`) & Cơ sở dữ liệu thuế
1. **Tra cứu trong file `I:\My Drive\DB NNT.xls` (sheet `NNT`):**
   - Tra cứu trực tiếp theo cột `Mã số thuế` hoặc `Tên người nộp thuế`.
   - Trích xuất thông tin ĐKKD: Số giấy chứng nhận ĐKKD / Quyết định thành lập (cột `Số giấy chứng nhận đăng ký kinh doanh`), Ngày cấp lần đầu (cột `Ngày cấp giấy chứng nhận ĐKKD` hoặc `Ngày cấp`), Ngày bắt đầu hoạt động kinh doanh (cột `Ngày bắt đầu kinh doanh`).
   - Trích xuất thông tin người đại diện: Họ tên Giám đốc / Đại diện pháp luật (cột `Tên giám đốc`), số CMND/CCCD (cột `Số CMND của người đại diện`).
   - Trích xuất địa chỉ trụ sở: Kết hợp cột `Số nhà/Đường phố`, `Tên Phường Xã`, `Tên Quận/Huyện`, `Tên Tỉnh/Thành Phố`.
   - Xác định đơn vị quản lý: `Tổ kiểm tra quản lý` (KT1, KT2, KT3) và Cán bộ quản lý (`CBQL`).
2. **Xác định trạng thái hoạt động & lịch sử pháp lý:**
   - *Đang hoạt động:* Xác nhận trạng thái tại cột `Trạng thái ĐKT tổ chức`, ghi nhận doanh nghiệp đang hoạt động tại địa chỉ đăng ký kinh doanh.
   - *Không hoạt động tại địa chỉ đăng ký / Đóng MST:* Lấy chính xác số Thông báo, ngày ban hành tại cột `Lý do đóng MST tổ chức` (ví dụ: *Thông báo số 20814/TB-CCTKV ngày 12/09/2024 của Chi cục Thuế khu vực Biên Hòa - Vĩnh Cửu*) và ngày có hiệu lực tại cột `Ngày đóng trạng thái tổ chức`.
   - *Doanh nghiệp thành lập giữa kỳ:* Nêu rõ ngày thành lập theo ĐKKD; giai đoạn trước ngày thành lập không phát sinh dữ liệu kê khai thuế.
3. **Tổng hợp hồ sơ khai thuế, BCTC & Hóa đơn:**
   - Thống kê chi tiết các kỳ kê khai thuế GTGT, TNDN, TNCN và các năm đã nộp Báo cáo tài chính (BCTC) trong giai đoạn yêu cầu.
   - Thống kê tình hình sử dụng hóa đơn: hóa đơn giấy, thời điểm chuyển đổi HĐĐT, thông báo phát hành và dữ liệu bảng kê HĐĐT mua vào/bán ra.
   - Rà soát tình hình nợ thuế, tiền chậm nộp, vi phạm hành chính, kết luận thanh tra/kiểm tra.

### Bước 3: Soạn thảo dự thảo Công văn trả lời (`TRA LOI CV ..._du thao.docx`)
Áp dụng nghiêm ngặt các quy chuẩn kỹ thuật thể thức sau:
- **Nguyên tắc nội dung:** Bám sát và căn cứ trực tiếp vào nội dung yêu cầu của công văn Công an để trả lời đầy đủ, chính xác từng nội dung yêu cầu cho từng doanh nghiệp.
- **Tên file:** Bắt buộc có hậu tố `_du thao` (ví dụ: `TRA LOI CV 29931_CA DONG NAI_du thao.docx`).
- **Phần đầu (Header Table - Bảng 2 cột không viền):**
  - *Cột 1 (Bên trái):*
    - Tên cơ quan cấp trên: `THUẾ THÀNH PHỐ ĐỒNG NAI` (cỡ chữ **13**, in hoa, chữ đứng).
    - Tên cơ quan trực tiếp ban hành: `THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI` (cỡ chữ **13**, in hoa, in đậm).
    - Đường gạch ngang dưới đơn vị ban hành: Căn giữa, độ dài bằng đúng **1/3** chiều dài dòng tên đơn vị ban hành.
    - Số văn bản: `Số:          ` (để nguyên chữ "Số", không thêm ký tự lạ).
    - Trích yếu: `V/v cung cấp thông tin, hồ sơ DN theo Công văn [Số CV].` (chữ **đứng**, không in nghiêng, không in đậm).
  - *Cột 2 (Bên phải):*
    - Quốc hiệu: `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` (cỡ chữ **13**, in hoa, in đậm, nằm trọn **1** dòng).
    - Tiêu ngữ: `Độc lập - Tự do - Hạnh phúc` (cỡ chữ **14**, in đậm).
    - Đường gạch dưới tiêu ngữ: Kẻ liền nét ngay dưới tiêu ngữ, độ dài vừa hết chiều dài dòng chữ.
    - Địa danh, ngày tháng: *Đồng Nai, ngày ... tháng ... năm ...* (chữ nghiêng).
  - *Độ rộng ô:* Căn chỉnh độ rộng 2 cột vừa đủ để chữ nằm trọn vẹn trên từng dòng, không bị ngắt rớt dòng.
- **Kính gửi:** `Kính gửi: Cơ quan Cảnh sát điều tra - Công an...` (căn **chính giữa** trang, chữ thường theo chuẩn).
- **Danh sách doanh nghiệp:** Đánh số thứ tự **1, 2, 3...**, **in đậm tên công ty**, thụt lề trái **1.25 cm**.
- **Quy chuẩn đoạn văn bản Công văn:**
  - Phông chữ: **Times New Roman**.
  - Cỡ chữ: **14**.
  - Giãn dòng (line spacing): **1.15**.
  - Khoảng cách đoạn trước (space before): **3pt**.
  - Thụt lề đầu dòng / lề trái: **1.25 cm**.
- **Bảng cuối trang (Nơi nhận & Chữ ký):**
  - Bảng 2 cột không viền vừa khít toàn bộ chiều rộng vùng in **16.5 cm** (lề trái **2.5 cm**, lề phải **2.0 cm**).
  - Căn hàng trên cùng (**top-aligned**) để dòng `Nơi nhận:` và chức danh ký nằm ngang hàng thẳng nhau.
  - *Cột trái:* `Nơi nhận:` (in đậm, nghiêng), danh sách cơ quan nhận (cỡ chữ 11-12, căn lề trái).
  - *Cột phải:* Chức danh `KT. TRƯỞNG THUẾ CƠ SỞ / PHÓ TRƯỞNG THUẾ CƠ SỞ` (in đậm, cỡ 13-14), cách **đúng 5 dòng trống** chữ ký, phía dưới là họ tên người ký (in đậm, cỡ **14**, căn giữa cột).

### Bước 4: Soạn thảo Tờ trình Lãnh đạo (`to trinh <Số CV>.docx`)
- **Trình bày:** Toàn bộ nội dung và 3 khối ký phải căn chỉnh gọn gàng trọn vẹn trong **01 trang in A4 duy nhất**.
- **Canh lề trang:** Lề trái **3.0 cm**, lề trên **2.5 cm**, lề dưới **1.0 cm**, lề phải **2.0 cm**.
- **Phần đầu (Header Table - Bảng 2 cột):**
  - Tự động canh độ rộng khung và cho phép mở rộng sang 2 bên lề (`tblInd` âm khoảng -1.2 cm đến -1.3 cm, cột 1 rộng **8.5 cm**, cột 2 rộng **10.1 cm**) để chữ không bị rớt dòng.
  - *Cột trái:* `THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI` (cỡ **13**, chữ thường, trọn **1** dòng) và `TỔ KIỂM TRA THUẾ SỐ 1` (cỡ **13**, in đậm), gạch ngang chính giữa bằng **1/3** dòng tên tổ.
  - *Cột phải:* `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` (cỡ **13**, in đậm, trọn **1** dòng) và `Độc lập - Tự do - Hạnh phúc` (cỡ **14**, in đậm, trọn **1** dòng) có gạch ngang liền dưới vừa hết phần chữ.
- **Quy chuẩn đoạn văn bản Tờ trình:**
  - Phông chữ: **Times New Roman**.
  - Cỡ chữ: **13**.
  - Giãn dòng (line spacing): **1.15**.
  - Khoảng cách đoạn trước (space before): **3pt**.
  - Thụt lề đầu dòng / lề trái: **1.25 cm**.
- **Tiêu đề & Trích dẫn:**
  - Tiêu đề `TỜ TRÌNH LÃNH ĐẠO` in đậm cỡ **14**.
  - Trích dẫn `Về việc: cung cấp thông tin doanh nghiệp` để chữ thường, **không in đậm**, không in nghiêng.
- **Khung 3 vị trí ký duyệt:**
  - Bố trí ở nửa bên phải trang giấy và căn chính giữa khối chữ ký.
  - Khối 1: Cán bộ đề xuất (**Nguyễn Thị Ánh Ngọc**).
  - Khối 2: Lãnh đạo Tổ KTR1 (Tổ trưởng **Nguyễn Trường Giang**).
  - Khối 3: Lãnh đạo phê duyệt (Phó Trưởng thuế cơ sở **Phạm Đồng Xuân**).
  - Khoảng cách giữa dòng chức danh và họ tên người ký là **đúng 5 dòng trống**.

### Bước 5: Lưu trữ hồ sơ & Dọn dẹp
1. Lưu toàn bộ file kết quả vào thư mục hồ sơ gốc tương ứng tại Google Drive.
2. Xóa bỏ các file nháp, file thừa, chỉ giữ lại đúng 2 file chuẩn cuối cùng:
   - `TRA LOI CV <Số CV>_<CQ>_du thao.docx`
   - `to trinh <Số CV>.docx`

## Bẫy & Lưu ý quan trọng (Pitfalls)
- **Trả lời không đúng phạm vi:** Bắt buộc căn cứ đúng các nội dung yêu cầu cụ thể của công văn Công an, không trả lời lệch hoặc thiếu nội dung yêu cầu.
- **Tự bịa số liệu:** Tuyệt đối không suy đoán số liệu ĐKKD, tình trạng hoạt động hay số liệu kê khai. Bắt buộc tra cứu trực tiếp từ `I:\My Drive\DB NNT.xls` (sheet `NNT`). Nếu DN mới thành lập, phải nêu rõ giai đoạn trước thành lập không phát sinh.
- **Bỏ sót số thông báo đóng MST:** Với doanh nghiệp ngừng hoạt động/khóa MST, bắt buộc trích xuất đúng số và ngày Thông báo từ cột `Lý do đóng MST tổ chức`.
- **Lỗi rớt dòng Header:** Nếu không chỉnh độ rộng ô bảng phần đầu và thiết lập mở rộng sang lề (`tblInd` âm và fixed column widths), Quốc hiệu hoặc Tên đơn vị sẽ rớt dòng gây mất thẩm mỹ.
- **Thiếu hậu tố `_du thao`:** Công văn trả lời luôn là bản dự thảo để lãnh đạo duyệt nên bắt buộc mang tên `_du thao.docx`.
- **Tờ trình tràn sang trang 2:** Cần tối ưu khoảng cách dòng (cỡ chữ 13, giãn dòng 1.15, before 3pt, thụt lề 1.25 cm, lề trên 2.5, lề dưới 1.0) để toàn bộ tờ trình cùng 3 khung ký nằm trọn vẹn trong 1 trang A4.
- **Lệch hàng bảng cuối trang:** Cột Nơi nhận và Cột Chữ ký phải top-aligned; không được để thụt thò so le nhau.

## Kiểm chứng (Verification)
- Công văn trả lời: đúng cỡ chữ 14 cho body, 13 cho header, 14 cho tiêu ngữ, giãn dòng 1.15, before 3pt, thụt lề 1.25 cm.
- Tờ trình lãnh đạo: lề trái 3.0 cm, trên 2.5 cm, dưới 1.0 cm, phải 2.0 cm; Header trên 1 dòng duy nhất; đúng cỡ chữ 13 cho body và header, 14 cho tiêu đề và tiêu ngữ, giãn dòng 1.15, before 3pt, thụt lề 1.25 cm, hiển thị đủ 3 khối ký trên 1 trang A4.
- Nội dung trả lời bám sát đúng từng nội dung yêu cầu của công văn Công an.
- Bảng Nơi nhận - Chữ ký vừa khít 16.5 cm, khoảng cách ký đúng 5 dòng trống.
- Dữ liệu pháp lý, ĐKKD, tình trạng hoạt động và số thông báo đóng MST khớp hoàn toàn với `I:\My Drive\DB NNT.xls`.
- File đã được lưu đúng vào thư mục gốc của công văn.

## Lịch sử (tự học)
- [2026-09-25] Khởi tạo skill chuẩn hóa quy trình dự thảo công văn và tờ trình trả lời Công an: quy định chi tiết thể thức bảng 2 cột, cỡ chữ 13-14, lề in 16.5cm, khoảng cách ký 5 dòng trống, tờ trình 1 trang A4, hậu tố `_du thao` và nhân sự phê duyệt chuẩn.
- [2026-09-25] Quy định rõ Bước 2 thu thập và đối chiếu dữ liệu người nộp thuế trực tiếp từ file cơ sở dữ liệu `I:\My Drive\DB NNT.xls` (sheet `NNT`): tra cứu MST, ĐKKD, người đại diện, địa chỉ, tổ/cán bộ quản lý, trạng thái hoạt động và trích xuất số/ngày thông báo đóng MST từ cột Lý do đóng MST tổ chức.
- [2026-09-25] Bổ sung nguyên tắc nội dung công văn trả lời căn cứ vào nội dung yêu cầu của công văn Công an; chuẩn hóa định dạng đoạn văn bản Tờ trình lãnh đạo (cỡ chữ 13, giãn dòng 1.15, space before 3pt, thụt lề 1.25 cm).
- [2026-09-25] Chuẩn hóa canh lề Tờ trình (trái 3.0 cm, trên 2.5 cm, dưới 1.0 cm, phải 2.0 cm), tự động canh khung header mở rộng sang lề (cột 8.5 cm & 10.1 cm) để Cơ quan ban hành và Quốc hiệu nằm trọn vẹn trên 1 dòng duy nhất, không bị rớt dòng.
