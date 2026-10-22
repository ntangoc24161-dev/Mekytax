---
name: Trả lời công văn Công an
description: Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho Công an theo quy chuẩn thể thức và dữ liệu DB NNT.xls
group: Tài chính
updated: '2026-09-30'
learned_updated: '2026-09-30'
---
## Khi nào dùng
- Soạn thảo dự thảo công văn trả lời cơ quan Công an (hoặc phúc đáp Thuế thành phố Đồng Nai khi có công văn chỉ đạo trung gian) và tờ trình phê duyệt của Tổ Kiểm tra (KT1).

## Chuẩn bị
- File dữ liệu `DB NNT.xls` tại `I:\My Drive\DB NNT.xls` (sheet `NNT` tra cứu hồ sơ doanh nghiệp, trạng thái đóng MST; sheet `KTR1` tra cứu cán bộ quản lý) để trích xuất thông tin NNT.
- Số công văn đến, ngày nhận, nội dung yêu cầu của cơ quan Công an.
- Nếu thư mục hồ sơ có công văn chỉ đạo của **Thuế thành phố Đồng Nai** (thường dạng `.../DON-KTr1`, V/v tạm ngừng / cung cấp theo yêu cầu Công an): lấy đủ **số, ngày** công văn Thuế TP và số/ngày công văn Công an gốc.
- Thư mục lưu trữ hồ sơ công văn tại `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\<STT>. CV...`.

## Quy trình
1. **Phân tích yêu cầu công văn Công an & trích xuất dữ liệu:**
   - Đọc kỹ nội dung yêu cầu của công văn Công an để phân loại:
     + *Có yêu cầu thông tin doanh nghiệp / pháp lý:* (yêu cầu cung cấp thông tin ĐKKD, người đại diện theo pháp luật, tình trạng hoạt động, cán bộ quản lý...). Khi đó mới trích xuất chi tiết từ `DB NNT.xls` để lập mục "Thông tin cung cấp:".
     + *Không yêu cầu thông tin doanh nghiệp:* (chỉ yêu cầu sao lục hồ sơ khai thuế, hóa đơn, BCTC, rà soát xử phạt vi phạm hành chính, xác minh hóa đơn cụ thể, tạm ngừng kê khai...). Tuyệt đối không lập mục "Thông tin cung cấp:", chỉ lập mục danh mục hồ sơ tài liệu hoặc trả lời đúng câu hỏi.
   - Nếu doanh nghiệp ngừng hoạt động: trích xuất chính xác số và ngày Thông báo không hoạt động tại địa chỉ đã đăng ký từ cột *Lý do đóng MST tổ chức*.
2. **Soạn thảo dự thảo Công văn trả lời (`TRA LOI CV <Số CV>_<CQ>_du thao.docx`):**
   - **Bảng Header 3 hàng (mở rộng ngoài định dạng trang in, tblInd âm khoảng -850 dxa / -1.5 cm, không đệm ô):**
     - *Đặc tính bảng:* Tổng chiều rộng 19.3 cm (10942 dxa), Cột trái 9.1 cm (5159 dxa), Cột phải 10.2 cm (5783 dxa), cố định fixed layout, không viền (`tblBorders: none`), không đệm ô (`top/bottom/left/right: 0`), thụt lề âm `tblInd = -850 dxa` (-1.5 cm) sang lề trái trang để các dòng dài không bị rớt dòng.
     - *Hàng 1:*
       - Cột trái (9.1 cm): Cơ quan ban hành gồm dòng trên "THUẾ THÀNH PHỐ ĐỒNG NAI" (cỡ 13, chữ đứng, căn giữa) và dòng dưới "THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI" (cỡ 13, in đậm, căn giữa). Đường kẻ ngang dưới đơn vị ban hành đặt chính giữa với độ dài bằng đúng 1/3 dòng tên đơn vị (khoảng 2.8 - 3.0 cm, 12 ký tự `─` ở cỡ 10pt). Cả 2 dòng tên đơn vị nằm trọn trên từng dòng riêng biệt, tuyệt đối không rớt dòng.
       - Cột phải (10.2 cm): Quốc hiệu "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" (cỡ 13, in đậm, căn giữa, nằm trọn trên 1 dòng) và Tiêu ngữ "Độc lập - Tự do - Hạnh phúc" (cỡ 14, in đậm, căn giữa, nằm trọn trên 1 dòng). Nét gạch ngang liền dưới đặt chính giữa với độ dài vừa hết chiều dài dòng chữ từ "Đ" đến "c" (khoảng 6.0 - 6.17 cm, 24 ký tự `─` ở cỡ 10pt), không rớt dòng.
     - *Hàng 2 (Số & Ngày tháng):*
       - Cột trái: `Số:          ` (chữ "Số:" giữ nguyên chuẩn, không chèn ký hiệu lạ hoặc dấu gạch chéo trơ trọi, để khoảng trống điền số văn bản, cỡ 13, **căn trái**, `left_indent = Cm(2)` cách lề trái **2 cm**, space before 4pt, space after 2pt). Không căn giữa ô.
       - Cột phải: *Đồng Nai, ngày       tháng       năm 2026* (cỡ 13, in nghiêng, căn giữa, thẳng hàng ngang với dòng Số, space before 4pt, space after 2pt).
     - *Hàng 3 (dòng phụ dưới Số):*
       - Ô trích yếu: cách lề trái trang **1 cm**, độ rộng ô **6 cm**; chữ `V/v cung cấp thông tin, hồ sơ DN theo Công văn....` đứng (không in nghiêng, không in đậm, cỡ 12-13), **canh giữa ô**, space before 2pt, space after 2pt. (Với header `tblInd = -1.5 cm` và lề trái trang 2.5 cm thì mép trái bảng đã sát mốc 1 cm; đặt `tcW` ô trích yếu = 6 cm / 3402 dxa.)
       - Phần còn lại hàng 3: để trống.
   - **Xác định người nhận & hướng trả lời (bắt buộc trước khi soạn):**
     + **Có công văn chỉ đạo của Thuế thành phố Đồng Nai** (file phối hợp `...DON-KTr1` / V/v cung cấp theo yêu cầu Công an trong thư mục CV): trả lời **Thuế thành phố Đồng Nai**, không kính gửi thẳng Công an.
     + **Không có** công văn Thuế TP (Công an gửi trực tiếp / chỉ có CV Công an): kính gửi đúng cơ quan Công an ban hành công văn gốc.
   - **Dòng Kính gửi:**
     - Căn giữa trang (`Center`), cỡ chữ 14 (Times New Roman), viết thường chuẩn quy cách, không in đậm.
     - Có CV Thuế TP → `Kính gửi: Thuế thành phố Đồng Nai`.
     - Không có CV Thuế TP → `Kính gửi: Cơ quan Cảnh sát điều tra - Công an thành phố...` (đúng tên cơ quan trên CV gốc).
     - Khoảng cách trước (space before) 18pt và khoảng cách sau (space after) 12-18pt, giãn dòng 1.15.
   - **Đoạn mở đầu / căn cứ (bắt buộc 2 đoạn):**
     + *Trường hợp có CV Thuế thành phố (mẫu CV 30291):*
       1. `Thuế cơ sở 3 thành phố Đồng Nai nhận được Công văn số <số DON-KTr1> ngày <ngày> của Thuế thành phố Đồng Nai về việc <trích yếu Thuế TP> theo yêu cầu của Cơ quan Công an tại Công văn số <số CV Công an> ngày <ngày>.`
       2. `Căn cứ dữ liệu quản lý thuế, Thuế cơ sở 3 thành phố Đồng Nai trả lời như sau:`
     + *Trường hợp trả lời thẳng Công an:*
       1. `Thuế cơ sở 3 thành phố Đồng Nai nhận được Công văn số ... ngày ... của Cơ quan Cảnh sát điều tra - Công an thành phố ... về việc đề nghị cung cấp thông tin, hồ sơ doanh nghiệp.`
       2. `Căn cứ dữ liệu quản lý thuế, Thuế cơ sở 3 thành phố Đồng Nai trả lời như sau:`
     + Trích yếu hàng 3 và câu kết phúc đáp phải khớp người nhận: có CV Thuế TP thì trích yếu / câu kết theo công văn Thuế TP; không thì theo công văn Công an.
   - **Thân bài & Cung cấp thông tin / hồ sơ Người nộp thuế:**
     - Phông chữ Times New Roman, cỡ chữ 14, giãn dòng 1.15, space before 3pt, space after 0pt.
     - Thụt lề đầu dòng 1.25 cm (`first_line_indent = Cm(1.25)`, không dùng `left_indent`), căn đều hai bên (`Justify`).
     - Dự thảo đủ **tất cả** doanh nghiệp trong công văn thuộc KT1, KT2, KT3 (đánh số liên tục), không chỉ riêng KT1.
     - Tên doanh nghiệp: In hoa, in đậm (space before 3pt).
     - **Chỉ lập mục Thông tin cung cấp khi có yêu cầu:** Bám sát đúng danh mục công văn Công an. Mẫu chuẩn (như CV 29938): Địa điểm ĐKKD; Tình trạng hoạt động (+ số/ngày TB nếu ngừng); Số tài khoản ngân hàng; Tình hình thực hiện nghĩa vụ thuế. **Không** tự thêm Giám đốc/CBQL nếu công văn không yêu cầu.
     - **Thanh tra / kiểm tra:** Nếu công văn yêu cầu cung cấp tài liệu liên quan thanh tra, kiểm tra (hoặc rà soát VPHC thuế/hóa đơn), bổ sung dòng `+ Tình hình thanh tra, kiểm tra, xử lý vi phạm về thuế:` (có bằng chứng QĐ/biên bản thì ghi đúng; không có → `Không có.`, tô đỏ nếu chưa xác minh).
     - **Nguồn dữ liệu bắt buộc (không bịa):**
       + Địa chỉ / tình trạng / TB đóng MST: `DB NNT.xls` + file `LICH SU <MST>.XLSX` trong thư mục CV (nếu có).
       + Số tài khoản ngân hàng: cột `Số tài khoản` + `Tên ngân hàng` trong `DB NNT.xls` (đăng ký với CQT). Trống → ghi "NNT không đăng ký với Cơ quan Thuế".
       + Nợ thuế / nghĩa vụ NSNN: kết xuất TMS hoặc sổ theo dõi nợ của đơn vị. Không có file/kết xuất → không bịa số; hỏi user hoặc ghi đúng theo bằng chứng vừa kiểm.
       + Thanh tra / kiểm tra / VPHC: hồ sơ quyết định, biên bản trong thư mục CV hoặc trên hệ thống. Không có → ghi "Không có".
     - **Đánh dấu chữ đỏ khi thiếu / chưa xác định (bắt buộc):**
       + Mọi giá trị còn thiếu, chưa chính xác, hoặc chưa có bằng chứng live phải tô **chữ đỏ** (RGB 255,0,0) trên file Word để tách khỏi phần đã kiểm chứng (chữ đen).
       + Chat Dashboard: dùng nhãn **`[CHƯA XÁC ĐỊNH]`** (in đậm) trước nội dung đó; khi ghi Word vẫn tô đỏ thật.
       + Ví dụ: câu nợ thuế chưa có TMS; "Không có" về thanh tra/VPHC khi chưa có QĐ/biên bản; trường chưa đối chiếu được từ DB.
       + Phần đã xác minh từ `DB NNT.xls` / `LICH SU` / hồ sơ trong thư mục CV giữ chữ đen.
     - Danh mục tài liệu cung cấp: Nhãn "Hồ sơ, tài liệu cung cấp bao gồm:" viết chữ thường, các mục chi tiết gạch đầu dòng; bám sát đúng danh mục và mốc thời gian công văn Công an yêu cầu. Không viết đoạn giải thích dài về thời kỳ yêu cầu vs ngày thành lập / ngày chuyển quản lý trong danh mục hồ sơ; chỉ ghi mốc quý/năm của từng loại tài liệu.
     - **Hóa đơn giấy (bắt buộc khi đủ điều kiện):** Nếu công văn yêu cầu cung cấp **thông báo phát hành hóa đơn** và/hoặc **báo cáo tình hình sử dụng hóa đơn**, và giai đoạn cung cấp có phần **từ năm 2022 về trước**, thì bổ sung vào danh mục:
       + Báo cáo tình hình sử dụng hóa đơn theo kỳ phát sinh (trong đoạn ≤ 2022, ghi theo **quý**);
       + Thông báo phát hành hóa đơn (nếu công văn yêu cầu hoặc thuộc giai đoạn còn dùng hóa đơn giấy).
       + Từ khi chuyển sang HĐĐT: dùng Thông báo chấp nhận đăng ký sử dụng HĐĐT và bảng kê HĐĐT mua vào - bán ra (theo quý).
     - **Phạm vi thời gian hồ sơ chỉ trong giai đoạn Thuế cơ sở 3 quản lý (bắt buộc):**
       + Nếu NNT chuyển về TCS3 giữa kỳ (có `LICH SU` / dữ liệu chuyển cơ quan): chỉ cung cấp tờ khai, BCTC, bảng kê HĐĐT từ kỳ bắt đầu quản lý của TCS3 đến quý/năm cuối cùng còn thuộc giai đoạn TCS3 quản lý (NNT đóng thì cắt theo quý/năm đóng; đang hoạt động thì đến quý/năm hiện tại hoặc mốc công văn yêu cầu).
       + **Cách ghi mốc thời gian (bắt buộc):** chỉ ghi **"đến quý..."** / **"đến năm..."**. **Không** viết thêm cụm "khóa MST", "thời điểm khóa MST", "tháng khóa MST".
       + **Quy đổi loại tờ khai:** GTGT và bảng kê HĐĐT / báo cáo sử dụng hóa đơn → theo **quý**; TNDN, TNCN, BCTC → theo **năm** (trong cùng giai đoạn cung cấp).
       + Ví dụ: chuyển từ tháng 07/2025, đóng tháng 03/2026 → GTGT/bảng kê từ **quý 03/2025 đến quý 01/2026**; TNDN/TNCN/BCTC **năm 2025**.
       + Ghi rõ câu: giai đoạn trước ngày chuyển, TCS3 không có dữ liệu cung cấp theo yêu cầu (khớp mẫu CV 29938) - đặt ở mục thông tin/tình trạng khi cần, không nhồi vào danh mục hồ sơ dưới dạng đoạn giải thích kép.
       + Nếu NNT thuộc TCS3 suốt thời đoạn công văn yêu cầu: cung cấp đúng mốc công văn, GTGT theo quý (vd quý 01/2020 đến quý 04/2025), TNDN/TNCN/BCTC theo năm (vd 2020 đến 2025).
   - **Bảng chữ ký cuối công văn:**
     - Bảng 2 cột không viền, rộng 16.5 cm, căn hàng trên (`top-aligned`).
     - Cột trái (7.5 cm): Nơi nhận (cỡ 12, in nghiêng đậm, các gạch đầu dòng cỡ 11 thường).
     - Cột phải (9.0 cm): Quyền hạn chức vụ (cỡ 13, in đậm, căn giữa), khoảng cách 5 dòng trống chữ ký, họ tên Lãnh đạo phê duyệt (Phó Trưởng thuế cơ sở Phạm Đồng Xuân, cỡ 14, in đậm, căn giữa).
3. **Soạn thảo Tờ trình lãnh đạo (`to trinh <Số CV>.docx`):**
   - Trình bày trọn vẹn trong 01 trang A4. Canh lề chuẩn: Trái 3.0 cm, Trên 2.5 cm, Dưới 1.0 cm, Phải 2.0 cm.
   - Khung Header mở rộng (cột 8.5 cm & 10.1 cm) để tên đơn vị và Quốc hiệu nằm trọn trên 1 dòng duy nhất.
   - Đường kẻ dưới tiêu ngữ chuẩn 5.5 - 5.6 cm (23 ký tự ─).
   - Tiêu đề `TỜ TRÌNH LÃNH ĐẠO` in đậm cỡ 14.
   - **Trích yếu tờ trình** dòng `Về việc: ...` cỡ 13, chữ đứng, **in đậm**.
   - Có CV Thuế TP: `Về việc` và mục `- Nội dung công việc:` nêu **Công văn Thuế TP** trước, rồi nêu Công văn Công an gốc là căn cứ yêu cầu cung cấp.
   - Nội dung cỡ chữ 13, giãn dòng 1.15, before 3pt, thụt lề đầu dòng 1.25 cm.
   - Các phần ý kiến của Tổ trưởng và Lãnh đạo là các dòng văn bản độc lập toàn dòng thụt lề 1.25 cm trước từng khối ký, tuyệt đối không chia cột.
   - Nhân sự ký duyệt: Cán bộ đề xuất Nguyễn Thị Ánh Ngọc, Tổ trưởng Tổ KT1 Nguyễn Trường Giang, Lãnh đạo phê duyệt Phó Trưởng thuế cơ sở Phạm Đồng Xuân.
4. **Quản lý tệp:** Lưu trực tiếp cả 2 file Word vào thư mục công văn tương ứng trên Drive, không tạo file sao lưu `.bak` tạm thời.

## Bẫy
- Không tự ý lập mục thông tin pháp lý doanh nghiệp khi công văn cơ quan Công an chỉ yêu cầu sao lục hồ sơ (tờ khai, BCTC, hóa đơn), rà soát nợ/xử phạt hoặc xác minh cụ thể.
- Tránh để đường kẻ dưới tiêu ngữ quá ngắn (< 5.8 cm) hoặc quá dài (> 6.3 cm) làm lệch tỷ lệ với chữ (chuẩn 6.0 - 6.17 cm, 24 ký tự `─` cỡ 10pt khớp đúng từ chữ "Đ" đến chữ "c").
- Tránh dùng `left_indent` cho thân bài làm thu hẹp lề in; phải dùng `first_line_indent = Cm(1.25)`.
- Không in đậm nhãn các mục như 'Mã số thuế', 'Giấy chứng nhận ĐKKD' trong phần liệt kê thông tin NNT.
- Không để rớt dòng tên cơ quan ban hành hoặc Quốc hiệu ở Header; phải mở rộng ngoài định dạng trang in (`tblInd = -850 dxa`, Cột 1 = 9.1 cm, Cột 2 = 10.2 cm).
- Không chèn ký hiệu lạ hoặc dấu gạch chéo trơ trọi sau mục 'Số:          '; không căn giữa dòng Số - phải căn trái và cách lề trái đúng 2 cm.
- Không dùng đoạn căn cứ kiểu liệt kê "Căn cứ Công văn...; Căn cứ CSDL...; phúc đáp như sau:" - bắt buộc đúng 2 đoạn mở đầu chuẩn ở trên.
- Có công văn Thuế thành phố Đồng Nai trong hồ sơ mà vẫn kính gửi / câu kết thẳng Công an là sai hướng trả lời.
- Trích yếu bắt buộc đặt chữ đứng, không in nghiêng, không in đậm, căn giữa ô ở hàng 3.
- Toàn bộ thân bài tuân thủ: cỡ chữ 14, giãn dòng 1.15, space before 3pt, thụt đầu dòng 1.25 cm.
- Không tạo hoặc để lại file `.bak` sao lưu tạm trong thư mục lưu trữ.
- Tờ trình không được tràn sang trang thứ 2; không chia cột dòng ý kiến Lãnh đạo.

## Kiểm chứng
- Mở kiểm tra hiển thị văn bản trên Word:
  + Công văn: Đúng người nhận (Thuế TP khi có DON-KTr1; Công an khi không có), căn cứ 2 đoạn khớp hướng trả lời, bám sát nội dung yêu cầu (chỉ có thông tin DN khi được yêu cầu), Header 3 hàng cân đối, không rớt dòng chữ cơ quan và Quốc hiệu, đường kẻ 1/3 và gạch dưới tiêu ngữ ôm khít chữ, trích yếu chữ đứng, dòng Kính gửi giãn 18pt/12pt căn giữa, các nhãn thông tin NNT viết thường, thụt đầu dòng 1.25 cm.
  + Tờ trình: Trọn vẹn trong đúng 01 trang A4, đúng 3 vị trí ký (Ngọc - Giang - Xuân).
  + Không tồn tại tệp `.bak` thừa trong thư mục.

## Lịch sử (tự học)
- [2026-09-25] Khởi tạo skill chuẩn hóa quy trình dự thảo công văn và tờ trình trả lời Công an: quy định chi tiết thể thức bảng 2 cột, cỡ chữ 13-14, lề in 16.5cm, khoảng cách ký 5 dòng trống, tờ trình 1 trang A4, hậu tố `_du thao` và nhân sự phê duyệt chuẩn.
- [2026-09-25] Quy định rõ Bước 2 thu thập và đối chiếu dữ liệu người nộp thuế trực tiếp từ file cơ sở dữ liệu `I:\My Drive\DB NNT.xls` (sheet `NNT`): tra cứu MST, ĐKKD, người đại diện, địa chỉ, tổ/cán bộ quản lý, trạng thái hoạt động và trích xuất số/ngày thông báo đóng MST từ cột Lý do đóng MST tổ chức.
- [2026-09-25] Bổ sung nguyên tắc nội dung công văn trả lời căn cứ vào nội dung yêu cầu của công văn Công an; chuẩn hóa định dạng đoạn văn bản Tờ trình lãnh đạo (cỡ chữ 13, giãn dòng 1.15, space before 3pt, thụt lề 1.25 cm).
- [2026-09-25] Chuẩn hóa canh lề Tờ trình (trái 3.0 cm, trên 2.5 cm, dưới 1.0 cm, phải 2.0 cm), tự động canh khung header mở rộng sang lề (cột 8.5 cm & 10.1 cm) để Cơ quan ban hành và Quốc hiệu nằm trọn vẹn trên 1 dòng duy nhất, không bị rớt dòng.
- [2026-09-25] Điều chỉnh cấu trúc Tờ trình: các phần ý kiến của Tổ trưởng và Lãnh đạo là các dòng văn bản độc lập thụt lề đầu dòng 1.25 cm trải rộng toàn dòng, không chia cột; các khối chữ ký được bố trí ở nửa bên phải căn giữa theo trục dọc.
- [2026-09-25] Chuẩn hóa Header Công văn trả lời: mở rộng ra ngoài định dạng trang in (cột 1 rộng 9.1 cm, cột 2 rộng 10.2 cm, tblInd -1.5 cm) để THUẾ THÀNH PHỐ ĐỒNG NAI (13, đứng), THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI (13, đậm), CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM (13, đậm) và Độc lập - Tự do - Hạnh phúc (14, đậm) nằm trọn vẹn trên từng dòng không rớt chữ; kẻ ngang đơn vị bằng đúng 1/3 dòng tên (12 ký tự ─); gạch dưới tiêu ngữ vừa hết chiều dài chữ từ "Đ" đến "c" (24 ký tự ─); trích yếu chữ đứng căn giữa; toàn bộ thân bài cỡ 14, giãn dòng 1.15, before 3pt, thụt lề đầu dòng 1.25 cm (`first_line_indent = Cm(1.25)`).
- [2026-09-25] Cập nhật chuẩn hóa chi tiết thể thức: mục Số giữ nguyên chuẩn không chèn ký hiệu lạ, dòng Kính gửi viết thường ở giữa trang, các mục thông tin NNT viết thường không in đậm và gạch dưới tiêu ngữ trọn vẹn chiều dài từ "Đ" đến "c".
- [2026-09-25] Hoàn thiện cấu trúc Header 3 hàng độc lập: tách hàng riêng cho trích yếu căn giữa ô chữ đứng (không in nghiêng/đậm), giữ mục Số ở hàng 2 tạo khoảng cách thoáng so với dòng tên đơn vị ở trên, loại bỏ hoàn toàn các ký hiệu lạ; áp dụng thụt đầu dòng chuẩn `first_line_indent = Cm(1.25)` cho thân bài và hoàn thiện bộ hồ sơ CV 21295/VPCQCSĐT.
- [2026-09-25] Cập nhật chuẩn hóa toàn diện theo chỉ đạo: Cơ quan ban hành (THUẾ THÀNH PHỐ ĐỒNG NAI chữ đứng cỡ 13, THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI in đậm cỡ 13, nằm trọn trên từng dòng; đường kẻ 1/3 đặt chính giữa); Quốc hiệu & Tiêu ngữ (CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM in đậm cỡ 13 trên 1 dòng; Độc lập - Tự do - Hạnh phúc in đậm cỡ 14 trên 1 dòng; nét gạch ngang liền dưới vừa hết chiều dài chữ; bảng header mở rộng ngoài định dạng trang in); Số công văn giữ nguyên chuẩn không chèn ký hiệu lạ; Trích yếu chữ đứng không in nghiêng không in đậm; Toàn bộ thân bài tuân thủ nghiêm ngặt cỡ chữ 14, giãn dòng 1.15, before 3pt, thụt đầu dòng 1.25 cm.
- [2026-09-25] Chuẩn hóa nguyên tắc nội dung: Chỉ lập mục thông tin người nộp thuế / doanh nghiệp khi công văn cơ quan Công an có yêu cầu cung cấp thông tin pháp lý / ĐKKD; nếu công văn chỉ yêu cầu sao lục hồ sơ, rà soát nợ / vi phạm hoặc xác minh hóa đơn thì không lập phần thông tin này mà chỉ cung cấp đúng tài liệu hoặc nội dung yêu cầu.
- [2026-09-29] Cập nhật theo CV 29931: dòng `Số:` căn trái, `left_indent = Cm(2)`; đoạn mở đầu đổi sang 2 câu chuẩn "nhận được Công văn... / Căn cứ dữ liệu quản lý thuế... trả lời như sau:"; dự thảo đủ toàn bộ DN thuộc KT1+KT2+KT3 trong cùng công văn; trạng thái hoạt động lấy live từ DB NNT.xls.
- [2026-09-29] Chuẩn nội dung theo mẫu 29938: không tự thêm Giám đốc/CBQL khi công văn không yêu cầu; TKNH lấy DB NNT; nợ thuế lấy TMS/sổ nợ (có bằng chứng mới ghi số); thanh tra lấy hồ sơ QĐ/biên bản hoặc "Không có"; ưu tiên file `LICH SU <MST>.XLSX` trong thư mục CV để cập nhật trạng thái/chuyển cơ quan quản lý.
- [2026-09-29] Preference: thông tin thiếu / chưa chính xác / chưa xác định tô chữ đỏ trên Word; chat dùng nhãn `[CHƯA XÁC ĐỊNH]`. Áp dụng ngay cho nợ thuế chưa TMS và thanh tra/VPHC chưa có QĐ-biên bản.
- [2026-09-29] Phạm vi hồ sơ cung cấp chỉ trong giai đoạn Thuế cơ sở 3 quản lý: NNT chuyển về giữa kỳ thì cắt từ kỳ bắt đầu quản lý (vd tháng 07/2025 → quý 03/2025 và năm 2025) đến quý/năm cuối thuộc giai đoạn TCS3; ghi rõ giai đoạn trước chuyển TCS3 không có dữ liệu.
- [2026-09-29] Mốc thời gian chỉ ghi đến quý/năm, không ghi "khóa MST"; GTGT/bảng kê theo quý, TNDN/TNCN/BCTC theo năm trong giai đoạn cung cấp.
- [2026-09-30] Không nhồi đoạn giải thích thời kỳ yêu cầu/ngày thành lập vào danh mục hồ sơ. Nếu CV yêu cầu thông báo phát hành hóa đơn / báo cáo tình hình sử dụng hóa đơn và giai đoạn cung cấp có phần từ 2022 về trước thì bổ sung các mục hóa đơn giấy tương ứng.
- [2026-09-30] Nếu CV yêu cầu tài liệu liên quan thanh tra/kiểm tra thì bổ sung dòng `+ Tình hình thanh tra, kiểm tra, xử lý vi phạm về thuế`.
- [2026-09-30] Ô trích yếu hàng 3: cách lề trái trang 1 cm, độ rộng 6 cm, chữ canh giữa ô; các phần header khác giữ nguyên.
- [2026-09-30] Trích yếu tờ trình dòng "Về việc: ..." tô đậm (cỡ 13).
- [2026-09-30] Nếu hồ sơ có công văn chỉ đạo Thuế thành phố Đồng Nai (`.../DON-KTr1`): kính gửi Thuế TP; căn cứ đoạn 1 là "nhận được Công văn Thuế TP ... theo yêu cầu Công an tại Công văn ..."; trích yếu, tờ trình và câu kết khớp Thuế TP (mẫu CV 30291).
