---
name: Trả lời công văn Công an
description: Dự thảo công văn phúc đáp và tờ trình phê duyệt cung cấp hồ sơ cho Công an theo quy chuẩn thể thức và dữ liệu DB NNT.xls
group: Tài chính
updated: '2026-11-02'
learned_updated: '2026-11-02'
---
## Khi nào dùng
- Soạn thảo dự thảo công văn trả lời cơ quan Công an (hoặc phúc đáp Thuế thành phố Đồng Nai khi có công văn chỉ đạo trung gian) và tờ trình phê duyệt của Tổ Kiểm tra (KT1).

## Chuẩn bị
- File dữ liệu `DB NNT.xls` tại `I:\My Drive\DB NNT.xls` (sheet `NNT` tra cứu hồ sơ doanh nghiệp, trạng thái đóng MST; sheet `KTR1` tra cứu cán bộ quản lý) để trích xuất thông tin NNT.
- Số công văn đến, ngày nhận, nội dung yêu cầu của cơ quan Công an.
- Nếu thư mục hồ sơ có công văn chỉ đạo của **Thuế thành phố Đồng Nai** (thường dạng `.../DON-KTr1`, V/v tạm ngừng / cung cấp theo yêu cầu Công an): lấy đủ **số, ngày** công văn Thuế TP và số/ngày công văn Công an gốc.
- Thư mục lưu trữ hồ sơ công văn tại `I:\My Drive\HS CONG AN-2025-2026\HS CUNG CAP CA 2026\1. CUNG CAP CONG AN NAM 2026\<STT>. CV...` (trên máy user ổ Drive có thể mang chữ `H:`; cùng cấu trúc).
- **Kiểm tra trên Drive (không chỉ local) trước khi soạn:** thư mục CV đã có dự thảo `TRA LOI CV <Số CV>_..._du thao.docx` chưa; có thư mục con (vd `CV <Số CV>`) chứa văn bản trả lời Công an trước đó của từng DN hay không.

## Quy trình
1. **Phân tích yêu cầu công văn Công an & trích xuất dữ liệu:**
   - **Phạm vi đối tượng cung cấp:** Chỉ cung cấp thông tin và hồ sơ của các doanh nghiệp thuộc địa bàn do **Thuế cơ sở 3** quản lý (phân công KT1, KT2, KT3). Tuyệt đối **không cung cấp doanh nghiệp ngoài địa bàn** quản lý (loại bỏ hoàn toàn các doanh nghiệp ngoài địa bàn khỏi danh mục báo cáo/dự thảo).
   - **Tên doanh nghiệp:** Giữ nguyên bản tên doanh nghiệp từ CSDL `DB NNT.xls`, tuyệt đối **không gắn kèm bất kỳ ghi chú khởi tố nào** (như `*(Đã khởi tố)*` hoặc chú thích tương tự).
   - **Tình trạng hoạt động:** Trích xuất và ghi đúng nguyên văn giá trị theo cột **"Trạng thái ĐKT tổ chức"** từ file `DB NNT.xls` (ví dụ: *NNT đang hoạt động*, *NNT không hoạt động tại địa chỉ đã đăng ký*, *NNT tạm ngừng KD có thời hạn*, *NNT ngừng HĐ nhưng chưa hoàn thành thủ tục chấm dứt hiệu lực*, *NNT ngừng hoạt động và đã hoàn thành thủ tục chấm dứt hiệu lực*...), tuyệt đối **không tự thêm nội dung giải thích hoặc ghi tình trạng hoạt động khác**.
   - Nếu doanh nghiệp ngừng hoạt động: trích xuất chính xác số và ngày Thông báo không hoạt động tại địa chỉ đã đăng ký từ cột *Lý do đóng MST tổ chức*.
   - **DN đã được trả lời Công an trước đó:** đọc văn bản trả lời cũ (thư mục con `CV <Số CV>` hoặc thư mục CV cũ của DN) làm nền nội dung, rồi đối chiếu lại `DB NNT.xls` để cập nhật phần đã thay đổi.
   - **Đã có dự thảo cũ:** sửa trên file đó (ghi đè, không tạo bản sao); rà lại tổ phụ trách theo DB NNT, DN có thật trong bảng của CV, hướng kính gửi và dòng Số.
   - Đọc kỹ nội dung yêu cầu của công văn Công an để phân loại:
     + *Có yêu cầu thông tin doanh nghiệp / pháp lý:* (yêu cầu cung cấp thông tin ĐKKD, người đại diện theo pháp luật, tình trạng hoạt động, cán bộ quản lý...). Khi đó mới trích xuất chi tiết từ `DB NNT.xls` để lập mục "Thông tin cung cấp:".
     + *Không yêu cầu thông tin doanh nghiệp:* (chỉ yêu cầu sao lục hồ sơ khai thuế, hóa đơn, BCTC, rà soát xử phạt vi phạm hành chính, xác minh hóa đơn cụ thể, tạm ngừng kê khai...). Tuyệt đối không lập mục "Thông tin cung cấp:", chỉ lập mục danh mục hồ sơ tài liệu hoặc trả lời đúng câu hỏi.
2. **Soạn thảo dự thảo Công văn trả lời (`TRA LOI CV <Số CV>_<CQ>_du thao.docx`):**
   - **Bảng Header 3 hàng (mở rộng ngoài định dạng trang in, tblInd âm khoảng -850 dxa / -1.5 cm, không đệm ô):**
     - *Đặc tính bảng:* Tổng chiều rộng 19.3 cm (10942 dxa), Cột trái 9.1 – 9.5 cm (5159 – 5386 dxa), Cột phải 10.2 cm (5783 dxa), cố định fixed layout, không viền (`tblBorders: none`), không đệm ô (`top/bottom/left/right: 0`), thụt lề âm `tblInd = -850 dxa` (-1.5 cm) sang lề trái trang để các dòng dài không bị rớt dòng.
     - *Hàng 1:*
       - Cột trái (9.1 – 9.5 cm): Cơ quan ban hành gồm dòng trên "THUẾ THÀNH PHỐ ĐỒNG NAI" (cỡ 13, chữ đứng, căn giữa) và dòng dưới "THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI" (cỡ 13, in đậm, căn giữa). **Cả 2 dòng tên đơn vị bắt buộc nằm trọn vẹn trên từng dòng riêng biệt, tuyệt đối không bị ngắt rớt dòng.** Đường kẻ ngang dưới đơn vị ban hành đặt chính giữa với độ dài bằng đúng 1/3 dòng tên đơn vị (khoảng 2.8 - 3.0 cm, 12 ký tự `─` ở cỡ 10pt).
       - Cột phải (10.2 cm): Quốc hiệu "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" (cỡ 13, in đậm, căn giữa, nằm trọn trên 1 dòng) và Tiêu ngữ "Độc lập - Tự do - Hạnh phúc" (cỡ 14, in đậm, căn giữa, nằm trọn trên 1 dòng). Nét gạch ngang liền dưới đặt chính giữa với độ dài vừa hết chiều dài dòng chữ từ "Đ" đến "c" (khoảng 6.0 - 6.17 cm, 24 ký tự `─` ở cỡ 10pt), không rớt dòng.
     - *Hàng 2 (Số & Ngày tháng):*
       - Cột trái: `Số:                 /TCS3-KTr1` (chữ "Số:" rồi 17 khoảng trắng để điền số văn bản, tiếp theo là **ký hiệu cố định `/TCS3-KTr1`**; cỡ 13, chữ đứng, không in đậm, **căn trái**, `left_indent = Cm(2)` cách lề trái **2 cm**, space before 4pt, space after 2pt). Không căn giữa ô.
       - Cột phải: *Trấn Biên, ngày       tháng       năm* (cỡ 13, in nghiêng, căn giữa, thẳng hàng ngang với dòng Số, space before 4pt, space after 2pt). Địa danh là **Trấn Biên**, **không ghi sẵn năm** - để trống cho người ký điền đủ ngày, tháng, năm. (Quy định bỏ năm này CHỈ áp cho công văn trả lời, không áp cho tờ trình.)
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
     + Trích yếu hàng 3 và câu kết phúc đáp: đối với công văn trả lời cho Thuế Thành phố Đồng Nai nếu căn cứ công văn của Công an thì phần trích yếu của công văn (`V/v ...`) ghi theo **số công văn của Công an**; câu kết phúc đáp và Kính gửi khớp Thuế thành phố Đồng Nai.
   - **Thân bài & Cung cấp thông tin / hồ sơ Người nộp thuế:**
     - Phông chữ Times New Roman, cỡ chữ 14, giãn dòng 1.15, space before 3pt, space after 0pt.
     - Thụt lề đầu dòng 1.25 cm (`first_line_indent = Cm(1.25)`, không dùng `left_indent`), căn đều hai bên (`Justify`).
     - Dự thảo đủ **tất cả** doanh nghiệp trong công văn thuộc địa bàn Thuế cơ sở 3 (KT1, KT2, KT3), đánh số liên tục; **tuyệt đối không đưa doanh nghiệp ngoài địa bàn vào dự thảo**.
     - Tên doanh nghiệp: In hoa, in đậm (space before 3pt), giữ nguyên tên từ CSDL, không gắn ghi chú khởi tố.
     - **Chỉ lập mục Thông tin cung cấp khi có yêu cầu:** Bám sát đúng danh mục công văn Công an. Mẫu chuẩn (như CV 29938): Địa điểm ĐKKD; Tình trạng hoạt động (+ số/ngày TB nếu ngừng); Số tài khoản ngân hàng; Tình hình thực hiện nghĩa vụ thuế. **Không** tự thêm Giám đốc/CBQL nếu công văn không yêu cầu.
     - **Tình trạng hoạt động:** Ghi đúng nguyên văn theo cột **"Trạng thái ĐKT tổ chức"** từ `DB NNT.xls`.
     - **Tình hình thanh tra, kiểm tra, rà soát xử lý vi phạm:** Khi công văn yêu cầu rà soát hoặc cung cấp tình hình thanh tra, kiểm tra, xử lý vi phạm về thuế và hóa đơn:
       + Đối với trường hợp chưa có vi phạm hoặc rà soát không phát hiện vi phạm: chuẩn hóa nội dung thành **"Công ty chưa bị xử phạt về thuế và hóa đơn."** (hoặc trong bảng: *Chưa bị xử phạt về thuế và hóa đơn*).
       + Đối với trường hợp có biên bản / quyết định xử phạt thực tế: trích dẫn chính xác số quyết định, ngày ban hành và số tiền xử phạt theo hồ sơ lưu.
       + Bổ sung dòng `+ Tình hình thanh tra, kiểm tra, xử lý vi phạm về thuế:` (có bằng chứng QĐ/biên bản thì ghi đúng; chưa có vi phạm → `Công ty chưa bị xử phạt về thuế và hóa đơn.`, tô đỏ nếu chưa xác minh).
     - **Nguồn dữ liệu bắt buộc (không bịa):**
       + Địa chỉ / tình trạng / TB đóng MST: `DB NNT.xls` + file `LICH SU <MST>.XLSX` trong thư mục CV (nếu có).
       + Số tài khoản ngân hàng: cột `Số tài khoản` + `Tên ngân hàng` trong `DB NNT.xls` (đăng ký với CQT). Trống → ghi "NNT không đăng ký với Cơ quan Thuế".
       + Nợ thuế / nghĩa vụ NSNN: kết xuất TMS hoặc sổ theo dõi nợ của đơn vị. Không có file/kết xuất → không bịa số; hỏi user hoặc ghi đúng theo bằng chứng vừa kiểm.
       + Thanh tra / kiểm tra / VPHC: hồ sơ quyết định, biên bản trong thư mục CV hoặc trên hệ thống. Chưa vi phạm → ghi "Công ty chưa bị xử phạt về thuế và hóa đơn.".
       + DN đã trả lời trước: văn bản trả lời cũ trong thư mục con `CV <Số CV>` hoặc thư mục CV trước đó; dữ liệu cá nhân lấy lại từ văn bản cũ mà DB NNT không có thì báo user kiểm lại trước khi ký.
     - **Đánh dấu chữ đỏ khi thiếu / chưa xác định (bắt buộc):**
       + Mọi giá trị còn thiếu, chưa chính xác, hoặc chưa có bằng chứng live phải tô **chữ đỏ** (RGB 255,0,0) trên file Word để tách khỏi phần đã kiểm chứng (chữ đen).
       + Chat Dashboard: dùng nhãn **`[CHƯA XÁC ĐỊNH]`** (in đậm) trước nội dung đó; khi ghi Word vẫn tô đỏ thật.
       + Ví dụ: câu nợ thuế chưa có TMS; chưa đối chiếu được từ DB.
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
   - **Trích yếu tờ trình** dòng `Về việc: ...` cỡ 13, chữ đứng, **viết thường, in đậm**; giữ nguyên nội dung theo mẫu hồ sơ, không tự đổi.
   - Có CV Thuế TP: `Về việc` và mục `- Nội dung công việc:` nêu **Công văn Thuế TP** trước, rồi nêu Công văn Công an gốc là căn cứ yêu cầu cung cấp.
   - Mục `- Nội dung công việc:` chỉ ghi căn cứ số/ngày công văn và trích yếu ngắn; **không** liệt kê dài danh sách doanh nghiệp dự thảo hay doanh nghiệp ngoài địa bàn trong tờ trình.
   - Nội dung cỡ chữ 13, giãn dòng 1.15, before 3pt, thụt lề đầu dòng 1.25 cm.
   - **Thứ tự khối ký (bắt buộc, mẫu CV 30291):** bảng 3 hàng gồm CÁN BỘ ĐỀ XUẤT → dòng **Ý kiến của lãnh đạo Tổ KTR1:** (ô gộp, nằm **ngay dưới** cán bộ đề xuất) → TỔ TRƯỞNG; sau đó đoạn **Ý kiến phê duyệt của Lãnh đạo TCS3:** rồi bảng chữ ký Phó Trưởng thuế cơ sở. Tuyệt đối không đặt ý kiến Tổ KTR1 lên trên cán bộ đề xuất.
   - **Khoảng cách chữ ký:** giữa chức danh (CÁN BỘ ĐỀ XUẤT, TỔ TRƯỞNG, Phó Trưởng thuế cơ sở) và họ tên người ký để đúng **4 dòng trống**.
   - **Dòng ngày tháng trong cả 3 khối ký:** `Trấn Biên, ngày     tháng     năm <YYYY>` (in nghiêng, căn giữa ô) - **giữ năm**, năm lấy theo năm tại thời điểm lập tờ trình (vd `năm 2026`).
   - Nhân sự ký duyệt: Cán bộ đề xuất Nguyễn Thị Ánh Ngọc, Tổ trưởng Tổ KT1 Nguyễn Trường Giang, Lãnh đạo phê duyệt Phó Trưởng thuế cơ sở Phạm Đồng Xuân.
4. **Quản lý tệp:** Lưu trực tiếp cả 2 file Word vào thư mục công văn tương ứng trên Drive, không tạo file sao lưu `.bak` tạm thời.

## Bẫy
- Không tự ý lập mục thông tin pháp lý doanh nghiệp khi công văn cơ quan Công an chỉ yêu cầu sao lục hồ sơ (tờ khai, BCTC, hóa đơn), rà soát nợ/xử phạt hoặc xác minh cụ thể.
- Tránh để đường kẻ dưới tiêu ngữ quá ngắn (< 5.8 cm) hoặc quá dài (> 6.3 cm) làm lệch tỷ lệ với chữ (chuẩn 6.0 - 6.17 cm, 24 ký tự `─` cỡ 10pt khớp đúng từ chữ "Đ" đến chữ "c").
- Tránh dùng `left_indent` cho thân bài làm thu hẹp lề in; phải dùng `first_line_indent = Cm(1.25)`.
- Không in đậm nhãn các mục như 'Mã số thuế', 'Giấy chứng nhận ĐKKD' trong phần liệt kê thông tin NNT.
- Không để rớt dòng tên cơ quan ban hành (`THUẾ THÀNH PHỐ ĐỒNG NAI`, `THUẾ CƠ SỞ 3 THÀNH PHỐ ĐỒNG NAI`) hoặc Quốc hiệu ở Header; phải mở rộng ngoài định dạng trang in (`tblInd = -850 dxa`, Cột 1 = 9.1 – 9.5 cm, Cột 2 = 10.2 cm, đệm ô = 0).
- Không cung cấp các doanh nghiệp ngoài địa bàn quản lý của Thuế cơ sở 3; chỉ cung cấp các DN thuộc địa bàn Thuế cơ sở 3 phụ trách (KT1, KT2, KT3).
- Không tự ý thêm ghi chú khởi tố vào tên doanh nghiệp (như `*(Đã khởi tố)*`).
- Không tự ý thêm bớt hay diễn giải sai khác nội dung tình trạng hoạt động; bắt buộc ghi đúng nguyên văn theo cột "Trạng thái ĐKT tổ chức" tại `DB NNT.xls`.
- Không ghi sai lệch nội dung rà soát xử lý vi phạm; chuẩn hóa câu: "Công ty chưa bị xử phạt về thuế và hóa đơn." khi chưa có vi phạm.
- Không bỏ trống ký hiệu sau dòng Số: bắt buộc `Số:                 /TCS3-KTr1`; không căn giữa dòng Số - phải căn trái và cách lề trái đúng 2 cm.
- Công văn trả lời: không ghi địa danh "Đồng Nai" và không điền sẵn năm ở dòng ngày tháng; đúng chuẩn là *Trấn Biên, ngày       tháng       năm*.
- Tờ trình: ngược lại với công văn, KHÔNG bỏ năm ở 3 khối ký; phải ghi năm lập tờ trình.
- Không dùng đoạn căn cứ kiểu liệt kê "Căn cứ Công văn...; Căn cứ CSDL...; phúc đáp như sau:" - bắt buộc đúng 2 đoạn mở đầu chuẩn ở trên.
- Có công văn Thuế thành phố Đồng Nai trong hồ sơ mà vẫn kính gửi / câu kết thẳng Công an là sai hướng trả lời.
- Trích yếu công văn (hàng 3 header) bắt buộc đặt chữ đứng, không in nghiêng, không in đậm, căn giữa ô.
- Toàn bộ thân bài công văn tuân thủ: cỡ chữ 14, giãn dòng 1.15, space before 3pt, thụt đầu dòng 1.25 cm.
- Không tạo hoặc để lại file `.bak` sao lưu tạm trong thư mục lưu trữ.
- Tờ trình không được tràn sang trang thứ 2; không chia cột dòng ý kiến Lãnh đạo.
- Không đặt dòng "Ý kiến của lãnh đạo Tổ KTR1" lên trên khối CÁN BỘ ĐỀ XUẤT; phải nằm ngay dưới cán bộ đề xuất rồi mới tới TỔ TRƯỞNG.
- Dòng `Về việc:` trên tờ trình phải in đậm (khác trích yếu công văn); không nhồi danh sách dài doanh nghiệp vào mục Nội dung công việc của tờ trình.
- Khoảng cách chữ ký tờ trình là 4 dòng trống, khác 5 dòng trống của bảng chữ ký công văn; đừng dùng chung một giá trị.
- Chỉ tra thư mục local rồi kết luận "chưa có dự thảo" là sai (sự cố CV 34662 ngày 02/11/2026: Drive đã có dự thảo từ 07/10/2026). Luôn tra thư mục CV trên Drive, kể cả thư mục con.
- Bỏ qua thư mục con `CV <Số CV>` chứa văn bản trả lời cũ của từng DN thì dự thảo lệch với nội dung đã cung cấp trước đó.
- Dự thảo cũ có thể ghi sai tổ (vd ghi KT1 trong khi DB NNT ghi KT2) hoặc chứa DN không có trong bảng của CV; phải đối chiếu lại trước khi giữ.
- Sửa skill này tại file chuẩn `skills/tra-loi-cong-van/SKILL.md`; chỉ sửa bản mirror `.claude/skills/...` thì có thể bị ghi đè khi đồng bộ.

## Kiểm chứng
- Mở kiểm tra hiển thị văn bản trên Word:
  + Công văn: Đúng người nhận (Thuế TP khi có DON-KTr1; Công an khi không có), căn cứ 2 đoạn khớp hướng trả lời, bám sát nội dung yêu cầu (chỉ có thông tin DN khi được yêu cầu), Header 3 hàng cân đối, không rớt dòng chữ cơ quan và Quốc hiệu, đường kẻ 1/3 và gạch dưới tiêu ngữ ôm khít chữ, trích yếu chữ đứng, dòng Kính gửi giãn 18pt/12pt căn giữa, các nhãn thông tin NNT viết thường, thụt đầu dòng 1.25 cm, dòng ngày tháng bỏ năm.
  + Tờ trình: Trọn vẹn trong đúng 01 trang A4, đúng 3 vị trí ký (Ngọc - Giang - Xuân), dòng `Về việc` in đậm, 4 dòng trống giữa chức danh và họ tên, cả 3 dòng ngày tháng có năm lập tờ trình.
  + Mọi DN trong dự thảo có trong bảng của CV Công an, đúng tổ theo DB NNT; DN đã trả lời trước khớp văn bản cũ trừ phần DB NNT đã đổi.
  + Không tồn tại tệp `.bak` thừa trong thư mục.

## Lịch sử (tự học)
- [2026-09-30] Có công văn chỉ đạo Thuế TP (`.../DON-KTr1`): kính gửi Thuế TP; căn cứ, trích yếu, tờ trình và câu kết khớp Thuế TP (mẫu CV 30291).
- [2026-10-02] Công văn trả lời Thuế TP căn cứ CV Công an thì trích yếu ghi theo số CV Công an.
- [2026-10-05] Header cột 1 9.1 - 9.5 cm; chỉ DN thuộc TCS3; tên DN nguyên bản; tình trạng hoạt động nguyên văn; câu chuẩn "Công ty chưa bị xử phạt về thuế và hóa đơn."
- [2026-10-05] Tờ trình: ý kiến lãnh đạo Tổ KTR1 nằm dưới cán bộ đề xuất (bảng 3 hàng mẫu 30291); mục Nội dung công việc không liệt kê dài danh sách DN.
- [2026-10-29] Header công văn: dòng Số kèm ký hiệu `/TCS3-KTr1`, căn trái cách lề 2 cm; dòng ngày tháng địa danh *Trấn Biên*, bỏ hẳn năm.
- [2026-10-29] Tờ trình theo yêu cầu user: dòng `Về việc: ...` cỡ 13, chữ đứng, viết thường, IN ĐẬM (thay "không in đậm"); 4 dòng trống giữa chức danh (CÁN BỘ ĐỀ XUẤT, TỔ TRƯỞNG, Phó Trưởng thuế cơ sở) và họ tên; dòng ngày tháng ở cả 3 khối ký GIỮ năm, lấy năm lập tờ trình (thay quy định bỏ năm cùng ngày). Công văn trả lời vẫn bỏ năm và giữ 5 dòng trống chữ ký.
- [2026-10-29] User yêu cầu 3 thay đổi cho tờ trình (Về việc in đậm, 4 dòng trống giữa chức danh và họ tên, giữ năm lập ở dòng ngày tháng 3 khối ký). Javis chỉ sửa bản mirror .claude/skills, còn file chuẩn skills/tra-loi-cong-van/SKILL.md vẫn ghi 'không in đậm' và 'bỏ hẳn năm' (dòng 96, 101, 124), nên làm theo bản
- [2026-10-29] Tờ trình theo yêu cầu user: dòng `Về việc: ...` cỡ 13, chữ đứng, viết thường, IN ĐẬM (thay "không in đậm"); 4 dòng trống giữa chức danh (CÁN BỘ ĐỀ XUẤT, TỔ TRƯỞNG, Phó Trưởng thuế cơ sở) và họ tên; dòng ngày tháng ở cả 3 khối ký GIỮ năm, lấy năm lập tờ trình. Công văn trả lời vẫn bỏ năm và giữ 5 dòng trống chữ ký.
- [2026-11-02] Trước khi soạn phải tra dự thảo có sẵn và thư mục con văn bản trả lời cũ trên Drive (sự cố CV 34662: báo nhầm "chưa có dự thảo" vì chỉ tra local).
- [2026-11-02] Ngày 02/11/2026 Javis báo thư mục 249 (CV 34662) 'chưa có dự thảo nào' vì chỉ tra thư mục local, trong khi Drive đã có dự thảo từ 07/10/2026; user còn chỉ ra thư mục con 'CV 34662' chứa văn bản đã trả lời trước của một số công ty mà Javis chưa đọc. Thêm bước tra dự thảo có sẵn trên Drive và đọc thư
