---
type: workflow
name: Dây chuyền xử lý công văn cảnh báo
slug: day-chuyen-xu-ly-cv-canh-bao
status: active
group: Tài chính
description: Thu thập, OCR, nạp sổ F0/F1, sắp xếp lưu trữ công văn cảnh báo. QC gác cổng từng bước, kết thúc sau sắp xếp lưu trữ.
steps:
  - agent: tra-cuu-tai-cv-ioffice
    task: |-
      THU THẬP. Đầu vào: {{input}}.

      {{input}} phải nêu rõ NGUỒN và PHẠM VI. Nếu không nêu rõ, trả về đúng một dòng
      "DUNG: dau vao khong neu ro nguon va pham vi" rồi kết thúc. KHÔNG hỏi lại, KHÔNG tự đoán.

      Ba nguồn hợp lệ:
      A. iOffice: đăng nhập VNPT iOffice Đồng Nai, tra theo trích yếu hoặc số CV hoặc khoảng ngày,
         tải tệp về D:\HO SO CONG VAN CANH BAO theo năm và tháng, giải nén zip.
      B. Google Drive: liệt kê tệp trong thư mục được chỉ định. HỘP THƯ ĐẾN chính thức, chốt 09/11/2026:
         G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\HO SO MOI
         Đường dẫn cũ 01_HO SO CONG AN CHUYEN DEN\2026\HO SO MOI ĐÃ KHÔNG CÒN TỒN TẠI, đừng dùng.
         Mặc định lấy TẤT CẢ tệp trong thư mục, không lọc theo số CV, trừ khi {{input}} ghi phạm vi hẹp hơn.
      C. Telegram hoặc Zalo: HIỆN CHƯA CÓ TOOL để agent tự kéo tệp về. Chỉ làm được khi Trưởng phòng
         đã đính kèm tệp và {{input}} ghi đường dẫn tuyệt đối của tệp đó. {{input}} chỉ nói "tệp gửi
         qua Zalo" mà không kèm đường dẫn thì trả "DUNG: nhanh Zalo chua co tool, can duong dan tuyet doi".

      BẮT BUỘC: liệt kê thư mục bằng tool rồi mới trả lời. Không được trả lời từ suy đoán.

      ĐẦU RA bắt buộc là bảng, mỗi tệp một dòng, đủ 4 cột:
      | Đường dẫn tuyệt đối | Tên tệp | Dung lượng | Loại (PDF text, PDF scan, XLSX) |

      Sau bảng ghi một dòng "TONG: <n> tep".

      CẤM trả lời "đã sẵn sàng", "sẵn sàng khi bạn gọi", "chờ lệnh", hoặc bất kỳ câu hỏi nào.
      Không có tệp nào thì trả "DUNG: thu muc rong".
      KHÔNG OCR, KHÔNG nạp sổ ở bước này.
    verify_agent: kiem-soat-chat-luong-du-lieu
    max_retries: 1
  - agent: so-hoa-ocr-cong-van
    task: |-
      SỐ HOÁ. Đầu vào là bảng tệp ở {{prev}}, phạm vi gốc ở {{input}}.

      Nếu {{prev}} không chứa bảng có đường dẫn tuyệt đối, trả về đúng một dòng
      "DUNG: khong co danh sach tep tu buoc thu thap" rồi kết thúc.

      Với mỗi tệp PDF hoặc ảnh trong bảng:
      1. Thử trích lớp chữ số hoá trước bằng pymupdf. Có chữ thì dùng luôn, nhanh hơn OCR nhiều lần.
      2. Không có lớp chữ thì OCR offline bằng skill pdf-sang-md.
      3. Ghi kết quả .md vào sources/ của brain này.
      4. Tệp .xlsx thì KHÔNG OCR, đọc trực tiếp bằng openpyxl rồi xuất bảng dữ liệu.

      Tệp scan nặng trên 5 MB thì xử lý sau cùng và ghi rõ thời gian đã tốn.

      ĐẦU RA bắt buộc, mỗi tệp một dòng:
      | Tệp gốc | Đường dẫn .md đã ghi | Số ký tự | PDF-da-trich hay PDF-chua-trich |

      CẤM hỏi lại, CẤM trả lời bằng câu hỏi, CẤM báo xong khi chưa có đường dẫn .md thật.
    verify_agent: kiem-soat-chat-luong-du-lieu
    max_retries: 1
  - agent: cập-nhật-công-văn-cảnh-báo-rủi-ro
    task: |-
      BÓC TÁCH VÀ NẠP SỔ. Đầu vào là các tệp .md ở {{prev}}.

      Nếu {{prev}} không chứa đường dẫn .md nào tồn tại thật, trả về đúng một dòng
      "DUNG: khong co file md de boc tach" rồi kết thúc.

      Sổ theo dõi: G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx

      CẤU TRÚC SỔ THẬT, kiểm trực tiếp 09/11/2026. Dùng đúng cái này, bỏ spec "25 cột cho cả hai sheet":
      - F0: 19 cột. Dòng 5-6 tiêu đề hiển thị, dòng 7 khoá máy, DỮ LIỆU TỪ DÒNG 8.
        Khoá cột: STT, MSTBan, TenBan, SoCV, NgayCV, DonViBH, NoiDung, TamNgung, NhomND, DiaChi,
        NgayNhan, TinhTrang, GhiChu, File, NamCV, Trung, KhoaCV, CVDau, MSTDau.
      - F1: 25 cột, cùng kiểu tiêu đề, dữ liệu từ dòng 8. Thêm TenMua, MSTMua, KyKK, SoHD,
        DoanhThu, ThueGTGT so với F0.
      - Mốc hiện tại: F0 17.075 dòng, F1 3.761 dòng.

      NGUỒN CHUẨN, chốt 09/11/2026: lô nào có tệp .xlsx hoặc .csv danh sách doanh nghiệp thì LẤY
      EXCEL LÀM CHUẨN cho MST, tên công ty, địa chỉ, ngành nghề. Bản OCR chỉ dùng để lấy SoCV,
      NgayCV, DonViBH, NoiDung. Lý do: OCR tiếng Việt đang mất dấu nặng nên tên doanh nghiệp lấy
      từ OCR là rác.

      Thứ tự BẮT BUỘC:
      1. Đếm và ghi lại số dòng F0 và F1 TRƯỚC khi ghi.
      2. Sao lưu sổ vào thư mục con "Sao lưu" với tên .bak_<YYYYMMDD_HHMMSS>.xlsx.
         KHÔNG ghi sổ khi chưa sao lưu xong.
      3. Bóc tách các trường chuẩn. Bên bán vào F0, bên mua vào F1.
      4. Lọc trùng: F0 theo SoCV cộng MSTBan. F1 theo SoCV cộng MSTBan cộng MSTMua.
      5. F1 BỎ QUA hóa đơn không có MST bên mua, không nạp dòng trống MSTMua. Đếm số dòng đã bỏ.
      6. Ghi dòng mới, rồi đếm lại số dòng F0 và F1.

      ĐẦU RA bắt buộc, đúng các dòng này:
      - Đường dẫn bản sao lưu vừa tạo
      - F0 truoc: <n> | F0 sau: <n> | da nap: <n> | bo trung: <n>
      - F1 truoc: <n> | F1 sau: <n> | da nap: <n> | bo trung: <n>
      - F1 bo qua vi thieu MST ben mua: <n> dong
      - Danh sách Số CV đã nạp

      CẤM bịa số liệu. Trường nào không đọc được trong .md thì để trống và ghi vào mục thiếu,
      không suy đoán MST hay tên công ty.
    verify_agent: kiem-soat-chat-luong-du-lieu
    max_retries: 1
  - agent: cập-nhật-công-văn-cảnh-báo-rủi-ro
    task: |-
      SẮP XẾP LƯU TRỮ. Đầu vào là kết quả nạp sổ ở {{prev}}.

      Nếu {{prev}} không xác định được cơ quan ban hành và số công ty F0, trả về đúng một dòng
      "DUNG: chua du thong tin de dat ten" rồi kết thúc.

      MẪU TÊN, chốt 09/11/2026: <Năm>_<Số CV>_<CO QUAN BAN HANH>_<SO DN F0>_<HAU TO>
      - Dấu "/" trong số CV đổi thành "_". Ví dụ 7268/CT-KTr thành 7268_CT-KTr.
      - <SO DN F0> là SỐ LƯỢNG doanh nghiệp bên bán nêu trong công văn, ví dụ 28.
      - <HAU TO> bắt buộc, để phân biệt các tệp cùng một công văn: CV, TT, TRIENKHAI, DS, PL.
        Không có hậu tố thì nhiều tệp cùng công văn sẽ trùng tên và đè nhau.
      - Ví dụ: 2026_7268_CT-KTr_CUCTHUE_28_CV.pdf, 2026_7268_CT-KTr_CUCTHUE_28_TT.pdf

      PHÂN NHÁNH THEO NGUỒN GỐC, KHÔNG THEO CƠ QUAN KÝ, chốt 09/11/2026:
      A. Hồ sơ có GỐC từ Công an, Viện kiểm sát, Tòa án, kể cả khi cơ quan thuế ký công văn
         chuyển tiếp: tạo thư mục mới theo mẫu tên trên trong
         G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\01_HO SO CONG AN CHUYEN DEN
         rồi chuyển toàn bộ tệp của lô vào đó, mỗi tệp đặt tên theo mẫu kèm hậu tố.
         Ví dụ: CV 7268 do Cục Thuế ký nhưng gốc là CV B1246 của CSĐT Công an Hà Tĩnh, đi nhánh A.
      B. Hồ sơ THUẦN nội bộ ngành thuế, không có công văn Công an làm gốc: chuyển tệp vào
         G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\02_NGUON RUI RO NGANH THUE
         và đổi tên theo mẫu trên, giữ đuôi tệp gốc.

      Di chuyển bằng cách COPY trước, kiểm tệp đích tồn tại và đúng dung lượng, SAU ĐÓ XOÁ TỆP NGUỒN.
      Trưởng phòng đã cho phép xoá tệp nguồn trên ổ G ngày 09/11/2026. Dung lượng đích lệch thì GIỮ
      tệp nguồn và báo lỗi.

      ĐẦU RA bắt buộc: mỗi tệp một dòng "<đường dẫn cũ>  ->  <đường dẫn mới>".

      ĐÂY LÀ BƯỚC CUỐI. QC bước này đạt thì công việc KẾT THÚC: tổng hợp số liệu nạp sổ và
      danh sách tệp đã sắp xếp rồi báo về Trưởng phòng.
      TUYỆT ĐỐI KHÔNG tự chạy đối chiếu NCC, KHÔNG tự soạn công văn, KHÔNG tự gọi quy trình
      tra-loi-ho-so-cong-an. Hai khâu đó chỉ chạy khi Trưởng phòng yêu cầu rõ bằng một lệnh riêng.
    verify_agent: kiem-soat-chat-luong-du-lieu
    max_retries: 1
updated: "2026-11-09"
---

Quy trình trục của phòng cảnh báo, gồm **4 bước làm việc thật**: thu thập, số hoá, nạp sổ F0/F1, sắp xếp lưu trữ.

### Điểm kết thúc công việc

Công việc **KẾT THÚC sau bước 4 khi QC đạt**. Nghĩa là: nạp sổ F0/F1 xong, tệp đã sắp xếp lưu trữ đúng thư mục và đúng tên, QC gác cổng cả hai bước trả `pass: true`. Lúc đó trợ lý tổng hợp số liệu và báo Trưởng phòng, rồi dừng.

**Đối chiếu NCC và soạn thảo công văn chỉ thực hiện khi Trưởng phòng yêu cầu rõ.** Không yêu cầu thì không chạy, không tự đề xuất chạy giữa dây chuyền, không coi là việc còn dở. Hai khâu đó nằm ở quy trình riêng `tra-loi-ho-so-cong-an`, khởi động bằng một lệnh riêng của Trưởng phòng.

### QC gác cổng bằng verify_agent

Mỗi bước có `verify_agent: kiem-soat-chat-luong-du-lieu`. Agent QC trả JSON `{"pass": true|false}`.
Trượt hết `max_retries` thì **workflow dừng thật**, không chạy tiếp các bước sau.

Đây là điểm khác hẳn thiết kế cũ. Trước đây QC là một bước thường, nên khi QC báo lỗi thì câu báo lỗi đó lại trở thành đầu vào `{{prev}}` của bước sau, và dây chuyền cứ chạy tiếp trên dữ liệu rỗng. Lần chạy ngày 09/11/2026 lúc 00:38 vỡ đúng theo cách này.

### Ba luật chống vỡ dây chuyền

1. **Trợ lý nền không hỏi lại được.** Không có ai trả lời lúc nó chạy. Thiếu đầu vào thì mỗi bước
   phải trả một dòng `DUNG: <lý do>` rồi kết thúc, tuyệt đối không viết câu hỏi vào đầu ra.
2. **Đầu ra phải là sản phẩm, không phải lời hứa.** Mỗi bước bắt buộc trả bảng hoặc con số cụ thể.
   Câu "đã sẵn sàng" hay "đã hoàn thành" mà không kèm đường dẫn hoặc số liệu bị QC tính là trượt.
3. **Đầu vào phải tuyệt đối.** Khi giao việc, `intent` phải ghi đường dẫn đầy đủ. Không để trợ lý
   tự đi tìm thư mục.

### Giao việc đúng cách

Giao cả dây chuyền: `javis_task op=add route="wf:day-chuyen-xu-ly-cv-canh-bao"`, `intent` ghi rõ nguồn và đường dẫn tuyệt đối.

Chỉ cần một khâu thì đừng gọi workflow, giao thẳng cho một agent với `route="auto"`, hoặc bảo Javis làm ngay trong chat.

### Điều kiện khác
Căn cứ luật, nghị định, mức phạt không nằm trong chuỗi này. Gọi trợ lý `phap-che-chinh-sach-thue` ở brain **Cung cấp thông tin công an**.
