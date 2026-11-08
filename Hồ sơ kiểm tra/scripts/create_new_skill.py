import os

skill_content = """---
name: Hồ sơ kiểm tra theo kế hoạch
description: "Lập và xuất bộ hồ sơ kiểm tra thuế theo kế hoạch (DI-01 đến DI-07, tổng hợp trình ký) căn cứ QĐ 1412, QĐ 1161, Luật 108/2025."
group: Tài chính
---

# Hồ sơ kiểm tra thuế theo kế hoạch

Quy trình chuẩn lập, cập nhật và xuất trọn bộ hồ sơ kiểm tra việc chấp hành pháp luật thuế tại trụ sở người nộp thuế theo kế hoạch năm (nhóm văn bản hồ sơ đi DI-01 đến DI-07 và file Word tổng hợp trình ký).

## Khi nào dùng

- Lập hồ sơ kiểm tra NNT thuộc danh sách kế hoạch kiểm tra năm (QĐ 1412/QĐ-DON).
- Xuất bộ hồ sơ đi (nhóm DI) hoặc xuất 1 file Word tổng hợp hồ sơ đi để trình ký lãnh đạo trước khi có Số QĐKT chính thức.
- Cập nhật thông tin NNT từ DB NNT, thời kỳ kiểm tra, nội dung rủi ro chuyên đề (giá vốn, chi phí quản lý, chi phí bán hàng...) lên sổ Excel JAVIS PRO và app.
- Phân biệt với: kiểm tra giải thể / đóng mã số thuế (thuộc quy trình giải thể tại bàn / tại trụ sở), kiểm tra hoàn thuế, xử phạt vi phạm hành chính.

## Căn cứ pháp lý áp dụng

1. **Kế hoạch năm:** Quyết định số 1412/QĐ-DON ngày 06/05/2026 của Trưởng Thuế Thành phố Đồng Nai về việc phê duyệt kế hoạch kiểm tra tại trụ sở người nộp thuế năm 2026.
2. **Quy trình kiểm tra:** Quyết định số 1161/QĐ-CT ngày 14/8/2026 của Cục Thuế ban hành Quy trình kiểm tra thuế (Mẫu 08/QTKT-TC phân tích rủi ro, Mẫu 12/QTKT-TC báo cáo kết quả...).
3. **Quy chế giám sát:** Quyết định số 1614/QĐ-TCT ngày 13/10/2022 của Tổng cục trưởng Tổng cục Thuế về Quy chế giám sát hoạt động của Đoàn thanh tra, kiểm tra thuế.
4. **Giao quyền:** Quyết định số 2618/QĐ-TCS3 ngày 20/05/2026 của Trưởng Thuế Cơ sở 3 về việc giao quyền ra Quyết định kiểm tra thuế tại trụ sở NNT.
5. **Luật Quản lý thuế:** Luật Quản lý thuế số 108/2025/QH15 ngày 10/12/2025 (Thời hạn kiểm tra tại trụ sở NNT không quá 20 ngày làm việc).

## Danh mục 7 mẫu Hồ sơ đi (Nhóm DI)

| Mã | Tên văn bản | Biểu mẫu / Căn cứ | Mục đích |
| :--- | :--- | :--- | :--- |
| **DI-02** | Tờ trình ban hành QĐ kiểm tra / giám sát | Mẫu tờ trình nội bộ | Trình Lãnh đạo Thuế cơ sở phê duyệt QĐKT và QĐGS |
| **DI-03** | Quyết định kiểm tra thuế tại trụ sở NNT | Mẫu 06.KTT / Luật 108/2025 | QĐ thành lập Đoàn KT, nội dung, thời kỳ, thời hạn KT |
| **DI-05** | Quyết định giám sát đoàn kiểm tra | Mẫu QĐGS / QĐ 1614/QĐ-TCT | Phân công người giám sát hoạt động của Đoàn KT |
| **DI-01** | Bản phân tích thông tin hồ sơ NNT | Mẫu 08/QTKT-TC (QĐ 1161) | Phân tích báo cáo tài chính, tờ khai, rủi ro trước kiểm |
| **DI-07** | Thư trao đổi thông tin với NNT | Mẫu thư hành chính | Hỗ trợ NNT, công khai đường dây nóng giám sát liêm chính |
| **DI-04** | Giấy xác nhận giao nhận QĐ kiểm tra | Mẫu 3.1 biên nhận | Giao trực tiếp QĐKT cho đại diện NNT khi công bố |
| **DI-06** | Kế hoạch giám sát hoạt động đoàn KT | Mẫu kèm QĐ 1614/QĐ-TCT | Lập kế hoạch tiến độ, nội dung và hình thức giám sát |

**File tổng hợp:** `MAU TONG HOP HO SO DI.docx` ghép toàn bộ 7 mẫu trên theo đúng thứ tự trình ký, ngăn cách bằng ngắt trang chuẩn.

## Quy trình 4 bước thực hiện

### Bước 1: Tra cứu thông tin và phân tích dữ liệu NNT
- Tra `I:\\My Drive\\DB NNT.xls` theo MST để lấy: Tên công ty, địa chỉ, người đại diện pháp luật (Giám đốc), CCCD, Kế toán trưởng, ngành nghề kinh doanh.
- Thu thập số liệu doanh thu, giá vốn, chi phí quản lý qua các năm để đưa vào Mục III Bản phân tích thông tin (DI-01).

### Bước 2: Cập nhật sổ theo dõi Excel JAVIS PRO
- Mở `Danh sach kiem tra quyet toan pro.xlsm`.
- Sheet làm việc: `01_DA_BAN_HANH` (nếu xuất trình ký hoặc đã có số) hoặc `02_DANG_THUC_HIEN`.
- Thời kỳ kiểm tra: `Từ năm YYYY đến năm YYYY`.
- Nội dung kiểm tra: trình bày xuống dòng, gạch đầu dòng, đưa **giá vốn hàng bán** lên trước chi phí quản lý / chi phí bán hàng:
  ```
  - Giá vốn hàng bán.
  - Chi phí bán hàng.
  - Chi phí quản lý doanh nghiệp.
  ```
- Phân công đoàn: Trưởng đoàn, Thành viên.

### Bước 3: Xuất bộ hồ sơ đi trình ký (Word)
- Kiểm tra app nội bộ (`D:\\01. Cong tac kiem tra\\HO SO KIEM TRA\\app`).
- Nếu chưa có Số QĐKT chính thức: sử dụng mã tạm `DU_THAO-TRINH-KY`.
- Xuất 7 file riêng lẻ (`DI-01` đến `DI-07`) và file tổng hợp `TONG HOP HO SO DI.docx` vào thư mục vụ việc:
  `G:\\Other computers\\My Laptop\\1. CONG TAC KIEM TRA\\1- KIEM TRA\\1. KIEM TRA TAI TRU SO NNT\\2026\\DANG KIEM TRA\\KE HOACH\\<TEN_DOANH_NGHIEP>\\0. BIEN BAN\\`
- Đóng gói file nén `hoso.zip` để lưu trữ và gửi kèm.

### Bước 4: Hoàn thiện sau khi phê duyệt
- In/trình duyệt bộ file Word (hoặc file tổng hợp).
- Khi Lãnh đạo ký duyệt và cấp Số QĐKT chính thức: cập nhật lại Số QĐKT, ngày ban hành vào sổ Excel và app để chuyển sang giai đoạn thực hiện kiểm tra (nhóm BB).

## Bẫy nghiệp vụ & Lưu ý

- **Tuyệt đối không nhầm với hồ sơ giải thể:** Hồ sơ giải thể căn cứ thông báo giải thể, NĐ 01/2021; còn hồ sơ kế hoạch căn cứ QĐ 1412/QĐ-DON phê duyệt kế hoạch năm.
- **Thời hạn kiểm tra:** Luật 108/2025/QH15 quy định không quá **20 ngày làm việc** (không phải 10 ngày như Luật cũ).
- **Mã tạm khi trình ký:** Luôn dùng `DU_THAO-TRINH-KY`, không tự ý bịa số QĐKT.
- **Thứ tự nội dung kiểm tra:** Giá vốn hàng bán luôn đứng trước Chi phí QLDN / Chi phí bán hàng.
"""

ref_so_sanh = """# So sánh Hồ sơ kiểm tra theo Kế hoạch vs Hồ sơ kiểm tra Giải thể

| Tiêu chí | Kiểm tra theo Kế hoạch năm | Kiểm tra Giải thể (đóng MST) |
| :--- | :--- | :--- |
| **Căn cứ quyết định** | Quyết định số 1412/QĐ-DON phê duyệt kế hoạch năm của Trưởng Thuế Tỉnh/TP | Thông báo giải thể của DN, đề nghị của bộ phận KK&KTT, NĐ 01/2021 |
| **Tiêu đề Tờ trình (DI-02)** | *"V/v: Ban hành quyết định kiểm tra thuế; Quyết định giám sát hoạt động của Đoàn kiểm tra"* | *"V/v: Ban hành quyết định kiểm tra quyết toán giải thể; Quyết định giám sát..."* |
| **Mục đích cuộc kiểm tra** | Đánh giá chấp hành pháp luật thuế theo chuyên đề rủi ro được phê duyệt hàng năm | Quyết toán toàn bộ nghĩa vụ thuế phát sinh đến thời điểm giải thể để đóng MST |
| **Phân tích rủi ro (DI-01)** | Tập trung tỷ lệ giá vốn/doanh thu, chi phí quản lý/doanh thu, biến động qua các năm | Toàn diện tình hình tài chính, công nợ, hàng tồn kho, tài sản thanh lý đến ngày giải thể |
| **Thời hạn kiểm tra** | Không quá 20 ngày làm việc tại trụ sở NNT (Luật 108/2025/QH15) | Thường kiểm tra tại bàn hoặc tại trụ sở theo quy trình giải thể |
| **Nhóm văn bản đi** | DI-01 đến DI-07 theo mẫu kế hoạch năm | DI-01 đến DI-07 theo mẫu quyết toán giải thể |
"""

ref_mau = """# Danh mục chi tiết 7 biểu mẫu Hồ sơ đi theo Kế hoạch

## 1. DI-02: Tờ trình ban hành QĐ kiểm tra / giám sát
- **Mã file sạch:** `DI-02.docx`
- **Người ký đề xuất:** Trưởng đoàn kiểm tra (Cán bộ đề xuất).
- **Người duyệt:** Tổ trưởng Tổ Kiểm tra số 1 -> Phó Trưởng Thuế Cơ sở 3 (Phạm Đồng Xuân).
- **Căn cứ chính:** QĐ 1614/QĐ-TCT, QĐ 1412/QĐ-DON, Luật 108/2025/QH15.

## 2. DI-03: Quyết định kiểm tra tại trụ sở NNT
- **Mã file sạch:** `DI-03.docx` (Mẫu 06.KTT)
- **Thẩm quyền ký:** Phó Trưởng Thuế Cơ sở 3 (theo QĐ giao quyền 2618/QĐ-TCS3).
- **Nội dung:** Ghi rõ từng chuyên đề kiểm tra (xuống dòng gạch đầu dòng), thời kỳ, thời hạn 20 ngày làm việc, thành phần đoàn.

## 3. DI-05: Quyết định giám sát hoạt động đoàn kiểm tra
- **Mã file sạch:** `DI-05.docx`
- **Người giám sát:** Ông Nguyễn Trường Giang – Tổ trưởng Tổ Kiểm tra thuế 1.
- **Căn cứ:** QĐ 1614/QĐ-TCT, QĐ 1412/QĐ-DON.

## 4. DI-01: Bản phân tích thông tin, xác định nội dung cuộc kiểm tra
- **Mã file sạch:** `DI-01.docx` (Mẫu 08/QTKT-TC theo QĐ 1161/QĐ-CT)
- **Nội dung:** Thông tin chung NNT, tình hình chấp hành pháp luật thuế, phân tích rủi ro BCTC/tờ khai theo kế hoạch, đề xuất nội dung kiểm tra.

## 5. DI-07: Thư trao đổi thông tin với NNT
- **Mã file sạch:** `DI-07.docx`
- **Mục đích:** Thông báo số điện thoại lãnh đạo phụ trách (Ông Phạm Đồng Xuân: 0913.327.749) và Người giám sát (Ông Nguyễn Trường Giang: 0916.369.389).

## 6. DI-04: Giấy xác nhận giao nhận Quyết định kiểm tra
- **Mã file sạch:** `DI-04.docx` (Mẫu 3.1)
- **Thành phần:** Đại diện đoàn kiểm tra giao trực tiếp cho Giám đốc/Đại diện pháp luật NNT.

## 7. DI-06: Kế hoạch giám sát hoạt động của Đoàn kiểm tra thuế
- **Mã file sạch:** `DI-06.docx` (Mẫu theo QĐ 1614/QĐ-TCT)
- **Nội dung:** Phê duyệt kế hoạch giám sát về trình tự, thủ tục, thời hạn, báo cáo.
"""

targets = [
    r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\skills\ho-so-kiem-tra-ke-hoach',
    r'D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\.claude\skills\ho-so-kiem-tra-ke-hoach',
    r'D:\SECOND BRAIN\javis-os\shared-skills\ho-so-kiem-tra-ke-hoach'
]

for t in targets:
    os.makedirs(t, exist_ok=True)
    os.makedirs(os.path.join(t, 'references'), exist_ok=True)
    with open(os.path.join(t, 'SKILL.md'), 'w', encoding='utf-8') as f:
        f.write(skill_content.strip())
    with open(os.path.join(t, 'references', 'so-sanh-ke-hoach-vs-giai-the.md'), 'w', encoding='utf-8') as f:
        f.write(ref_so_sanh.strip())
    with open(os.path.join(t, 'references', 'mau-ho-so-di-ke-hoach.md'), 'w', encoding='utf-8') as f:
        f.write(ref_mau.strip())
    print('Created skill at:', t)
