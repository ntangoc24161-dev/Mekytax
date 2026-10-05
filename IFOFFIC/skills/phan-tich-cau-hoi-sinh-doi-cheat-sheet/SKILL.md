---
name: Phân tích câu hỏi sinh đôi và Cheat Sheet ôn thi
description: "Phân tích câu hỏi sinh đôi, tìm keyword cốt lõi, bóc tách bẫy đáp án, lập Cheat Sheet mapping và tóm tắt siêu tốc < 300 từ dạng Word/Excel ôn thi."
group: Năng suất
---

# Phân tích câu hỏi sinh đôi, Lập Cheat Sheet & Tóm tắt siêu tốc ôn thi

Kỹ năng chuyên sâu phục vụ rà soát, xử lý và đóng gói ngân hàng câu hỏi trắc nghiệm ôn thi (kiểm tra thuế, công chức, sát hạch nghiệp vụ): nhận diện các cụm câu hỏi "anh em sinh đôi", trích xuất từ khóa cốt lõi (KEYWORD vàng), bóc tách bẫy phương án sai, xây dựng bộ quy tắc phản xạ nhanh (Cheat Sheet mapping từ khóa) và biên soạn tóm tắt siêu tốc "Hack Não" dưới 300 từ cho từng chuyên đề.

## Khi nào sử dụng skill

- Người dùng yêu cầu: "phân tích câu hỏi sinh đôi", "tìm câu hỏi gần giống nhau", "so sánh và tìm điểm khác biệt giữa các câu hỏi", "tìm keyword vàng", "lập cheat sheet mapping từ khóa", "xây dựng bảng mẹo nhận diện nhanh", "tóm tắt siêu tốc hack não dưới 300 từ", "tô đậm phương án đúng trong các file đề thi và ngân hàng".
- Đã có ngân hàng câu hỏi tổng hợp hoặc danh sách bộ đề thi trắc nghiệm (định dạng Excel/Word/JSON).
- Cần tối ưu hóa tốc độ học tập: chuyển từ việc học thuộc lòng thụ động sang ghi nhớ quy luật, phản xạ từ khóa trong 3 giây và tránh bẫy tinh vi của đề thi.
- Cần chuẩn hóa hình thức các file đề thi: chèn tiêu đề chuẩn cỡ 16 in đậm, tô màu đáp án đúng (Bold + nền xanh lá `#E2EFDA`), xuất file mục lục tinh gọn (chỉ lấy Câu 1).

## Quy trình 4 bước thực hiện

### Bước 1: Rà soát & Chuẩn hóa ngân hàng đề thi
1. **Khử trùng lặp**: Đối soát giữa các file đề thi, xác định tỷ lệ trùng câu hỏi và loại bỏ các đề trùng khớp 100%.
2. **Chuẩn hóa Header & Đóng băng**: Chèn Dòng 1 tại tất cả các file đề với định dạng Tên File, Mã Đề, Font chữ Calibri/Arial 16pt, In đậm (Bold), căn giữa, Freeze Panes để cuộn dễ dàng.
3. **Định dạng Highlight đáp án đúng**:
   - Quét cột Đáp án đúng của từng câu (A, B, C hoặc D).
   - Thiết lập thuộc tính Font Bold và Background Fill màu xanh lá nhạt (`#E2EFDA` / `Fill.PatternType = solid`) cho ô phương án tương ứng.
4. **Trích xuất Mục lục tinh gọn (Câu 1 Index)**: Tạo bảng tổng hợp chỉ gồm: STT, Mã Đề, Tên File Excel, Mã Câu 1 (ID), Chuyên đề nghiệp vụ, Nội dung Câu 1 và Đáp án đúng; lược bỏ toàn bộ các cột phương án A, B, C, D để dễ tra cứu nhanh.

### Bước 2: Phân tích Cụm câu hỏi "Anh em sinh đôi" (Twin Questions)
1. **Thuật toán gom cụm câu hỏi gần giống**:
   - Sử dụng độ đo tương đồng văn bản (Levenshtein distance / SequenceMatcher / TF-IDF) kết hợp nhóm theo Chuyên đề và Chủ thể pháp lý.
   - Lọc ra các câu hỏi có độ tương đồng cấu trúc >= 70% hoặc cùng đề cập đến một quy định nhưng có điều kiện ràng buộc khác nhau.
2. **Trích xuất KEYWORD vàng**:
   - Xác định 01 - 03 từ khóa cốt lõi (chủ thể, hành vi, mốc thời gian, điều kiện loại trừ, số tiền...) làm thay đổi hoàn toàn bản chất câu hỏi và đảo chiều đáp án.
   - Tô màu vàng (`#FFF2CC`) cho KEYWORD trên bảng tính Excel và đóng khung nổi bật trong báo cáo.
3. **Phân tích Bẫy đề thi & Mẹo nhận diện**:
   - Bóc tách sự khác biệt tinh vi giữa 4 phương án (A, B, C, D).
   - Nêu rõ bẫy tâm lý thường gặp (nhầm lẫn giữa thủ tục vs số tiền thuế, tự ý vs cho phép, nhẹ vs nặng, trực tiếp vs liên đới).
   - Đưa ra mẹo nhận diện chuẩn xác ngay khi vừa đọc lướt qua câu hỏi.

### Bước 3: Xây dựng Bộ quy tắc nhận diện nhanh (Cheat Sheet Mapping)
1. **Cấu trúc Bảng Mapping từ khóa ➡️ Đáp án**:
   - **Chuyên đề**: Nhóm nghiệp vụ / văn bản quy phạm pháp luật.
   - **Từ khóa nhận diện câu hỏi (Trigger)**: Cụm từ nhận biết trong đề thi.
   - **Từ khóa đáp án đúng buộc chọn (Action)**: Cụm từ bắt buộc phải xuất hiện trong phương án đúng.
   - **Bẫy cần loại trừ ngay (Trap)**: Các phương án gây nhiễu, luật cũ hoặc số liệu sai cần gạch bỏ trong 1 giây.
   - **Công thức / Mẹo nhớ nhanh**: Quy tắc vần điệu, số học hoặc logic để ghi nhớ tức thì.
2. **Quy tắc phản xạ 3 giây**: Thiết kế sao cho thí sinh chỉ cần nhìn thấy Trigger là tự động ánh xạ sang Action mà không cần đọc hết 4 phương án dài.

### Bước 4: Tóm tắt siêu tốc "Hack Não" (< 300 từ/chuyên đề)
1. **Tiêu chuẩn cô đọng kiến thức**:
   - Mỗi chuyên đề nghiệp vụ được nén dưới **300 từ**.
   - Trực diện vào số liệu, thời hạn, tỷ lệ, thẩm quyền, thứ tự ưu tiên và công thức cốt lõi.
   - Kết thúc mỗi chuyên đề bằng một dòng **"Mẹo Hack Não"** súc tích.
2. **Đóng gói đa định dạng**:
   - **File Word (`.docx`)**: Định dạng chuẩn văn bản, tiêu đề rõ ràng, bullet points khoa học, phù hợp in ấn học liệu hoặc đọc trên điện thoại.
   - **File Excel (`.xlsx`)**: Đa sheet (Sheet Twin Questions, Sheet Cheat Sheet, Sheet Index Câu 1, Sheet Master Bank).

## Mẫu kịch bản tự động hóa Python tham khảo

```python
# Kịch bản gom cụm câu hỏi sinh đôi và tạo Cheat Sheet
import re
from difflib import SequenceMatcher
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()

# Highlight phương án đúng
green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
bold_font = Font(name="Calibri", size=11, bold=True)

def standardize_exam_file(file_path, exam_name):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    # Chèn dòng 1 tiêu đề
    ws.insert_rows(1)
    ws.cell(row=1, column=1, value=f"BỘ ĐỀ ÔN THI KIỂM TRA THUẾ - {exam_name}")
    ws.cell(row=1, column=1).font = Font(name="Calibri", size=16, bold=True)
    # Tô đậm phương án đúng
    # ... duyệt các dòng và áp dụng green_fill cho cột đáp án đúng
    wb.save(file_path)
```

## Tiêu chuẩn đầu ra (Deliverables)

1. `Phan_Tich_Cau_Hoi_Sinh_Doi_Va_Cheat_Sheet.xlsx`: Chứa toàn bộ các cặp câu hỏi sinh đôi, bẫy đáp án và bảng mapping từ khóa.
2. `Tom_Tat_Sieu_Toc_Hack_Nao_On_Thi_*.docx`: Tài liệu Word tóm tắt toàn bộ chuyên đề (mỗi chuyên đề < 300 từ).
3. `Danh_Sach_Tong_Hop_40_De_Va_Cau_1.xlsx`: Bảng mục lục tinh gọn tra cứu Câu 1 của các đề.
4. Toàn bộ các file đề thi gốc được chuẩn hóa Dòng 1 cỡ 16 in đậm và tô đậm phương án đúng kèm nền xanh nhạt.
