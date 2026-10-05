---
name: Chuẩn hóa và xuất ngân hàng trắc nghiệm
description: "Chuẩn hóa ngân hàng trắc nghiệm Excel sang Word: gom nhóm chuyên đề, sắp câu sinh đôi, tô vàng keyword, in đậm đáp án chuẩn, format in ấn."
group: Năng suất
---

# Chuẩn hóa và xuất ngân hàng trắc nghiệm (Excel ➡️ Word & Excel chuẩn)

Kỹ năng tổng quát phục vụ xử lý, chuẩn hóa, sắp xếp câu hỏi tương đồng và chuyển đổi ngân hàng đề thi trắc nghiệm từ bảng tính Excel sang tài liệu Word in ấn và bảng Excel Master khoa học cho mọi kỳ thi.

## Khi nào dùng

- Người dùng yêu cầu: "chuyển file excel ngân hàng câu hỏi sang word", "format đề thi trắc nghiệm in đậm đáp án đúng", "sắp xếp câu hỏi theo chuyên đề và gom câu sinh đôi", "tô vàng từ khóa keyword trong câu hỏi", "tạo file word ôn tập trắc nghiệm", "chuẩn hóa bảng excel đề thi".
- Khi cần xử lý hàng loạt câu hỏi trắc nghiệm (100 đến 2.000+ câu) từ nhiều nguồn/nhiều đề thi thành một tài liệu học tập hoặc ngân hàng đề chuẩn mực.
- Cần tối ưu bố cục in ấn: font chữ rõ ràng, giãn dòng hợp lý, phân tách chuyên đề trực quan, nhận diện nhanh phương án đúng.

## Chuẩn bị & Cấu trúc dữ liệu đầu vào

File Excel nguồn cần tối thiểu các trường thông tin (hoặc có thể trích xuất được):
- Nội dung câu hỏi.
- 4 phương án lựa chọn: Phương án A, B, C, D.
- Đáp án đúng / Phương án chuẩn (ký tự A, B, C, D hoặc nội dung phương án).
- (Tùy chọn) Chuyên đề / Phân loại nghiệp vụ, Mã câu hỏi hệ thống.

## Quy trình thực hiện 4 bước

### Bước 1: Chuẩn hóa bảng tính Excel Master
1. **Cấu trúc 8 - 9 cột chuẩn**:
   - `STT`: Đánh số thứ tự tăng dần từ 1 đến hết.
   - `Mã Câu`: Giữ lại mã định danh hệ thống (nếu có) để đối soát.
   - `Chuyên Đề`: Gán nhóm chủ đề / môn học tương ứng.
   - `Nội Dung Câu Hỏi`: Văn bản câu hỏi đã làm sạch khoảng trắng thừa.
   - `Phương Án A`, `Phương Án B`, `Phương Án C`, `Phương Án D`: Giữ nguyên đúng thứ tự các phương án.
   - `Đáp Án Chuẩn`: Ghi rõ ký tự (A, B, C, D).
2. **Định dạng màu sắc trực quan**:
   - Ô Đáp Án Chuẩn và ô Phương án đúng tương ứng được tô màu **chữ xanh lá đậm (#006100)** trên **nền xanh lá nhạt (#E2EFDA)**.
   - Dòng tiêu đề: Cỡ chữ 11-12pt, In đậm (Bold), nền xám nhạt hoặc xanh đậm, cố định dòng (Freeze Panes Dòng 1).
3. **Thống kê chuyên đề**: Tự động tạo thêm sheet `Thong_Ke_Chuyen_De` tính tổng số câu và tỷ trọng (%) cho từng chuyên đề.

### Bước 2: Gom nhóm Chuyên đề & Sắp xếp câu hỏi sinh đôi (Twin Questions)
1. **Phân loại chuyên đề**: Sắp xếp toàn bộ câu hỏi theo từng chuyên đề/chương bài học.
2. **Gom cụm câu hỏi tương đồng / sinh đôi**:
   - Trong từng chuyên đề, tính toán độ tương đồng văn bản (Levenshtein ratio >= 0.70 hoặc trùng cụm từ khóa chính).
   - Đặt các câu hỏi có cùng cấu trúc câu, cùng tình huống nhưng khác điều kiện/mốc thời gian/số liệu nằm **liền kề nhau**.
   - Giúp người học nhận diện ngay sự tương phản và phân biệt bẫy đề thi khi đọc liên tiếp.

### Bước 3: Nhận diện & Highlight Từ khóa cốt lõi (KEYWORD)
1. **Bộ quy tắc nhận diện Keyword**:
   - **Mốc thời gian / Thời hạn**: "30 ngày", "05 năm", "chậm nhất", "kể từ ngày".
   - **Định mức / Số liệu**: "20 triệu", "10%", "từ 02 người trở lên".
   - **Chủ thể / Thẩm quyền**: Tên cơ quan, chức danh ban hành, đối tượng áp dụng.
   - **Điều kiện loại trừ / Phủ định**: "KHÔNG ĐƯỢC", "NGOẠI TRỪ", "ĐƯỢC MIỄN", "TRỪ TRƯỜNG HỢP".
2. **Quy cách tô màu**:
   - Trên Excel: Tô nền vàng (`#FFF2CC`) hoặc định dạng chữ nổi bật.
   - Trên Word: Sử dụng tính năng Highlight màu vàng (Yellow Highlight) cho từ khóa trong câu hỏi và phương án đúng.

### Bước 4: Xuất tài liệu Word (.docx) chuyên nghiệp
1. **Tiêu đề tài liệu**: Đặt ở đầu trang, cỡ chữ 16pt, In đậm (Bold), căn giữa (ví dụ: `TỔNG HỢP ĐỀ THI ÔN TẬP`).
2. **Thanh phân tách Chuyên đề**: Mỗi chuyên đề có thanh tiêu đề riêng: `CHUYÊN ĐỀ [STT]: [TÊN CHUYÊN ĐỀ] ([SỐ LƯỢNG] CÂU)` (13pt, In đậm, màu xanh Navy hoặc xám viền).
3. **Quy cách từng câu hỏi**:
   - Tiêu đề câu: **`Câu hỏi [STT]: `** in đậm kèm nội dung câu hỏi.
   - Các phương án A, B, C, D: Mỗi phương án nằm trên **một dòng riêng biệt**, thụt đầu dòng (indent 0.25 inch).
   - **In đậm toàn bộ phương án đúng** theo cột đáp án chuẩn.
   - Font chữ: **Times New Roman** hoặc Arial, cỡ chữ 11-12pt, giãn dòng 1.15 line, cách đoạn 2-3pt.

## Kịch bản Python mẫu (openpyxl & python-docx)

```python
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def create_exam_word_doc(questions, output_path, title="TỔNG HỢP ĐỀ THI ÔN TẬP"):
    doc = docx.Document()
    
    # Thiết lập lề trang chuẩn A4
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        
    # Tiêu đề tài liệu
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(title)
    title_run.font.name = "Times New Roman"
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    
    current_topic = None
    for q in questions:
        # Phân tách chuyên đề mới
        if q.get("chuyen_de") != current_topic:
            current_topic = q.get("chuyen_de")
            topic_p = doc.add_paragraph()
            topic_p.paragraph_format.space_before = Pt(12)
            topic_p.paragraph_format.space_after = Pt(6)
            topic_run = topic_p.add_run(f"=== {current_topic.upper()} ===")
            topic_run.font.name = "Times New Roman"
            topic_run.font.size = Pt(13)
            topic_run.font.bold = True
            topic_run.font.color.rgb = RGBColor(0, 32, 96)
            
        # Nội dung câu hỏi
        q_p = doc.add_paragraph()
        q_p.paragraph_format.space_before = Pt(4)
        q_p.paragraph_format.space_after = Pt(2)
        q_label = q_p.add_run(f"Câu hỏi {q['stt']}: ")
        q_label.font.name = "Times New Roman"
        q_label.font.bold = True
        
        q_text = q_p.add_run(q['noi_dung'])
        q_text.font.name = "Times New Roman"
        
        # 4 phương án A, B, C, D
        for opt_key in ["A", "B", "C", "D"]:
            opt_text = q.get(f"pa_{opt_key}", "")
            opt_p = doc.add_paragraph()
            opt_p.paragraph_format.left_indent = Inches(0.25)
            opt_p.paragraph_format.space_before = Pt(1)
            opt_p.paragraph_format.space_after = Pt(1)
            
            run = opt_p.add_run(f"{opt_key}. {opt_text}")
            run.font.name = "Times New Roman"
            
            # In đậm phương án đúng
            if q.get("dap_an", "").strip().upper() == opt_key:
                run.font.bold = True
                
    doc.save(output_path)
```

## Bẫy & Lưu ý cần tránh

- **Không đổi thứ tự phương án**: Khi xử lý dữ liệu, tuyệt đối giữ nguyên thứ tự 4 phương án A, B, C, D ban đầu của đề gốc.
- **Không bỏ sót câu hỏi**: Luôn đối chiếu số lượng câu hỏi trước và sau khi xử lý (ví dụ: tổng 1.204 câu phải đủ 1.204 câu).
- **Làm sạch ký tự rác**: Loại bỏ các ký tự xuống dòng bất thường `\r\n` bên trong một phương án để tránh bị ngắt dòng sai.
- **Xử lý đáp án khuyết**: Với các câu chưa có đáp án chính thức, ghi chú rõ nguồn hoặc trạng thái thay vì tự ý chọn đáp án khi chưa kiểm chứng.

## Tiêu chuẩn nghiệm thu sản phẩm

1. **File Word (`.docx`)**: Đầy đủ 100% câu hỏi, tiêu đề in đậm căn giữa, phân tách chuyên đề rõ nét, phương án đúng được in đậm rõ ràng, sẵn sàng in ấn không bị lỗi tràn trang.
2. **File Excel (`.xlsx`)**: Đầy đủ các cột chuẩn, Freeze Panes dòng đầu, tô màu xanh lá (`#E2EFDA`) cho đáp án đúng, có sheet thống kê số lượng theo chuyên đề.
