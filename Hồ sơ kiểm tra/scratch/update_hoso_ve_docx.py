import docx
import re

doc_path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\1. TD\12. TT_BIEN_HOA_3603883505\0. BIEN BAN\TONG HOP.docx"
doc = docx.Document(doc_path)

# Số liệu:
# Truy thu TNDN: 14.023.597 đồng
# Truy thu GTGT: 0 đồng
# Tổng truy thu: 14.023.597 đồng
# Tiền chậm nộp TNDN: 1.022.320 đồng
# Tiền chậm nộp GTGT: 0 đồng
# Tổng chậm nộp: 1.022.320 đồng
# Phạt 20%: 2.804.719 đồng
# Phạt thủ tục: 0 đồng
# Tổng phạt: 2.804.719 đồng
# Tổng số tiền truy thu, phạt, chậm nộp: 17.850.636 đồng (Mười bảy triệu tám trăm năm mươi nghìn sáu trăm ba mươi sáu đồng)
# Giảm lỗ: 0 đồng
# Giảm khấu trừ: 0 đồng

# Cập nhật các đoạn văn bản cụ thể
for p in doc.paragraphs:
    # 1. Giữ nội dung Chi phí quản lý kinh doanh nếu có chỗ nào còn bị đổi
    # (Người dùng yêu cầu giữ nội dung Chi phí quản lý kinh doanh)
    if "Chi phí quản lý doanh nghiệp" in p.text:
        p.text = p.text.replace("Chi phí quản lý doanh nghiệp", "Chi phí quản lý kinh doanh")

    # 2. Thay thế số tiền BBKT (P325 - P336)
    if "Phạt tiền với số tiền: - đồng (Bằng chữ -)" in p.text:
        p.text = p.text.replace("Phạt tiền với số tiền: - đồng (Bằng chữ -)", "Phạt tiền với số tiền: 2.804.719 đồng (Bằng chữ: Hai triệu tám trăm lẻ bốn nghìn bảy trăm mười chín đồng)")
    if "+ Phạt về hành vi vi phạm thủ tục thuế: - đồng" in p.text:
        p.text = p.text.replace("+ Phạt về hành vi vi phạm thủ tục thuế: - đồng", "+ Phạt về hành vi vi phạm thủ tục thuế: 0 đồng")
    if "+ Phạt về hành vi khai sai dẫn đến thiếu số tiền thuế phải nộp (theo tỷ lệ 20% trên số thuế truy thu): - đồng" in p.text:
        p.text = p.text.replace("+ Phạt về hành vi khai sai dẫn đến thiếu số tiền thuế phải nộp (theo tỷ lệ 20% trên số thuế truy thu): - đồng", "+ Phạt về hành vi khai sai dẫn đến thiếu số tiền thuế phải nộp (theo tỷ lệ 20% trên số thuế truy thu): 2.804.719 đồng")
    if "-  Số tiền thuế truy thu: - đồng (Bằng chữ: -)" in p.text:
        p.text = p.text.replace("-  Số tiền thuế truy thu: - đồng (Bằng chữ: -)", "-  Số tiền thuế truy thu: 14.023.597 đồng (Bằng chữ: Mười bốn triệu không trăm hai mươi ba nghìn năm trăm chín mươi bảy đồng)")
    if "+ Thuế TNDN (TM 1052), số tiền: - đồng" in p.text:
        p.text = p.text.replace("+ Thuế TNDN (TM 1052), số tiền: - đồng", "+ Thuế TNDN (TM 1052), số tiền: 14.023.597 đồng")
    if "+ Thuế GTGT (TM 1701), số tiền: - đồng" in p.text:
        p.text = p.text.replace("+ Thuế GTGT (TM 1701), số tiền: - đồng", "+ Thuế GTGT (TM 1701), số tiền: 0 đồng")
    if "Giảm lỗ: - đồng" in p.text:
        p.text = p.text.replace("Giảm lỗ: - đồng", "Giảm lỗ: 0 đồng")
    if "Giảm khấu trừ: - đồng" in p.text:
        p.text = p.text.replace("Giảm khấu trừ: - đồng", "Giảm khấu trừ: 0 đồng")

    # 3. Thay thế số tiền BCKQ (P400 - P407)
    if "- Kiến nghị xử lý truy thu: -" in p.text:
        p.text = p.text.replace("- Kiến nghị xử lý truy thu: -					 đồng", "- Kiến nghị xử lý truy thu: 14.023.597 đồng")
    if "+ Thuế TNDN:  	- đồng." in p.text:
        p.text = p.text.replace("+ Thuế TNDN:  	- đồng.", "+ Thuế TNDN: 14.023.597 đồng.")
    if "+ Thuế GTGT: 	- đồng." in p.text:
        p.text = p.text.replace("+ Thuế GTGT: 	- đồng.", "+ Thuế GTGT: 0 đồng.")
    if "- Kiến nghị xử phạt: Tổng cộng số tiền phạt: - đồng" in p.text:
        p.text = p.text.replace("- Kiến nghị xử phạt: Tổng cộng số tiền phạt: - đồng", "- Kiến nghị xử phạt: Tổng cộng số tiền phạt: 2.804.719 đồng")
    if "+ Tiền chậm nộp: Xử phạt 0,03% mỗi ngày tính trên số tiền thuế chậm nộp được quy định tại Khoản 2 Điều 16 Luật Quản lý thuế 108/2025/QH15 của Quốc Hội, được tính từ ngày tiếp sau ngày cuối cùng của thời hạn nộp thuế đến ngày ký Biên bản kiểm tra, với số tiền:  - đồng." in p.text:
        p.text = p.text.replace("với số tiền:  - đồng.", "với số tiền: 1.022.320 đồng.")
    if "+ Phạt 20% đối với hành vi khai sai được quy định tại Khoản 1 Điều 16 Chương 1 Nghị Định 125/2020/NĐ-CP ngày 19/10/2020 của Chính phủ và Nghị định 310/2025/NĐ-CP ngày 02/12/2025 của Chính phủ, với số tiền: - đồng." in p.text:
        p.text = p.text.replace("với số tiền: - đồng.", "với số tiền: 2.804.719 đồng.")
    if "+ Phạt vi phạm hành chính về hành vi khai sai tại Khoản 3 Điều 12 Mục 1 Chương II Nghị định số 125/2020/NĐ-CP ngày 19/10/2020 của Chính phủ và Nghị định 310/2025/NĐ-CP ngày 02/12/2025 của Chính phủ: - đồng" in p.text:
        p.text = p.text.replace(": - đồng;", ": 0 đồng;")

    # 4. Thay thế số tiền QĐXPHC (P597 - P605)
    if "+ Thuế TNDN (TM 1052), số tiền:  	- đồng;" in p.text:
        p.text = p.text.replace("+ Thuế TNDN (TM 1052), số tiền:  	- đồng;", "+ Thuế TNDN (TM 1052), số tiền: 14.023.597 đồng;")
    if "+ Thuế GTGT (TM 1701), số tiền:  - 	đồng;" in p.text:
        p.text = p.text.replace("+ Thuế GTGT (TM 1701), số tiền:  - 	đồng;", "+ Thuế GTGT (TM 1701), số tiền: 0 đồng;")
    if "+ Tiền chậm nộp thuế TNDN (TM 4918), số tiền:  - đồng;" in p.text:
        p.text = p.text.replace("+ Tiền chậm nộp thuế TNDN (TM 4918), số tiền:  - đồng;", "+ Tiền chậm nộp thuế TNDN (TM 4918), số tiền: 1.022.320 đồng;")
    if "+ Tiền chậm nộp thuế GTGT (TM 4931), số tiền: - đồng;" in p.text:
        p.text = p.text.replace("+ Tiền chậm nộp thuế GTGT (TM 4931), số tiền: - đồng;", "+ Tiền chậm nộp thuế GTGT (TM 4931), số tiền: 0 đồng;")
    if "Tổng số tiền truy thu, phạt: - đồng (bằng chữ: -)." in p.text:
        p.text = p.text.replace("Tổng số tiền truy thu, phạt: - đồng (bằng chữ: -).", "Tổng số tiền truy thu, phạt: 17.850.636 đồng (bằng chữ: Mười bảy triệu tám trăm năm mươi nghìn sáu trăm ba mươi sáu đồng).")

    # 5. Cập nhật mốc ngày 30/11/2026 cho BBKT và tính chậm nộp
    if "(Tính đến ngày ..../..../2026)" in p.text:
        p.text = p.text.replace("(Tính đến ngày ..../..../2026)", "(Tính đến ngày 30/11/2026)")
    if "Số tiền chậm nộp tiền thuế nêu trên được tính đến hết ngày ..../..../2026" in p.text:
        p.text = p.text.replace("Số tiền chậm nộp tiền thuế nêu trên được tính đến hết ngày ..../..../2026", "Số tiền chậm nộp tiền thuế nêu trên được tính đến hết ngày 30/11/2026")
    if "kể từ sau ngày ..../..../2026" in p.text:
        p.text = p.text.replace("kể từ sau ngày ..../..../2026", "kể từ sau ngày 30/11/2026")

# Cập nhật trong các table nếu có chuỗi "Chi phí quản lý doanh nghiệp"
for t in doc.tables:
    for row in t.rows:
        for cell in row.cells:
            if "Chi phí quản lý doanh nghiệp" in cell.text:
                cell.text = cell.text.replace("Chi phí quản lý doanh nghiệp", "Chi phí quản lý kinh doanh")

doc.save(doc_path)
print("Updated TONG HOP.docx with numbers and Chi phi quan ly kinh doanh successfully!")
