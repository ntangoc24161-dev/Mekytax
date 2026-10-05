import openpyxl
import json
from pathlib import Path

clean_mapping = {
    '1143/TCS26-QLDN': {'ten': 'CÔNG TY TNHH SUNGKWANG SG VINA', 'mst': '3703088765', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '1423/TCS3-QLDN': {'ten': 'CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ BOM BAY', 'mst': '5801512834', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '18301, 18319, 18305, 18297, 18278/TCS28-QLDN': {'ten': 'CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ LA ĐÙ VINA', 'mst': '3703248095', 'nd': 'Không hoạt động tại địa chỉ đăng ký; Chênh lệch số liệu tờ khai và HĐĐT'},
    '2220/TCS5-QLDN2': {'ten': 'CÔNG TY TNHH TMDV NI PHẠM', 'mst': '0315951217', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '2626/2545/TCS3-QLDN1': {'ten': 'CÔNG TY TNHH TMDV HỢP TÁC THỊNH VƯỢNG', 'mst': '0319579381', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '27719/TCS2-QLDN2': {'ten': 'CÔNG TY TNHH LỘC KHA TRIP', 'mst': '4201996571', 'nd': 'Không hoạt động tại địa chỉ đăng ký; Chuyển tin báo CQĐT'},
    '2801/TCS18-QLDN1': {'ten': 'CÔNG TY TNHH ĐẦU TƯ SẢN XUẤT GỖ BẠCH DƯƠNG', 'mst': '0318331490', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '2874/TCS7-QLDN1': {'ten': 'CÔNG TY TNHH TMDV MỘC LAM', 'mst': '0319201800', 'nd': 'Không hoạt động tại địa chỉ đăng ký; Kèm bảng kê bán ra'},
    '2977/TCS13-QLDN1': {'ten': 'CÔNG TY TNHH NHÀ HÀNG KHÁCH SẠN TRƯỜNG GIANG HOTEL', 'mst': '0318735415', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '3445/TCS12-QLDN3': {'ten': 'CÔNG TY TNHH PACITRANS', 'mst': '0314380264', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '3448/TCS12-QLDN3': {'ten': 'CÔNG TY TNHH CƠ KHÍ THỦY LỰC DŨNG SĨ', 'mst': '0315972217', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '3448/TCS15-QLDN3': {'ten': 'CÔNG TY TNHH CƠ KHÍ THỦY LỰC DŨNG SĨ', 'mst': '0315972217', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '4153/TCS18-QLDN1': {'ten': 'CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ BAO BÌ QUANG KHẢI', 'mst': '0316938850', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '523/TCS4-QLDN': {'ten': 'CÔNG TY TNHH SX TM QUỐC BÌNH L.A', 'mst': '1102041180', 'nd': 'Chênh lệch số liệu tờ khai và HĐĐT'},
    '528/TCS4-QLDN': {'ten': 'CÔNG TY TNHH TMDV KINH DOANH GỖ HOÀNG GIA NAM', 'mst': '1102154610', 'nd': 'Không hoạt động tại địa chỉ đăng ký'},
    '536/QLDN-TCS4': {'ten': 'CÔNG TY TNHH TM XD NGỌC ANH', 'mst': '1102092403', 'nd': 'Chênh lệch số liệu tờ khai và HĐĐT'},
}

wb_path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\00_SO THEO DOI & BAO CAO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
if not Path(wb_path).is_file():
    wb_path = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\2- HOA DON\1. XU LY RUI RO\Theo doi CVCBRR\SoTheoDoi_CV_CanhBaoRuiRo_Thue.xlsx"
wb = openpyxl.load_workbook(wb_path)
ws = wb["THEO_DOI"]

for r in range(28154, ws.max_row + 1):
    scv = str(ws.cell(r, 2).value or "").strip()
    if scv in clean_mapping:
        info = clean_mapping[scv]
        ws.cell(r, 5, info["ten"])
        ws.cell(r, 6, info["mst"])
        if info["nd"]:
            ws.cell(r, 7, info["nd"])
        print(f"Updated row {r}: {scv} -> {info['ten']} ({info['mst']})")

wb.save(wb_path)
print("Successfully saved cleaned company names to Excel!")

# Update records_2026_09_scan.json as well
json_path = Path(r"D:\SECOND BRAIN\javis-os\brains\Cập nhật công văn cảnh báo\skills\theo-doi-cv-canh-bao-rui-ro\scripts\records_2026_09_scan.json")
if json_path.exists():
    with open(json_path, "r", encoding="utf-8") as f:
        recs = json.load(f)
    for rec in recs:
        scv = rec.get("so_cv")
        if scv in clean_mapping:
            info = clean_mapping[scv]
            rec["ten_ban"] = info["ten"]
            rec["mst_ban"] = info["mst"]
            rec["noi_dung"] = info["nd"]
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(recs, f, ensure_ascii=False, indent=2)
    print("Successfully updated records_2026_09_scan.json!")
