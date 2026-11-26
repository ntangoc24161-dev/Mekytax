"""Điền KQ1 (Giai đoạn 1 - kiểm kê & đánh giá) cho MST 3603698904 từ kết quả kiểm kê.
Chỉ đọc file nguồn/mẫu; ghi file KQ1 ra scratch rồi sao sang thư mục '4. KET QUA PHAN TICH'.
"""
import json
import re
import shutil
from pathlib import Path

import openpyxl

MST = "3603698904"
NGAY = "2026-11-25"
VAULT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế")
SCR = VAULT / "scratch"
MAU = VAULT / "thu-muc-mau-ho-so-doanh-nghiep"
TPL = MAU / "MAU_FILE_TONG_HOP" / "cu-v1.2" / "KQ1_MAU_Kiem_ke_danh_gia.xlsx"
STD = MAU / "Bo_bieu_mau_chuan_phan_tich_ho_so_thue_v2.xlsx"
INV = SCR / "gd1_inventory_3603698904.json"
OUT_DIR = SCR / "out_3603698904"
OUT = OUT_DIR / "KQ1_3603698904.xlsx"
CASE_OUT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Dang kiem tra\KE HOACH\PHUC_NGUYEN_TIEN_3603698904-chua so\4. KET QUA PHAN TICH"
)

NAMES = {
    "IN00": "Manifest nguồn (kiểm kê)",
    "IN01": "Hóa đơn điện tử mua vào/bán ra (HĐĐT)",
    "IN02": "Bảng kê mua vào/bán ra của DN",
    "IN03": "Tờ khai GTGT",
    "IN04": "BCTC / Quyết toán TNDN",
    "IN05": "Danh sách NLĐ / TNCN",
    "IN06": "Quyết toán TNCN",
    "IN07": "Danh sách NCC / CVCBRR / danh bạ",
    "IN08": "Sổ kế toán (nhật ký chung, sổ cái)",
    "IN09": "Kho và giá thành",
    "IN10": "Chứng từ tiền mặt TK 1111",
    "IN11": "Sao kê ngân hàng TK 1121",
    "IN12": "Giao dịch liên kết (GDLK)",
    "IN13": "Sổ chi tiết TK 511/632",
    "IN14": "Thông báo / lịch sử nghĩa vụ thuế (TMS)",
}
ALL_IN = [f"IN{i:02d}" for i in range(15)]


def codes_in(text):
    return set(re.findall(r"IN\d{2}", text or ""))


def req_codes(text):
    """Chuyển 'IN01–IN08' thành tập IN01..IN08; 'IN02, IN07' thành {IN02, IN07}."""
    out = set()
    for a, b in re.findall(r"IN(\d{2})(?:\s*[–-]\s*IN(\d{2}))?", text or ""):
        lo = int(a)
        hi = int(b) if b else lo
        out |= {f"IN{i:02d}" for i in range(lo, hi + 1)}
    return out


def data_rows(sheets):
    if not isinstance(sheets, dict):
        return 0
    return sum(v for k, v in sheets.items() if isinstance(v, int) and not k.startswith("_"))


def nguon_goc(tops):
    m = {
        "1. DU LIEU THO NGANH THUE": "CQT",
        "2. HO SO NNT CUNG CAP": "DN",
        "3. DU LIEU BO SUNG & TRA CUU QLT": "CQT/Tra cứu",
        "4. KET QUA PHAN TICH": "Kết quả cũ (không thuộc IN)",
        "ket-qua": "Kết quả cũ (không thuộc IN)",
        "7. Sao luu tong hop": "Sao lưu (không thuộc IN)",
    }
    return ", ".join(sorted({m.get(t, t) for t in tops}))


def main():
    inv = json.loads(INV.read_text(encoding="utf-8"))

    groups = {}
    for r in inv:
        g = groups.setdefault(
            r["in_code"],
            {"files": 0, "rows": 0, "sha": [], "tops": set(), "ext": set(), "err": 0},
        )
        g["files"] += 1
        g["rows"] += data_rows(r.get("sheets"))
        g["sha"].append(r.get("sha256", ""))
        g["tops"].add(r["top"])
        g["ext"].add(r["ext"])
        if "error" in r:
            g["err"] += 1

    def has_content(g):
        return g["files"] > 0 and (g["rows"] > 0 or g["ext"] & {".xml", ".pdf", ".docx"})

    present = {"IN00"}  # manifest chính là kết quả kiểm kê này
    for label, g in groups.items():
        if has_content(g):
            present |= codes_in(label)

    # ---- KQ1_Kiem_ke ----
    wb = openpyxl.load_workbook(TPL)
    ws = wb["KQ1_Kiem_ke"]
    ws["A2"] = f"MST: {MST} | Kỳ: theo từng file nguồn | Giai đoạn: 1 | Phiên bản biểu mẫu: 1.2 | Ngày chạy: {NGAY} | Phiên bản KQ: v1"
    ws.delete_rows(5, ws.max_row)

    def order_key(label):
        c = sorted(codes_in(label))
        return (c[0] if c else "ZZ", label)

    row = 5
    stt = 1
    for label in sorted(groups, key=order_key):
        g = groups[label]
        c = sorted(codes_in(label))
        loai = "; ".join(NAMES.get(x, x) for x in c) if c else "Không xếp được IN (rà thủ công)"
        if not c:
            ket = "Chưa xếp" if label == "Chua xep" else "Không thuộc IN"
        elif g["files"] == 0:
            ket = "Chưa nhận"
        elif g["rows"] == 0 and ".pdf" in g["ext"]:
            ket = "PDF chưa trích"
        elif g["rows"] == 0 and ".xml" in g["ext"]:
            ket = "Đủ"
        elif g["rows"] == 0:
            ket = "Thiếu"
        elif g["rows"] < 10:
            ket = "Mỏng"
        else:
            ket = "Đủ"
        sha = g["sha"][0] if g["files"] == 1 else f"{g['files']} file (xem gd1_inventory JSON)"
        ghi = f"{g['files']} file; lỗi đọc {g['err']}; phân loại tự động theo tên file, cần rà soát"
        vals = [
            stt, label, loai, nguon_goc(g["tops"]), ", ".join(sorted(g["ext"])) or "—",
            "Theo file nguồn", g["rows"], "Chưa kiểm", "Chưa xác định", sha, ket, ghi,
        ]
        for col, v in enumerate(vals, start=1):
            ws.cell(row=row, column=col, value=v)
        row += 1
        stt += 1

    # IN chưa có nguồn nào: ghi là 'Chưa nhận'
    labels_codes = set()
    for label in groups:
        labels_codes |= codes_in(label)
    for code in ALL_IN:
        if code in labels_codes or code == "IN00":
            continue
        vals = [stt, code, NAMES[code], "Chưa xác định", "—", "Theo file nguồn", 0, "—", "—", "—",
                "Chưa nhận", "Không có file trong thư mục hồ sơ"]
        for col, v in enumerate(vals, start=1):
            ws.cell(row=row, column=col, value=v)
        row += 1
        stt += 1

    # ---- Danh mục quy tắc (lấy từ bộ chuẩn v2) ----
    std = openpyxl.load_workbook(STD, data_only=True)["04_Quy_tac_rui_ro"]
    rules = []
    for r in std.iter_rows(values_only=True):
        rid = str(r[0] or "").strip()
        if re.fullmatch(r"[A-Z]{1,3}\d{1,2}", rid):
            rules.append({
                "id": rid, "nhom": r[1] or "", "ten": r[2] or "", "dulieu": r[6] or "",
                "goi_y": r[8] or "", "skill": r[11] or "",
            })

    ws2 = wb["KQ1_Kha_nang_phan_tich"]
    ws2["A2"] = f"MST: {MST} | Kỳ: theo từng file nguồn | Giai đoạn: 1 | Phiên bản biểu mẫu: 1.1 | Ngày chạy: {NGAY} | Phiên bản KQ: v1"
    ws2.delete_rows(5, ws2.max_row)

    missing_count = {}
    summary = {"Chạy ngay": 0, "Chạy một phần": 0, "Bị chặn": 0, "Không xác định": 0}
    row = 5
    for rule in rules:
        req = req_codes(rule["dulieu"]) or codes_in(rule["dulieu"])
        if not req:
            status = "Không xác định"
            missing = set()
        else:
            missing = req - present
            if not missing:
                status = "Chạy ngay"
            elif missing == req:
                status = "Bị chặn"
            else:
                status = "Chạy một phần"
        for m in missing:
            missing_count.setdefault(m, []).append(rule["id"])
        summary[status] += 1
        stage_text = rule["skill"] + " " + rule["goi_y"]
        stage = "3" if re.search(r"giai-doan-3|GĐ3|GD3|giai đoạn 3", stage_text, re.I) else "2"
        dl = ", ".join(sorted(req)) if req else rule["dulieu"]
        thieu = "; ".join(f"{m} ({NAMES.get(m, m)})" for m in sorted(missing)) or "—"
        vals = [rule["id"], rule["nhom"], rule["ten"], stage, dl, status, thieu, rule["goi_y"]]
        for col, v in enumerate(vals, start=1):
            ws2.cell(row=row, column=col, value=v)
        row += 1

    # ---- KQ1_De_nghi_bo_sung ----
    ws3 = wb["KQ1_De_nghi_bo_sung"]
    ws3["A2"] = f"MST: {MST} | Kỳ: theo từng file nguồn | Giai đoạn: 1 | Phiên bản biểu mẫu: 1.1 | Ngày chạy: {NGAY} | Phiên bản KQ: v1"
    ws3.delete_rows(5, ws3.max_row)
    ranked = sorted(missing_count.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    row = 5
    for idx, (m, ids) in enumerate(ranked, start=1):
        vals = [idx, NAMES.get(m, m), m, len(ids), ", ".join(ids),
                f"Chưa có file {m} trong thư mục hồ sơ (kiểm kê {NGAY})"]
        for col, v in enumerate(vals, start=1):
            ws3.cell(row=row, column=col, value=v)
        row += 1

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print("saved", OUT)
    print("rules:", len(rules), summary)
    print("present:", sorted(present))
    print("missing:", [(m, len(ids)) for m, ids in ranked])

    try:
        CASE_OUT.mkdir(parents=True, exist_ok=True)
        shutil.copy2(OUT, CASE_OUT / OUT.name)
        print("copied to", CASE_OUT / OUT.name)
    except OSError as e:
        print("COPY FAILED:", e)


if __name__ == "__main__":
    main()
