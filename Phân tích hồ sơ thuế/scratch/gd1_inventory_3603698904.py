"""Giai đoạn 1 - kiểm kê manifest cho MST 3603698904 (PHUC NGUYEN TIEN).
Quét toàn bộ thư mục hồ sơ, băm SHA-256, đếm dòng (xlsx/xls), xếp IN00-IN14.
Chỉ đọc; không sửa file nguồn. Kết quả: JSON trong scratch.
"""
import hashlib, json, os, re, sys
from pathlib import Path

CASE_REAL = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\2026\Dang kiem tra\KE HOACH\PHUC_NGUYEN_TIEN_3603698904-chua so"
CASE = Path(r"\\?\\" + CASE_REAL)  # tiền tố đường dẫn dài (>260 ký tự)
OUT = Path(r"D:\SECOND BRAIN\javis-os\brains\Phân tích hồ sơ thuế\scratch\gd1_inventory_3603698904.json")


def sha256(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def count_rows(p: Path):
    ext = p.suffix.lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            res = {ws.title: (ws.max_row or 0) for ws in wb.worksheets}
            wb.close()
            return res
        if ext == ".xls":
            import xlrd
            wb = xlrd.open_workbook(p, on_demand=True)
            res = {s: wb.sheet_by_name(s).nrows for s in wb.sheet_names()}
            wb.release_resources()
            return res
    except Exception as e:  # ghi nhận lỗi đọc, không bỏ qua
        return {"_LOI_DOC": str(e)[:200]}
    return None


def classify(rel: str):
    n = rel.lower()
    if n.startswith("1. du lieu") or n.startswith("1. du lieu tho"):
        return "CQT"
    if n.startswith("2. ho so") or "cty gui" in n:
        return "DN"
    if n.startswith("3. du lieu"):
        return "CQT/Tra cuu"
    if n.startswith("4. ket qua") or n.startswith("ket-qua") or n.startswith("7. sao luu"):
        return "Ket qua (ban phan tich truoc)"
    return "Khac"


def in_code(rel: str) -> str:
    n = rel.lower()
    base = os.path.basename(n)
    if re.search(r"xml|tờ khai|to khai|gtgt|tndn|tncn|bctc|qtt|tb0|bc26|ttS|tts|lpmb|mbai|tb01", base) and "1.1" in n:
        return "IN03/IN04/IN06 (tờ khai XML)"
    if "hddt" in n or "hoa don" in n or "hoá đơn" in n or "hd mv" in n or "hd ban" in n:
        return "IN01_HDDT"
    if "bang ke" in n or "bảng kê" in n or "bang_ke" in n:
        return "IN02_BangKe"
    if "1.3" in n or "thong bao" in n or "thông báo" in n:
        return "IN14_TMS_Nghia_vu"
    if "1.4" in n or "tncn" in n and "nld" in n or "danh sach nld" in n or "nld" in n:
        return "IN05_NLD_TNCN"
    if "3.2" in n or "danh ba" in n or "danh sach doi tac" in n or "danh sách ncc" in n or "ncc" in n:
        return "IN07_NCC_CanhBao"
    if "3.1" in n or "cvcbrr" in n or "cong an" in n:
        return "IN07_NCC_CanhBao"
    if "ket-qua" in n or "ket qua" in n or "4. ket qua" in n or "7. sao luu" in n or "sao-luu" in n:
        return "Ket qua (khong thuoc IN)"
    if "511" in n or "632" in n or re.search(r"tk 5(11|12|18)|tk 911", base):
        return "IN13_So_511_632"
    if "2.2" in n or re.search(r"tk 15[256]|tk 154|kho|xnt|giá thành|gia thanh", base):
        return "IN09_Kho_GiaThanh"
    if "2.3" in n or re.search(r"tien.gui|tiền gửi|tk 1121|sao ke|sao kê|1121", base):
        return "IN11_Sao_ke_NH"
    if re.search(r"tk 1111|1111", base):
        return "IN10_Chi_tien_mat"
    if "2.4" in n or re.search(r"l.ng|khau hao|khấu hao|s09|tk 242|tk 334", base):
        return "IN05/IN08 (luong, TSCD, CP tra truoc)"
    if "2.1" in n or re.search(r"nhat ky|nhật ký|so cai|sổ cái|tk \d+|cong no|công nợ", base):
        return "IN08_So_KeToan"
    return "Chua xep"


def main():
    rows = []
    for root, dirs, files in os.walk(CASE):
        for fn in files:
            p = Path(root) / fn
            rel = os.path.relpath(p, CASE)
            top = rel.split(os.sep)[0]
            rec = {
                "rel": rel,
                "top": top,
                "src_group": classify(rel),
                "in_code": in_code(rel),
                "ext": p.suffix.lower(),
            }
            try:
                rec["size"] = p.stat().st_size
                rec["sha256"] = sha256(p)
                rec["sheets"] = count_rows(p)
            except Exception as e:
                rec["error"] = f"{type(e).__name__}: {str(e)[:160]}"
            rows.append(rec)
    OUT.write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    print("files:", len(rows), "->", OUT)


if __name__ == "__main__":
    main()
