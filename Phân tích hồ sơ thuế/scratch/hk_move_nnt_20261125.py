# -*- coding: utf-8 -*-
"""Di chuyển dữ liệu NNT Huy Khang vào cây 4 bước. Không xóa nguồn trước khi đối soát."""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(
    r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT"
    r"\2026\Ke hoach nam\Bo sung\HUY_KHANG_3603755214"
)
SRC = ROOT / "DU LIEU HUY KHANG GOI THUE"
B2 = ROOT / "2. HO SO NNT CUNG CAP"
DEST = {
    "2.1": B2 / "2.1. So ke toan tong hop va chi tiet",
    "2.2": B2 / "2.2. Kho - Chi phi - Gia thanh",
    "2.3": B2 / "2.3. Ngan hang & Dong tien",
    "2.4": B2 / "2.4. Lao dong - Tien luong - TSCD",
    "2.5": B2 / "2.5. Bang ke mua vao ban ra",
}
LOG = Path(r"C:\Users\NGOC\AppData\Local\Temp\huykhang_move_log.json")

TK_KHO = {"152", "153", "154", "155", "156"}
TK_TSCD = {"211", "214", "242", "244"}


def year_of(rel: str) -> str:
    m = re.search(r"NAM\s+(20\d{2})", rel, re.I)
    return m.group(1) if m else "khac"


def tk_from_name(name: str) -> str | None:
    m = re.search(r"(?:tk|tai khoan|tài khoản)\s*(\d{3})", name, re.I)
    if m:
        return m.group(1)
    m = re.search(r"(\d{3})\.xlsx?$", name, re.I)
    if m:
        return m.group(1)
    return None


def classify(rel: str, name: str) -> str:
    n = name.casefold()
    d = rel.replace("/", "\\").casefold()
    parent = str(Path(rel).parent).casefold()

    if any(k in parent for k in ["bang ke", "bangke", "dau ra dau vao"]):
        return "2.5"
    if any(k in n for k in ["bang ke", "bangke", "bảng kê"]):
        return "2.5"

    if any(k in parent for k in ["ton kho", "nhap xuat ton", "\\kho\\", "\\kho "]):
        return "2.2"
    if any(k in n for k in ["ton_kho", "ton kho", "nhap xuat ton", "nhập xuất tồn"]):
        return "2.2"

    if any(k in parent for k in ["bang luong", "luong", "ccdc", "cptt", "tscd"]):
        return "2.4"
    if any(
        k in n
        for k in [
            "bang luong",
            "bảng lương",
            "luong co",
            "lương",
            "khau hao",
            "khấu hao",
            "tai san co dinh",
            "tài sản cố định",
            "cong cu dung cu",
            "công cụ dụng cụ",
            "chi phi tra truoc",
            "chi phí trả trước",
            "phan bo chi phi",
            "phân bổ",
            "so_congcu",
            "s21 dn",
        ]
    ):
        return "2.4"

    tk = tk_from_name(name)
    if tk in TK_KHO:
        return "2.2"
    if tk in TK_TSCD:
        return "2.4"
    return "2.1"


def sha256(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def unique_dest(folder: Path, name: str) -> Path:
    dest = folder / name
    if not dest.exists():
        return dest
    stem, suf = Path(name).stem, Path(name).suffix
    i = 2
    while True:
        cand = folder / f"{stem}__{i}{suf}"
        if not cand.exists():
            return cand
        i += 1


def main():
    if not SRC.exists():
        print("STOP: source missing", SRC)
        sys.exit(1)
    files = [p for p in SRC.rglob("*") if p.is_file()]
    print("SOURCE FILES", len(files), "BYTES", sum(p.stat().st_size for p in files))
    plan = []
    counts = {"2.1": 0, "2.2": 0, "2.3": 0, "2.4": 0, "2.5": 0}
    for p in files:
        rel = str(p.relative_to(SRC))
        code = classify(rel, p.name)
        year = year_of(rel)
        dest_dir = DEST[code] / year
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = unique_dest(dest_dir, p.name)
        plan.append(
            {
                "src": str(p),
                "rel": rel,
                "dest": str(dest),
                "code": code,
                "year": year,
                "size": p.stat().st_size,
            }
        )
        counts[code] += 1
    print("PLAN", counts)

    moved = []
    errors = []
    for item in plan:
        src = Path(item["src"])
        dest = Path(item["dest"])
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dest))
            if not dest.exists():
                raise FileNotFoundError(dest)
            if dest.stat().st_size != item["size"]:
                raise ValueError(f"size mismatch {dest.stat().st_size} != {item['size']}")
            item["ok"] = True
            moved.append(item)
        except Exception as e:
            item["ok"] = False
            item["err"] = str(e)
            errors.append(item)
            print("ERR", item["rel"], e)

    print("MOVED", len(moved), "ERR", len(errors))
    remain = [p for p in SRC.rglob("*") if p.is_file()]
    print("REMAIN IN SOURCE", len(remain))
    if remain:
        for p in remain:
            print("  leftover", p.relative_to(SRC))

    if not remain and not errors:
        # remove empty source tree
        shutil.rmtree(SRC, ignore_errors=False)
        print("REMOVED SOURCE FOLDER", SRC.exists())
    else:
        print("KEEP SOURCE because leftover or error")

    LOG.write_text(json.dumps({"counts": counts, "moved": len(moved), "errors": errors, "plan": plan}, ensure_ascii=False, indent=2), encoding="utf-8")
    print("LOG", LOG)

    # dest counts
    for code, folder in DEST.items():
        n = len([p for p in folder.rglob("*") if p.is_file()])
        print("DEST", code, n)


if __name__ == "__main__":
    main()
