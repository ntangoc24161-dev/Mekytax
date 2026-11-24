"""CLI: đọc thư mục hồ sơ XML/Excel, xuất findings.json + báo cáo .md/.xlsx."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


ROOT = _app_dir()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
# PyInstaller unpacks to _MEIPASS
_meipass = getattr(sys, "_MEIPASS", None)
if _meipass and _meipass not in sys.path:
    sys.path.insert(0, _meipass)

from analyze.m1_hddt import run_m1
from analyze.m2_bctc import run_m2
from analyze.m3_score import run_m3
from ingest.hoso import scan_dir
from report.markdown import to_markdown, write_outputs


def _pick_folders() -> tuple[Path, Path] | None:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception:
        print("ERROR: thiếu --input-dir. Kéo thả thư mục hồ sơ vào Chay.bat", file=sys.stderr)
        return None
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    inp = filedialog.askdirectory(title="Chọn thư mục hồ sơ (XML tờ khai/BCTC + Excel hóa đơn)")
    if not inp:
        return None
    src = Path(inp)
    out = filedialog.askdirectory(title="Chọn thư mục ghi kết quả (Hủy = tạo ket-qua trong hồ sơ)")
    dest = Path(out) if out else src / "ket-qua"
    messagebox.showinfo("Phân tích hồ sơ thuế", f"Đọc:\n{src}\n\nGhi:\n{dest}")
    return src, dest


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    p = argparse.ArgumentParser(prog="phan-tich-ho-so-thue")
    p.add_argument("input_pos", nargs="?", help="Thư mục hồ sơ (kéo thả vào Chay.bat).")
    p.add_argument("--input-dir", default=None, help="Thư mục XML tờ khai/BCTC + Excel bảng kê HĐ")
    p.add_argument("--out", default=None, help="Thư mục ghi .xlsx .md .json (mặc định: <hồ sơ>/ket-qua)")
    p.add_argument("--json-only", action="store_true")
    args = p.parse_args(argv)
    src_s = args.input_dir or args.input_pos
    dest_s = args.out
    if not src_s:
        picked = _pick_folders()
        if not picked:
            return 2
        src, dest = picked
    else:
        src = Path(src_s)
        dest = Path(dest_s) if dest_s else src / "ket-qua"
    if not src.is_dir():
        print(f"ERROR: không thấy thư mục {src}", file=sys.stderr)
        return 2
    hs = scan_dir(src)
    if not hs.to_khai and not hs.hoa_don:
        print("ERROR: không đọc được XML tờ khai hay bảng kê HĐ.", file=sys.stderr)
        for s in hs.skipped[:10]:
            print(" ", s, file=sys.stderr)
        return 1
    f1 = run_m1(hs)
    f2 = run_m2(hs)
    f3 = run_m3(hs, f1 + f2)
    findings = f1 + f2 + f3
    md = to_markdown(hs, findings)
    if args.json_only:
        print(json.dumps([x.to_dict() for x in findings], ensure_ascii=False, indent=2))
        return 0
    md_path, js_path, xlsx_path = write_outputs(dest, hs, findings, md)
    print(f"OK  MST {hs.mst}  {len(hs.to_khai)} XML  {len(hs.hoa_don)} HD  {len(findings)} findings")
    print(f"    {xlsx_path}")
    print(f"    {md_path}")
    print(f"    {js_path}")
    if not args.input_dir and not args.input_pos:
        try:
            input("Nhan Enter de dong...")
        except EOFError:
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
