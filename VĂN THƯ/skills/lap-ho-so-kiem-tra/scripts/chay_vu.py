"""Gọi app hồ sơ kiểm tra qua HTTP. Chỉ dùng thư viện Python có sẵn."""

from __future__ import annotations

import argparse
import os
import re
import sys
import urllib.error
import urllib.request
from http.cookiejar import CookieJar
from pathlib import Path
from urllib.parse import quote, unquote, urlencode

APP = Path(r"D:\01. Cong tac kiem tra\HO SO KIEM TRA\app")
if not APP.exists():
    APP = Path(r"D:\SET UP BY AI\HO SO KIEM TRA\app")
BASE = "http://127.0.0.1:8000"
USER = os.environ.get("HOSO_USER", "canbo")
PASSWORD = os.environ.get("HOSO_PASS", "canbo123")
OUTPUT = APP / "output"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def opener():
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(CookieJar()))


def fetch(op, url: str, data: bytes | None = None, headers: dict | None = None, timeout: int = 60):
    req = urllib.request.Request(url, data=data, headers=headers or {})
    try:
        with op.open(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8", "replace")
            return resp.status, resp.geturl(), body
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", "replace")
        return exc.code, url, body


def strip_tags(html: str) -> str:
    text = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", text).strip()


def alerts(html: str, kind: str) -> list[str]:
    found = re.findall(rf'<div class="alert {kind}">(.*?)</div>', html, re.S)
    return [strip_tags(x) for x in found if strip_tags(x)]


def login(op) -> None:
    status, final, html = fetch(
        op,
        BASE + "/login",
        data=urlencode({"username": USER, "password": PASSWORD}).encode(),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    if status >= 400 or "Sai tên đăng nhập" in html or final.rstrip("/").endswith("/login"):
        raise SystemExit("Đăng nhập thất bại. Kiểm tra HOSO_USER và HOSO_PASS.")


def cmd_status(_args) -> int:
    try:
        req = urllib.request.Request(BASE + "/login")
        with urllib.request.urlopen(req, timeout=5) as resp:
            html = resp.read().decode("utf-8", "replace")
            code = resp.status
    except (urllib.error.URLError, TimeoutError, OSError):
        print("chua chay")
        print(f"Mo: {APP / 'chay.bat'}")
        return 2
    if code == 200 and 'name="username"' in html:
        print("dang chay")
        print(BASE)
        return 0
    print("chua chay")
    print(f"Cong 8000 tra ma {code}, khong phai trang dang nhap.")
    return 2


def encode_xlsx(path: Path) -> tuple[bytes, str]:
    boundary = "----HoSoKiemTra7f3a"
    raw = path.read_bytes()
    head = (
        f"--{boundary}\r\n"
        'Content-Disposition: form-data; name="file"; filename="import.xlsx"\r\n'
        "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
    ).encode()
    tail = f"\r\n--{boundary}--\r\n".encode()
    return head + raw + tail, f"multipart/form-data; boundary={boundary}"


def cmd_import(args) -> int:
    path = Path(args.file)
    if not path.is_file():
        print(f"Khong thay file: {path}")
        return 2
    op = opener()
    login(op)
    body, ctype = encode_xlsx(path)
    status, _url, html = fetch(
        op,
        BASE + "/excel/import",
        data=body,
        headers={"Content-Type": ctype},
        timeout=180,
    )
    errs = alerts(html, "err")
    warns = alerts(html, "warn")
    ok = alerts(html, "ok")
    for line in errs:
        print("LOI:", line)
    for line in warns:
        print("CANH BAO:", line)
    created = updated = None
    if ok:
        m = re.search(r"Đã tạo (\d+) vụ, cập nhật (\d+) vụ", ok[0])
        if m:
            created, updated = m.group(1), m.group(2)
    print(f"tao={created if created is not None else '?'}")
    print(f"cap_nhat={updated if updated is not None else '?'}")
    if errs or status >= 400 or created is None:
        return 4
    print("OK")
    return 0


def list_vus(op, mst: str) -> list[tuple[int, str]]:
    status, _url, html = fetch(op, BASE + "/?q=" + quote(mst))
    if status >= 400:
        return []
    rows = []
    for m in re.finditer(r'<a class="vu-item[^"]*" href="/vu/(\d+)[^"]*">(.*?)</a>', html, re.S):
        rows.append((int(m.group(1)), strip_tags(m.group(2))))
    return rows


def pick_vu(rows: list[tuple[int, str]], mst: str, qd: str) -> list[int]:
    want_mst = re.sub(r"\s+", "", mst)
    want_qd = re.sub(r"\s+", "", qd)
    hits = []
    for vid, text in rows:
        mm = re.search(r"MST\s+(\S+)", text)
        qm = re.search(r"QĐ\s+([^·]+)", text)
        got_mst = re.sub(r"\s+", "", mm.group(1)) if mm else ""
        got_qd = re.sub(r"\s+", "", qm.group(1)) if qm else ""
        if got_mst == want_mst and got_qd == want_qd:
            hits.append(vid)
    return hits


def checked_docs(html: str) -> list[str]:
    docs = []
    for m in re.finditer(r"<input\s+type=\"checkbox\"\s+name=\"(doc_[^\"]+)\"([^>]*)>", html):
        attrs = m.group(2)
        if "checked" in attrs and "disabled" not in attrs:
            docs.append(m.group(1)[4:])
    return docs


def cmd_xuat(args) -> int:
    op = opener()
    login(op)
    rows = list_vus(op, args.mst)
    hits = pick_vu(rows, args.mst, args.qd)
    if not hits:
        print("Khong thay vu khop MST va so quyet dinh.")
        for vid, text in rows[:8]:
            print(f"  id={vid} {text}")
        return 3
    if len(hits) > 1:
        print("Nhieu vu cung MST va so quyet dinh:", ", ".join(str(x) for x in hits))
        return 3
    vu_id = hits[0]
    status, _url, page = fetch(op, f"{BASE}/vu/{vu_id}/xuat")
    if status >= 400:
        print("Khong mo duoc trang xuat.")
        return 4
    if args.docs:
        docs = [x.strip() for x in args.docs.split(",") if x.strip()]
    else:
        docs = checked_docs(page)
    if not docs:
        print("Khong co mau nao de xuat.")
        return 4
    print("mau=" + ",".join(docs))
    form = urlencode([(f"doc_{doc}", "on") for doc in docs]).encode()
    status, _url, html = fetch(
        op,
        f"{BASE}/vu/{vu_id}/xuat",
        data=form,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=180,
    )
    for line in alerts(html, "err"):
        print("LOI:", line)
    for line in alerts(html, "warn"):
        print("CANH BAO:", line)
    if "Xác nhận và xuất" in html or status >= 400 or "Đã xuất hồ sơ" not in html:
        print("CHUA XUAT. Can user xem loi hoac canh bao. Khong tu xac nhan.")
        return 5
    print("OK")
    print(f"vu_id={vu_id}")
    for href in re.findall(r'href="(/file/[^"]+)"', html):
        rel = unquote(href[len("/file/") :])
        print(OUTPUT / Path(rel))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Van hanh app ho so kiem tra.")
    sub = parser.add_subparsers(dest="lenh", required=True)
    sub.add_parser("status")
    p_imp = sub.add_parser("import")
    p_imp.add_argument("file")
    p_xuat = sub.add_parser("xuat")
    p_xuat.add_argument("--mst", required=True)
    p_xuat.add_argument("--qd", required=True)
    p_xuat.add_argument("--docs", default="")
    args = parser.parse_args()
    if args.lenh == "status":
        return cmd_status(args)
    if args.lenh == "import":
        return cmd_import(args)
    return cmd_xuat(args)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except urllib.error.URLError:
        print("chua chay")
        print(f"Mo: {APP / 'chay.bat'}")
        raise SystemExit(2)
