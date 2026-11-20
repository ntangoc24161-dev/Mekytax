#!/usr/bin/env python3
"""Cửa kiểm sổ Bao_cao trước bóc tách/ghi (chốt 20/11/2026).

Luồng VPS (Javis Docker):
1) Nhắc chủ qua bot CA @Kiem_soat_dieu_phoi_bot
2) Best-effort gọi đóng Excel trên MÁY ĐANG CHẠY Javis (javis_app_close) - trên VPS Docker
   tool này thường báo không điều khiển được máy Windows của chủ
3) KHÔNG chờ chủ trả lời chat; chờ ngắn để Drive sync (mặc định 60s)
4) Exit 0 để quy trình bóc tách/ghi sổ chạy tiếp (ghi đè đúng 1 file sống Drive)

Máy Windows chạy Javis native: bước 2 đóng Excel thật được thì tốt hơn.
"""
from __future__ import annotations

import argparse
import asyncio
import sys
import time
from pathlib import Path

DEFAULT_WAIT = 60
CHAT = "8799384849"
ACC = "acc_656f8aaa4e"


def _send_telegram(text: str) -> bool:
    sys.path.insert(0, "/app/server")
    try:
        import channel_accounts
        from telegram_bot import TelegramBot
    except Exception as e:
        print(f"telegram import fail: {e}")
        return False
    tok = channel_accounts.get_token(ACC) or ""
    if not tok or str(tok).startswith("enc:"):
        print("telegram token missing for bot CA")
        return False

    async def _go():
        bot = TelegramBot(str(tok), CHAT, None, giau_trang_thai=True)
        ok, err = await bot.send_text(CHAT, text)
        print(f"telegram ok={ok} err={err or ''}".strip())
        return bool(ok)

    return bool(asyncio.run(_go()))


def _try_close_excel() -> str:
    """Best-effort close Excel on the machine running Javis. Returns status text."""
    # 1) Prefer in-process plugin (works when Javis runs on Windows host)
    try:
        sys.path.insert(0, "/app/system/plugins/desktop-apps")
        sys.path.insert(0, "/app/server")
        import plugin as desktop_apps  # type: ignore

        fn = None
        for item in getattr(desktop_apps, "TOOLS", []) or getattr(desktop_apps, "tools", []) or []:
            name = getattr(item, "name", None) or (item.get("name") if isinstance(item, dict) else None)
            if name == "javis_app_close":
                fn = getattr(item, "fn", None) or getattr(item, "handler", None) or (item.get("fn") if isinstance(item, dict) else None)
                break
        if fn is None and hasattr(desktop_apps, "javis_app_close"):
            fn = desktop_apps.javis_app_close
        if fn is not None:
            out = fn(name="excel", force=False)
            print("app_close plugin:", str(out)[:800])
            return str(out)
    except Exception as e:
        print(f"plugin close path: {type(e).__name__}: {e}")

    # 2) Fallback hub: search then javis_run_tool
    try:
        import json as _json
        import urllib.request
        import mcp_hub

        hub = mcp_hub.hub_url()
        tok = mcp_hub.hub_token()
        vault = "/brains/Cung cấp thông tin công an"

        def rpc(method, params=None, _id=1):
            body = {"jsonrpc": "2.0", "id": _id, "method": method, "params": params or {}}
            req = urllib.request.Request(
                hub,
                data=_json.dumps(body).encode(),
                headers={
                    "Authorization": f"Bearer {tok}",
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/event-stream",
                    "X-Javis-Vault": vault.encode("utf-8").decode("latin-1"),
                    "X-Javis-Mode": "full",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=60) as resp:
                raw = resp.read().decode()
            if "data:" in raw:
                datas = [line[5:].strip() for line in raw.splitlines() if line.startswith("data:")]
                raw = datas[-1] if datas else raw
            return _json.loads(raw)

        rpc(
            "initialize",
            {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "cua-kiem-dong-bao-cao", "version": "1"},
            },
        )
        req = urllib.request.Request(
            hub,
            data=_json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}}).encode(),
            headers={
                "Authorization": f"Bearer {tok}",
                "Content-Type": "application/json",
                "X-Javis-Vault": vault.encode("utf-8").decode("latin-1"),
                "X-Javis-Mode": "full",
            },
            method="POST",
        )
        urllib.request.urlopen(req, timeout=30).read()
        rpc("tools/call", {"name": "javis_search_tools", "arguments": {"query": "javis_app_close"}}, _id=8)
        res = rpc(
            "tools/call",
            {
                "name": "javis_run_tool",
                "arguments": {"name": "javis_app_close", "args": {"name": "excel", "force": False}},
            },
            _id=9,
        )
        texts = [
            c.get("text") or ""
            for c in ((res.get("result") or {}).get("content") or [])
            if c.get("type") == "text"
        ]
        out = "\n".join(texts).strip() or str(res)[:500]
        print("app_close hub:", out[:800])
        return out
    except Exception as e:
        msg = f"app_close skip/error: {type(e).__name__}: {e}"
        print(msg)
        return msg


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--wait", type=int, default=DEFAULT_WAIT, help="Giây chờ Drive sync sau nhắc (mặc định 60)")
    ap.add_argument("--skip-wait", action="store_true", help="Không chờ sync")
    ap.add_argument("--no-telegram", action="store_true")
    ap.add_argument("--no-close", action="store_true")
    args = ap.parse_args()

    msg = (
        "[CỬA KIỂM SỔ] Hệ thống sắp bóc tách/ghi Bao_cao_cong_van.xlsx. "
        "Anh đóng Excel nếu đang mở file này. VPS sẽ tự chạy tiếp sau khoảng "
        f"{args.wait}s (không chờ trả lời), ghi đè đúng 1 file sống trên Drive."
    )
    if not args.no_telegram:
        _send_telegram(msg)
    else:
        print("skip telegram")

    if not args.no_close:
        _try_close_excel()
    else:
        print("skip app_close")

    if not args.skip_wait and args.wait > 0:
        print(f"wait {args.wait}s for Drive sync...")
        time.sleep(args.wait)
    print("READY continue boc-tach/ghi-so")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
