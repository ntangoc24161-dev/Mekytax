#!/usr/bin/env python3
"""Gửi tin/file xem trước qua bot Telegram CA @Kiem_soat_dieu_phoi_bot.

Dùng thay POST /telegram/send-file (bot Javis chính) cho hồ sơ Công an.
Token lấy từ channel_accounts acc_656f8aaa4e - không hard-code.

Ví dụ:
  python scripts/gui-telegram-bot-ca.py --text "[XEM TRƯỚC] ..."
  python scripts/gui-telegram-bot-ca.py --file /abs/path.pdf --caption "[XEM TRƯỚC] ..."
  python scripts/gui-telegram-bot-ca.py --file a.pdf --file b.jpg --caption "..."
"""
from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path

ACC_ID = "acc_656f8aaa4e"
DEFAULT_CHAT = "8799384849"


def _load_token() -> str:
    sys.path.insert(0, "/app/server")
    import channel_accounts  # noqa: WPS433

    tok = channel_accounts.get_token(ACC_ID) or ""
    if not tok or str(tok).startswith("enc:"):
        raise SystemExit(
            f"Không lấy được token bot CA ({ACC_ID}). Kiểm tra trang Chatbot → kênh Telegram CA."
        )
    return str(tok)


async def _run(chat_id: str, text: str, files: list[str], caption: str) -> int:
    from telegram_bot import TelegramBot  # noqa: WPS433

    bot = TelegramBot(_load_token(), chat_id, None, giau_trang_thai=True)
    code = 0

    if text:
        ok, err = await bot.send_text(chat_id, text)
        print(f"text ok={ok} err={err or ''}".strip())
        if not ok:
            code = 1

    cap = caption or ""
    for i, path in enumerate(files):
        p = Path(path)
        if not p.is_file():
            print(f"file missing: {path}")
            code = 1
            continue
        # Caption chỉ gắn file đầu để khỏi spam; file sau gửi không caption.
        this_cap = cap if i == 0 else ""
        ok, err = await bot.send_file(str(p.resolve()), this_cap, chat=chat_id)
        print(f"file {p.name} ok={ok} err={err or ''}".strip())
        if not ok:
            code = 1
    return code


def main() -> int:
    ap = argparse.ArgumentParser(description="Gửi tin/file qua @Kiem_soat_dieu_phoi_bot")
    ap.add_argument("--chat-id", default=DEFAULT_CHAT, help="Telegram chat_id chủ (mặc định 8799384849)")
    ap.add_argument("--text", default="", help="Tin chữ (tuỳ chọn)")
    ap.add_argument("--file", action="append", default=[], help="Đường dẫn file tuyệt đối (lặp được)")
    ap.add_argument("--caption", default="", help="Caption cho file đầu")
    args = ap.parse_args()
    if not args.text and not args.file:
        ap.error("Cần --text và/hoặc --file")
    return asyncio.run(_run(str(args.chat_id).strip(), args.text, args.file, args.caption))


if __name__ == "__main__":
    raise SystemExit(main())
