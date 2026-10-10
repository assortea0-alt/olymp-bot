# -*- coding: utf-8 -*-
"""
Узнать свой chat_id.

Как пользоваться:
  1. В config.json впиши bot_token (chat_id пока оставь как есть).
  2. Открой своего бота в Telegram и отправь ему /start.
  3. python get_chat_id.py  -> выведет chat_id, его и вставь в config.json.
"""

import json
import sys
from pathlib import Path

import requests

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError):
        pass

cfg_path = Path(__file__).resolve().parent / "config.json"
if not cfg_path.exists():
    sys.exit("Сначала скопируй config.example.json в config.json и впиши bot_token.")

token = json.loads(cfg_path.read_text(encoding="utf-8")).get("bot_token", "")
if not token or "СЮДА" in token:
    sys.exit("В config.json не вписан bot_token.")

r = requests.get("https://api.telegram.org/bot%s/getUpdates" % token, timeout=30)
data = r.json()

if not data.get("ok"):
    sys.exit("Telegram не принял токен: %s" % data.get("description"))

seen = {}
for upd in data.get("result", []):
    msg = upd.get("message") or upd.get("edited_message") or {}
    chat = msg.get("chat") or {}
    if chat.get("id"):
        who = chat.get("username") or chat.get("first_name") or chat.get("title") or ""
        seen[chat["id"]] = "%s (%s)" % (who, chat.get("type"))

if not seen:
    print("Апдейтов нет. Открой бота в Telegram, нажми /start и запусти скрипт снова.")
else:
    print("Найденные чаты:")
    for cid, who in seen.items():
        print("  chat_id = %s   %s" % (cid, who))
    print("\nВставь нужный chat_id в config.json.")
