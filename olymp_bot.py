# -*- coding: utf-8 -*-
"""
Мониторинг олимпиад по информатике -> Telegram.

За один запуск:
  1. Качает Google-таблицу «Перечневые олимпиады» и разбирает лист-календарь.
  2. Сравнивает со снимком прошлого запуска -> сообщает об изменениях
     (новые этапы, сдвиг дат, открытие/закрытие регистрации).
  3. Обходит официальные сайты тех олимпиад, где сроки ещё не объявлены,
     и сообщает, если на странице появились новые даты.
  4. Напоминает о дедлайнах регистрации и о начале этапов за N дней.

  python olymp_bot.py            # обычный запуск (для планировщика)
  python olymp_bot.py --digest   # прислать полную картину прямо сейчас
  python olymp_bot.py --dry-run  # ничего не отправлять, только показать в консоли
"""

import argparse
import datetime as dt
import html
import json
import os
import re
import sys
import time
from pathlib import Path

import openpyxl
import requests
from bs4 import BeautifulSoup

HERE = Path(__file__).resolve().parent
CONFIG_PATH = HERE / "config.json"
STATE_PATH = HERE / "state.json"
SHEET_CACHE = HERE / ".sheet.xlsx"
LOG_PATH = HERE / "bot.log"

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36")

DATE_FULL = re.compile(r"\b(\d{1,2})[.\-/](\d{1,2})[.\-/](20\d{2})\b")
DATE_SHORT = re.compile(r"\b(\d{1,2})\.(\d{1,2})\b(?!\s*[.\-/]?\s*\d)")
MONTH_WORDS = [
    ("январ", 1), ("феврал", 2), ("март", 3), ("апрел", 4), ("мая", 5), ("май", 5),
    ("июн", 6), ("июл", 7), ("август", 8), ("сентябр", 9), ("октябр", 10),
    ("ноябр", 11), ("декабр", 12),
]
DATE_WORD = re.compile(
    r"\b(\d{1,2})\s+("
    r"январ\w*|феврал\w*|март\w*|апрел\w*|мая|май|июн\w*|июл\w*|"
    r"август\w*|сентябр\w*|октябр\w*|ноябр\w*|декабр\w*)",
    re.IGNORECASE)
RELEVANT = re.compile(r"регистрац|заявк|приём|прием|срок|тур|этап|отбороч|заключительн",
                      re.IGNORECASE)

SKIP_REG = ("ЗАКРЫТА", "НЕ В ПЕРЕЧНЕ")


# --------------------------------------------------------------------------- io

class _Tee:
    """Пишет сразу в консоль и в лог-файл; молча терпит сломанный поток."""

    def __init__(self, *streams):
        self.streams = [s for s in streams if s is not None]

    def write(self, data):
        for s in self.streams:
            try:
                s.write(data)
            except Exception:
                pass
        return len(data)

    def flush(self):
        for s in self.streams:
            try:
                s.flush()
            except Exception:
                pass


def setup_output():
    """Трасса всегда идёт в bot.log — под планировщиком консоли нет вообще.

    Заодно вывод переводится в UTF-8: консоль Windows по умолчанию в cp1251,
    падает на эмодзи и портит кириллицу.
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass
    try:
        if LOG_PATH.exists() and LOG_PATH.stat().st_size > 1_000_000:
            LOG_PATH.unlink()
        logfile = open(LOG_PATH, "a", encoding="utf-8", buffering=1)
    except OSError:
        return                                  # без лога, но работать продолжим
    sys.stdout = _Tee(sys.stdout, logfile)
    sys.stderr = _Tee(sys.stderr, logfile)
    print("\n=== запуск %s ===" % dt.datetime.now().isoformat(timespec="seconds"))


def _env_flag(name, default):
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    return raw.strip().lower() in ("1", "true", "yes", "да", "on")


def load_config():
    """Настройки из config.json, но переменные окружения главнее.

    Локально удобнее файл, а в облаке (GitHub Actions) секреты приходят
    только через окружение, и файла с токеном там нет вовсе.
    """
    cfg = {}
    if CONFIG_PATH.exists():
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))

    if os.environ.get("TG_BOT_TOKEN"):
        cfg["bot_token"] = os.environ["TG_BOT_TOKEN"].strip()
    if os.environ.get("TG_CHAT_ID"):
        cfg["chat_id"] = os.environ["TG_CHAT_ID"].strip()
    if os.environ.get("SHEET_ID"):
        cfg["sheet_id"] = os.environ["SHEET_ID"].strip()
    if os.environ.get("MY_OLYMPIADS"):
        cfg["my_olympiads"] = [s.strip() for s in
                               os.environ["MY_OLYMPIADS"].split(",") if s.strip()]
    cfg["watch_sites"] = _env_flag("WATCH_SITES", cfg.get("watch_sites", True))
    cfg["only_in_perechen"] = _env_flag("ONLY_IN_PERECHEN",
                                        cfg.get("only_in_perechen", True))

    token, chat = str(cfg.get("bot_token", "")), str(cfg.get("chat_id", ""))
    if not token or not chat:
        sys.exit("Не задан bot_token или chat_id — ни в config.json, "
                 "ни в переменных TG_BOT_TOKEN / TG_CHAT_ID.")
    if "СЮДА" in token or "СЮДА" in chat:
        sys.exit("В config.json ещё стоят заглушки. Впиши настоящие bot_token "
                 "(от @BotFather) и chat_id (его покажет get_chat_id.py).")
    cfg.setdefault("sheet_id", "16wKSRx8b7BC8YQeT9ulPl3gHKJ7hHWq0")
    cfg.setdefault("remind_days_before", [14, 7, 3, 1, 0])
    cfg.setdefault("stage_remind_days", [3, 1, 0])
    cfg.setdefault("watch_sites", True)
    cfg.setdefault("only_in_perechen", True)
    cfg.setdefault("my_olympiads", [])
    return cfg


def load_state():
    if STATE_PATH.exists():
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except ValueError:
            pass
    return {}


def save_state(state):
    state["last_run"] = dt.datetime.now().isoformat(timespec="seconds")
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=1),
                          encoding="utf-8")


# ---------------------------------------------------------------------- telegram

def tg_send(cfg, text, dry_run=False):
    """Отправляет сообщение, разбивая по лимиту Telegram в 4096 символов.

    Возвращает True, только если Telegram принял все куски. Вызывающий
    обязан проверить результат: если отправка не удалась, состояние
    сохранять нельзя, иначе напоминание будет считаться доставленным
    и больше никогда не повторится.
    """
    chunks, cur = [], ""
    for line in text.split("\n"):
        if len(cur) + len(line) + 1 > 3800:
            chunks.append(cur)
            cur = ""
        cur += line + "\n"
    if cur.strip():
        chunks.append(cur)

    ok = True
    for chunk in chunks:
        if dry_run:
            print("--- telegram ---")
            print(re.sub(r"<[^>]+>", "", chunk))
            continue
        url = "https://api.telegram.org/bot%s/sendMessage" % cfg["bot_token"]
        try:
            r = requests.post(url, timeout=30, data={
                "chat_id": cfg["chat_id"],
                "text": chunk,
                "parse_mode": "HTML",
                "disable_web_page_preview": "true",
            })
            if r.status_code != 200:
                ok = False
                print("ОШИБКА: Telegram ответил %s: %s"
                      % (r.status_code, r.text[:300]))
                if r.status_code == 401:
                    print("  -> неверный bot_token (секрет TG_BOT_TOKEN)")
                elif r.status_code == 400:
                    print("  -> похоже на неверный chat_id (секрет TG_CHAT_ID) "
                          "или боту не нажимали /start")
        except requests.RequestException as e:
            ok = False
            print("ОШИБКА: не удалось отправить в Telegram: %s" % e)
        time.sleep(0.4)
    return ok


# ------------------------------------------------------------------------- даты

def season_start_year(today):
    """Учебный год: сентябрь 2026 и январь 2027 относятся к сезону 2026/27."""
    return today.year if today.month >= 8 else today.year - 1


def parse_dates(raw, season):
    """Все календарные даты из строки вида «01.10.2026 – 12.01.2027»."""
    if not raw:
        return []
    out = []
    for m in DATE_FULL.finditer(raw):
        try:
            out.append(dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))))
        except ValueError:
            pass
    rest = DATE_FULL.sub(" ", raw)
    for m in DATE_SHORT.finditer(rest):
        day, mon = int(m.group(1)), int(m.group(2))
        if not 1 <= mon <= 12:
            continue
        year = season if mon >= 8 else season + 1
        try:
            out.append(dt.date(year, mon, day))
        except ValueError:
            pass
    return sorted(set(out))


def normalized_dates(text):
    """Нормализованные строки-даты со страницы — подпись для сравнения."""
    found = set()
    for m in DATE_FULL.finditer(text):
        found.add("%02d.%02d.%s" % (int(m.group(1)), int(m.group(2)), m.group(3)))
    for m in DATE_WORD.finditer(text):
        word = m.group(2).lower()
        for stem, num in MONTH_WORDS:
            if word.startswith(stem):
                found.add("%02d.%02d" % (int(m.group(1)), num))
                break
    return found


def human_date(d):
    return d.strftime("%d.%m.%Y")


# ---------------------------------------------------------------------- таблица

def fetch_sheet(cfg):
    """Качает таблицу в xlsx и возвращает список строк календаря."""
    url = ("https://docs.google.com/spreadsheets/d/%s/export?format=xlsx"
           % cfg["sheet_id"])
    r = requests.get(url, timeout=60, headers={"User-Agent": UA})
    r.raise_for_status()
    if r.content[:2] != b"PK":
        raise RuntimeError("Вместо xlsx пришёл не тот файл — возможно, "
                           "у таблицы закрыли доступ по ссылке.")
    SHEET_CACHE.write_bytes(r.content)

    wb = openpyxl.load_workbook(SHEET_CACHE, data_only=True)
    ws = wb.worksheets[0]
    rows = []
    for row in ws.iter_rows(min_row=2):
        cells = [(c.value if c.value is not None else "") for c in row]
        cells += [""] * (9 - len(cells))
        name = str(cells[0]).strip()
        if not name:
            continue
        link = ""
        if len(row) > 6 and row[6].hyperlink:
            link = row[6].hyperlink.target or ""
        rows.append({
            "olympiad": name,
            "level": str(cells[1]).strip(),
            "stage": str(cells[2]).strip(),
            "date": str(cells[3]).strip(),
            "date_status": str(cells[4]).strip(),
            "note": str(cells[5]).strip(),
            "link": link,
            "reg": str(cells[7]).strip(),
        })
    return rows


def row_key(r):
    return "%s | %s" % (r["olympiad"], r["stage"])


def is_reg_deadline(r, dates):
    """Конец срока в строке — это правда дедлайн записи, а не дата тура.

    Отдельная строка «Регистрация» — да. Диапазон вида «01.10 – 12.01» у
    отборочного — тоже да: запись и отбор идут одновременно. А одиночная дата
    тура («I отборочный 18.10.2026») дедлайном записи не является.
    """
    if "егистрация" in r["stage"]:
        return True
    return len(dates) > 1


def diff_sheet(old_rows, new_rows):
    """Что изменилось в таблице с прошлого запуска."""
    old = dict(old_rows or {})
    msgs = []
    seen = set()
    for r in new_rows:
        key = row_key(r)
        seen.add(key)
        prev = old.get(key)
        if prev is None:
            msgs.append("\U0001F195 <b>Новый этап в таблице</b>\n%s\nДата: %s\nРегистрация: %s"
                        % (html.escape(key), html.escape(r["date"] or "—"),
                           html.escape(r["reg"] or "—")))
            continue
        changes = []
        if prev.get("date") != r["date"]:
            changes.append("дата: <s>%s</s> → <b>%s</b>"
                           % (html.escape(prev.get("date") or "—"),
                              html.escape(r["date"] or "—")))
        if prev.get("reg") != r["reg"]:
            changes.append("регистрация: <s>%s</s> → <b>%s</b>"
                           % (html.escape(prev.get("reg") or "—"),
                              html.escape(r["reg"] or "—")))
        if prev.get("date_status") != r["date_status"]:
            changes.append("статус: %s → <b>%s</b>"
                           % (html.escape(prev.get("date_status") or "—"),
                              html.escape(r["date_status"] or "—")))
        if changes:
            body = "\U0001F4E3 <b>Изменение в таблице</b>\n%s\n%s" % (
                html.escape(key), "\n".join("• " + c for c in changes))
            if r["link"]:
                body += '\n<a href="%s">сайт олимпиады</a>' % html.escape(r["link"])
            msgs.append(body)
    for key in old:
        if key not in seen:
            msgs.append("➖ <b>Этап убрали из таблицы</b>\n%s" % html.escape(key))
    return msgs


# ------------------------------------------------------------------------ сайты

def page_text(url):
    headers = {"User-Agent": UA, "Accept-Language": "ru,en"}
    try:
        r = requests.get(url, timeout=25, headers=headers)
    except requests.exceptions.SSLError as e:
        # бывает у вузовских сайтов: не собирается цепочка сертификатов.
        # Проверку не отключаем — лучше честно сообщить и смотреть руками.
        raise RuntimeError("сертификат сайта не проверился (%s)" % e.__class__.__name__)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or r.encoding
    soup = BeautifulSoup(r.text, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()
    return re.sub(r"[ \t\xa0]+", " ", soup.get_text("\n"))


def snippets_for(text, dates, limit=3):
    """Фразы со страницы, в которых встретились новые даты."""
    out = []
    for line in (l.strip() for l in text.split("\n")):
        if len(line) < 12 or len(line) > 300:
            continue
        if not RELEVANT.search(line):
            continue
        if any(d.split(".2")[0] in line or d in line for d in dates):
            out.append(line)
        if len(out) >= limit:
            break
    return out


def watch_sites(cfg, rows, state, notify):
    """Обходит сайты олимпиад с необъявленными сроками, ищет новые даты."""
    targets = {}
    for r in rows:
        if not r["link"]:
            continue
        if any(s in r["reg"].upper() for s in SKIP_REG):
            continue
        unknown = ("НЕ ОБЪЯВЛЕН" in r["reg"].upper()
                   or "Ожидается" in r["date_status"])
        if unknown:
            targets.setdefault(r["link"], set()).add(r["olympiad"])

    sites = state.setdefault("sites", {})
    for url, names in sorted(targets.items()):
        entry = sites.setdefault(url, {"dates": [], "fails": 0})
        try:
            text = page_text(url)
        except Exception as e:                      # сеть, 403, кривой сертификат
            entry["fails"] = entry.get("fails", 0) + 1
            print("не открылся %s (%s), подряд неудач: %d" % (url, e, entry["fails"]))
            if entry["fails"] == 5:
                notify.append("⚠️ <b>Сайт не открывается 5 запусков подряд</b>\n%s\n%s"
                              % (html.escape(", ".join(sorted(names))),
                                 html.escape(url)))
            continue
        entry["fails"] = 0
        found = normalized_dates(text)
        known = set(entry.get("dates") or [])
        fresh = sorted(found - known)
        if known and fresh:
            body = ["\U0001F50E <b>На сайте появились новые даты</b>",
                    html.escape(", ".join(sorted(names))),
                    "Даты: <b>%s</b>" % html.escape(", ".join(fresh[:12]))]
            for s in snippets_for(text, fresh):
                body.append("· <i>%s</i>" % html.escape(s))
            body.append('<a href="%s">открыть сайт</a>' % html.escape(url))
            notify.append("\n".join(body))
        entry["dates"] = sorted(found)
        time.sleep(1.0)


# ------------------------------------------------------------------- напоминания

def interesting(cfg, r):
    """Стоит ли вообще напоминать про этот этап."""
    reg = r["reg"].upper()
    if any(s in reg for s in SKIP_REG) and cfg["only_in_perechen"]:
        return False
    if cfg["my_olympiads"]:
        return any(m.lower() in r["olympiad"].lower() for m in cfg["my_olympiads"])
    return True


def reminders(cfg, rows, state, today):
    """Напоминания о дедлайнах регистрации и о начале этапов."""
    season = season_start_year(today)
    sent = set(state.setdefault("sent", []))
    msgs = []

    def fire(key, when, days, title, r):
        stamp = "%s|%s|%s|%s" % (key, when.isoformat(), days, title)
        if stamp in sent:
            return
        if days == 0:
            lead = "<b>сегодня</b>"
        elif days == 1:
            lead = "<b>завтра</b>"
        else:
            lead = "через <b>%d дн.</b>" % days
        body = ["⏰ <b>%s</b> — %s (%s)" % (title, lead, human_date(when)),
                "%s — %s" % (html.escape(r["olympiad"]), html.escape(r["stage"]))]
        if r["level"] and r["level"] not in ("—", ""):
            body.append("Уровень %s" % html.escape(r["level"]))
        if r["date"]:
            body.append("Срок по таблице: %s" % html.escape(r["date"]))
        if r["note"]:
            body.append("<i>%s</i>" % html.escape(r["note"]))
        if r["link"]:
            body.append('<a href="%s">страница олимпиады</a>' % html.escape(r["link"]))
        msgs.append("\n".join(body))
        sent.add(stamp)

    for r in rows:
        if not interesting(cfg, r):
            continue
        dates = parse_dates(r["date"], season)
        if not dates:
            continue
        key = row_key(r)
        reg = r["reg"].upper()

        if "ОТКРЫТА" in reg and is_reg_deadline(r, dates):
            deadline = max(dates)
            days = (deadline - today).days
            if days in cfg["remind_days_before"]:
                fire(key, deadline, days, "Дедлайн регистрации", r)
        elif "ОТКРОЕТСЯ" in reg:
            # дата открытия записи часто стоит в самой колонке «Регистрация»
            # («ОТКРОЕТСЯ 10.10»), а в колонке даты — сроки самого тура
            reg_dates = parse_dates(r["reg"], season)
            opening = min(reg_dates) if reg_dates else min(dates)
            days = (opening - today).days
            if days in cfg["remind_days_before"]:
                fire(key, opening, days, "Открывается регистрация", r)

        if "егистрация" not in r["stage"]:
            start = min(dates)
            days = (start - today).days
            if days in cfg["stage_remind_days"]:
                fire(key, start, days, "Начало этапа", r)

    state["sent"] = sorted(sent)[-4000:]
    return msgs


# ---------------------------------------------------------------------- дайджест

def digest(cfg, rows, today):
    season = season_start_year(today)
    open_now, soon, closed, waiting = [], [], [], []
    for r in rows:
        reg = r["reg"].upper()
        if "НЕ В ПЕРЕЧНЕ" in reg and cfg["only_in_perechen"]:
            continue
        dates = parse_dates(r["date"], season)
        line = "• <b>%s</b> (ур. %s) — %s" % (
            html.escape(r["olympiad"]), html.escape(r["level"] or "—"),
            html.escape(r["stage"]))
        if "ОТКРЫТА" in reg:
            if not dates:
                open_now.append((9999, line + " — <b>срок не объявлен</b>, "
                                              "записаться можно сейчас"))
            elif is_reg_deadline(r, dates):
                left = (max(dates) - today).days
                open_now.append((left, line + " — запись до <b>%s</b> (%d дн.)"
                                 % (human_date(max(dates)), left)))
            else:
                left = (min(dates) - today).days
                open_now.append((left, line + " — тур <b>%s</b> (%d дн.)"
                                 % (human_date(min(dates)), left)))
        elif "ОТКРОЕТСЯ" in reg:
            reg_dates = parse_dates(r["reg"], season) or dates
            if reg_dates:
                opening = min(reg_dates)
                soon.append((opening, line + " — запись откроется <b>%s</b>"
                             % human_date(opening)))
        elif "ЗАКРЫТА" in reg:
            closed.append(line)
        elif "НЕ ОБЪЯВЛЕН" in reg or "Ожидается" in r["date_status"]:
            waiting.append(line)

    parts = ["\U0001F4CB <b>Олимпиады по информатике — срез на %s</b>" % human_date(today)]
    if open_now:
        parts.append("\n\U0001F7E2 <b>Регистрация открыта</b>")
        parts += [l for _, l in sorted(open_now)]
    if soon:
        parts.append("\n\U0001F535 <b>Скоро откроется</b>")
        parts += [l for _, l in sorted(soon)]
    if waiting:
        parts.append("\n\U0001F7E1 <b>Сроки не объявлены (слежу за сайтами)</b>")
        parts += sorted(set(waiting))
    if closed:
        parts.append("\n⚫ <b>Регистрация закрыта</b>")
        parts += sorted(set(closed))
    return "\n".join(parts)


# --------------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="не отправлять в Telegram, печатать в консоль")
    ap.add_argument("--digest", action="store_true",
                    help="прислать полную картину по всем олимпиадам")
    ap.add_argument("--no-sites", action="store_true",
                    help="не обходить сайты, только таблица и напоминания")
    ap.add_argument("--check-sites", action="store_true",
                    help="самопроверка: какие сайты открываются и сколько дат видно")
    args = ap.parse_args()

    setup_output()

    cfg = load_config()
    state = load_state()
    today = dt.date.today()
    notify = []

    try:
        rows = fetch_sheet(cfg)
    except Exception as e:
        tg_send(cfg, "⚠️ <b>Не смог прочитать таблицу</b>\n%s" % html.escape(str(e)),
                args.dry_run)
        sys.exit("Таблица не прочитана: %s" % e)
    print("строк в таблице: %d" % len(rows))

    if args.check_sites:
        links = {}
        for r in rows:
            if r["link"]:
                links.setdefault(r["link"], set()).add(r["olympiad"])
        print("\nПроверяю %d сайтов...\n" % len(links))
        ok = 0
        for url in sorted(links):
            try:
                found = normalized_dates(page_text(url))
                ok += 1
                print("  OK    дат: %-4d %s" % (len(found), url))
            except Exception as e:
                print("  СБОЙ  %s\n          %s" % (url, str(e)[:160]))
            time.sleep(1.0)
        print("\nОткрылось %d из %d." % (ok, len(links)))
        return

    first_run = not state.get("sheet_rows")
    if not first_run:
        notify += diff_sheet(state.get("sheet_rows"), rows)

    if cfg["watch_sites"] and not args.no_sites:
        watch_sites(cfg, rows, state, notify)

    notify += reminders(cfg, rows, state, today)

    state["sheet_rows"] = {row_key(r): r for r in rows}

    delivered = True
    if first_run or args.digest:
        delivered &= tg_send(cfg, digest(cfg, rows, today), args.dry_run)

    if notify:
        delivered &= tg_send(cfg, "\n\n".join(notify), args.dry_run)
        print("сообщений к отправке: %d" % len(notify))
    else:
        print("изменений нет")

    # Состояние — это память о доставленном. Сохранять его при сбое отправки
    # нельзя: напоминание окажется помечено как отправленное и не повторится
    # никогда. Лучше упасть, повторить в следующий раз и получить от GitHub
    # письмо о красном запуске.
    if not delivered:
        print("состояние НЕ сохранено — на следующем запуске попробуем снова")
        sys.exit("Telegram не принял сообщения, см. ошибки выше")

    if args.dry_run:
        # репетиция не должна глушить настоящую отправку: если записать
        # напоминание в состояние, бот сочтёт его доставленным
        print("пробный запуск: состояние не изменено")
        return

    save_state(state)
    print("готово, состояние сохранено")


if __name__ == "__main__":
    main()
