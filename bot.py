#!/usr/bin/env python3
"""Gavhar Gold — тилло фактура назорати Telegram боти (Railway учун).

Муҳит ўзгарувчилари (Railway → Variables):
  BOT_TOKEN      — @BotFather берган токен (мажбурий)
  ALLOWED_USERS  — ботдан фойдаланишга рухсат этилган Telegram ID лар, вергул билан (мас. 123456,789012)
  DATA_DIR       — маълумот папкаси (Railway Volume шу ерга уланади). Стандарт: /data
"""
import asyncio
import datetime as dt
import filecmp
import json
import logging
import os
import re
import shutil
import tempfile
from pathlib import Path

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ChatAction
from telegram.ext import (Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters)

import nazorat

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
log = logging.getLogger("gavhar-bot")

BASE = Path(__file__).resolve().parent
DATA = Path(os.environ.get("DATA_DIR", "/data"))
TRANSH = DATA / "transhlar.json"
LEDGER = DATA / "ledger.csv"
BACKUP = DATA / "backup"
ALLOWED = {int(x) for x in re.findall(r"\d+", os.environ.get("ALLOWED_USERS", ""))}
LOCK = asyncio.Lock()

TYPE_ALIASES = [
    ("Ван Клиф", ("ван клиф", "ванклиф", "van cleef", "van clif", "van klif", "клиф")),
    ("Бриллиант", ("бриллиант", "брилиант", "brilliant", "brilyant")),
    ("Фианит", ("фианит", "fianit")),
    ("Дорика", ("дорика", "dorika", "занжир")),
    ("Браслет", ("браслет", "braslet", "билак")),
    ("Детский", ("детский", "детски", "detskiy", "болалар")),
    ("Кузсиз", ("кузсиз", "кўзсиз", "kuzsiz", "ko'zsiz")),
    ("Вакс", ("вакс", "waks", "vaks")),
]

HELP = (
    "🟡 *Gavhar — тилло нарх назорати*\n\n"
    "📎 *Фактура* — тизимдан олинган Excel (.xlsx) экспортини шу ерга ташланг. Бот ҳар позицияни "
    "транш нархи билан солиштиради, FIFO бўйича транш қолдиғини янгилайди ва ҳисобот қайтаради.\n\n"
    "*Буйруқлар:*\n"
    "/holat — пул қолдиғи, траншлар, фарқлар + ҳисобот файли\n"
    "/narxlar — траншлар бўйича 1 гр нархлари\n"
    "/transh — янги пул (транш) қўшиш. Буйруқдан кейин шу кўринишда ёзинг:\n"
    "`/transh\n13.10.2026\nБиржа нархи 134.81$\nБриллиант 89.97$\nВакс+кузсиз 87.85$\nДорика 90.49$\n"
    "Фианит 86.27$\nБраслет 90.49$\nДетский 90.49$\nВан Клиф 90.49$\nСумма 707000$`\n"
    "/yop 1 2026-10-13 — 1-траншни шу санадан ёпиш (қолдиғи бўлса ҳам кейинги товарлар унга ёзилмайди)\n"
    "/och 1 — ёпилган траншни қайта очиш\n"
    "/chegara 0.3 — рухсат этилган фарқ, $/гр\n"
    "/bekor — охирги юкланган фактурани бекор қилиш\n"
    "/reestr — барча приход қаторлари (CSV)\n"
    "/myid — Telegram ID ингиз"
)


# ---------------- ёрдамчи ----------------
def ensure_data():
    DATA.mkdir(parents=True, exist_ok=True)
    BACKUP.mkdir(exist_ok=True)
    if not TRANSH.exists():
        src = BASE / "seed_transhlar.json"
        if src.exists():
            shutil.copy(src, TRANSH)
        else:  # бошланғич файл йўқ бўлса — бўш созлама (транш кейин mini app дан қўшилади)
            TRANSH.write_text(json.dumps({"chegara_usd_gr": 0.3, "transhlar": []}, ensure_ascii=False, indent=2), encoding="utf-8")
    if not LEDGER.exists():
        src = BASE / "seed_ledger.csv"
        if src.exists():
            shutil.copy(src, LEDGER)
        else:
            nazorat.save_ledger(str(LEDGER), [])


def load_cfg():
    return json.loads(TRANSH.read_text(encoding="utf-8"))


def save_cfg(cfg):
    backup(TRANSH)
    tmp = TRANSH.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(TRANSH)


def backup(path: Path):
    if path.exists():
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy(path, BACKUP / f"{path.stem}_{stamp}{path.suffix}")
        olds = sorted(BACKUP.glob(f"{path.stem}_*{path.suffix}"))
        for o in olds[:-30]:  # охирги 30 тасини сақлаймиз
            o.unlink()


def allowed(update: Update) -> bool:
    u = update.effective_user
    return bool(u and u.id in ALLOWED)


async def deny(update: Update):
    uid = update.effective_user.id if update.effective_user else "?"
    await update.effective_message.reply_text(
        f"⛔ Рухсат йўқ. Сизнинг ID: {uid}\nRailway → Variables → ALLOWED_USERS га шу ID ни қўшинг.")


def guard(fn):
    async def wrap(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        if not allowed(update):
            return await deny(update)
        return await fn(update, ctx)
    return wrap


async def send_long(msg, text):
    for i in range(0, len(text), 3900):
        await msg.reply_text(text[i:i + 3900])


async def run_and_reply(update: Update, files=(), caption="Ҳисобот"):
    msg = update.effective_message
    await msg.chat.send_action(ChatAction.TYPING)
    async with LOCK:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / f"Gavhar_nazorat_{dt.date.today():%Y-%m-%d}.xlsx"
            text = await asyncio.to_thread(nazorat.run, str(TRANSH), str(LEDGER), str(out), list(files))
            await send_long(msg, text)
            with open(out, "rb") as f:
                await msg.reply_document(f, filename=out.name, caption=caption)


# ---------------- транш матнини ўқиш ----------------
NUM = r"(\d+(?:[.,]\d+)?)"


def parse_transh(text: str):
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    lines = [l for l in lines if not l.startswith("/transh")]
    res = {"sana": None, "birja": None, "lom": None, "summa": None, "narxlar": {}}
    for l in lines:
        low = l.lower().replace("ё", "е")
        m = re.search(r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})", l) or re.search(r"(\d{4})-(\d{2})-(\d{2})", l)
        if m and not res["sana"]:
            g = m.groups()
            d = dt.date(int(g[0]), int(g[1]), int(g[2])) if len(g[0]) == 4 else dt.date(int(g[2]), int(g[1]), int(g[0]))
            res["sana"] = d.isoformat()
            continue
        nums = [float(x.replace(",", ".")) for x in re.findall(NUM, l.replace(" ", ""))]
        if not nums:
            continue
        if "биржа" in low or "birja" in low or "999" in low and "нарх" in low:
            res["birja"] = nums[-1]
            continue
        if "лом" in low or "lom" in low:
            res["lom"] = nums[-1]
            continue
        hit = False
        for name, keys in TYPE_ALIASES:
            if any(k in low for k in keys):
                res["narxlar"][name] = nums[-1]
                hit = True
        if hit:
            if "вакс" in low and ("кузсиз" in low or "кўзсиз" in low):
                res["narxlar"]["Вакс"] = res["narxlar"]["Кузсиз"] = nums[-1]
            continue
        big = max(nums)
        if big >= 1000:
            res["summa"] = big
    return res


def check_transh(p, cfg):
    errs, notes = [], []
    if not p["sana"]:
        errs.append("сана топилмади (мас. 13.10.2026)")
    if not p["birja"]:
        errs.append("биржа нархи топилмади (мас. «Биржа нархи 134.81»)")
    if not p["summa"]:
        errs.append("пул суммаси топилмади (мас. «Сумма 707000»)")
    if not p["narxlar"]:
        errs.append("тур нархлари топилмади")
    if p["birja"]:
        lom = p["birja"] * 585 / 999
        notes.append(f"Лом (биржа×585/999) = {lom:.2f}" + (f", сиз ёзгансиз {p['lom']:.2f}" if p["lom"] else ""))
        if p["lom"] and abs(lom - p["lom"]) > 0.05:
            notes.append("⚠ Ёзилган лом ҳисобдан фарқ қилади!")
        prev = cfg["transhlar"][-1] if cfg["transhlar"] else None
        for n, v in p["narxlar"].items():
            line = f"{n}: {v:.2f} (ишчилик+НДС {v - lom:.2f}"
            if prev and n in prev["narxlar"]:
                pl = prev["narxlar"][n] - prev["birja_999"] * 585 / 999
                d = (v - lom) - pl
                line += f", олдинги транш {pl:.2f}" + (f", ⚠ {d:+.2f}" if abs(d) > 0.02 else "")
            notes.append(line + ")")
        if prev:
            for n in prev["narxlar"]:
                if n not in p["narxlar"]:
                    notes.append(f"⚠ «{n}» нархи берилмади — бу турдаги товар шу траншда ҳисобланмайди (Дорика нархи олинади)")
    return errs, notes


# ---------------- ҳандлерлар ----------------
async def start(update: Update, ctx):
    if not allowed(update):
        return await deny(update)
    await update.message.reply_text(HELP, parse_mode="Markdown")


async def myid(update: Update, ctx):
    await update.message.reply_text(f"Сизнинг Telegram ID: {update.effective_user.id}")


@guard
async def holat(update: Update, ctx):
    await run_and_reply(update)


@guard
async def narxlar(update: Update, ctx):
    cfg = load_cfg()
    L = [f"Рухсат чегараси: ±{cfg.get('chegara_usd_gr', 0.3)} $/гр"]
    for t in cfg["transhlar"]:
        st = f" — ёпилган {t['yopilgan_sana']}" if t.get("yopilgan_sana") else ""
        L.append(f"\nТранш №{t['no']} | {t['sana']} | {t['summa_usd']:,.0f} $ | биржа {t['birja_999']}{st}")
        for n, v in t["narxlar"].items():
            L.append(f"   {n}: {v:.2f}")
    await send_long(update.message, "\n".join(L))


@guard
async def transh(update: Update, ctx):
    p = parse_transh(update.message.text or "")
    cfg = load_cfg()
    errs, notes = check_transh(p, cfg)
    if errs:
        return await update.message.reply_text("❌ " + "; ".join(errs) + "\n\nНамуна учун /start")
    ctx.user_data["pending_transh"] = p
    no = max((t["no"] for t in cfg["transhlar"]), default=0) + 1
    text = (f"Янги транш №{no}\nСана: {p['sana']}\nСумма: {p['summa']:,.0f} $\nБиржа 999: {p['birja']}\n\n" + "\n".join(notes)
            + "\n\nСақлайми?")
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("✅ Сақлаш", callback_data="tr_ok"),
                                InlineKeyboardButton("❌ Бекор", callback_data="tr_no")]])
    await update.message.reply_text(text, reply_markup=kb)


async def transh_cb(update: Update, ctx):
    q = update.callback_query
    if not allowed(update):
        return await q.answer("Рухсат йўқ", show_alert=True)
    await q.answer()
    p = ctx.user_data.pop("pending_transh", None)
    if q.data == "tr_no" or not p:
        return await q.edit_message_text("Бекор қилинди.")
    async with LOCK:
        cfg = load_cfg()
        no = max((t["no"] for t in cfg["transhlar"]), default=0) + 1
        cfg["transhlar"].append({"no": no, "sana": p["sana"], "summa_usd": p["summa"], "birja_999": p["birja"],
                                 "narxlar": p["narxlar"], "manba": f"Telegram бот, {dt.date.today()}", "yopilgan_sana": None})
        cfg["transhlar"].sort(key=lambda t: t["sana"])
        save_cfg(cfg)
    await q.edit_message_text(q.message.text.replace("Сақлайми?", f"✅ Транш №{no} сақланди."))
    await run_and_reply(update, caption="Янги транш билан ҳисобот")


@guard
async def yop(update: Update, ctx):
    a = ctx.args
    if len(a) < 2 or not a[0].isdigit():
        return await update.message.reply_text("Мисол: /yop 1 2026-10-13")
    sana = nazorat.parse_date(a[1])
    if not sana:
        return await update.message.reply_text("Сана нотўғри. Мисол: /yop 1 2026-10-13 ёки 13.10.2026")
    async with LOCK:
        cfg = load_cfg()
        t = next((t for t in cfg["transhlar"] if t["no"] == int(a[0])), None)
        if not t:
            return await update.message.reply_text("Бундай транш йўқ.")
        t["yopilgan_sana"] = sana.isoformat()
        save_cfg(cfg)
    await update.message.reply_text(f"Транш №{a[0]} {sana:%d.%m.%Y} дан ёпилди.")
    await run_and_reply(update)


@guard
async def och(update: Update, ctx):
    a = ctx.args
    if not a or not a[0].isdigit():
        return await update.message.reply_text("Мисол: /och 1")
    async with LOCK:
        cfg = load_cfg()
        t = next((t for t in cfg["transhlar"] if t["no"] == int(a[0])), None)
        if not t:
            return await update.message.reply_text("Бундай транш йўқ.")
        t["yopilgan_sana"] = None
        save_cfg(cfg)
    await update.message.reply_text(f"Транш №{a[0]} қайта очилди.")


@guard
async def chegara(update: Update, ctx):
    try:
        v = float(ctx.args[0].replace(",", "."))
    except (IndexError, ValueError):
        return await update.message.reply_text("Мисол: /chegara 0.3")
    async with LOCK:
        cfg = load_cfg()
        cfg["chegara_usd_gr"] = v
        save_cfg(cfg)
    await update.message.reply_text(f"Чегара: ±{v} $/гр")


@guard
async def bekor(update: Update, ctx):
    lb = DATA / "last_backup.txt"
    last = lb.read_text().strip() if lb.exists() else None
    if not last or not Path(last).exists():
        return await update.message.reply_text("Бекор қиладиган охирги юклама топилмади.")
    async with LOCK:
        shutil.copy(last, LEDGER)
        lb.unlink(missing_ok=True)
    await update.message.reply_text("↩️ Охирги фактура юкламаси бекор қилинди.")
    await run_and_reply(update)


@guard
async def reestr(update: Update, ctx):
    with open(LEDGER, "rb") as f:
        await update.message.reply_document(f, filename="gavhar_reestr.csv")


@guard
async def document(update: Update, ctx):
    doc = update.message.document
    name = doc.file_name or "faktura.xlsx"
    if not name.lower().endswith((".xlsx", ".xlsm")):
        return await update.message.reply_text("Фақат .xlsx файл қабул қилинади (Excel да «Сохранить как → .xlsx»).")
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / name
        f = await doc.get_file()
        await f.download_to_drive(path)
        # юкламадан олдинги реестр нусхаси — /bekor учун
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        bk = BACKUP / f"ledger_{stamp}.csv"
        shutil.copy(LEDGER, bk)
        try:
            rows = await asyncio.to_thread(nazorat.read_invoice, str(path))
        except Exception as e:
            log.exception("read error")
            return await update.message.reply_text(f"❌ Файлни ўқиб бўлмади: {e}")
        if not rows:
            return await update.message.reply_text(
                "❌ Файлда Gavhar қаторлари топилмади. Керакли устунлар: «Товар номи», «Товар категорияси», "
                "«Нарх» ёки «Қиймати», «Таъминотчи фактура санаси».")
        await run_and_reply(update, files=[str(path)], caption=f"Ҳисобот — {name}")
        if filecmp.cmp(bk, LEDGER, shallow=False):
            bk.unlink()  # ҳеч нарса қўшилмади — бекор қилиш нуқтаси ўзгармайди
        else:
            (DATA / "last_backup.txt").write_text(str(bk))


async def on_error(update, ctx):
    log.exception("Хато", exc_info=ctx.error)
    if isinstance(update, Update) and update.effective_message:
        await update.effective_message.reply_text(f"❌ Хатолик: {ctx.error}")


def main():
    token = os.environ.get("BOT_TOKEN")
    if not token:
        raise SystemExit("BOT_TOKEN берилмаган")
    ensure_data()
    if not ALLOWED:
        log.warning("ALLOWED_USERS бўш — бот ҳеч кимга жавоб бермайди (фақат /myid)")
    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler(["start", "help"], start))
    app.add_handler(CommandHandler("myid", myid))
    app.add_handler(CommandHandler("holat", holat))
    app.add_handler(CommandHandler("narxlar", narxlar))
    app.add_handler(CommandHandler("transh", transh))
    app.add_handler(CommandHandler("yop", yop))
    app.add_handler(CommandHandler("och", och))
    app.add_handler(CommandHandler("chegara", chegara))
    app.add_handler(CommandHandler("bekor", bekor))
    app.add_handler(CommandHandler("reestr", reestr))
    app.add_handler(CallbackQueryHandler(transh_cb, pattern="^tr_"))
    app.add_handler(MessageHandler(filters.Document.ALL, document))
    app.add_error_handler(on_error)
    log.info("Бот ишга тушди. DATA_DIR=%s", DATA)
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
