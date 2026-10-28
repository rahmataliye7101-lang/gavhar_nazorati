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

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update, WebAppInfo
from telegram.constants import ChatAction
from telegram.ext import (Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters)

import nazorat

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
log = logging.getLogger("gavhar-bot")

BASE = Path(__file__).resolve().parent
DATA = Path(os.environ.get("DATA_DIR") or os.environ.get("RAILWAY_VOLUME_MOUNT_PATH") or "/data")  # Volume қаерга уланса ўша ер
TRANSH = DATA / "transhlar.json"
LEDGER = DATA / "ledger.csv"
VOZVRAT = DATA / "vozvrat.csv"
BACKUP = DATA / "backup"
WEBAPP_URL = os.environ.get("WEBAPP_URL") or (
    f"https://{os.environ['RAILWAY_PUBLIC_DOMAIN']}" if os.environ.get("RAILWAY_PUBLIC_DOMAIN") else "")
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
    "/vozvrat — возврат рўйхатини юклаш\n"
    "/xatolar — хатоликлар Excel да (/xatolar GVR-27102026 — битта фактура)\n"
    "/tozalash — базани тозалаш (нолдан бошлаш)\n"
    "/reestr — барча приход қаторлари (CSV)\n"
    "/admin — фойдаланувчилар (одам қўшиш/ўчириш, фақат админ)\n"
    "/myid — Telegram ID ингиз"
)


# ---------------- ёрдамчи ----------------
VOZVRAT_2 = {"no": 2, "sana": "2026-10-27", "summa_usd": 208422, "birja": 149.94,
             "narxlar": {"Бриллиант": 99.52, "Вакс": 97.41, "Кузсиз": 97.41, "Дорика": 100.05, "Фианит": 95.82,
                         "Браслет": 100.05, "Детский": 100.05, "Ван Клиф": 100.05},
             "manba": "Фойдаланувчи 27.10.2026: қўшимча возврат 208 422 $, биржа 149.94"}


def volume_yoq():
    """Railway да Volume уланмаган бўлса True — ҳар деплойда маълумот ўчиб кетади."""
    import os
    return bool(os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("RAILWAY_PROJECT_ID")) and \
        not os.environ.get("RAILWAY_VOLUME_MOUNT_PATH")


VOLUME_OGOH = ("⚠️ ДИҚҚАТ: Railway да Volume уланмаган!\n"
               "GitHub га янги файл юклаганингизда (ёки сервер қайта ишга тушганда) траншлар, фактуралар ва возвратлар ЎЧИБ КЕТАДИ.\n"
               "Тузатиш: Railway → сервис → «Attach volume» → Mount path: /data → Deploy.")


def ensure_data():
    DATA.mkdir(parents=True, exist_ok=True)
    BACKUP.mkdir(exist_ok=True)
    if not TRANSH.exists():
        src = BASE / "seed_transhlar.json"
        if src.exists():
            shutil.copy(src, TRANSH)
        else:  # бошланғич файл йўқ бўлса — бўш созлама (транш кейин mini app дан қўшилади)
            TRANSH.write_text(json.dumps({"chegara_usd_gr": 0.3, "transhlar": []}, ensure_ascii=False, indent=2), encoding="utf-8")
    cfg = json.loads(TRANSH.read_text(encoding="utf-8"))
    if "vozvrat_qismlar" not in cfg:  # бир марталик: возврат 2-қисм (27.10.2026 да берилган)
        cfg["vozvrat_qismlar"] = [dict(VOZVRAT_2)]
        TRANSH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
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
    BACKUP.mkdir(parents=True, exist_ok=True)
    if path.exists():
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy(path, BACKUP / f"{path.stem}_{stamp}{path.suffix}")
        olds = sorted(BACKUP.glob(f"{path.stem}_*{path.suffix}"))
        for o in olds[:-30]:  # охирги 30 тасини сақлаймиз
            o.unlink()


# ---------------- фойдаланувчилар ва админлар ----------------
USERS = DATA / "users.json"


def load_users():
    try:
        u = json.loads(USERS.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        u = {}
    u.setdefault("users", [])
    u.setdefault("invites", {})
    u.setdefault("pending", [])
    return u


def save_users(u):
    USERS.parent.mkdir(parents=True, exist_ok=True)
    tmp = USERS.with_suffix(".tmp")
    tmp.write_text(json.dumps(u, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(USERS)


def user_rec(uid):
    return next((x for x in load_users()["users"] if int(x["id"]) == int(uid)), None)


def allowed_id(uid) -> bool:
    return uid is not None and (int(uid) in ALLOWED or user_rec(uid) is not None)


def admin_id(uid) -> bool:
    if uid is None:
        return False
    r = user_rec(uid)
    return int(uid) in ALLOWED or bool(r and r.get("role") == "admin")


def admin_ids():
    return set(ALLOWED) | {int(x["id"]) for x in load_users()["users"] if x.get("role") == "admin"}


def allowed(update: Update) -> bool:
    u = update.effective_user
    return bool(u and allowed_id(u.id))


def is_admin(update: Update) -> bool:
    u = update.effective_user
    return bool(u and admin_id(u.id))


def add_user(uid, role="user", name="", by=None, perms=None):
    u = load_users()
    name = name or u.get("names", {}).get(str(uid), "")
    rec = next((x for x in u["users"] if int(x["id"]) == int(uid)), None)
    if rec:
        rec["role"] = role
        if name:
            rec["name"] = name
    else:
        rec = {"id": int(uid), "role": role, "name": name, "added_by": by, "sana": str(dt.date.today())}
        u["users"].append(rec)
    if perms is not None:
        rec["perms"] = [p for p in PERMS if p in perms]
    u["pending"] = [p for p in u["pending"] if int(p) != int(uid)]
    save_users(u)


def remove_user(uid):
    u = load_users()
    u["users"] = [x for x in u["users"] if int(x["id"]) != int(uid)]
    save_users(u)


def uname(user):
    if not user:
        return ""
    return (user.full_name or "") + (f" (@{user.username})" if user.username else "")


async def deny(update: Update, ctx=None):
    user = update.effective_user
    uid = user.id if user else "?"
    await update.effective_message.reply_text(
        "🟡 Gavhar — тилло нарх назорати\n\n"
        "✅ Сўровингиз админга юборилди. Админ тасдиқлагач, сизга шу ерда хабар келади — кейин /start ни босинг.")
    if not user:
        return
    u = load_users()
    if int(uid) in [int(p) for p in u["pending"]]:
        return  # аввал хабар берилган
    u["pending"].append(int(uid))
    u.setdefault("names", {})[str(uid)] = uname(user)
    save_users(u)
    bot = ctx.bot if ctx is not None else update.get_bot()
    rows_kb = [[InlineKeyboardButton("🔍 Текширувчи", callback_data=f"ad_ok:{uid}:user:tekshiruvchi"),
                InlineKeyboardButton("👁 Кўрувчи", callback_data=f"ad_ok:{uid}:user:koruvchi")],
               [InlineKeyboardButton("👑 Админ", callback_data=f"ad_ok:{uid}:admin"),
                InlineKeyboardButton("❌ Рад этиш", callback_data=f"ad_no:{uid}")]]
    if WEBAPP_URL:
        from telegram import WebAppInfo
        rows_kb.append([InlineKeyboardButton("📋 Admin бўлимида кўриш", web_app=WebAppInfo(url=WEBAPP_URL + "#admin"))])
    kb = InlineKeyboardMarkup(rows_kb)
    for aid in admin_ids():
        try:
            await bot.send_message(aid, f"🔔 Янги одам ботга кирмоқчи:\n{uname(user)}\nID: {uid}\n\nРухсат берасизми?",
                                   reply_markup=kb)
        except Exception:
            log.warning("admin %s га хабар юборилмади", aid)


def guard(fn):
    async def wrap(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        if not allowed(update):
            return await deny(update, ctx)
        return await fn(update, ctx)
    return wrap


def admin_only(fn):
    async def wrap(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
        if not allowed(update):
            return await deny(update, ctx)
        if not is_admin(update):
            return await update.effective_message.reply_text("⛔ Бу бўлим фақат админлар учун.")
        return await fn(update, ctx)
    return wrap


# ---------------- рухсатлар (ким нимани кўради) ----------------
PERMS = {
    "holat": "📊 Ҳолат (қолдиқ, умумий)",
    "faktura": "🧾 Фактура бўлими",
    "xatolar": "⚠️ Хатоликлар",
    "transhlar": "💳 Траншлар ва нархлар (кўриш)",
    "faktura_yuklash": "📥 Фактура юклаш",
    "vozvrat_yuklash": "↩️ Возврат рўйхатини юклаш",
    "transh_qoshish": "➕ Транш қўшиш / ёпиш, чегара",
    "hisobot": "📑 Excel ҳисобот олиш",
}
DEFAULT_PERMS = ["holat", "faktura", "xatolar", "transhlar", "faktura_yuklash", "hisobot"]
# Тайёр рухсат тўпламлари
PRESETS = {
    "tekshiruvchi": ("🔍 Текширувчи — юклайди ва текширади",
                     ["holat", "faktura", "xatolar", "transhlar", "faktura_yuklash", "vozvrat_yuklash", "hisobot"]),
    "koruvchi": ("👁 Кўрувчи — фақат кўради", ["holat", "faktura", "xatolar", "transhlar"]),
    "boshqaruvchi": ("💳 Бошқарувчи — пул ҳам киритади", list(PERMS)),
}


def preset_of(perms):
    ps = set(perms or [])
    return next((k for k, (_, v) in PRESETS.items() if set(v) == ps), "")


def inv_ok(v, now=None):
    """Таклиф ҳаволаси амалдами (муддатсиз — exp йўқ)."""
    e = v.get("exp")
    return not e or e > (now or dt.datetime.now().isoformat())


def user_perms(uid):
    if admin_id(uid):
        return list(PERMS)
    r = user_rec(uid)
    if not r:
        return []
    return list(r.get("perms", DEFAULT_PERMS))


def has_perm(uid, p):
    return p in user_perms(uid)


def set_perms(uid, perms):
    u = load_users()
    for x in u["users"]:
        if int(x["id"]) == int(uid):
            x["perms"] = [p for p in PERMS if p in perms]
    save_users(u)


def need(perm):
    def deco(fn):
        async def wrap(update: Update, ctx):
            if not allowed(update):
                return await deny(update, ctx)
            if not has_perm(update.effective_user.id, perm):
                return await update.effective_message.reply_text(f"⛔ Сизда бу учун рухсат йўқ: {PERMS[perm]}. Админга мурожаат қилинг.")
            return await fn(update, ctx)
        return wrap
    return deco


def perm_kb(uid):
    cur = set(user_perms(uid))
    pk = preset_of(cur)
    rows = [[InlineKeyboardButton(("● " if pk == k else "") + t.split(" — ")[0], callback_data=f"ad_ps:{uid}:{k}")]
            for k, (t, _) in PRESETS.items()]
    rows += [[InlineKeyboardButton(("✅ " if p in cur else "⬜ ") + t, callback_data=f"ad_pt:{uid}:{p}")] for p, t in PERMS.items()]
    rows.append([InlineKeyboardButton("⬅️ Рўйхатга қайтиш", callback_data="ad_back")])
    return InlineKeyboardMarkup(rows)


def perm_text(uid):
    r = user_rec(uid) or {}
    return (f"⚙️ Рухсатлар: {r.get('name') or uid} (ID {uid})\n"
            "Тепадаги тайёр тўпламдан бирини танланг ёки пастда биттама-битта ёқинг/ўчиринг — дарҳол сақланади.\n"
            + "\n".join(f"{t}" for t, _ in PRESETS.values()))


ROLE_L = {"admin": "👑 админ", "user": "👤 фойдаланувчи"}


def admin_panel():
    u = load_users()
    lines = ["👥 Фойдаланувчилар", ""]
    for aid in sorted(ALLOWED):
        lines.append(f"👑 {aid} — асосий админ (Railway)")
    for x in u["users"]:
        if int(x["id"]) in ALLOWED:
            continue
        pk = preset_of(x.get("perms", DEFAULT_PERMS)) if x.get("role") != "admin" else ""
        lines.append(f"{ROLE_L.get(x.get('role'), x.get('role'))} {x.get('name') or '—'} · {x['id']}"
                     + (f" · {PRESETS[pk][0].split(' — ')[0]}" if pk else (" · махсус рухсат" if x.get("role") != "admin" else "")))
    if len(lines) == 2:
        lines.append("(ҳали ҳеч ким қўшилмаган)")
    rows = []
    for x in u["users"]:
        if int(x["id"]) in ALLOWED:
            continue
        nm = (x.get("name") or str(x["id"]))[:22]
        other = "user" if x.get("role") == "admin" else "admin"
        rows.append([InlineKeyboardButton(f"⚙️ {nm}", callback_data=f"ad_perm:{x['id']}"),
                     InlineKeyboardButton("🔁 " + ("фойд." if other == "user" else "админ"), callback_data=f"ad_role:{x['id']}:{other}"),
                     InlineKeyboardButton("🗑", callback_data=f"ad_del:{x['id']}")])
    lines += ["", "Ботга янги одам кирса, сизга рухсат сўрови келади — Текширувчи / Кўрувчи / Админ ни танлаб тасдиқлайсиз."]
    return "\n".join(lines), InlineKeyboardMarkup(rows)


@admin_only
async def admin(update: Update, ctx):
    text, kb = admin_panel()
    await update.effective_message.reply_text(text, reply_markup=kb)


async def admin_cb(update: Update, ctx):
    q = update.callback_query
    if not is_admin(update):
        return await q.answer("Фақат админлар учун", show_alert=True)
    await q.answer()
    parts = q.data.split(":")
    act = parts[0]
    if act == "ad_link":
        me = await ctx.bot.get_me()
        return await q.message.reply_text(
            f"🔗 Ботга кириш ҳаволаси (ўчиб кетмайди):\nhttps://t.me/{me.username}?start=sorov\n\n"
            "Шуни одамга юборинг. У «Start» ни босса, сизга рухсат сўрови келади — тасдиқлаб, нима кўришини танлайсиз.")
    if act == "ad_inv":
        role = parts[1]
        preset = parts[2] if len(parts) > 2 and parts[2] in PRESETS else ("" if role == "admin" else "tekshiruvchi")
        import secrets
        code = secrets.token_urlsafe(8)
        u = load_users()
        u["invites"] = {k: v for k, v in u["invites"].items() if inv_ok(v)}
        u["invites"][code] = {"role": role, "preset": preset, "exp": None, "by": update.effective_user.id}
        save_users(u)
        me = await ctx.bot.get_me()
        link = f"https://t.me/{me.username}?start=inv_{code}"
        nom = ROLE_L[role] if role == "admin" else PRESETS[preset][0]
        return await q.message.reply_text(
            f"🔗 Таклиф ҳаволаси ({nom}).\nМуддатсиз — ўчиб кетмайди, бир марта ишлайди:\n{link}\n\n"
            "Шу ҳаволани қўшмоқчи бўлган одамга юборинг — у босиб «Start» ни босса, ўзи қўшилади.")
    if act == "ad_id":
        ctx.user_data["await_admin_id"] = True
        return await q.message.reply_text("Қўшиладиган одамнинг Telegram ID сини ёзинг (масалан: 512334871).\n"
                                          "ID ни у ботга /myid ёзиб билади.")
    if act == "ad_ok":
        uid, role = int(parts[1]), parts[2]
        preset = parts[3] if len(parts) > 3 and parts[3] in PRESETS else ""
        add_user(uid, role, by=update.effective_user.id, perms=PRESETS[preset][1] if preset else None)
        await q.edit_message_text(q.message.text + f"\n\n✅ Қўшилди ({PRESETS[preset][0] if preset else ROLE_L[role]}).",
                                  reply_markup=None if role == "admin" else InlineKeyboardMarkup(
                                      [[InlineKeyboardButton("⚙️ Нималарни кўришини белгилаш", callback_data=f"ad_perm:{uid}")]]))
        try:
            await ctx.bot.send_message(uid, f"✅ Сизга рухсат берилди ({ROLE_L[role]}). /start ни босинг.")
        except Exception:
            pass
        return
    if act == "ad_no":
        uid = int(parts[1])
        u = load_users()
        u["pending"] = [p for p in u["pending"] if int(p) != uid]
        save_users(u)
        return await q.edit_message_text(q.message.text + "\n\n❌ Рад этилди.")
    if act == "ad_perm":
        uid = int(parts[1])
        if admin_id(uid):
            return await q.message.reply_text("👑 Админда ҳамма рухсатлар бор. Чеклаш учун аввал 🔁 билан фойдаланувчи қилинг.")
        return await q.message.reply_text(perm_text(uid), reply_markup=perm_kb(uid))
    if act == "ad_ps":
        uid, k = int(parts[1]), parts[2]
        if k in PRESETS:
            set_perms(uid, PRESETS[k][1])
        return await q.edit_message_reply_markup(reply_markup=perm_kb(uid))
    if act == "ad_pt":
        uid, p = int(parts[1]), parts[2]
        cur = set(user_perms(uid))
        cur.symmetric_difference_update({p})
        set_perms(uid, cur)
        return await q.edit_message_reply_markup(reply_markup=perm_kb(uid))
    if act == "ad_back":
        text, kb = admin_panel()
        return await q.edit_message_text(text, reply_markup=kb)
    if act == "ad_del":
        remove_user(int(parts[1]))
    if act == "ad_role":
        rec = user_rec(int(parts[1]))
        if rec:
            add_user(int(parts[1]), parts[2], rec.get("name", ""))
    text, kb = admin_panel()
    await q.edit_message_text(text, reply_markup=kb)


async def accept_invite(update: Update, ctx, code) -> bool:
    u = load_users()
    inv = u["invites"].get(code)
    if not inv or not inv_ok(inv):
        await update.effective_message.reply_text("⛔ Таклиф ҳаволаси эскирган ёки ишлатилган. Админдан янгисини сўранг.")
        return False
    del u["invites"][code]
    save_users(u)
    user = update.effective_user
    pr = inv.get("preset") or ""
    add_user(user.id, inv["role"], uname(user), by=inv.get("by"), perms=PRESETS[pr][1] if pr in PRESETS else None)
    for aid in admin_ids():
        if aid != user.id:
            try:
                await ctx.bot.send_message(
                    aid, f"✅ {uname(user)} (ID {user.id}) ҳавола орқали қўшилди — {ROLE_L[inv['role']]}.\n"
                         + ((f"Рухсат: {PRESETS[pr][0]}. Ўзгартириш:" if pr in PRESETS else "Нималарни кўришини белгиланг:")
                            if inv["role"] != "admin" else ""),
                    reply_markup=None if inv["role"] == "admin" else InlineKeyboardMarkup(
                        [[InlineKeyboardButton("⚙️ Нималарни кўришини белгилаш", callback_data=f"ad_perm:{user.id}")]]))
            except Exception:
                pass
    return True


async def send_long(msg, text):
    for i in range(0, len(text), 3900):
        await msg.reply_text(text[i:i + 3900])


async def run_and_reply(update: Update, files=(), caption="Ҳисобот", umumiy=False):
    msg = update.effective_message
    await msg.chat.send_action(ChatAction.TYPING)
    async with LOCK:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / f"Gavhar_nazorat_{dt.date.today():%Y-%m-%d}.xlsx"
            text = await asyncio.to_thread(nazorat.run, str(TRANSH), str(LEDGER), str(out), list(files), str(VOZVRAT))
            await send_long(msg, text)
            if update.effective_user and not has_perm(update.effective_user.id, "hisobot"):
                return
            for fp in sorted(Path(td).glob("Faktura_*.xlsx")):  # ҳар бир янги фактура (дўкон) бўйича алоҳида текширув
                with open(fp, "rb") as f:
                    await msg.reply_document(f, filename=fp.name, caption="📑 Батафсил — " + fp.stem[8:].replace("_", " "))
            jp = Path(td) / "Fakturalar_jami.xlsx"
            if jp.exists():
                with open(jp, "rb") as f:
                    await msg.reply_document(f, filename=f"Fakturalar_jami_{dt.date.today():%Y-%m-%d}.xlsx",
                                             caption="📑 Ҳамма фактура ва SKU — битта файлда")
            if umumiy:
                with open(out, "rb") as f:
                    await msg.reply_document(f, filename=out.name, caption="📑 Умумий ҳисобот")


# ---------------- транш матнини ўқиш ----------------
NUM = r"(\d+(?:[.,]\d+)?)"


def parse_transh(text: str):
    # бир қаторга ёзилган бўлса: «… 137.92$ Лом нархи 80.76$ Бриллиант …» → ҳар нарх алоҳида қатор
    text = re.sub(r"(\d)\s*\$?[ \t]+(?=[^\d\s$])", r"\1\n", text or "")
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
            l = (l[:m.start()] + " " + l[m.end():]).strip()
            low = l.lower().replace("ё", "е")
            if not l:
                continue
        nums = [float(x.replace(",", ".")) for x in re.findall(NUM, l.replace(" ", ""))]
        if not nums:
            continue
        if "биржа" in low or "birja" in low or "тилло нарх" in low or "tillo narx" in low or "999" in low and "нарх" in low:
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
        if p["lom"] and abs(lom - p["lom"]) > 0.05:
            notes.append(f"⚠ Лом нотўғри: {p['lom']:.2f} ёзилган, ҳисоб бўйича {lom:.2f} бўлиши керак")
        notes.append(f"Лом: {lom:.2f} $\n\n1 гр нархлари:")
        prev = cfg["transhlar"][-1] if cfg["transhlar"] else None
        ogoh = []
        for n, v in p["narxlar"].items():
            notes.append(f"• {n}: {v:.2f} $ (лом + {v - lom:.2f})")
            if prev and n in prev["narxlar"]:
                d = (v - lom) - (prev["narxlar"][n] - prev["birja_999"] * 585 / 999)
                if abs(d) > 0.02:
                    ogoh.append(f"⚠ {n}: ишчилик+НДС олдинги траншдан {d:+.2f} $ фарқ қилади")
        if prev:
            for n in prev["narxlar"]:
                if n not in p["narxlar"]:
                    ogoh.append(f"⚠ «{n}» нархи йўқ — бу товарларга Дорика нархи олинади")
        notes += ([""] + ogoh) if ogoh else ["", "✅ Ишчилик+НДС олдинги транш билан бир хил"]
    return errs, notes


# ---------------- ҳандлерлар ----------------
async def start(update: Update, ctx):
    if not allowed(update):
        return await deny(update, ctx)
    await update.message.reply_text(HELP, parse_mode="Markdown")


async def myid(update: Update, ctx):
    await update.message.reply_text(f"Сизнинг Telegram ID: {update.effective_user.id}")


@need("holat")
async def holat(update: Update, ctx):
    await run_and_reply(update, umumiy=True)


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


async def confirm_transh(update: Update, ctx, p):
    """Транш ёки возврат қисмини ДАРҲОЛ сақлайди (тасдиқ сўрамайди). Хато бўлса «↩️ Бекор қилиш» тугмаси."""
    cfg = load_cfg()
    errs, notes = check_transh(p, cfg)
    if errs:
        return await update.effective_message.reply_text("❌ Тушунмадим: " + "; ".join(errs) + ".\nGavhar нарх хабарини тўлиқ forward қилинг.")
    sana = nazorat._dmy(p["sana"]) if p.get("sana") else "—"
    async with LOCK:
        cfg = load_cfg()
        if p.get("vozvrat"):
            vq = cfg.setdefault("vozvrat_qismlar", [])
            dup = next((q for q in vq if abs(q["birja"] - p["birja"]) < 0.005), None)
            if dup:
                dup["summa_usd"] += p["summa"]
                undo, head = f"tr_undo:vq:{dup['no']}:{p['summa']}", (
                    f"↩️ Возврат {dup['no']}-қисмга {p['summa']:,.0f} $ қўшилди → жами {dup['summa_usd']:,.0f} $")
            else:
                no = max((q["no"] for q in vq), default=1) + 1
                vq.append({"no": no, "sana": p["sana"], "summa_usd": p["summa"], "birja": p["birja"],
                           "narxlar": p["narxlar"], "manba": f"Telegram бот, {dt.date.today()}"})
                undo, head = f"tr_undo:vqnew:{no}:0", f"↩️ Возврат {no}-қисм қўшилди: {p['summa']:,.0f} $ · биржа {p['birja']}"
            body = "\n".join(f"• {n}: {v:.2f} $" for n, v in p["narxlar"].items())
        else:
            dup = next((t for t in cfg["transhlar"] if t["sana"] == p["sana"] and abs(t["birja_999"] - p["birja"]) < 0.005), None)
            if dup:
                dup["summa_usd"] += p["summa"]
                undo, head = f"tr_undo:tr:{dup['no']}:{p['summa']}", (
                    f"💳 Транш №{dup['no']} ({sana}) га {p['summa']:,.0f} $ қўшилди → жами {dup['summa_usd']:,.0f} $")
                body = "Нархлар ўзгармади (ўша транш нархлари)."
            else:
                no = max((t["no"] for t in cfg["transhlar"]), default=0) + 1
                eng_yangi = max((t["sana"] for t in cfg["transhlar"] if not t.get("faqat_tekshiruv")), default="")
                eski = bool(eng_yangi and p["sana"] and p["sana"] < eng_yangi)
                cfg["transhlar"].append({"no": no, "sana": p["sana"], "summa_usd": p["summa"], "birja_999": p["birja"],
                                         "narxlar": p["narxlar"], "manba": f"Telegram бот, {dt.date.today()}", "yopilgan_sana": None,
                                         "faqat_tekshiruv": False})
                cfg["transhlar"].sort(key=lambda t: t["no"])
                undo, head = f"tr_undo:trnew:{no}:0", f"💳 Янги транш №{no} қўшилди\n📅 {sana} · 💵 {p['summa']:,.0f} $ · биржа {p['birja']}"
                body = "\n".join(notes)
                if eski:
                    body += (f"\n\n📅 Бу пул санаси охирги пулдан ({nazorat._dmy(eng_yangi)}) эски — навбатда санаси бўйича ОЛДИНГА "
                             "турди (эски пул биринчи ишлатилади). Агар у фақат ўтган ойлар приходини текшириш учун бўлса — пастдаги тугмани босинг.")
                    kb2 = InlineKeyboardMarkup([[InlineKeyboardButton("📅 Фақат эски приход текшируви учун", callback_data=f"tr_navbat:{no}:0")],
                                                [InlineKeyboardButton("↩️ Бекор қилиш", callback_data=undo)]])
        save_cfg(cfg)
    kb = locals().get("kb2") or InlineKeyboardMarkup([[InlineKeyboardButton("↩️ Бекор қилиш", callback_data=undo)]])
    await update.effective_message.reply_text(f"✅ {head}\n{body}", reply_markup=kb)
    await run_and_reply(update)


DATE_RE = re.compile(r"\b\d{1,2}[./-]\d{1,2}[./-]\d{4}\b|\b\d{4}-\d{2}-\d{2}\b")


def split_transh_text(text):
    """Битта хабарда бир нечта транш бўлса (ҳар бири ўз санаси билан) — бўлакларга ажратади."""
    pos = [m.start() for m in DATE_RE.finditer(text or "")]
    if len(pos) < 2:
        return [text]
    parts = [text[a:b] for a, b in zip(pos, pos[1:] + [len(text)])]
    parts[0] = text[:pos[0]] + parts[0]  # биринчи санагача ёзилган сўзлар (мас. «Возврат») биринчи бўлакка
    ok = [p for p in parts if (lambda q: q["birja"] and len(q["narxlar"]) >= 2)(parse_transh(p))]
    return ok if len(ok) >= 2 else [text]


async def transh_flow(update: Update, ctx, text):
    """Матндан транш(лар) ўқиш. Бир хабарда бир нечта транш бўлса — ҳар бирини санаси бўйича алоҳида қўшади."""
    blocks = split_transh_text(text or "")
    if len(blocks) == 1:
        return await transh_flow_one(update, ctx, text)
    ps = [parse_transh(b) for b in blocks]
    if not has_perm(update.effective_user.id, "transh_qoshish"):
        await update.effective_message.reply_text(f"⛔ Сизда бу учун рухсат йўқ: {PERMS['transh_qoshish']}.")
        return True
    miss = [nazorat._dmy(p["sana"]) for p in ps if not p["summa"]]
    if miss:
        await update.effective_message.reply_text(
            f"ℹ️ Хабарда {len(ps)} та транш бор, лекин {', '.join(miss)} учун сумма йўқ.\n"
            "Ҳар бир санадан кейин «Сумма 707739» деб ёзинг ёки траншларни алоҳида-алоҳида юборинг.")
        return True
    await update.effective_message.reply_text(
        f"📥 Хабарда {len(ps)} та транш топилди: " + ", ".join(f"{nazorat._dmy(p['sana'])} — {p['summa']:,.0f} $" for p in ps)
        + ".\nЁзилган тартибда қўшаман: тепадаги — 1-навбат (биринчи тугайди).")
    for b in blocks:
        await transh_flow_one(update, ctx, b)
    return True


async def transh_flow_one(update: Update, ctx, text):
    """Матндан битта транш (пул + нархлар) ўқиш. Транш матни бўлмаса False қайтаради."""
    p = parse_transh(text or "")
    if not p["birja"] or len(p["narxlar"]) < 2:
        return False
    if not has_perm(update.effective_user.id, "transh_qoshish"):
        await update.effective_message.reply_text(f"⛔ Сизда бу учун рухсат йўқ: {PERMS['transh_qoshish']}.")
        return True
    p["vozvrat"] = bool(re.search(r"возв|возрат|vozvrat", text or "", re.I))
    if p["vozvrat"] and not p["sana"]:
        p["sana"] = dt.date.today().isoformat()
    cfg = load_cfg()
    dup = next((t for t in cfg["transhlar"] if t["sana"] == p["sana"] and abs(t["birja_999"] - p["birja"]) < 0.005), None)
    if dup and not p["summa"] and not p["vozvrat"]:
        ctx.user_data["await_summa"] = p
        await update.effective_message.reply_text(
            f"ℹ️ Бу нархлар Транш №{dup['no']} да бор ({nazorat._dmy(dup['sana'])}, {dup['summa_usd']:,.0f} $).\n"
            "Шу нархда яна пул берган бўлсангиз — суммасини ёзинг (масалан: 150000), шу траншга қўшиб қўяман.")
        return True
    if not p["summa"]:
        ctx.user_data["await_summa"] = p
        await update.effective_message.reply_text(
            f"✅ Нархлар ўқилди ({nazorat._dmy(p['sana']) if p.get('sana') else '—'}, биржа {p['birja']}).\n"
            + ("Возврат суммаси қанча? Ёзинг (масалан: 208422)" if p.get("vozvrat") else "Қанча пул берилди? Суммани ёзинг (масалан: 707000)"))
        return True
    await confirm_transh(update, ctx, p)
    return True


@need("transh_qoshish")
async def transh(update: Update, ctx):
    if not await transh_flow(update, ctx, update.message.text):
        await update.message.reply_text("❌ Транш матни ўқилмади. Намуна учун /start")


@guard
async def on_text(update: Update, ctx):
    text = update.message.text or ""
    if ctx.user_data.get("await_admin_id") and is_admin(update):
        m = re.fullmatch(r"\s*(\d{5,15})\s*", text)
        if m:
            ctx.user_data.pop("await_admin_id", None)
            add_user(int(m.group(1)), "user", by=update.effective_user.id)
            await update.message.reply_text(f"✅ {m.group(1)} фойдаланувчи сифатида қўшилди. Админ қилиш учун /admin.")
            try:
                await ctx.bot.send_message(int(m.group(1)), "✅ Сизга Gavhar назорат ботидан фойдаланишга рухсат берилди. /start ни босинг.")
            except Exception:
                pass
            return
    aw = ctx.user_data.get("await_summa")
    if aw:
        m = re.fullmatch(r"\s*([\d\s.,]+)\s*(\$|доллар|доллор|usd)?\s*", text, re.I)
        if m:
            num = float(re.sub(r"[^\d.]", "", m.group(1).replace(",", "")) or 0)
            if num >= 1000:
                ctx.user_data.pop("await_summa", None)
                aw["summa"] = num
                return await confirm_transh(update, ctx, aw)
    if await transh_flow(update, ctx, text):
        return
    await update.message.reply_text(
        "📎 Фактура (Excel) ташланг — текшириб бераман.\n"
        "💳 Янги пул берилса — Gavhar нарх хабарини forward қилинг.\n"
        "📊 Ҳолат: /holat")


async def transh_cb(update: Update, ctx):
    q = update.callback_query
    if not allowed(update):
        return await q.answer("Рухсат йўқ", show_alert=True)
    await q.answer()
    if not has_perm(update.effective_user.id, "transh_qoshish"):
        return await q.edit_message_text("⛔ Сизда транш қўшиш рухсати йўқ.")
    if q.data.startswith("tr_navbat:"):
        _, no, v = q.data.split(":")
        async with LOCK:
            cfg = load_cfg()
            t = next((t for t in cfg["transhlar"] if t["no"] == int(no)), None)
            if t:
                t["faqat_tekshiruv"] = v != "1"
                save_cfg(cfg)
        await q.edit_message_text(q.message.text + ("\n\n💳 Кундалик навбатга қўшилди." if v == "1" else "\n\n📅 Фақат текширув учун."))
        return await run_and_reply(update)
    if q.data.startswith("tr_undo:"):
        _, kind, no, amt = q.data.split(":")
        no, amt = int(no), float(amt)
        async with LOCK:
            cfg = load_cfg()
            if kind in ("tr", "trnew"):
                lst = cfg["transhlar"]
                t = next((t for t in lst if t["no"] == no), None)
            else:
                lst = cfg.setdefault("vozvrat_qismlar", [])
                t = next((x for x in lst if x["no"] == no), None)
            if not t:
                return await q.edit_message_text(q.message.text + "\n\n(аллақачон бекор қилинган)")
            if kind.endswith("new"):
                lst.remove(t)
            else:
                t["summa_usd"] -= amt
            save_cfg(cfg)
        await q.edit_message_text(q.message.text + "\n\n↩️ Бекор қилинди.")
        return await run_and_reply(update)
    p = ctx.user_data.pop("pending_transh", None)
    if q.data == "tr_no" or not p:
        return await q.edit_message_text("Бекор қилинди.")
    if p.get("vozvrat"):
        async with LOCK:
            cfg = load_cfg()
            vq = cfg.setdefault("vozvrat_qismlar", [])
            no = max((q["no"] for q in vq), default=1) + 1
            vq.append({"no": no, "sana": dt.date.today().isoformat(), "summa_usd": p["summa"], "birja": p["birja"],
                       "narxlar": p["narxlar"], "manba": f"Telegram бот, {dt.date.today()}"})
            save_cfg(cfg)
        await q.edit_message_text(q.message.text.replace("Сақлайми?", f"✅ Возврат {no}-қисм сақланди."))
        return await run_and_reply(update)
    async with LOCK:
        cfg = load_cfg()
        no = max((t["no"] for t in cfg["transhlar"]), default=0) + 1
        cfg["transhlar"].append({"no": no, "sana": p["sana"], "summa_usd": p["summa"], "birja_999": p["birja"],
                                 "narxlar": p["narxlar"], "manba": f"Telegram бот, {dt.date.today()}", "yopilgan_sana": None})
        cfg["transhlar"].sort(key=lambda t: t["no"])
        save_cfg(cfg)
    await q.edit_message_text(q.message.text.replace("Сақлайми?", f"✅ Транш №{no} сақланди."))
    await run_and_reply(update)


@need("transh_qoshish")
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
        t["yopilgan_qator"] = len(nazorat.load_ledger(str(LEDGER)))
        save_cfg(cfg)
    await update.message.reply_text(f"🔒 Транш №{a[0]} ёпилди. Энди янги фактуралар фақат жорий транш нархида текширилади.")
    await run_and_reply(update)


def vozvrat_qism_yop(no, yop=True):
    """Возврат қисмини якунлаш / қайта очиш. Хато матни ёки None қайтаради."""
    cfg = load_cfg()
    if no == 1:
        q = cfg.setdefault("vozvrat_1", {})
    else:
        q = next((x for x in cfg.get("vozvrat_qismlar", []) if x["no"] == no), None)
    if q is None:
        return "Бундай возврат қисми йўқ."
    q["yopilgan_sana"] = dt.date.today().isoformat() if yop else None
    q["yopilgan_qator"] = len(nazorat.load_ledger(str(LEDGER))) if yop else None
    save_cfg(cfg)
    return None


@need("transh_qoshish")
async def vozvrat_yop(update: Update, ctx):
    a = ctx.args
    if not a or not a[0].isdigit():
        cfg = load_cfg()
        nos = ", ".join(str(q["no"]) for q in cfg.get("vozvrat_qismlar", [])) or "йўқ"
        return await update.message.reply_text(f"Мисол: /vozvrat_yop 2\n1 — возврат рўйхати (617 минг); суммали қисмлар: {nos}")
    async with LOCK:
        err = vozvrat_qism_yop(int(a[0]), True)
    if err:
        return await update.message.reply_text(err)
    await update.message.reply_text(
        f"🔒 Возврат {a[0]}-қисм якунланди. Энди янги фактуралар "
        + ("фақат пуллар ва бошқа очиқ возвратлар" if a[0] == "1" else "пуллар ва возврат рўйхати (1-қисм)") + " бўйича текширилади.\n"
        f"Қайта очиш: /vozvrat_och {a[0]}")
    await run_and_reply(update)


@need("transh_qoshish")
async def vozvrat_och(update: Update, ctx):
    a = ctx.args
    if not a or not a[0].isdigit():
        return await update.message.reply_text("Мисол: /vozvrat_och 2")
    async with LOCK:
        err = vozvrat_qism_yop(int(a[0]), False)
    await update.message.reply_text(err or f"🔓 Возврат {a[0]}-қисм қайта очилди.")


async def _pul_faol(update, ctx, faol):
    a = ctx.args
    if not a or not a[0].isdigit():
        cfg = load_cfg()
        lst = "\n".join(f"   №{t['no']} ({nazorat._dmy(t['sana'])}) — {'✅ текширувда' if t.get('faol', True) else '⏸ ўчирилган'}"
                        for t in cfg["transhlar"]) or "   йўқ"
        return await update.message.reply_text(f"Мисол: /pul_ochir 1 ёки /pul_yoq 1\nПуллар:\n{lst}")
    async with LOCK:
        cfg = load_cfg()
        t = next((t for t in cfg["transhlar"] if t["no"] == int(a[0])), None)
        if not t:
            return await update.message.reply_text("Бундай транш йўқ.")
        t["faol"] = faol
        save_cfg(cfg)
    await update.message.reply_text(
        f"{'✅' if faol else '⏸'} Транш №{a[0]} ({nazorat._dmy(t['sana'])}) текширувдан {'ёқилди' if faol else 'ўчирилди'}.\n"
        + ("Энди у навбатда ва нарх текширишда қатнашади." if faol else "Энди у на навбатда, на нарх текширишда ҳисобга олинмайди."))
    await run_and_reply(update)


@need("transh_qoshish")
async def pul_ochir(update: Update, ctx):
    await _pul_faol(update, ctx, False)


@need("transh_qoshish")
async def pul_yoq(update: Update, ctx):
    await _pul_faol(update, ctx, True)


@need("holat")
async def oy_yakuni(update: Update, ctx):
    """/oy_yakuni [2026-10 ёки 10.2026] — ой бўйича пуллар: ой бошида, ишлатилди, кейинги ойга ўтди."""
    cfg = load_cfg()
    rows, _ = nazorat.compute(cfg, nazorat.load_ledger(str(LEDGER)), nazorat.load_returns(str(VOZVRAT)))
    a = " ".join(ctx.args or []).strip()
    m = re.match(r"(\d{4})-(\d{1,2})", a) or re.match(r"(\d{1,2})[./](\d{4})", a)
    if m:
        y, mo = (m.group(1), m.group(2)) if len(m.group(1)) == 4 else (m.group(2), m.group(1))
        oy = f"{int(y):04d}-{int(mo):02d}"
    else:
        oy = max(((r.get("sana") or "")[:7] for r in rows if r.get("sana")), default=dt.date.today().strftime("%Y-%m"))
    await update.message.reply_text(nazorat.oy_pullar_matn(rows, cfg, oy) + "\n\nБошқа ой: /oy_yakuni 2026-11")


@need("faktura_yuklash")
async def oylik(update: Update, ctx):
    ctx.user_data["oylik_mode"] = True
    await update.message.reply_text(
        "🔎 Текшириладиган файлни (ойлик / 15 кунлик приход ёки эски фактура) ташланг.\n"
        "Бот ҳар товарни санасига ва нархига қараб ўзи текширади: ёқилган пуллар → возвратлар → бошқа пуллар. "
        "Реестрга қўшилмайди, пул айирилмайди.")


async def handle_oylik_file(update: Update, path):
    msg = update.effective_message
    await msg.chat.send_action(ChatAction.TYPING)
    async with LOCK:
        cfg = load_cfg()
        ledger = nazorat.load_ledger(str(LEDGER))
        rets = nazorat.load_returns(str(VOZVRAT))
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / f"Oylik_tekshiruv_{dt.date.today():%Y-%m-%d}.xlsx"
        try:
            mr, text = await asyncio.to_thread(nazorat.oylik_tekshir, cfg, ledger, rets, [str(path)], str(out))
        except Exception:
            log.exception("oylik")
            return await msg.reply_text("❌ Файлни ўқиб бўлмади. Excel (.xlsx) приход эканини текширинг.")
        await send_long(msg, text)
        if mr and out.exists() and has_perm(update.effective_user.id, "hisobot"):
            with open(out, "rb") as f:
                await msg.reply_document(f, filename=out.name, caption="📑 Ойлик текширув")
            bad = [r for r in mr if str(r.get("holat", "")).startswith("⚠")]
            if bad:
                ep = Path(td) / f"Oylik_xatoliklar_{dt.date.today():%Y-%m-%d}.xlsx"
                await asyncio.to_thread(nazorat.write_errors_report, str(ep), mr, "Ойлик приход — хатоликлар", rets)
                with open(ep, "rb") as f:
                    await msg.reply_document(f, filename=ep.name, caption="❌ Ойлик хатоликлар (Gavhar га эътироз учун)")


@need("transh_qoshish")
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
        t["yopilgan_qator"] = None
        save_cfg(cfg)
    await update.message.reply_text(f"🔓 Транш №{a[0]} қайта очилди.")


@need("transh_qoshish")
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
    await update.message.reply_text("↩️ Охирги юкланган фактура бекор қилинди.")
    await run_and_reply(update)


@need("hisobot")
async def reestr(update: Update, ctx):
    with open(LEDGER, "rb") as f:
        await update.message.reply_document(f, filename="gavhar_reestr.csv")


RET_WORDS = ("возврат", "возрат", "vozvrat", "qaytar", "қайтар", "кайтар")


def looks_like_returns(path, name):
    if any(w in name.lower() for w in RET_WORDS):
        return True
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        for ws in wb.worksheets:
            if any(w in ws.title.lower() for w in RET_WORDS):
                return True
            for r in ws.iter_rows(max_row=6, values_only=True):
                if any(w in str(c or "").lower() for c in r for w in RET_WORDS):
                    return True
    except Exception:
        pass
    return False


@need("vozvrat_yuklash")
async def vozvrat(update: Update, ctx):
    ctx.user_data["vozvrat_mode"] = True
    st = nazorat.returns_stats(nazorat.load_returns(str(VOZVRAT)))
    await update.message.reply_text(
        f"↩️ Возврат рўйхатида ҳозир: {st['n']} та товар.\n"
        "Янги возврат Excel файлини ташланг.\n\n"
        "Рўйхатсиз (фақат сумма ва нархлар) возврат бўлса — «Возврат» сўзи билан нархларни ёзинг, масалан:\n"
        "Возврат\nБиржа нархи 149.94\nБриллиант 99.52\nВакс 97.41\n...\nСумма 208422")


async def handle_returns_file(update: Update, path, name):
    try:
        rows = await asyncio.to_thread(nazorat.read_returns, str(path))
    except Exception as e:
        log.exception("returns read error")
        return await update.message.reply_text(f"❌ Возврат файлини ўқиб бўлмади: {e}")
    if not rows:
        return await update.message.reply_text(
            "❌ Возврат рўйхати ўқилмади. Устунлар: «Товар номи» ёки «Артикул», «Грамм», «Сумма» (ёки «1 гр нархи»).")
    old = [r for r in nazorat.load_returns(str(VOZVRAT)) if r.get("chiqdi_qator") is None and not r.get("chiqdi_sana")]
    nk = {nazorat.ret_key(r) for r in rows}
    gone = sum(1 for r in old if nazorat.ret_key(r) not in nk)
    if old and gone > 0.3 * len(old):  # кўпи йўқолди — тўлиқ рўйхатми ёки қўшимчами, сўраймиз
        RET_PENDING[update.effective_user.id] = rows
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("🔄 Бу тўлиқ рўйхат (+ ва −)", callback_data="rv_full")],
                                   [InlineKeyboardButton("➕ Фақат қўшимча (қўшиш)", callback_data="rv_add")],
                                   [InlineKeyboardButton("❌ Бекор", callback_data="rv_no")]])
        return await update.message.reply_text(
            f"↩️ Файлда {len(rows)} та товар. Ҳозирги рўйхатдаги {len(old)} тадан {gone} таси бу файлда йўқ.\n"
            "Бу файл возвратнинг ТЎЛИҚ ЖОРИЙ рўйхатими (йўқлари чиқарилади) ёки фақат ҚЎШИМЧАми?", reply_markup=kb)
    await apply_returns(update, rows, False)


RET_PENDING = {}


async def apply_returns(update: Update, rows, faqat_qosh):
    msg = update.effective_message
    async with LOCK:
        backup(VOZVRAT)
        st = nazorat.sync_returns(str(VOZVRAT), rows, len(nazorat.load_ledger(str(LEDGER))), faqat_qosh)
    L = ["✅ Возврат рўйхати (1-қисм) янгиланди:"]
    if st["plus"]:
        L.append(f"   ➕ қўшилди: {st['plus']} та ({st['plus_s']:,.0f} $)")
    if st["minus"]:
        L.append(f"   ➖ чиқарилди: {st['minus']} та ({st['minus_s']:,.0f} $) — файлда йўқ")
    if st["back"]:
        L.append(f"   ↩️ қайта қўшилди: {st['back']} та")
    if not (st["plus"] or st["minus"] or st["back"]):
        L.append("   ўзгариш йўқ — рўйхат олдингидек")
    L.append(f"📦 Ҳозир рўйхатда: {st['n']} та товар · {st['summa']:,.0f} $")
    L.append("Чиқарилган товарлар олдин юкланган фактураларда возврат бўлиб қолади, янги фактураларда эса ҳисобланмайди.")
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("↩️ Бекор қилиш", callback_data="rv_undo")]])
    await msg.reply_text("\n".join(L), reply_markup=kb)
    await run_and_reply(update)


async def returns_cb(update: Update, ctx):
    q = update.callback_query
    if not allowed(update) or not has_perm(update.effective_user.id, "vozvrat_yuklash"):
        return await q.answer("Рухсат йўқ", show_alert=True)
    await q.answer()
    if q.data == "rv_undo":
        olds = sorted(BACKUP.glob(f"{VOZVRAT.stem}_*{VOZVRAT.suffix}"))
        if not olds:
            return await q.edit_message_text(q.message.text + "\n\n(нусха топилмади)")
        async with LOCK:
            shutil.copy(olds[-1], VOZVRAT)
            olds[-1].unlink()
        await q.edit_message_text(q.message.text + "\n\n↩️ Бекор қилинди — олдинги рўйхат қайтди.")
        return await run_and_reply(update)
    rows = RET_PENDING.pop(update.effective_user.id, None)
    if q.data == "rv_no" or not rows:
        return await q.edit_message_text("Бекор қилинди.")
    await q.edit_message_text(q.message.text + ("\n\n→ Тўлиқ рўйхат" if q.data == "rv_full" else "\n\n→ Фақат қўшиш"))
    await apply_returns(update, rows, q.data == "rv_add")


@guard
async def document(update: Update, ctx):
    doc = update.message.document
    name = doc.file_name or "faktura.xlsx"
    if not name.lower().endswith((".xlsx", ".xlsm")):
        return await update.message.reply_text("❌ Фақат Excel (.xlsx) файл ташланг.")
    if (doc.file_size or 0) > 20 * 1024 * 1024:  # Telegram бот 20 МБ дан каттасини юклаб ололмайди
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("📤 Mini app орқали юклаш", web_app=WebAppInfo(url=WEBAPP_URL + "#yuklash"))]]) \
            if WEBAPP_URL else None
        return await update.message.reply_text(
            f"📦 Файл катта ({doc.file_size / 1024 / 1024:.1f} МБ). Telegram бот 20 МБ гача файлни олади.\n"
            "Mini app → Faktura → «📤 Excel yuklash» орқали юкланг — у ерда чеклов йўқ.", reply_markup=kb)
    with tempfile.TemporaryDirectory() as td:
        path = Path(td) / name
        f = await doc.get_file()
        await f.download_to_drive(path)
        uid = update.effective_user.id
        cap = (update.message.caption or "").lower()
        if ctx.user_data.pop("oylik_mode", False) or re.search(r"ойлик|oylik|текшир|tekshir", cap):
            if not has_perm(uid, "faktura_yuklash"):
                return await update.message.reply_text(f"⛔ Сизда рухсат йўқ: {PERMS['faktura_yuklash']}.")
            return await handle_oylik_file(update, path)
        if ctx.user_data.pop("vozvrat_mode", False) or await asyncio.to_thread(looks_like_returns, str(path), name):
            if not has_perm(uid, "vozvrat_yuklash"):
                return await update.message.reply_text(f"⛔ Сизда рухсат йўқ: {PERMS['vozvrat_yuklash']}.")
            return await handle_returns_file(update, path, name)
        if not has_perm(uid, "faktura_yuklash"):
            return await update.message.reply_text(f"⛔ Сизда рухсат йўқ: {PERMS['faktura_yuklash']}.")
        # юкламадан олдинги реестр нусхаси — /bekor учун
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        BACKUP.mkdir(parents=True, exist_ok=True)
        bk = BACKUP / f"ledger_{stamp}.csv"
        shutil.copy(LEDGER, bk)
        try:
            rows = await asyncio.to_thread(nazorat.read_invoice, str(path))
        except Exception as e:
            log.exception("read error")
            return await update.message.reply_text("❌ Файлни очиб бўлмади. Excel (.xlsx) эканини текширинг.")
        if rows:
            invs = {(r.get("faktura"), r.get("filial")) for r in rows}
            ds = sorted({r.get("sana") for r in rows if r.get("sana")})
            span = (dt.date.fromisoformat(ds[-1]) - dt.date.fromisoformat(ds[0])).days if ds else 0
            eski = nazorat.eski_faktura(load_cfg(), rows)
            if eski:  # кундалик пуллардан олдинги сана — эски фактура: фақат текширамиз
                bk.unlink(missing_ok=True)
                await update.message.reply_text(
                    f"📅 Бу фактура санаси ({nazorat._dmy(eski)}) кундалик пуллардан олдин — эски фактура.\n"
                    "Реестрга ҚЎШМАЙМАН ва ҳозирги пулдан айирмайман — ўша давр пуллари ва возвратлар бўйича текшираман.")
                return await handle_oylik_file(update, path)
            if len(invs) > 5 or span > 3:  # кўп фактура / узоқ давр — бу ойлик (15 кунлик) приход, реестрга қўшмаймиз
                bk.unlink(missing_ok=True)
                await update.message.reply_text(
                    f"📅 Бу ойлик/давр приходи кўринади ({len(invs)} та фактура, {nazorat._dmy(ds[0])} – {nazorat._dmy(ds[-1])}).\n"
                    "Реестрга ҚЎШМАЙМАН — фақат текшириб хулоса бераман.")
                return await handle_oylik_file(update, path)
        if not rows:
            return await update.message.reply_text(
                "❌ Бу файлда фактура топилмади.\n"
                "Возврат рўйхати бўлса — аввал /vozvrat ёзинг, кейин файлни ташланг.")
        # бир нечта фактура кетма-кет ташланса — 4 сония кутиб, ҲАММАСИНИ БИРГА текширамиз (битта жавоб)
        chat = update.effective_chat.id
        b = BATCH.get(chat)
        if b is None:
            b = BATCH[chat] = {"dir": tempfile.mkdtemp(), "files": [], "bk": bk, "task": None}
        else:
            bk.unlink(missing_ok=True)  # реестр ҳали ўзгармаган — биринчи нусха етарли
        dst = Path(b["dir"]) / f"{len(b['files']):03d}_{name}"
        shutil.copy(path, dst)
        b["files"].append(str(dst))
        b["update"] = update
        if b["task"]:
            b["task"].cancel()
        b["task"] = asyncio.create_task(_batch_run(chat))


BATCH = {}


async def _batch_run(chat):
    try:
        await asyncio.sleep(4)
    except asyncio.CancelledError:
        return
    b = BATCH.pop(chat, None)
    if not b:
        return
    update, files, bk = b["update"], b["files"], b["bk"]
    try:
        if len(files) > 1:
            await update.effective_message.reply_text(f"📥 {len(files)} та файл қабул қилинди — ҳаммасини бирга текширяпман…")
        await run_and_reply(update, files=files, caption="Ҳисобот")
        if filecmp.cmp(bk, LEDGER, shallow=False):
            bk.unlink()  # ҳеч нарса қўшилмади — бекор қилиш нуқтаси ўзгармайди
        else:
            (DATA / "last_backup.txt").write_text(str(bk))
    except Exception:
        log.exception("batch")
        await update.effective_message.reply_text("❌ Текширишда хато чиқди. Файлларни қайта ташлаб кўринг.")
    finally:
        shutil.rmtree(b["dir"], ignore_errors=True)


@admin_only
async def tozalash(update: Update, ctx):
    cfg = load_cfg()
    n_tr = len(cfg["transhlar"])
    n_rows = len(nazorat.load_ledger(str(LEDGER)))
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("🗑 Ҳаммасини ўчириш (транш + фактура + возврат)", callback_data="rs_all")],
        [InlineKeyboardButton("🧾 Фақат фактураларни ўчириш", callback_data="rs_inv")],
        [InlineKeyboardButton("↩️ Фақат возврат рўйхатини ўчириш", callback_data="rs_ret")],
        [InlineKeyboardButton("❌ Бекор", callback_data="rs_no")]])
    await update.message.reply_text(
        f"Ҳозир базада: {n_tr} та транш, {n_rows} та фактура қатори, "
        f"{len(nazorat.load_returns(str(VOZVRAT)))} та возврат.\n\n"
        "Нимани ўчирай? (ўчиришдан олдин нусха сақланади)", reply_markup=kb)


async def tozalash_cb(update: Update, ctx):
    q = update.callback_query
    if not is_admin(update):
        return await q.answer("Фақат админлар учун", show_alert=True)
    await q.answer()
    if q.data == "rs_no":
        return await q.edit_message_text("Бекор қилинди. Ҳеч нарса ўчирилмади.")
    if q.data == "rs_ret":
        async with LOCK:
            backup(VOZVRAT)
            nazorat.save_returns(str(VOZVRAT), [])
            cfg = load_cfg(); cfg["vozvrat_qismlar"] = []; cfg.pop("vozvrat_1", None); save_cfg(cfg)
        return await q.edit_message_text("✅ Возврат рўйхати ўчирилди. Янги рўйхатни /vozvrat орқали ташланг.")
    async with LOCK:
        backup(LEDGER)
        nazorat.save_ledger(str(LEDGER), [])
        if q.data == "rs_all":
            backup(VOZVRAT)
            nazorat.save_returns(str(VOZVRAT), [])
            cfg = load_cfg()
            cfg["transhlar"] = []
            cfg["vozvrat_qismlar"] = []
            save_cfg(cfg)  # save_cfg ўзи нусха олади
        (DATA / "last_backup.txt").unlink(missing_ok=True)
    if q.data == "rs_all":
        await q.edit_message_text(
            "✅ Ҳаммаси ўчирилди. База бўш.\n\n"
            "Энди биринчи транш (пул) ни киритинг: Gavhar юборган нарх хабарини шу ерга forward қилинг "
            "ва сўралганда суммани ёзинг (ёки mini app → Transhlar → «+ Yangi transh»).\n"
            "Кейин фактураларни ташлайверасиз.")
    else:
        await q.edit_message_text("✅ Фактуралар ўчирилди, траншлар қолди. Фактураларни қайтадан ташлашингиз мумкин.")


@need("hisobot")
async def xatolar_cmd(update: Update, ctx):
    """/xatolar — ҳамма хатоликлар Excel да. /xatolar GVR-27102026 — битта фактура бўйича."""
    arg = " ".join(ctx.args or []).strip()
    async with LOCK:
        cfg = load_cfg()
        rets = nazorat.load_returns(str(VOZVRAT))
        rows, _ = nazorat.compute(cfg, nazorat.load_ledger(str(LEDGER)), rets)
    if arg:
        rows = [r for r in rows if arg.lower() in f"{r['faktura']} {r['filial']}".lower()]
    bad = [r for r in rows if str(r.get("holat", "")).startswith("⚠")]
    if not bad:
        return await update.message.reply_text("✅ Хатолик йўқ" + (f" ({arg})" if arg else "") + ".")
    with tempfile.TemporaryDirectory() as td:
        nm = f"Xatoliklar_{re.sub(r'[^\w-]+', '_', arg)[:40] or 'barchasi'}_{dt.date.today():%Y-%m-%d}.xlsx"
        out = Path(td) / nm
        st = await asyncio.to_thread(nazorat.write_errors_report, str(out), rows, f"Фильтр: {arg or 'барча хатоликлар'}", rets)
        with open(out, "rb") as f:
            await update.message.reply_document(f, filename=nm,
                                                caption=f"❌ Хатолар: {st['faktura']} та фактура, {st['n']} та товар, жами {st['farq']:+,.2f} $")


async def on_error(update, ctx):
    log.exception("Хато", exc_info=ctx.error)
    if isinstance(update, Update) and update.effective_message:
        await update.effective_message.reply_text("❌ Нимадир хато кетди. Яна бир марта уриниб кўринг.")


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
