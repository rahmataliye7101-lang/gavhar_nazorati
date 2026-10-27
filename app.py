#!/usr/bin/env python3
"""Gavhar nazorat — Telegram Mini App + бот (битта Railway сервиси).

Муҳит ўзгарувчилари:
  BOT_TOKEN       — @BotFather токени
  ALLOWED_USERS   — рухсат этилган Telegram ID лар (вергул билан)
  DATA_DIR        — Volume папкаси (стандарт /data)
  WEBAPP_URL      — mini app манзили (бўш бўлса RAILWAY_PUBLIC_DOMAIN дан олинади)
  PORT            — Railway ўзи беради
  DEV_NO_AUTH=1   — фақат локал синов учун: Telegram текширувисиз (DEV_USER ID билан)
"""
import asyncio
import datetime as dt
import hashlib
import hmac
import json
import logging
import os
import re
import tempfile
import time
from pathlib import Path
from urllib.parse import parse_qsl

from aiohttp import web

import bot as B
import nazorat as N

log = logging.getLogger("gavhar-app")
BASE = Path(__file__).resolve().parent
TOKEN = os.environ.get("BOT_TOKEN", "")
PORT = int(os.environ.get("PORT", "8080"))
DEV = os.environ.get("DEV_NO_AUTH") == "1"
WEBAPP_URL = os.environ.get("WEBAPP_URL") or (
    f"https://{os.environ['RAILWAY_PUBLIC_DOMAIN']}" if os.environ.get("RAILWAY_PUBLIC_DOMAIN") else "")
TG_APP = None  # telegram Application (ҳисобот юбориш учун)


# ---------------- Telegram initData текшируви ----------------
def verify_init_data(init_data: str):
    if DEV:
        return int(os.environ.get("DEV_USER", "0") or 0)
    if not init_data or not TOKEN:
        return None
    pairs = dict(parse_qsl(init_data, keep_blank_values=True))
    h = pairs.pop("hash", None)
    if not h:
        return None
    dcs = "\n".join(f"{k}={v}" for k, v in sorted(pairs.items()))
    secret = hmac.new(b"WebAppData", TOKEN.encode(), hashlib.sha256).digest()
    calc = hmac.new(secret, dcs.encode(), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(calc, h):
        return None
    if time.time() - int(pairs.get("auth_date", "0")) > 7 * 86400:
        return None
    try:
        return int(json.loads(pairs.get("user", "{}")).get("id"))
    except (TypeError, ValueError):
        return None


@web.middleware
async def auth_mw(request, handler):
    if request.path.startswith("/api/"):
        uid = verify_init_data(request.headers.get("X-Init-Data", ""))
        if uid is None:
            return web.json_response({"error": "Telegram орқали очинг"}, status=401)
        if not DEV and uid not in B.ALLOWED:
            return web.json_response({"error": f"Рухсат йўқ. ID: {uid}"}, status=403)
        request["uid"] = uid
    return await handler(request)


# ---------------- ҳисоб ----------------
def state():
    cfg = B.load_cfg()
    rows, trs = N.compute(cfg, N.load_ledger(str(B.LEDGER)))
    return cfg, rows, trs, N.stats(rows, trs)


def r2(x, n=2):
    return None if x is None else round(float(x), n)


def row_json(r):
    return dict(sana=r["sana"], faktura=r["faktura"], filial=r["filial"], kod=r["kod"], nomi=r["nomi"],
                kategoriya=r["kategoriya"], tur=r["tur"], gramm=r2(r["gramm"]), summa=r2(r["summa"]),
                fakt_1gr=r2(r.get("fakt_1gr")), hisob_1gr=r2(r.get("hisob_1gr")), farq_gr=r2(r.get("farq_gr")),
                farq=r2(r.get("farq")), holat=r.get("holat"), izoh=r.get("izoh"), transh=r.get("transh"))


def transh_no(r):
    m = re.search(r"№(\d+)", r.get("transh") or "")
    return int(m.group(1)) if m else None


def invoices(rows, cfg):
    """Фактуралар (фактура № + филиал + сана) — транш бўйича гуруҳланган."""
    inv = {}
    for r in rows:
        k = (r["faktura"], r["filial"], r["sana"])
        a = inv.get(k)
        if a is None:
            a = inv[k] = dict(key="|".join(k), faktura=r["faktura"], filial=r["filial"], sana=r["sana"], transh=transh_no(r),
                              n=0, gramm=0.0, fakt=0.0, hisob=0.0, bad=0, qimmat=0.0, turlar={})
        a["n"] += 1
        a["gramm"] += r["gramm"] or 0
        a["fakt"] += r["summa"] or 0
        a["hisob"] += r.get("hisob") or 0
        isbad = str(r.get("holat", "")).startswith("⚠")
        if isbad:
            a["bad"] += 1
            if (r.get("farq") or 0) > 0:
                a["qimmat"] += r["farq"]
        t = a["turlar"].setdefault(r["tur"], dict(tur=r["tur"], n=0, gramm=0.0, fakt=0.0, hisob=0.0, bad=0))
        t["n"] += 1; t["gramm"] += r["gramm"] or 0; t["fakt"] += r["summa"] or 0; t["hisob"] += r.get("hisob") or 0
        t["bad"] += 1 if isbad else 0
    out = []
    for a in inv.values():
        a["farq"] = a["fakt"] - a["hisob"]
        tl = []
        for t in sorted(a["turlar"].values(), key=lambda t: -t["gramm"]):
            g = t["gramm"] or 1
            tl.append(dict(tur=t["tur"], n=t["n"], gramm=r2(t["gramm"]), fakt=r2(t["fakt"]), hisob=r2(t["hisob"]),
                           farq=r2(t["fakt"] - t["hisob"]), fakt_1gr=r2(t["fakt"] / g), hisob_1gr=r2(t["hisob"] / g), bad=t["bad"]))
        a["turlar"] = tl
        for f in ("gramm", "fakt", "hisob", "farq", "qimmat"):
            a[f] = r2(a[f])
        out.append(a)
    out.sort(key=lambda a: (a["sana"] or "", a["faktura"]), reverse=True)
    return out


async def api_fakturalar(request):
    cfg, rows, trs, s = await asyncio.to_thread(state)
    inv = invoices(rows, cfg)
    groups = []
    for t in sorted(trs, key=lambda t: t["sana"], reverse=True):
        items = [a for a in inv if a["transh"] == t["no"]]
        groups.append(dict(no=t["no"], sana=t["sana"], summa=t["summa"], birja=t["birja"], holat=t["holat"],
                           used=r2(t["used"]), left=r2(t["left"]), pct=r2(t["used"] / t["summa"] * 100, 1) if t["summa"] else 0,
                           count=len(items), n=sum(a["n"] for a in items), gramm=r2(sum(a["gramm"] for a in items)),
                           fakt=r2(sum(a["fakt"] for a in items)), hisob=r2(sum(a["hisob"] for a in items)),
                           farq=r2(sum(a["farq"] for a in items)), bad=sum(1 for a in items if a["bad"]), items=items))
    other = [a for a in inv if a["transh"] not in {t["no"] for t in trs}]
    return web.json_response({"chegara": cfg.get("chegara_usd_gr", 0.3), "groups": groups, "other": other})


async def api_holat(request):
    cfg, rows, trs, s = await asyncio.to_thread(state)
    inv = invoices(rows, cfg)
    ok = s["ok"]
    days = sorted(s["by_day"])[-30:]
    agg = {}
    for r in ok:
        if r["sana"] in days:
            a = agg.setdefault(r["sana"], [0.0, 0.0, 0])
            a[0] += r["summa"]; a[1] += r["hisob"]; a[2] += 1
    turlar = {}
    for r in rows:
        t = turlar.setdefault(r["tur"], {"tur": r["tur"], "n": 0, "gramm": 0.0, "bad": 0, "farq": 0.0})
        t["n"] += 1; t["gramm"] += r["gramm"] or 0
        if str(r.get("holat", "")).startswith("⚠"):
            t["bad"] += 1; t["farq"] += r.get("farq") or 0
    act = s["active"]
    dates = [r["sana"] for r in rows if r["sana"]]
    return web.json_response(dict(
        given=s["tot_given"], used=r2(s["tot_used"]), left=r2(s["left_total"]),
        active=dict(no=act["no"], sana=act["sana"], birja=act["birja"]) if act else None,
        gramm999=r2(s["left_total"] / act["birja"], 1) if act else None,
        temp=r2(s["temp"], 0), tugash=s["tugash"].isoformat() if s["tugash"] else None,
        fakt=r2(s["fakt_sum"]), hisob=r2(s["hisob_sum"]), farq=r2(s["fakt_sum"] - s["hisob_sum"]),
        qimmat=r2(s["qimmat"]), arzon=r2(s["arzon"]), n=len(rows), bad_n=len(s["bad"]),
        filiallar=len({r["filial"] for r in rows}), sana_from=min(dates, default=None), sana_to=max(dates, default=None),
        chegara=cfg.get("chegara_usd_gr", 0.3),
        transhlar=[dict(no=t["no"], sana=t["sana"], summa=t["summa"], birja=t["birja"], used=r2(t["used"]),
                        left=r2(t["left"]), holat=t["holat"], pct=r2(t["used"] / t["summa"] * 100, 1) if t["summa"] else 0)
                   for t in trs],
        kunlik=[dict(sana=d, fakt=r2(agg[d][0]), hisob=r2(agg[d][1]), farq=r2(agg[d][0] - agg[d][1]), n=agg[d][2])
                for d in days if d in agg],
        turlar=sorted(({**t, "gramm": r2(t["gramm"]), "farq": r2(t["farq"])} for t in turlar.values()), key=lambda t: -t["n"]),
        faktura_soni=len(inv), oxirgi=inv[0] if inv else None,
    ))


async def api_xatolar(request):
    cfg, rows, trs, s = await asyncio.to_thread(state)
    bad = sorted(s["bad"], key=lambda r: -abs(r.get("farq") or 0))
    return web.json_response({"chegara": cfg.get("chegara_usd_gr", 0.3), "rows": [row_json(r) for r in bad]})


async def api_transhlar(request):
    cfg, rows, trs, s = await asyncio.to_thread(state)
    st = {t["no"]: t for t in trs}
    out = []
    for t in cfg["transhlar"]:
        x = st.get(t["no"], {})
        out.append(dict(no=t["no"], sana=t["sana"], summa=t["summa_usd"], birja=t["birja_999"], narxlar=t["narxlar"],
                        yopilgan_sana=t.get("yopilgan_sana"), used=r2(x.get("used")), left=r2(x.get("left")),
                        holat=x.get("holat"), lom=r2(t["birja_999"] * 585 / 999)))
    return web.json_response({"chegara": cfg.get("chegara_usd_gr", 0.3), "transhlar": out})


def transh_from_body(b):
    narx = {}
    for k, v in (b.get("narxlar") or {}).items():
        try:
            if v not in (None, ""):
                narx[k] = float(str(v).replace(",", "."))
        except ValueError:
            pass
    if "Вакс" in narx and "Кузсиз" not in narx:
        narx["Кузсиз"] = narx["Вакс"]
    d = N.parse_date(b.get("sana"))
    num = lambda x: float(str(x).replace(" ", "").replace(",", ".")) if x not in (None, "") else None
    return {"sana": d.isoformat() if d else None, "birja": num(b.get("birja")), "lom": None,
            "summa": num(b.get("summa")), "narxlar": narx}


async def api_transh_tekshir(request):
    p = transh_from_body(await request.json())
    errs, notes = B.check_transh(p, B.load_cfg())
    return web.json_response({"errors": errs, "notes": notes, "parsed": p})


async def api_transh_saqla(request):
    p = transh_from_body(await request.json())
    async with B.LOCK:
        cfg = B.load_cfg()
        errs, _ = B.check_transh(p, cfg)
        if errs:
            return web.json_response({"errors": errs}, status=400)
        no = max((t["no"] for t in cfg["transhlar"]), default=0) + 1
        cfg["transhlar"].append({"no": no, "sana": p["sana"], "summa_usd": p["summa"], "birja_999": p["birja"],
                                 "narxlar": p["narxlar"], "manba": f"Mini app, {dt.date.today()}", "yopilgan_sana": None})
        cfg["transhlar"].sort(key=lambda t: t["sana"])
        B.save_cfg(cfg)
    return web.json_response({"ok": True, "no": no})


async def api_transh_yop(request):
    b = await request.json()
    d = N.parse_date(b.get("sana")) if b.get("sana") else None
    async with B.LOCK:
        cfg = B.load_cfg()
        t = next((t for t in cfg["transhlar"] if t["no"] == int(b.get("no", 0))), None)
        if not t:
            return web.json_response({"errors": ["Бундай транш йўқ"]}, status=404)
        t["yopilgan_sana"] = d.isoformat() if d else None
        B.save_cfg(cfg)
    return web.json_response({"ok": True})


async def api_chegara(request):
    b = await request.json()
    try:
        v = float(str(b.get("v")).replace(",", "."))
        assert 0 < v < 50
    except Exception:
        return web.json_response({"errors": ["Чегара нотўғри"]}, status=400)
    async with B.LOCK:
        cfg = B.load_cfg()
        cfg["chegara_usd_gr"] = v
        B.save_cfg(cfg)
    return web.json_response({"ok": True})


async def api_hisobot(request):
    if TG_APP is None:
        return web.json_response({"errors": ["Бот уланмаган"]}, status=503)
    uid = request["uid"]
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / f"Gavhar_nazorat_{dt.date.today():%Y-%m-%d}.xlsx"
        async with B.LOCK:
            await asyncio.to_thread(N.run, str(B.TRANSH), str(B.LEDGER), str(out), [])
        with open(out, "rb") as f:
            await TG_APP.bot.send_document(uid, f, filename=out.name, caption="Ҳисобот (mini app дан)")
    return web.json_response({"ok": True})


async def index(request):
    return web.FileResponse(BASE / "index.html", headers={"Cache-Control": "no-cache"})


def make_web():
    app = web.Application(middlewares=[auth_mw])
    app.router.add_get("/", index)
    app.router.add_get("/api/holat", api_holat)
    app.router.add_get("/api/xatolar", api_xatolar)
    app.router.add_get("/api/fakturalar", api_fakturalar)
    app.router.add_get("/api/transhlar", api_transhlar)
    app.router.add_post("/api/transh/tekshir", api_transh_tekshir)
    app.router.add_post("/api/transh", api_transh_saqla)
    app.router.add_post("/api/transh/yop", api_transh_yop)
    app.router.add_post("/api/chegara", api_chegara)
    app.router.add_post("/api/hisobot", api_hisobot)
    app.router.add_get("/static/chart.umd.min.js", lambda r: web.FileResponse(BASE / "chart.umd.min.js"))
    app.router.add_get("/health", lambda r: web.Response(text="ok"))
    return app


# ---------------- бот ----------------
async def start_cmd(update, ctx):
    from telegram import InlineKeyboardButton, InlineKeyboardMarkup, WebAppInfo
    if not B.allowed(update):
        return await B.deny(update)
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("🟡 Назоратни очиш", web_app=WebAppInfo(url=WEBAPP_URL))]]) if WEBAPP_URL else None
    await update.message.reply_text(
        "Gavhar — тилло нарх назорати.\n\n"
        "• Ҳолат, хатоликлар, траншлар — «Назоратни очиш» тугмаси ёки пастдаги «Назорат» менюси орқали.\n"
        "• Фактура — тизим экспортини (.xlsx) шу чатга ташланг, бот текшириб ҳисобот қайтаради.\n"
        "• /bekor — охирги юкламани бекор қилиш.", reply_markup=kb)


def make_bot():
    from telegram import Update
    from telegram.ext import Application, CommandHandler, MessageHandler, filters
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler(["start", "help"], start_cmd))
    app.add_handler(CommandHandler("myid", B.myid))
    app.add_handler(CommandHandler("bekor", B.bekor))
    app.add_handler(CommandHandler("holat", B.holat))
    app.add_handler(CommandHandler("reestr", B.reestr))
    app.add_handler(MessageHandler(filters.Document.ALL, B.document))
    app.add_error_handler(B.on_error)
    return app, Update


async def main():
    global TG_APP
    B.ensure_data()
    runner = web.AppRunner(make_web())
    await runner.setup()
    await web.TCPSite(runner, "0.0.0.0", PORT).start()
    log.info("Mini app: порт %s, URL %s", PORT, WEBAPP_URL or "(WEBAPP_URL йўқ)")
    if TOKEN:
        from telegram import MenuButtonWebApp, WebAppInfo
        TG_APP, Update = make_bot()
        await TG_APP.initialize()
        await TG_APP.start()
        await TG_APP.updater.start_polling(allowed_updates=Update.ALL_TYPES)
        if WEBAPP_URL:
            await TG_APP.bot.set_chat_menu_button(menu_button=MenuButtonWebApp("Назорат", WebAppInfo(url=WEBAPP_URL)))
        log.info("Бот ишга тушди")
    else:
        log.warning("BOT_TOKEN йўқ — фақат веб сервер")
    await asyncio.Event().wait()


if __name__ == "__main__":
    asyncio.run(main())
