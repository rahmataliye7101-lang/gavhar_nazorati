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


PERM_MAP = {("GET", "/api/holat"): "holat", ("GET", "/api/fakturalar"): "faktura", ("GET", "/api/xatolar"): "xatolar",
            ("GET", "/api/transhlar"): "transhlar", ("POST", "/api/transh/tekshir"): "transh_qoshish",
            ("POST", "/api/transh"): "transh_qoshish", ("POST", "/api/transh/yop"): "transh_qoshish",
            ("POST", "/api/chegara"): "transh_qoshish", ("POST", "/api/hisobot"): "hisobot",
            ("POST", "/api/xatolar/excel"): "hisobot"}


@web.middleware
async def auth_mw(request, handler):
    if request.path.startswith("/api/"):
        uid = verify_init_data(request.headers.get("X-Init-Data", ""))
        if uid is None:
            return web.json_response({"error": "Telegram орқали очинг"}, status=401)
        if not DEV and not B.allowed_id(uid):
            return web.json_response({"error": f"Рухсат йўқ. ID: {uid}"}, status=403)
        request["uid"] = uid
        need = PERM_MAP.get((request.method, request.path))
        if need and not DEV and not B.has_perm(uid, need):
            return web.json_response({"error": f"Рухсат йўқ: {B.PERMS[need]}"}, status=403)
    return await handler(request)


# ---------------- ҳисоб ----------------
def state():
    cfg = B.load_cfg()
    rets = N.load_returns(str(B.VOZVRAT))
    rows, trs = N.compute(cfg, N.load_ledger(str(B.LEDGER)), rets)
    st = N.stats(rows, trs)
    st["vozvrat"] = N.returns_stats(rets)
    return cfg, rows, trs, st


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
                              n=0, gramm=0.0, fakt=0.0, hisob=0.0, bad=0, qimmat=0.0, vozvrat=0, turlar={})
        if a["transh"] is None:
            a["transh"] = transh_no(r)
        a.setdefault("trs", set()).update(int(x) for x in re.findall(r"№(\d+)", str(r.get("transh") or "")))
        a["n"] += 1
        if r.get("transh") == "возврат":
            a["vozvrat"] += 1
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
    tinfo = {t["no"]: t for t in cfg.get("transhlar", [])}
    out = []
    for a in inv.values():
        a["farq"] = a["fakt"] - a["hisob"]
        a["trs"] = [dict(no=n, sana=tinfo[n]["sana"], birja=tinfo[n].get("birja_999"))
                    for n in sorted(a.pop("trs", set())) if n in tinfo]
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


OY_NOMI = ["", "Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun", "Iyul", "Avgust", "Sentabr", "Oktabr", "Noyabr", "Dekabr"]


def parse_filter(request):
    tr = {int(x) for x in re.findall(r"\d+", request.query.get("tr", ""))}
    oy = request.query.get("oy", "")[:7]
    return tr, oy


def apply_filter(rows, trs, tr, oy):
    """Транш(лар) ва ой бўйича фильтр. Возврат қаторлари ўз фактурасининг траншига тегишли ҳисобланади."""
    inv_tr = {}
    for r in rows:
        k = (r["faktura"], r["filial"], r["sana"])
        if inv_tr.get(k) is None:
            inv_tr[k] = transh_no(r)

    def keep(r):
        if oy and not (r["sana"] or "").startswith(oy):
            return False
        if tr:
            n = transh_no(r) or inv_tr.get((r["faktura"], r["filial"], r["sana"]))
            if n not in tr:
                return False
        return True
    return [r for r in rows if keep(r)], [t for t in trs if not tr or t["no"] in tr]


async def api_filtr(request):
    cfg, rows, trs, s = await asyncio.to_thread(state)
    months = sorted({(r["sana"] or "")[:7] for r in rows if r["sana"]}, reverse=True)
    return web.json_response(dict(
        transhlar=[dict(no=t["no"], sana=t["sana"], summa=t["summa"], birja=t["birja"]) for t in sorted(trs, key=lambda t: t["sana"])],
        oylar=[dict(oy=m, nomi=f"{OY_NOMI[int(m[5:7])]} {m[:4]}") for m in months]))


async def api_fakturalar(request):
    cfg, rows, trs, s = await asyncio.to_thread(state)
    _, oy = parse_filter(request)
    inv = [a for a in invoices(rows, cfg) if not oy or (a["sana"] or "").startswith(oy)]
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
    cfg, rows_all, trs_all, s_all = await asyncio.to_thread(state)
    tr_f, oy = parse_filter(request)
    rows, trs = apply_filter(rows_all, trs_all, tr_f, oy)
    s = N.stats(rows, trs)
    s["vozvrat"] = s_all["vozvrat"]
    inv = invoices(rows, cfg)
    ok = s["ok"]
    davr = None
    if tr_f or oy:
        rv = [r for r in rows if r.get("transh") == "возврат"]
        okn = [r for r in ok if r.get("transh") != "возврат"]
        davr = dict(oy=oy, oy_nomi=f"{OY_NOMI[int(oy[5:7])]} {oy[:4]}" if oy else "", tr=sorted(tr_f),
                    faktura_soni=len(inv), n=len(rows), gramm=r2(sum(r["gramm"] or 0 for r in rows)),
                    fakt=r2(sum(r["summa"] or 0 for r in ok)), hisob=r2(sum(r["hisob"] or 0 for r in ok)),
                    farq=r2(s["fakt_sum"] - s["hisob_sum"]), qimmat=r2(s["qimmat"]), arzon=r2(s["arzon"]),
                    bad_n=len(s["bad"]), vozvrat_n=len(rv), vozvrat_summa=r2(sum(r["summa"] or 0 for r in rv)),
                    olingan=r2(sum(r["hisob"] or 0 for r in okn)), filiallar=len({r["filial"] for r in rows}))
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
        vozvrat={k: r2(v) if isinstance(v, float) else v for k, v in s["vozvrat"].items()},
        davr=davr,
    ))


async def api_xatolar(request):
    cfg, rows, trs, s = await asyncio.to_thread(state)
    tr_f, oy = parse_filter(request)
    rows, _ = apply_filter(rows, trs, tr_f, oy)
    bad = sorted([r for r in rows if str(r.get("holat", "")).startswith("⚠")], key=lambda r: -abs(r.get("farq") or 0))
    return web.json_response({"chegara": cfg.get("chegara_usd_gr", 0.3), "rows": [row_json(r) for r in bad]})


async def api_xatolar_excel(request):
    """Танланган хатоликларни Excel қилиб чатга юбориш (фильтрлар билан)."""
    if TG_APP is None:
        return web.json_response({"errors": ["Бот уланмаган"]}, status=503)
    b = await request.json()
    cfg, rows, trs, s = await asyncio.to_thread(state)
    tr = {int(x) for x in b.get("tr") or []}
    oy = (b.get("oy") or "")[:7]
    rows, _ = apply_filter(rows, trs, tr, oy)
    k, fil, tur_, q = b.get("k", "all"), b.get("filial", ""), b.get("tur", ""), (b.get("q") or "").strip().lower()
    faktura = b.get("faktura", "")

    def keep(r):
        h = str(r.get("holat", ""))
        if not h.startswith("⚠"):
            return False
        if k == "q" and "ҚИММАТ" not in h: return False
        if k == "a" and "арзон" not in h: return False
        if k == "o" and ("ҚИММАТ" in h or "арзон" in h): return False
        if fil and r["filial"] != fil: return False
        if tur_ and r["tur"] != tur_: return False
        if faktura and r["faktura"] != faktura: return False
        if q and q not in f"{r['kod']} {r['faktura']} {r['nomi']}".lower(): return False
        return True
    sel = [r for r in rows if keep(r)]
    if not sel:
        return web.json_response({"errors": ["Танланган хатолик йўқ"]}, status=400)
    parts = []
    if tr: parts.append("транш " + ", ".join(f"№{n}" for n in sorted(tr)))
    if oy: parts.append(f"{OY_NOMI[int(oy[5:7])]} {oy[:4]}")
    if faktura: parts.append(f"фактура {faktura}")
    if fil: parts.append(fil)
    if tur_: parts.append(tur_)
    if k != "all": parts.append({"q": "қиммат", "a": "арзон", "o": "бошқа"}.get(k, k))
    if q: parts.append(f"қидирув «{q}»")
    title = "Фильтр: " + (" · ".join(parts) if parts else "барча хатоликлар")
    with tempfile.TemporaryDirectory() as td:
        nm = "Xatoliklar_" + (re.sub(r"[^\w\-]+", "_", faktura or fil or oy or "barchasi")[:40]) + f"_{dt.date.today():%Y-%m-%d}.xlsx"
        out = Path(td) / nm
        st = await asyncio.to_thread(N.write_errors_report, str(out), sel, title, N.load_returns(str(B.VOZVRAT)))
        with open(out, "rb") as f:
            await TG_APP.bot.send_document(request["uid"], f, filename=nm,
                                           caption=f"⚠ Хатоликлар: {st['n']} позиция, {st['faktura']} фактура, фарқ {st['farq']:+,.2f} $\n{title}")
    return web.json_response({"ok": True, "n": st["n"], "farq": round(st["farq"], 2)})


async def api_transhlar(request):
    cfg, rows, trs, s = await asyncio.to_thread(state)
    st = {t["no"]: t for t in trs}
    out = []
    for t in cfg["transhlar"]:
        x = st.get(t["no"], {})
        out.append(dict(no=t["no"], sana=t["sana"], summa=t["summa_usd"], birja=t["birja_999"], narxlar=t["narxlar"],
                        yopilgan_sana=t.get("yopilgan_sana"), used=r2(x.get("used")), left=r2(x.get("left")),
                        holat=x.get("holat"), lom=r2(t["birja_999"] * 585 / 999), tarkib=N.narx_tarkibi(t, cfg)))
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
        t["yopilgan_qator"] = len(N.load_ledger(str(B.LEDGER))) if d else None
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
            await asyncio.to_thread(N.run, str(B.TRANSH), str(B.LEDGER), str(out), [], str(B.VOZVRAT))
        with open(out, "rb") as f:
            await TG_APP.bot.send_document(uid, f, filename=out.name, caption="Ҳисобот (mini app дан)")
    return web.json_response({"ok": True})


# ---------------- админ (mini app) ----------------
def is_admin_req(request):
    return DEV or B.admin_id(request["uid"])


async def api_me(request):
    uid = request["uid"]
    rec = B.user_rec(uid) if uid else None
    perms = list(B.PERMS) if is_admin_req(request) else B.user_perms(uid)
    return web.json_response({"id": uid, "admin": is_admin_req(request), "name": (rec or {}).get("name", ""), "perms": perms,
                              "perm_names": B.PERMS})


def admin_state():
    u = B.load_users()
    names = u.get("names", {})
    users = [dict(id=aid, name="Asosiy admin (Railway)", role="admin", asosiy=True) for aid in sorted(B.ALLOWED)]
    users += [dict(id=int(x["id"]), name=x.get("name") or "", role=x.get("role", "user"), asosiy=False, sana=x.get("sana"),
                   perms=x.get("perms", B.DEFAULT_PERMS))
              for x in u["users"] if int(x["id"]) not in B.ALLOWED]
    pending = [dict(id=int(p), name=names.get(str(p), "")) for p in u["pending"]]
    now = dt.datetime.now().isoformat()
    inv = sum(1 for v in u["invites"].values() if v["exp"] > now)
    return dict(users=users, pending=pending, invites=inv, perm_names=B.PERMS)


async def notify(uid, text):
    if TG_APP is not None:
        try:
            await TG_APP.bot.send_message(int(uid), text)
        except Exception:
            log.warning("xabar yuborilmadi: %s", uid)


async def api_admin(request):
    if not is_admin_req(request):
        return web.json_response({"errors": ["Фақат админлар учун"]}, status=403)
    return web.json_response(admin_state())


async def api_admin_act(request):
    if not is_admin_req(request):
        return web.json_response({"errors": ["Фақат админлар учун"]}, status=403)
    b = await request.json()
    act, uid, role = b.get("act"), b.get("id"), b.get("role", "user")
    role = role if role in ("user", "admin") else "user"
    async with B.LOCK:
        if act in ("approve", "add"):
            try:
                uid = int(str(uid).strip())
            except ValueError:
                return web.json_response({"errors": ["ID нотўғри"]}, status=400)
            B.add_user(uid, role, by=request["uid"])
            await notify(uid, f"✅ Сизга Gavhar назорат ботидан фойдаланишга рухсат берилди ({B.ROLE_L[role]}). /start ни босинг.")
        elif act == "reject":
            u = B.load_users()
            u["pending"] = [p for p in u["pending"] if int(p) != int(uid)]
            B.save_users(u)
        elif act == "role":
            if int(uid) in B.ALLOWED:
                return web.json_response({"errors": ["Асосий админни ўзгартириб бўлмайди"]}, status=400)
            rec = B.user_rec(int(uid))
            if rec:
                B.add_user(int(uid), role, rec.get("name", ""))
        elif act == "perms":
            B.set_perms(int(uid), b.get("perms") or [])
        elif act == "delete":
            if int(uid) in B.ALLOWED:
                return web.json_response({"errors": ["Асосий админни ўчириб бўлмайди"]}, status=400)
            if int(uid) == int(request["uid"] or 0):
                return web.json_response({"errors": ["Ўзингизни ўчира олмайсиз"]}, status=400)
            B.remove_user(int(uid))
        elif act == "invite":
            import secrets
            code = secrets.token_urlsafe(8)
            u = B.load_users()
            now = dt.datetime.now()
            u["invites"] = {k: v for k, v in u["invites"].items() if v["exp"] > now.isoformat()}
            u["invites"][code] = {"role": role, "exp": (now + dt.timedelta(hours=24)).isoformat(), "by": request["uid"]}
            B.save_users(u)
            uname = None
            if TG_APP is not None:
                uname = (await TG_APP.bot.get_me()).username
            st = admin_state()
            st["link"] = f"https://t.me/{uname or 'BOT_USERNAME'}?start=inv_{code}"
            st["link_role"] = role
            return web.json_response(st)
        else:
            return web.json_response({"errors": ["Номаълум амал"]}, status=400)
    return web.json_response(admin_state())


async def index(request):
    return web.FileResponse(BASE / "index.html", headers={"Cache-Control": "no-cache"})


def make_web():
    app = web.Application(middlewares=[auth_mw])
    app.router.add_get("/", index)
    app.router.add_get("/api/holat", api_holat)
    app.router.add_get("/api/xatolar", api_xatolar)
    app.router.add_get("/api/fakturalar", api_fakturalar)
    app.router.add_get("/api/filtr", api_filtr)
    app.router.add_get("/api/me", api_me)
    app.router.add_post("/api/xatolar/excel", api_xatolar_excel)
    app.router.add_get("/api/admin", api_admin)
    app.router.add_post("/api/admin", api_admin_act)
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
    if ctx.args and ctx.args[0].startswith("inv_") and not B.allowed(update):
        if not await B.accept_invite(update, ctx, ctx.args[0][4:]):
            return
    if not B.allowed(update):
        return await B.deny(update, ctx)
    rec = B.user_rec(update.effective_user.id)
    if rec is not None and not rec.get("name"):
        B.add_user(update.effective_user.id, rec.get("role", "user"), B.uname(update.effective_user))
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("🟡 Назоратни очиш", web_app=WebAppInfo(url=WEBAPP_URL))]]) if WEBAPP_URL else None
    await update.message.reply_text(
        "🟡 Gavhar — тилло нарх назорати\n\n"
        "📎 Фактура (Excel) ташланг — текшириб бераман\n"
        "💳 Янги пул — Gavhar нарх хабарини forward қилинг\n"
        "↩️ Возврат рўйхати — /vozvrat\n"
        "📊 Ҳолат — /holat\n"
        "❌ Хатолар Excel да — /xatolar\n"
        "↩️ Охирги фактурани бекор қилиш — /bekor\n\n"
        "Батафсил — «Назоратни очиш» тугмаси.", reply_markup=kb)


def make_bot():
    from telegram import Update
    from telegram.ext import Application, CallbackQueryHandler, CommandHandler, MessageHandler, filters
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler(["start", "help"], start_cmd))
    app.add_handler(CommandHandler("myid", B.myid))
    app.add_handler(CommandHandler("bekor", B.bekor))
    app.add_handler(CommandHandler("holat", B.holat))
    app.add_handler(CommandHandler("reestr", B.reestr))
    app.add_handler(CommandHandler("transh", B.transh))
    app.add_handler(CommandHandler("tozalash", B.tozalash))
    app.add_handler(CommandHandler("admin", B.admin))
    app.add_handler(CallbackQueryHandler(B.admin_cb, pattern="^ad_"))
    app.add_handler(CommandHandler("vozvrat", B.vozvrat))
    app.add_handler(CommandHandler("xatolar", B.xatolar_cmd))
    app.add_handler(CallbackQueryHandler(B.tozalash_cb, pattern="^rs_"))
    app.add_handler(CallbackQueryHandler(B.transh_cb, pattern="^tr_"))
    app.add_handler(MessageHandler(filters.Document.ALL, B.document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, B.on_text))
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
