#!/usr/bin/env python3
"""Gavhar Gold zavod — тилло фактура нарх назорати ва транш (FIFO) ҳисоби.

Ишлатиш:
  python3 gavhar_nazorat.py --transh gavhar_transhlar.json --ledger gavhar_ledger.csv \
      --out hisobot.xlsx [yangi_faktura1.xlsx yangi_faktura2.xlsx ...]

- Янги фактура файллари (тизим экспорти: 'Товар номи', 'Товар категорияси', 'Нарх'/'Қиймати',
  'Таъминотчи фактура рақами', 'Таъминотчи фактура санаси', 'Ташкилот' устунлари) реестрга қўшилади.
  Такрорий қаторлар (бир хил ID ёки фактура+код+грамм+сумма) иккинчи марта қўшилмайди.
- Бутун реестр ҳар сафар сана бўйича тартибланиб, FIFO билан траншларга қайта тақсимланади.
- Ҳар қатор: фактура 1 гр нархи ↔ транш нарх жадвалидаги шу тур нархи. Фарқ > chegara бўлса ⚠.
- Натижа: ҳисобот xlsx + қисқа матн (stdout) + янгиланган реестр csv.
"""
import argparse, csv, json, re, sys, datetime as dt
from pathlib import Path
from collections import defaultdict, OrderedDict

LEDGER_COLS = ["id", "sana", "faktura", "filial", "kod", "nomi", "kategoriya", "soni", "gramm", "summa"]
K = 585 / 999
# Ишчилик нормаси (НДС сиз), $/гр — тош тури бўйича (Нархлар варағидан). Созламада "ishchilik" билан ўзгартириш мумкин.
ISHCHILIK = {"Бриллиант": 10.5, "Вакс": 8.5, "Кузсиз": 8.5, "Дорика": 11.0, "Фианит": 7.0, "Браслет": 11.0,
             "Детский": 11.0, "Ван Клиф": 11.0}


def narx_tarkibi(t, cfg=None):
    """Транш нархини таркибга ажратиш: лом + ишчилик + НДС = 1 гр нархи."""
    norm_ = dict(ISHCHILIK)
    norm_.update((cfg or {}).get("ishchilik") or {})
    lom = t["birja_999"] * K
    out = []
    for tur_, jami in t["narxlar"].items():
        if tur_ == "Кузсиз" and "Вакс" in t["narxlar"]:
            continue
        ish = norm_.get(tur_, 11.0)
        out.append(dict(tur=tur_, lom=round(lom, 2), ishchilik=ish, qqs=round(jami - lom - ish, 2), jami=round(jami, 2)))
    return out


# ---------------- ўқиш ----------------
def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def parse_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 30000 < v < 80000:  # Excel serial
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    s = str(v).strip()[:10]
    for f in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    m = re.fullmatch(r"(\d{2})(\d{2})(\d{4})", s)  # 24072026
    if m:
        try:
            return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def gramm_from_name(name):
    m = re.search(r"vazni\s*[-–:]?\s*([0-9]+(?:[.,][0-9]+)?)", name, re.I)
    return float(m.group(1).replace(",", ".")) if m else None


def to_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(" ", "").replace(",", "."))
    except ValueError:
        return None


GV_CAT = (("билак", "Билак узуклар"), ("билаг", "Билак узуклар"), ("браслет", "Билак узуклар"), ("bilak", "Билак узуклар"),
          ("bilag", "Билак узуклар"), ("braslet", "Билак узуклар"), ("zanjir", "Занжирлар"), ("узук", "Узуклар"), ("зирак", "Зираклар"),
          ("занжир", "Занжирлар"), ("кулон", "Кулонлар"), ("подвеска", "Кулонлар"))
RU_MONTH = {"январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма": 5, "июн": 6, "июл": 7, "август": 8,
            "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12}


def read_gavhar_sheet(ws, warnings):
    """Gavhar Jewel Group нинг ўз фактураси (INVOICE №GVR-..., «Нархи 1 грамм учун» устуни)."""
    rows = list(ws.iter_rows(values_only=True))
    hi = None
    for i, r in enumerate(rows[:20]):
        if any("нархи 1 грамм" in norm(c).lower() for c in r):
            hi = i
            break
    if hi is None:
        return None
    hdr = [norm(c).lower() for c in rows[hi]]
    sub = [norm(c).lower() for c in rows[hi + 1]] if hi + 1 < len(rows) else []

    def col(key, where):
        for j, c in enumerate(where):
            if key in c:
                return j
        return None

    c_name = col("tovar nomi", hdr)
    c_vazni = col("vazni", sub)
    c_tosh = col("toshi", sub)
    c_proba = col("proba", sub)
    c_art = col("артикул", hdr)
    c_narx = col("нархи 1 грамм", hdr)
    c_jami = col("умумий нарх", hdr)
    head = " ".join(norm(c) for r in rows[:hi] for c in r if c)
    m = re.search(r"№\s*([A-Za-zА-Яа-я0-9][\w\-/]*)", head)
    faktura = m.group(1) if m else ""
    sana = None
    m = re.search(r"от\s+(\d{1,2})\s+([А-Яа-я]+)\s+(\d{4})", head)
    if m:
        mon = next((v for k, v in sorted(RU_MONTH.items(), key=lambda kv: -len(kv[0])) if m.group(2).lower().startswith(k)), None)
        if mon:
            sana = dt.date(int(m.group(3)), mon, int(m.group(1)))
    if not sana:
        m = re.search(r"(\d{8})", faktura)
        sana = parse_date(m.group(1)) if m else None
    m = re.search(r"\(([^)]+)\)", head)
    filial = (m.group(1).strip() if m else ws.title.strip())
    out, sum_g, sum_s, jami = [], 0.0, 0.0, None
    for r in rows[hi + 2:]:
        g = lambda j: (r[j] if j is not None and j < len(r) else None)
        name = norm(g(c_name))
        if not name:
            continue
        if "жами" in name.lower() or "итого" in name.lower():
            jami = (to_num(g(c_vazni)), to_num(g(c_jami)))
            continue
        vazni, narx, summa = to_num(g(c_vazni)), to_num(g(c_narx)), to_num(g(c_jami))
        if not vazni and not summa:
            continue
        if summa is None and vazni and narx:
            summa = vazni * narx
        if vazni and narx and summa is not None and abs(vazni * narx - summa) > 0.1:
            warnings.append(f"{name}: {vazni} гр × {narx} = {vazni * narx:,.2f}, фактурада {summa:,.2f}")
        stone = norm(g(c_tosh))
        first = name.split()[0].lower() if name.split() else ""
        kat = next((v for k, v in GV_CAT if first.startswith(k)), first.capitalize())
        nomi = f"{name} Proba - {norm(g(c_proba)) or '585'}" + (f" Tosh turi - {stone}" if stone else "") + f" Vazni - {vazni}"
        art = norm(g(c_art))
        out.append(dict(id=("GV" + art) if art else "", sana=sana.isoformat() if sana else "", faktura=faktura, filial=filial,
                        kod=art or name.split()[-1], nomi=nomi, kategoriya=kat, soni=1, gramm=vazni or "",
                        summa=summa if summa is not None else ""))
        sum_g += vazni or 0
        sum_s += summa or 0
    if jami:
        if jami[0] is not None and abs(jami[0] - sum_g) > 0.01:
            warnings.append(f"ЖАМИ грамм {jami[0]}, қаторлар йиғиндиси {sum_g:.2f}")
        if jami[1] is not None and abs(jami[1] - sum_s) > 0.5:
            warnings.append(f"ЖАМИ сумма {jami[1]:,.2f}, қаторлар йиғиндиси {sum_s:,.2f}")
    if not sana:
        warnings.append(f"{ws.title}: фактура санаси топилмади")
    return out


def read_invoice(path, warnings=None):
    import openpyxl
    if warnings is None:
        warnings = []
    wb = openpyxl.load_workbook(path, data_only=True)
    rows_out = []
    for ws in wb.worksheets:
        gvr = read_gavhar_sheet(ws, warnings)
        if gvr is not None:
            rows_out.extend(gvr)
            continue
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, r in enumerate(rows[:15]):
            cells = [norm(c).lower() for c in r]
            if "товар номи" in cells and "товар категорияси" in cells and ("нарх" in cells or "қиймати" in cells):
                hdr_i = i
                break
        if hdr_i is None:
            continue
        hdr = [norm(c).lower() for c in rows[hdr_i]]

        def col(*names):
            for n in names:
                if n in hdr:
                    return hdr.index(n)
            return None

        c = dict(
            id=col("id"), kod=col("товар коди"), nomi=col("товар номи"), kat=col("товар категорияси"),
            soni=col("сони"), narx=col("нарх"), qiymat=col("қиймати", "қиймат"), filial=col("ташкилот", "филиал"),
            taminotchi=col("таъминотчи"), faktura=col("таъминотчи фактура рақами", "фактура рақами"),
            fsana=col("таъминотчи фактура санаси", "фактура санаси"), asana=col("амалиёт санаси"),
            gramm=col("грамм"),
        )
        for r in rows[hdr_i + 1:]:
            g = lambda k: (r[c[k]] if c[k] is not None and c[k] < len(r) else None)
            nomi = norm(g("nomi"))
            if not nomi:
                continue
            tam = norm(g("taminotchi"))
            if tam and "gavhar" not in tam.lower():
                continue
            soni = to_num(g("soni")) or 1
            gr = to_num(g("gramm")) or gramm_from_name(nomi)
            summa = to_num(g("qiymat"))
            if summa is None and to_num(g("narx")) is not None:
                summa = to_num(g("narx")) * soni
            sana = parse_date(g("fsana")) or parse_date(g("asana"))
            rows_out.append(dict(
                id=norm(g("id")), sana=sana.isoformat() if sana else "", faktura=norm(g("faktura")),
                filial=norm(g("filial")), kod=norm(g("kod")), nomi=nomi, kategoriya=norm(g("kat")),
                soni=soni, gramm=(gr * soni) if gr else "", summa=summa if summa is not None else "",
            ))
    return rows_out


def load_ledger(path):
    try:
        with open(path, encoding="utf-8") as f:
            return list(csv.DictReader(f))
    except FileNotFoundError:
        return []


def save_ledger(path, rows):
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=LEDGER_COLS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in LEDGER_COLS})


def key(r):
    if r.get("id"):
        return ("id", str(r["id"]))
    return ("x", r.get("faktura"), r.get("kod"), str(r.get("gramm")), str(r.get("summa")))


# ---------------- возврат ----------------
RET_COLS = ["id", "sana", "filial", "kod", "model", "nomi", "gramm", "summa", "hujjat"]
MODEL_RE = re.compile(r"\b([A-Z]{2,4}\d{5,7})\b")


def model_of(name):
    m = MODEL_RE.search(str(name or "").upper())
    return m.group(1) if m else ""


def read_returns(path):
    """Возврат рўйхати (Excel) — устунлар номи бўйича эркин аниқланади."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    keys = dict(
        kod=("артикул", "артикл", "товар коди", "товар код", "artikul", "kod", "код"),
        model=("модель", "model"),
        segment=("сегмент", "тош", "камень", "toshi"),
        sana=("сана", "дата", "sana", "date"),
        filial=("филиал", "дўкон", "дукон", "ташкилот", "магазин", "filial", "do'kon", "dokon"),
        gramm=("грамм", "вес", "vazni", "gramm", "вазни"),
        narx1=("1 гр", "1гр", "грамм учун", "за грамм"),
        summa=("сумма", "қиймат", "киймат", "стоим", "summa", "qiymat", "нарх", "narx", "цена"),
        hujjat=("ҳужжат", "хужжат", "документ", "фактура", "hujjat", "накладная"),
        nomi=("номи", "наимен", "nomi", "товар", "tovar", "mahsulot"),
    )
    out = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hi, col = None, {}
        for i, r in enumerate(rows[:25]):
            cells = [norm(c).lower() for c in r]
            found, usedc = {}, set()
            for k, subs in keys.items():
                for j, c in enumerate(cells):
                    if j in usedc or not c:
                        continue
                    if any(sb in c for sb in subs):
                        found[k] = j
                        usedc.add(j)
                        break
            if ("nomi" in found or "kod" in found) and ("summa" in found or "narx1" in found):
                hi, col = i, found
                break
        if hi is None:
            continue
        for r in rows[hi + 1:]:
            g = lambda k: (r[col[k]] if k in col and col[k] < len(r) else None)
            nomi = norm(g("nomi"))
            kod = norm(g("kod"))
            mdl = norm(g("model"))
            if not nomi and not kod:
                continue
            if any(w in (nomi + kod).lower() for w in ("жами", "итого", "jami", "всего")):
                continue
            if mdl and mdl not in nomi:
                nomi = f"{nomi} {mdl}".strip()
            if norm(g("segment")):
                nomi = f"{nomi} ({norm(g('segment'))})"
            if False:
                continue
            gr = to_num(g("gramm")) or gramm_from_name(nomi)
            summa = to_num(g("summa"))
            n1 = to_num(g("narx1"))
            if n1 and gr and (summa is None or "summa" not in col):
                summa = n1 * gr
            if not gr or not summa:
                continue
            d = parse_date(g("sana"))
            out.append(dict(id="", sana=d.isoformat() if d else "", filial=norm(g("filial")) or ws.title.strip(), kod=kod,
                            model=(mdl.upper() if mdl else "") or model_of(nomi) or model_of(kod), nomi=nomi, gramm=gr, summa=summa, hujjat=norm(g("hujjat"))))
    return out


def load_returns(path):
    if not path:
        return []
    try:
        with open(path, encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        return []
    for r in rows:
        r["gramm"] = to_num(r["gramm"]) or 0.0
        r["summa"] = to_num(r["summa"]) or 0.0
    return rows


def save_returns(path, rows):
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=RET_COLS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in RET_COLS})


def add_returns(path, new_rows):
    old = load_returns(path)
    seen = {(r["kod"], r["model"], r["sana"], round(float(r["gramm"]), 2), round(float(r["summa"]), 2)) for r in old}
    added = 0
    for r in new_rows:
        k = (r["kod"], r["model"], r["sana"], round(float(r["gramm"]), 2), round(float(r["summa"]), 2))
        if k in seen:
            continue
        seen.add(k)
        old.append(r)
        added += 1
    save_returns(path, old)
    return added, len(new_rows) - added


def find_return(r, g, returns, f1=None, tol=0.3):
    """Возврат рўйхатидан мос товар: (топилган, нархи мосми). Аввал артикул, кейин модель+грамм;
    бир нечта номзод бўлса — нархи фактурага мос келгани олинади."""
    kod, model = str(r.get("kod") or "").strip(), model_of(r.get("nomi"))
    cands = [x for x in returns if not x.get("_used") and kod and x.get("kod") and str(x["kod"]).strip() == kod]
    if not cands and model:
        cands = [x for x in returns if not x.get("_used") and x.get("model") == model and abs(float(x["gramm"]) - g) <= 0.011]
    if not cands:
        return None, False
    cands.sort(key=lambda x: abs(float(x["gramm"]) - g))  # энг яқин грамм биринчи
    if f1 is not None:
        for x in cands:
            if abs(float(x["summa"]) / float(x["gramm"]) - f1) <= tol:
                return x, True
    return cands[0], False


def returns_stats(returns):
    used = [r for r in returns if r.get("_used")]
    rest = [r for r in returns if not r.get("_used")]
    sm = lambda rs: sum(float(r["summa"]) for r in rs)
    return dict(n=len(returns), summa=sm(returns), berilgan_n=len(used), berilgan_summa=sm(used),
                qolgan_n=len(rest), qolgan_summa=sm(rest))


# ---------------- тур аниқлаш ----------------
BRASLET_RE = re.compile(r"билак|билаг|браслет|bilak|bilag|braslet|bracelet|б/узук")
ZANJIR_RE = re.compile(r"занжир|zanjir|цеп|chain")
VANCLIF = ("van cl", "vancl", "van-cl", "ван клиф", "ванклиф", "клиф", "klif", "cleef", "clef")


def tur(nomi, kat):
    """Тур аниқлаш. Билак узук (браслет) ва занжир — ТОШ ТУРИДАН ҚАТЪИ НАЗАР категория бўйича нархланади."""
    n, k = (nomi or "").lower(), (kat or "").lower()
    first = n.split()[0] if n.split() else ""
    nom = n.split(" proba")[0]  # товар номи (тош/вазн қисмисиз)
    # 1) категория бўйича (тошга қарамайди): билак узук / браслет → Браслет нархи, занжир → Дорика нархи
    if BRASLET_RE.search(k) or BRASLET_RE.search(nom):
        return "Браслет"
    if ZANJIR_RE.search(k) or ZANJIR_RE.match(first):
        return "Дорика"
    # 2) Ван Клиф
    if any(v in n for v in VANCLIF):
        return "Ван Клиф"
    # 3) тош тури бўйича
    if "brilliant" in n or "brilyant" in n or "бриллиант" in n:
        return "Бриллиант"
    if "fianit" in n or "фианит" in n:
        return "Фианит"
    if "waks" in n or "vaks" in n or "вакс" in n:
        return "Вакс"
    m = re.search(r"tosh\s*turi\s*[-–:]?\s*([A-Za-zА-Яа-яЎўҚқҒғҲҳ']*)", nomi, re.I)
    if not m or not m.group(1) or m.group(1).lower().startswith(("o'lcham", "vazni", "olcham")):
        return "Детский"
    return "❓ " + m.group(1)


# ---------------- асосий ҳисоб ----------------
def compute(cfg, ledger, returns=None):
    returns = returns or []
    for ret in returns:
        ret.pop("_used", None)
    chegara = cfg.get("chegara_usd_gr", 0.3)
    tr = sorted(cfg["transhlar"], key=lambda t: t["sana"])
    left = [float(t["summa_usd"]) for t in tr]
    used = [0.0] * len(tr)
    first_date = [None] * len(tr)
    last_date = [None] * len(tr)

    order = sorted(range(len(ledger)), key=lambda i: (ledger[i]["sana"] or "9999", ledger[i]["faktura"], i))
    out = []
    for i in order:
        r = dict(ledger[i])
        r["tur"] = tur(r["nomi"], r["kategoriya"])
        g = to_num(r["gramm"]) or 0.0
        s = to_num(r["summa"])
        r["gramm"] = g
        r["summa"] = s
        r["izoh"] = ""
        if g <= 0:
            r.update(holat="⚠ грамм йўқ", hisob=None, hisob_1gr=None, fakt_1gr=None, farq_gr=None, farq=None, transh="")
            out.append(r)
            continue
        if not tr:
            r.update(holat="⚠ транш йўқ", hisob=None, hisob_1gr=None, fakt_1gr=None, farq_gr=None, farq=None,
                     transh="", izoh="Аввал транш (пул ва нархлар) киритинг")
            out.append(r)
            continue
        if r["tur"].startswith("❓"):
            r["izoh"] = "Тур аниқланмади — Детский нормаси олинди"
        def alloc(commit):
            lf = left if commit else list(left)
            rem_g, cost, parts = g, 0.0, []
            for ti, t in enumerate(tr):
                if rem_g <= 1e-9:
                    break
                if lf[ti] <= 1e-6:
                    continue
                ys, yq = t.get("yopilgan_sana"), t.get("yopilgan_qator")
                if yq is not None:  # ёпилгандан кейин юкланган фактуралар фақат жорий транш нархида
                    if i >= yq:
                        continue
                elif ys and r["sana"] and r["sana"] >= ys:  # эски усул: шу санадан ёпилган
                    continue
                p = t["narxlar"].get(r["tur"]) or t["narxlar"].get("Детский" if r["tur"].startswith("❓") else "Дорика")
                take = min(rem_g, lf[ti] / p)
                c = take * p
                lf[ti] -= c
                cost += c
                rem_g -= take
                parts.append((t["no"], take, p))
                if commit:
                    used[ti] += c
                    first_date[ti] = first_date[ti] or r["sana"]
                    last_date[ti] = r["sana"]
            if rem_g > 1e-9:  # пул етмади — охирги транш нархида
                t = tr[-1]
                p = t["narxlar"].get(r["tur"]) or t["narxlar"].get("Дорика")
                cost += rem_g * p
                parts.append(("пул етмади", rem_g, p))
            return cost, parts

        cost, parts = alloc(False)
        # нарх транш бўйича тўғри келмаса — возврат рўйхатидан қидирамиз
        if s is not None and returns and abs(s / g - cost / g) > chegara:
            ret, price_ok = find_return(r, g, returns, s / g, chegara)
            if ret is not None and price_ok:
                ret["_used"] = True
                ret["_faktura"] = r.get("faktura")
                ret["_filial2"] = r.get("filial")
                ret["_sana2"] = r.get("sana")
                r1 = float(ret["summa"]) / float(ret["gramm"])
                exp = r1 * g
                r.update(transh="возврат", hisob=exp, hisob_1gr=r1, fakt_1gr=s / g, farq_gr=s / g - r1, farq=s - exp,
                         holat="✓ возврат товар",
                         izoh=(f"Возврат товар: {ret.get('filial') or '—'} дан {_dmy(ret.get('sana'))} да {r1:.2f} $ дан қайтарилган — "
                               f"шу нархда берилди, тўғри."))
                out.append(r)
                continue
            if ret is not None:
                r["_ret_hint"] = (f"Возвратда ўхшаш товар бор ({ret.get('filial')}, {_dmy(ret.get('sana'))}), лекин у "
                                  f"{float(ret['summa']) / float(ret['gramm']):.2f} $ дан қайтарилган — бу ҳам мос эмас.")
        cost, parts = alloc(True)
        r["transh"] = " + ".join(f"№{n}: {gg:.2f} гр × {p:.2f}" for n, gg, p in parts) if len(parts) > 1 else f"№{parts[0][0]}"
        r["hisob"] = cost
        r["hisob_1gr"] = cost / g
        if s is None:
            r.update(holat="⚠ фактура нархи йўқ", fakt_1gr=None, farq_gr=None, farq=None)
        else:
            r["fakt_1gr"] = s / g
            r["farq_gr"] = r["fakt_1gr"] - r["hisob_1gr"]
            r["farq"] = s - cost
            if abs(r["farq_gr"]) <= chegara:
                r["holat"] = "✓ тўғри"
            else:
                r["holat"] = "⚠ ҚИММАТ" if r["farq_gr"] > 0 else "⚠ арзон"
                r["izoh"] = sabab(r, tr, parts)
                if r.get("_ret_hint"):
                    r["izoh"] += " " + r.pop("_ret_hint")
        if any(n == "пул етмади" for n, _, _ in parts):
            r["holat"] = r.get("holat", "") + " | пул етмади"
        out.append(r)

    trs = []
    for ti, t in enumerate(tr):
        if left[ti] <= 1:
            holat = "Тугаган"
        elif t.get("yopilgan_sana"):
            holat = f"Ёпилган {_dmy(t['yopilgan_sana'])}"
        else:
            holat = "Фаол" if used[ti] > 0 else "Ишлатилмаган"
        trs.append(dict(no=t["no"], sana=t["sana"], summa=t["summa_usd"], birja=t["birja_999"],
                        used=used[ti], left=left[ti], holat=holat, birinchi=first_date[ti], oxirgi=last_date[ti]))
    return out, trs


def _dmy(iso):
    return f"{iso[8:10]}.{iso[5:7]}.{iso[:4]}" if iso and len(iso) >= 10 else (iso or "")


def sabab(r, tr, parts):
    """Хатонинг содда ва тушунарли сабаби (бир жумла)."""
    f, kerak = r["fakt_1gr"], r["hisob_1gr"]
    tno = parts[0][0]
    t = next((x for x in tr if x["no"] == tno), tr[-1])
    for name, p in t["narxlar"].items():  # бошқа тур нархи билан ёзилганми
        if name != r["tur"] and name != "Кузсиз" and abs(f - p) <= 0.05:
            return f"Gavhar «{name}» нархини қўйган. Бу товар «{r['tur']}» — нархи {kerak:.2f} $ бўлиши керак."
    for x in tr:  # бошқа транш нархими
        if x["no"] != tno:
            p = x["narxlar"].get(r["tur"])
            if p and abs(f - p) <= 0.05:
                return (f"Бошқа пул ({_dmy(x['sana'])}) нархи қўйилган. Ҳозир {_dmy(t['sana'])} даги пул ишлатиляпти — "
                        f"нархи {kerak:.2f} $.")
    labor = kerak - t["birja_999"] * K
    impl = (f - labor) / K
    if abs(impl - t["birja_999"]) > 0.3:
        return (f"Gavhar нархни бошқа биржа нархи (≈{impl:.1f}) билан ҳисоблаган. "
                f"Пулингиз {_dmy(t['sana'])} да {t['birja_999']:g} биржада берилган — шунинг учун нархи {kerak:.2f} $.")
    return f"Фактурада {f:.2f} $, тўғри нарх {kerak:.2f} $."


# ---------------- ҳисобот ----------------
def stats(rows, trs):
    """Умумий кўрсаткичлар (ҳисобот, бот ва mini app учун)."""
    ok = [r for r in rows if r.get("fakt_1gr") is not None]
    bad = [r for r in rows if str(r.get("holat", "")).startswith("⚠")]
    tot_given = sum(t["summa"] for t in trs)
    tot_used = sum(t["used"] for t in trs)
    fakt_sum = sum(r["summa"] for r in ok)
    hisob_sum = sum(r["hisob"] for r in ok)
    qimmat = sum(r["farq"] for r in ok if r["holat"].startswith("⚠ ҚИММАТ"))
    arzon = sum(r["farq"] for r in ok if r["holat"].startswith("⚠ арзон"))
    active = next((t for t in trs if t["holat"] == "Фаол"), None) or next((t for t in trs if t["holat"] == "Ишлатилмаган"), None)
    left_total = sum(t["left"] for t in trs if not str(t["holat"]).startswith("Ёпилган"))
    # темп: охирги 30 кун
    by_day = defaultdict(float)
    for r in ok:
        if r["sana"] and r.get("transh") != "возврат":
            by_day[r["sana"]] += r["hisob"]
    tugash = None
    temp = None
    if by_day:
        last = max(by_day)
        d_last = dt.date.fromisoformat(last)
        win = [d for d in by_day if (d_last - dt.date.fromisoformat(d)).days < 30]
        span = min(30, (d_last - dt.date.fromisoformat(min(by_day))).days + 1)  # маълумот 30 кундан қисқа бўлса — бор кунлар
        temp = sum(by_day[d] for d in win) / max(span, 1)
        if temp > 0:
            tugash = d_last + dt.timedelta(days=left_total / temp)

    return dict(ok=ok, bad=bad, tot_given=tot_given, tot_used=tot_used, fakt_sum=fakt_sum, hisob_sum=hisob_sum,
                qimmat=qimmat, arzon=arzon, active=active, left_total=left_total, by_day=by_day, tugash=tugash, temp=temp)


def write_report(path, cfg, rows, trs, returns=None):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    F = "Arial"
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    red = PatternFill("solid", fgColor="F8CBAD")
    green = PatternFill("solid", fgColor="E2EFDA")
    wb = openpyxl.Workbook()

    def table(ws, start, headers, data, widths=None, fmt=None):
        for j, h in enumerate(headers, 1):
            c = ws.cell(start, j, h)
            c.font = Font(name=F, bold=True, color="FFFFFF")
            c.fill = hdr_fill
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for i, d in enumerate(data, start + 1):
            for j, v in enumerate(d, 1):
                c = ws.cell(i, j, v)
                c.font = Font(name=F)
                if fmt and fmt.get(j):
                    c.number_format = fmt[j]
        if widths:
            for j, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(j)].width = w
        ws.freeze_panes = ws.cell(start + 1, 1)

    st = stats(rows, trs)
    ok, bad, tot_given, tot_used = st["ok"], st["bad"], st["tot_given"], st["tot_used"]
    fakt_sum, hisob_sum, qimmat, arzon = st["fakt_sum"], st["hisob_sum"], st["qimmat"], st["arzon"]
    active, left_total, tugash, temp = st["active"], st["left_total"], st["tugash"], st["temp"]

    ws = wb.active
    ws.title = "Назорат"
    ws["A1"] = "GAVHAR — ТИЛЛО НАРХ ВА ПУЛ НАЗОРАТИ"
    ws["A1"].font = Font(name=F, bold=True, size=14)
    ws["A2"] = f"Ҳисобот санаси: {dt.date.today():%d.%m.%Y}. Чегара: ±{cfg.get('chegara_usd_gr',0.3)} $/гр. FIFO: олдинги транш тугамагунча кейингиси ишлатилмайди."
    lines = [
        ("ПУЛ", None, None),
        ("Жами берилган пул, $", tot_given, "#,##0"),
        ("Жами олинган товар (ҳисоб бўйича), $", tot_used, "#,##0"),
        ("Таъминотчидаги қолдиқ, $", left_total, "#,##0"),
        ("Фаол транш", f"№{active['no']} ({active['sana']}, биржа {active['birja']})" if active else "—", None),
        ("Қолдиқ, грамм 999 экв. (фаол транш биржа нархида)", (left_total / active["birja"]) if active else None, "#,##0.0"),
        ("Ўртача кунлик сарф (охирги 30 кун), $", temp, "#,##0"),
        ("Таҳминий тугаш санаси", tugash.strftime("%d.%m.%Y") if tugash else "—", None),
        ("НАРХ НАЗОРАТИ", None, None),
        ("Фактура суммаси, $", fakt_sum, "#,##0.00"),
        ("Ҳисоб суммаси, $", hisob_sum, "#,##0.00"),
        ("Фарқ (фактура − ҳисоб), $", fakt_sum - hisob_sum, "#,##0.00"),
        ("Шундан ҚИММАТ ёзилган (бизга зарар), $", qimmat, "#,##0.00"),
        ("Шундан арзон ёзилган, $", arzon, "#,##0.00"),
        ("Жами позиция", len(rows), "#,##0"),
        ("Тўғри позиция", sum(1 for r in rows if r.get("holat", "").startswith("✓")), "#,##0"),
        ("Фарқли / муаммоли позиция", len(bad), "#,##0"),
        ("Сана оралиғи", f"{min((r['sana'] for r in rows if r['sana']), default='')} — {max((r['sana'] for r in rows if r['sana']), default='')}", None),
        ("Филиаллар сони", len({r['filial'] for r in rows}), "#,##0"),
    ]
    for i, (a, b, f) in enumerate(lines, 4):
        ws.cell(i, 1, a).font = Font(name=F, bold=b is None and f is None and a.isupper())
        if b is not None:
            c = ws.cell(i, 2, b)
            c.font = Font(name=F, bold=True)
            if f:
                c.number_format = f
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 28

    ws = wb.create_sheet("Транш")
    table(ws, 1, ["№", "Сана", "Сумма, $", "Биржа 999, $/гр", "Ишлатилган, $", "Қолдиқ, $", "Ишлатилган, %", "Ҳолати", "Биринчи товар", "Охирги товар"],
          [[t["no"], t["sana"], t["summa"], t["birja"], t["used"], t["left"], (t["used"] / t["summa"]) if t["summa"] else None,
            t["holat"], t["birinchi"], t["oxirgi"]] for t in trs],
          [6, 12, 14, 14, 14, 14, 12, 16, 14, 14], {3: "#,##0", 5: "#,##0", 6: "#,##0", 7: "0.0%"})

    ws = wb.create_sheet("Хатоликлар")
    bad_s = sorted(bad, key=lambda r: -abs(r.get("farq") or 0))
    table(ws, 1, ["№", "Сана", "Фактура №", "Филиал", "Товар коди", "Тур", "Грамм", "Фактура $/гр", "Керак $/гр", "Фарқ $/гр", "Фарқ $", "Ҳолати", "Сабаб", "Товар номи"],
          [[i, r["sana"], r["faktura"], r["filial"], r["kod"], r["tur"], r["gramm"], r.get("fakt_1gr"), r.get("hisob_1gr"), r.get("farq_gr"),
            r.get("farq"), r["holat"], r["izoh"], r["nomi"]] for i, r in enumerate(bad_s, 1)],
          [5, 11, 11, 20, 14, 11, 8, 11, 11, 10, 10, 14, 60, 60], {7: "0.00", 8: "0.00", 9: "0.00", 10: "0.00", 11: "#,##0.00"})
    for i, r in enumerate(bad_s, 2):
        ws.cell(i, 12).fill = red

    ws = wb.create_sheet("Приход")
    table(ws, 1, ["Сана", "Фактура №", "Филиал", "Товар коди", "Товар номи", "Категория", "Тур", "Грамм", "Фактура сумма, $", "Фактура $/гр",
                  "Транш", "Ҳисоб сумма, $", "Ҳисоб $/гр", "Фарқ $/гр", "Фарқ $", "Ҳолати", "Изоҳ"],
          [[r["sana"], r["faktura"], r["filial"], r["kod"], r["nomi"], r["kategoriya"], r["tur"], r["gramm"], r["summa"], r.get("fakt_1gr"),
            r.get("transh"), r.get("hisob"), r.get("hisob_1gr"), r.get("farq_gr"), r.get("farq"), r.get("holat"), r.get("izoh")] for r in rows],
          [11, 11, 20, 14, 55, 14, 11, 8, 12, 10, 10, 12, 10, 9, 9, 14, 50],
          {8: "0.00", 9: "#,##0.00", 10: "0.00", 12: "#,##0.00", 13: "0.00", 14: "0.00", 15: "#,##0.00"})
    for i, r in enumerate(rows, 2):
        ws.cell(i, 16).fill = green if str(r.get("holat")).startswith("✓") else red
    ws.auto_filter.ref = f"A1:Q{len(rows)+1}"

    ws = wb.create_sheet("Кунлик")
    agg = OrderedDict()
    for r in sorted(ok, key=lambda r: r["sana"]):
        a = agg.setdefault(r["sana"], [0, 0.0, 0.0, 0.0, 0])
        a[0] += 1; a[1] += r["gramm"]; a[2] += r["summa"]; a[3] += r["hisob"]
        a[4] += 1 if r["holat"].startswith("⚠") else 0
    table(ws, 1, ["Сана", "Позиция", "Грамм", "Фактура, $", "Ҳисоб, $", "Фарқ, $", "Фарқли позиция"],
          [[d, a[0], a[1], a[2], a[3], a[2] - a[3], a[4]] for d, a in agg.items()],
          [12, 10, 10, 14, 14, 12, 14], {3: "#,##0.00", 4: "#,##0.00", 5: "#,##0.00", 6: "#,##0.00"})

    ws = wb.create_sheet("Нархлар")
    types = []
    for t in cfg["transhlar"]:
        for n in t["narxlar"]:
            if n not in types:
                types.append(n)
    table(ws, 1, ["Тур"] + [f"Транш №{t['no']} ({t['sana']}, биржа {t['birja_999']})" for t in cfg["transhlar"]],
          [[n] + [t["narxlar"].get(n) for t in cfg["transhlar"]] for n in types], [14] + [26] * len(cfg["transhlar"]),
          {j: "0.00" for j in range(2, 2 + len(cfg["transhlar"]))})
    r0 = len(types) + 3
    ws.cell(r0, 1, "Тур аниқлаш тартиби: Билак узуклар → Браслет; номида Van Cleef/Клиф → Ван Клиф; brilliant → Бриллиант; fianit → Фианит; "
                   "Занжирлар → Дорика; waks → Вакс; Tosh turi йўқ/бўш → Детский.").font = Font(name=F, italic=True)
    if returns:
        ws = wb.create_sheet("Возврат")
        table(ws, 1, ["Возврат санаси", "Қайтарган дўкон", "Артикул / код", "Модель", "Товар номи", "Грамм", "Возврат сумма, $", "1 гр, $",
                      "Ҳолати", "Қайта берилган дўкон", "Сана", "Фактура №"],
              [[r.get("sana"), r.get("filial"), r.get("kod"), r.get("model"), r.get("nomi"), float(r["gramm"]), float(r["summa"]),
                float(r["summa"]) / float(r["gramm"]) if float(r["gramm"]) else None,
                "Қайта берилган" if r.get("_used") else "Gavhar да (берилмаган)", r.get("_filial2"), r.get("_sana2"), r.get("_faktura")]
               for r in returns],
              [12, 18, 14, 12, 45, 8, 13, 9, 20, 18, 11, 14], {6: "0.00", 7: "#,##0.00", 8: "0.00"})
    wb.save(path)
    return dict(tot_given=tot_given, tot_used=tot_used, left=left_total, fakt=fakt_sum, hisob=hisob_sum, qimmat=qimmat, arzon=arzon,
                bad=len(bad), n=len(rows), active=active, tugash=tugash, temp=temp, bad_s=bad_s)


def vozvrat_holati(r, returns_by_model):
    """Хато товар возврат рўйхатида борми — изоҳ учун."""
    m = model_of(r.get("nomi"))
    c = returns_by_model.get(m, []) if m else []
    g = float(r.get("gramm") or 0)
    if not c:
        return "Возвратда йўқ", None
    e = [x for x in c if abs(float(x["gramm"]) - g) <= 0.05]
    if e:
        x = min(e, key=lambda x: abs(float(x["gramm"]) - g))
        p = float(x["summa"]) / float(x["gramm"])
        return f"Модель+грамм бор: {x.get('filial')} {x.get('sana')}, возврат {p:.2f} $/гр", p
    return f"Фақат модели бор (бошқа грамм: {', '.join(sorted({str(round(float(x['gramm']), 2)) for x in c})[:4])})", None


def write_errors_report(path, rows, title="", returns=None):
    """Фақат хатоликлар (фарқли позициялар) — Excel. Gavhar га эътироз учун."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    F = "Arial"
    bym = {}
    for x in returns or []:
        bym.setdefault(x.get("model"), []).append(x)
    bad = sorted([r for r in rows if str(r.get("holat", "")).startswith("⚠")],
                 key=lambda r: (r.get("sana") or "", r.get("faktura") or "", -abs(r.get("farq") or 0)))
    wb = openpyxl.Workbook()
    hdr_fill = PatternFill("solid", fgColor="1F3864")

    def header(ws, row, cols, widths):
        for j, h in enumerate(cols, 1):
            c = ws.cell(row, j, h)
            c.font = Font(name=F, bold=True, color="FFFFFF")
            c.fill = hdr_fill
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for j, wd in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = wd

    # 1) хулоса фактура бўйича
    ws = wb.active
    ws.title = "Фактура бўйича"
    ws["A1"] = "ХАТОЛИКЛАР — Gavhar га эътироз учун"
    ws["A1"].font = Font(name=F, bold=True, size=14)
    ws["A2"] = title or "Барча хатоликлар"
    ws["A2"].font = Font(name=F, italic=True)
    inv = OrderedDict()
    for r in bad:
        a = inv.setdefault((r.get("faktura"), r.get("filial"), r.get("sana")), [0, 0.0, 0.0, 0.0, 0.0])
        a[0] += 1; a[1] += r.get("gramm") or 0; a[2] += r.get("summa") or 0; a[3] += r.get("hisob") or 0; a[4] += r.get("farq") or 0
    header(ws, 4, ["Сана", "Фактура №", "Дўкон", "Хато позиция", "Грамм", "Фактура $", "Керак $", "Ортиқча $"], [12, 16, 22, 12, 10, 14, 14, 14])
    i = 4
    for (fno, fil, sana), a in inv.items():
        i += 1
        for j, v in enumerate([sana, fno, fil, a[0], round(a[1], 2), a[2], a[3], a[4]], 1):
            c = ws.cell(i, j, v)
            c.font = Font(name=F)
            if j >= 6:
                c.number_format = "#,##0.00"
    i += 1
    tot = [sum(a[k] for a in inv.values()) for k in range(5)]
    for j, v in enumerate(["ЖАМИ", "", f"{len(inv)} фактура", tot[0], round(tot[1], 2), tot[2], tot[3], tot[4]], 1):
        c = ws.cell(i, j, v)
        c.font = Font(name=F, bold=True)
        if j >= 6:
            c.number_format = "#,##0.00"
    ws.freeze_panes = "A5"

    # 2) ҳар бир позиция
    ws = wb.create_sheet("Хатоликлар")
    cols = ["№", "Сана", "Фактура №", "Дўкон", "Артикул", "Модель", "Тур", "Грамм", "Фактура $/гр", "Керак $/гр", "Фарқ $/гр",
            "Фактура $", "Керак $", "Фарқ $", "Ҳолати", "Сабаб", "Возврат рўйхатида", "Возврат нархида фарқ $"]
    header(ws, 1, cols, [5, 11, 15, 18, 11, 12, 11, 8, 11, 11, 10, 11, 11, 10, 14, 70, 45, 12])
    red, amb = PatternFill("solid", fgColor="F8CBAD"), PatternFill("solid", fgColor="FCE4D6")
    for n, r in enumerate(bad, 1):
        vh, vp = vozvrat_holati(r, bym)
        vfarq = (r["summa"] - vp * r["gramm"]) if (vp and r.get("summa") is not None) else None
        vals = [n, r.get("sana"), r.get("faktura"), r.get("filial"), r.get("kod"), model_of(r.get("nomi")), r.get("tur"), r.get("gramm"),
                r.get("fakt_1gr"), r.get("hisob_1gr"), r.get("farq_gr"), r.get("summa"), r.get("hisob"), r.get("farq"), r.get("holat"),
                r.get("izoh"), vh, vfarq]
        for j, v in enumerate(vals, 1):
            c = ws.cell(n + 1, j, v)
            c.font = Font(name=F)
            if j in (8, 9, 10, 11):
                c.number_format = "0.00"
            if j in (12, 13, 14, 18):
                c.number_format = "#,##0.00"
        ws.cell(n + 1, 15).fill = red if (r.get("farq") or 0) > 0 else amb
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:R{len(bad) + 1}"
    wb.save(path)
    return dict(n=len(bad), farq=sum(r.get("farq") or 0 for r in bad), faktura=len(inv))


def write_invoice_report(path, fno, fil, sana_t, rs, trn):
    """Битта фактура (дўкон) бўйича текширув Excel и."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    F = "Arial"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Текширув"
    nb = [r for r in rs if r["holat"].startswith("⚠")]
    fakt = sum((r["summa"] or 0) for r in rs)
    hisob = sum((r["hisob"] or 0) for r in rs if r.get("hisob") is not None)
    farq = sum((r.get("farq") or 0) for r in nb)
    head = [("Фактура", fno), ("Дўкон (филиал)", fil), ("Сана", sana_t), ("Транш", ", ".join(trn) or "—"),
            ("Позиция", len(rs)), ("Грамм", round(sum(r["gramm"] for r in rs), 2)), ("Фактура суммаси, $", round(fakt, 2)),
            ("Керакли сумма, $", round(hisob, 2)), ("Фарқ, $", round(farq, 2)),
            ("ХУЛОСА", "✅ ТЎҒРИ" if not nb else f"❌ ХАТО БОР — {len(nb)} та позиция, {farq:+,.2f} $")]
    for i, (k, v) in enumerate(head, 1):
        ws.cell(i, 1, k).font = Font(name=F, bold=True)
        c = ws.cell(i, 2, v)
        c.font = Font(name=F, bold=k == "ХУЛОСА", color=("C00000" if (k == "ХУЛОСА" and nb) else "1E7B3C" if k == "ХУЛОСА" else "000000"))
    start = len(head) + 2
    hdr = ["№", "Артикул / код", "Товар номи", "Тур", "Грамм", "Фактура $/гр", "Керак $/гр", "Фарқ $/гр", "Фактура $", "Керак $", "Фарқ $", "Ҳолати", "Изоҳ"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(start, j, h)
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(wrap_text=True)
    red, green = PatternFill("solid", fgColor="F8CBAD"), PatternFill("solid", fgColor="E2EFDA")
    for i, r in enumerate(sorted(rs, key=lambda r: (not r["holat"].startswith("⚠"), -abs(r.get("farq") or 0))), 1):
        vals = [i, r["kod"], r["nomi"], r["tur"], r["gramm"], r.get("fakt_1gr"), r.get("hisob_1gr"), r.get("farq_gr"), r["summa"],
                r.get("hisob"), r.get("farq"), r["holat"], r.get("izoh")]
        for j, v in enumerate(vals, 1):
            c = ws.cell(start + i, j, v)
            c.font = Font(name=F)
            if j in (5, 6, 7, 8):
                c.number_format = "0.00"
            if j in (9, 10, 11):
                c.number_format = "#,##0.00"
        ws.cell(start + i, 12).fill = red if r["holat"].startswith("⚠") else green
    for j, w in enumerate([5, 12, 45, 11, 8, 11, 11, 10, 11, 11, 10, 16, 70], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.column_dimensions["A"].width = 18
    ws.freeze_panes = ws.cell(start + 1, 1)
    wb.save(path)


def run(transh_path, ledger_path, out_path, files=(), returns_path=None):
    """Фактураларни реестрга қўшади, ҳамма ҳисобни қайта қилади, ҳисобот ёзади. Қисқа матн қайтаради."""
    cfg = json.load(open(transh_path, encoding="utf-8"))
    ledger = load_ledger(ledger_path)
    seen = {key(r) for r in ledger}
    added, dup, new_rows = 0, 0, []
    warns = []
    for f in files:
        for r in read_invoice(f, warns):
            k = key(r)
            if k in seen:
                dup += 1
                continue
            seen.add(k)
            ledger.append(r)
            new_rows.append(r)
            added += 1
    if added:
        save_ledger(ledger_path, ledger)
    returns = load_returns(returns_path)
    rows, trs = compute(cfg, ledger, returns)
    s = write_report(out_path, cfg, rows, trs, returns)

    L = []
    new_keys = {key(r) for r in new_rows}
    new_res = [r for r in rows if key(r) in new_keys]
    if files and dup and not added:
        L.append("ℹ️ Бу фактура олдин юкланган — қайта қўшилмади.")
    elif files and dup:
        L.append(f"ℹ️ {dup} та товар олдин юкланган — улар қайта қўшилмади.")
    if new_res and not cfg["transhlar"]:
        L.append(f"🧾 Янги фактура: {len(new_res)} позиция, {sum(r['gramm'] for r in new_res):,.2f} гр, "
                 f"{sum((r['summa'] or 0) for r in new_res):,.2f} $ — реестрга қўшилди.\n"
                 "⚠ Ҳали транш (пул ва нархлар) киритилмаган, шунинг учун нарх текширилмади.\n"
                 "Gavhar нарх хабарини шу ерга forward қилинг ва суммани ёзинг — кейин фактура автомат текширилади.")
        new_res = []
    if new_res:
        inv = OrderedDict()
        for r in new_res:
            inv.setdefault((r["faktura"], r["filial"], r["sana"]), []).append(r)
        for (fno, fil, sana), rs in inv.items():
            nb = [r for r in rs if r["holat"].startswith("⚠")]
            rv = [r for r in rs if r.get("transh") == "возврат"]
            fakt = sum((r["summa"] or 0) for r in rs)
            hisob = sum((r["hisob"] or 0) for r in rs if r.get("hisob") is not None)
            farq = sum((r.get("farq") or 0) for r in nb)
            trn = sorted({str(r["transh"]).split(":")[0] for r in rs if r.get("transh") and r["transh"] != "возврат"})
            sana_t = f"{sana[8:10]}.{sana[5:7]}.{sana[:4]}" if sana else "—"
            tmap = {t["no"]: t for t in cfg["transhlar"]}
            tnos = sorted({int(x) for r in rs for x in re.findall(r"№(\d+)", str(r.get("transh") or ""))})
            tline = " + ".join(f"Транш №{n} ({_dmy(tmap[n]['sana'])}, биржа {tmap[n].get('birja_999')})" for n in tnos if n in tmap)
            L.append("━━━━━━━━━━━━━━━")
            L.append(f"🧾 {fno or '—'} · {fil or '—'} · {sana_t}")
            L.append(f"   {len(rs)} та товар · {sum(r['gramm'] for r in rs):,.2f} гр")
            L.append(f"💳 {tline} нархида текширилди" if tline else
                     ("💳 Ҳаммаси возврат товар — возврат нархида текширилди" if rv and len(rv) == len(rs) else "💳 Транш топилмади"))
            if rv:
                L.append(f"↩️ {len(rv)} та товар возвратдан — возврат нархида берилган, тўғри")
            if not nb:
                L.append(f"\n✅ ТЎҒРИ. Фактура {fakt:,.2f} $ — нархлар тўғри.")
            else:
                diff = fakt - hisob
                L.append(f"\n❌ ХАТО БОР!\n"
                         f"   Фактура: {fakt:,.2f} $\n"
                         f"   Тўғриси: {hisob:,.2f} $")
                if diff > 0:
                    L.append(f"👉 {diff:,.2f} $ ОРТИҚЧА ёзилган ({len(nb)} та товарда)")
                else:
                    L.append(f"👉 {abs(diff):,.2f} $ КАМ ёзилган ({len(nb)} та товарда)")
                grp = OrderedDict()
                for r in nb:
                    k = (r["tur"], round(r.get("hisob_1gr") or 0, 2), (r.get("farq") or 0) > 0)
                    g_ = grp.setdefault(k, {"n": 0, "g": 0.0, "f": 0.0, "s": 0.0, "izoh": r["izoh"]})
                    g_["n"] += 1; g_["g"] += r["gramm"] or 0; g_["f"] += r.get("farq") or 0; g_["s"] += r["summa"] or 0
                L.append("\nҚаерда хато:")
                items = sorted(grp.items(), key=lambda kv: -abs(kv[1]["f"]))
                for (tr_, hk, _q), g_ in items[:8]:
                    fk = g_["s"] / g_["g"] if g_["g"] else 0
                    L.append(f"• {tr_}: {g_['n']} дона, {g_['g']:.2f} гр — {fk:.2f} $ қўйилган, {hk:.2f} $ бўлиши керак → {g_['f']:+,.2f} $")
                if len(items) > 8:
                    L.append(f"• ... яна {len(items) - 8} гуруҳ (Excel файлда)")
                sab = []
                for _, g_ in items:
                    z = (g_["izoh"] or "").split(". ")[0].rstrip(".")
                    z = re.sub(r" \(≈[\d.]+\)", "", z)
                    if z and z not in sab:
                        sab.append(z)
                if sab:
                    L.append("\nСабаб: " + ". ".join(sab[:3]) + ".")
            try:
                safe = re.sub(r"[^\w\-]+", "_", f"{fno}_{fil}").strip("_")[:60] or "faktura"
                write_invoice_report(str(Path(out_path).with_name(f"Faktura_{safe}.xlsx")), fno, fil, sana_t, rs, trn)
            except Exception as e:  # ҳисобот ёзилмаса ҳам матн чиқсин
                L.append(f"   (фактура Excel ёзилмади: {e})")
    if warns:
        L.append("🧮 Фактура арифметикасида хато:\n" + "\n".join("   • " + w for w in warns[:15]))
    if new_res:
        act = next((t for t in trs if t["holat"] in ("Фаол", "Ишлатилмаган")), None)
        L.append("━━━━━━━━━━━━━━━")
        if act:
            L.append(f"💰 Жорий транш №{act['no']} ({_dmy(act['sana'])}): қолдиқ {act['left']:,.0f} $ (ишлатилди {act['used']:,.0f} $)")
        L.append(f"💰 Gavhar да жами қолдиқ: {s['left']:,.0f} $")
        L.append("Батафсил: Excel файл ёки «Назорат» mini app.")
        return "\n".join(L)
    L += holat_matn(rows, trs, s)
    return "\n".join(L)


def holat_matn(rows, trs, s):
    """Содда ҳолат матни (бот учун)."""
    L = [f"📊 ҲОЛАТ · {dt.date.today():%d.%m.%Y}", ""]
    if not trs:
        return L + ["Ҳали транш (пул) киритилмаган.", "Gavhar нарх хабарини шу ерга forward қилинг — бот суммани сўрайди."]
    L.append(f"💰 Gavhar да пулингиз: {s['left']:,.0f} $")
    act = next((t for t in trs if t["holat"] in ("Фаол", "Ишлатилмаган")), None)
    for t in trs:
        nom = f"Транш №{t['no']} ({_dmy(t['sana'])}, биржа {t['birja']})"
        if t is act:
            L.append(f"💳 Ҳозир ишлаяпти: {nom} — қолдиқ {t['left']:,.0f} $")
        elif t["holat"] == "Ишлатилмаган":
            L.append(f"   Навбатда: {nom} — {t['summa']:,.0f} $")
        else:
            L.append(f"   {nom} — {'тугаган' if t['holat'] == 'Тугаган' else 'ёпилган'}")
    inv = {(r["faktura"], r["filial"], r["sana"]) for r in rows}
    if not inv:
        L += ["", "🧾 Ҳали фактура юкланмаган."]
        return L
    fakt = sum((r["summa"] or 0) for r in rows)
    L += ["", f"🧾 Фактуралар: {len(inv)} та · {fakt:,.0f} $"]
    bad = [r for r in rows if str(r.get("holat", "")).startswith("⚠")]
    qim = sum((r.get("farq") or 0) for r in bad if (r.get("farq") or 0) > 0)
    arz = sum((r.get("farq") or 0) for r in bad if (r.get("farq") or 0) < 0)
    if not bad:
        L.append("✅ Ҳамма нархлар тўғри")
    else:
        binv = {(r["faktura"], r["filial"]) for r in bad}
        if qim:
            L.append(f"❌ Ортиқча ёзилган: {qim:,.2f} $ ({len(binv)} та фактурада)")
        if arz:
            L.append(f"⚠ Кам ёзилган: {abs(arz):,.2f} $")
        L.append("   Рўйхат: /xatolar")
    rv = sum(1 for r in rows if r.get("transh") == "возврат")
    if rv:
        L.append(f"↩️ Возвратдан берилган: {rv} та товар (возврат нархида)")
    if s.get("tugash"):
        L.append(f"⏳ Пул тахминан {s['tugash']:%d.%m.%Y} гача етади")
    return L


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--transh", required=True)
    ap.add_argument("--ledger", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--returns", default=None)
    ap.add_argument("files", nargs="*")
    a = ap.parse_args()
    print(run(a.transh, a.ledger, a.out, a.files, a.returns))


if __name__ == "__main__":
    main()
