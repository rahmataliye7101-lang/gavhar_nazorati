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
from collections import defaultdict, OrderedDict

LEDGER_COLS = ["id", "sana", "faktura", "filial", "kod", "nomi", "kategoriya", "soni", "gramm", "summa"]
K = 585 / 999


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


GV_CAT = (("билак", "Билак узуклар"), ("браслет", "Билак узуклар"), ("узук", "Узуклар"), ("зирак", "Зираклар"),
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


# ---------------- тур аниқлаш ----------------
VANCLIF = ("van cl", "vancl", "van-cl", "ван клиф", "ванклиф", "клиф", "klif", "cleef", "clef")


def tur(nomi, kat):
    n, k = nomi.lower(), kat.lower()
    if "билак" in k:
        return "Браслет"
    if any(v in n for v in VANCLIF):
        return "Ван Клиф"
    if "brilliant" in n or "brilyant" in n or "бриллиант" in n:
        return "Бриллиант"
    if "fianit" in n or "фианит" in n:
        return "Фианит"
    if "занжир" in k:
        return "Дорика"
    if "waks" in n or "vaks" in n or "вакс" in n:
        return "Вакс"
    m = re.search(r"tosh\s*turi\s*[-–:]?\s*([A-Za-zА-Яа-яЎўҚқҒғҲҳ']*)", nomi, re.I)
    if not m or not m.group(1) or m.group(1).lower().startswith(("o'lcham", "vazni", "olcham")):
        return "Детский"
    return "❓ " + m.group(1)


# ---------------- асосий ҳисоб ----------------
def compute(cfg, ledger):
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
        if r["tur"].startswith("❓"):
            r["izoh"] = "Тур аниқланмади — Детский нормаси олинди"
        rem_g, cost, parts = g, 0.0, []
        for ti, t in enumerate(tr):
            if rem_g <= 1e-9:
                break
            if left[ti] <= 1e-6:
                continue
            ys = t.get("yopilgan_sana")
            if ys and r["sana"] and r["sana"] >= ys:  # транш шу санадан ёпилган
                continue
            p = t["narxlar"].get(r["tur"]) or t["narxlar"].get("Детский" if r["tur"].startswith("❓") else "Дорика")
            take = min(rem_g, left[ti] / p)
            c = take * p
            left[ti] -= c
            used[ti] += c
            cost += c
            rem_g -= take
            parts.append((t["no"], take, p))
            first_date[ti] = first_date[ti] or r["sana"]
            last_date[ti] = r["sana"]
        if rem_g > 1e-9:  # пул етмади — охирги транш нархида
            t = tr[-1]
            p = t["narxlar"].get(r["tur"]) or t["narxlar"].get("Дорика")
            cost += rem_g * p
            parts.append(("пул етмади", rem_g, p))
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
        if any(n == "пул етмади" for n, _, _ in parts):
            r["holat"] = r.get("holat", "") + " | пул етмади"
        out.append(r)

    trs = []
    for ti, t in enumerate(tr):
        if left[ti] <= 1:
            holat = "Тугаган"
        elif t.get("yopilgan_sana"):
            holat = f"Ёпилган {t['yopilgan_sana']} (қолдиқ реестрда кўринмаган)"
        else:
            holat = "Фаол" if used[ti] > 0 else "Ишлатилмаган"
        trs.append(dict(no=t["no"], sana=t["sana"], summa=t["summa_usd"], birja=t["birja_999"],
                        used=used[ti], left=left[ti], holat=holat, birinchi=first_date[ti], oxirgi=last_date[ti]))
    return out, trs


def sabab(r, tr, parts):
    f = r["fakt_1gr"]
    tno = parts[0][0]
    t = next((x for x in tr if x["no"] == tno), tr[-1])
    for name, p in t["narxlar"].items():  # бошқа тур нархи билан ёзилганми
        if name != r["tur"] and abs(f - p) <= 0.05:
            return f"Таъминотчи «{name}» нархини қўйган ({p:.2f}), керак «{r['tur']}» ({r['hisob_1gr']:.2f})"
    for x in tr:  # бошқа транш нархими
        if x["no"] != tno:
            p = x["narxlar"].get(r["tur"])
            if p and abs(f - p) <= 0.05:
                return f"Бошқа транш (№{x['no']}) нархи қўйилган ({p:.2f}); FIFO бўйича №{tno}"
    labor = r["hisob_1gr"] - t["birja_999"] * K  # ишчилик + НДС (шу тур, шу транш)
    impl = (f - labor) / K
    if abs(impl - t["birja_999"]) > 0.3:
        return (f"Таъминотчи биржа ≈{impl:.2f} бўйича ҳисоблаган, керак транш №{tno} фиксацияси {t['birja_999']:.2f} "
                f"(1 гр да {f - r['hisob_1gr']:+.2f} $)")
    return f"Нарх фарқи: фактура {f:.2f}, керак {r['hisob_1gr']:.2f}"


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
        if r["sana"]:
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


def write_report(path, cfg, rows, trs):
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
    wb.save(path)
    return dict(tot_given=tot_given, tot_used=tot_used, left=left_total, fakt=fakt_sum, hisob=hisob_sum, qimmat=qimmat, arzon=arzon,
                bad=len(bad), n=len(rows), active=active, tugash=tugash, temp=temp, bad_s=bad_s)


def run(transh_path, ledger_path, out_path, files=()):
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
    rows, trs = compute(cfg, ledger)
    s = write_report(out_path, cfg, rows, trs)

    L = []
    new_keys = {key(r) for r in new_rows}
    new_res = [r for r in rows if key(r) in new_keys]
    if files:
        L.append(f"📥 Янги қатор: {added} (такрорий ўтказилди: {dup}). Реестрда жами: {s['n']}")
    if new_res:
        nb = [r for r in new_res if r["holat"].startswith("⚠")]
        L.append(f"🧾 Янги фактура: {len(new_res)} позиция, {sum(r['gramm'] for r in new_res):,.2f} гр\n"
                 f"   фактура {sum((r['summa'] or 0) for r in new_res):,.2f} $ | ҳисоб {sum((r['hisob'] or 0) for r in new_res):,.2f} $\n"
                 f"   {'✅ ҳаммаси тўғри' if not nb else f'⚠ фарқли: {len(nb)} та, фарқ {sum((r.get('farq') or 0) for r in nb):+,.2f} $'}")
        grp = OrderedDict()
        for r in sorted(nb, key=lambda r: -abs(r.get("farq") or 0)):
            k = (r["tur"], round(r.get("fakt_1gr") or 0, 1), round(r.get("hisob_1gr") or 0, 2))
            a = grp.setdefault(k, {"n": 0, "g": 0.0, "f": 0.0, "izoh": r["izoh"], "fil": set()})
            a["n"] += 1; a["g"] += r["gramm"] or 0; a["f"] += r.get("farq") or 0; a["fil"].add(r["filial"])
        for (tr_, fk, hk), a in sorted(grp.items(), key=lambda kv: -abs(kv[1]["f"]))[:12]:
            L.append(f"   ⚠ {tr_}: {a['n']} дона, {a['g']:.2f} гр — фактура ≈{fk:.2f} ≠ керак {hk:.2f} $/гр → {a['f']:+,.2f} $\n"
                     f"      {a['izoh']}")
    if warns:
        L.append("🧮 Фактура арифметикасида хато:\n" + "\n".join("   • " + w for w in warns[:15]))
    L.append(f"💰 Берилган: {s['tot_given']:,.0f} $ | Олинган: {s['tot_used']:,.0f} $ | Қолдиқ: {s['left']:,.0f} $")
    for t in trs:
        L.append(f"   Транш №{t['no']} ({t['sana']}): {t['summa']:,.0f} $ → ишлатилган {t['used']:,.0f}, қолдиқ {t['left']:,.0f} — {t['holat']}")
    L.append(f"📊 Жами фарқ (фактура−ҳисоб): {s['fakt']-s['hisob']:+,.2f} $ (қиммат {s['qimmat']:+,.2f}, арзон {s['arzon']:+,.2f}); фарқли позиция: {s['bad']}")
    if s["tugash"]:
        L.append(f"⏳ Темп: {s['temp']:,.0f} $/кун → таҳминий тугаш {s['tugash']:%d.%m.%Y}")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--transh", required=True)
    ap.add_argument("--ledger", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("files", nargs="*")
    a = ap.parse_args()
    print(run(a.transh, a.ledger, a.out, a.files))


if __name__ == "__main__":
    main()
