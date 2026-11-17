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
from collections import defaultdict, OrderedDict, Counter

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



# ---------------- ҚҚС (НДС) таҳлили ----------------
# ПҚ-207 (26.06.2025), 6-банд, 01.01.2026 – 01.01.2028: ҚҚС базаси = сотиш нархи − қимматбаҳо металл/тош қиймати
# (Солиқ қўмитаси эълон қилган нархларда); бюджетга тўланган ҚҚС нинг 80% и қайтарилади (даромаднинг ≥80% и заргарликдан бўлса).
QQS_DEFAULT = dict(stavka=12.0, qonun_qaytish=50.0, gavhar_qaytish=50.0, manba="Солиқ қўмитаси эълон қилган нарх")


def qqs_cfg(cfg):
    d = dict(QQS_DEFAULT)
    d.update({k: v for k, v in ((cfg or {}).get("qqs") or {}).items() if v not in (None, "")})
    return d


def qqs_k(stavka, qaytish):
    """1 гр нархга қўшиладиган ҚҚС улуши: (нарх − металл) × k.  k = ставка/(1+ставка) × (1 − қайтиш)."""
    s = stavka / 100
    return s / (1 + s) * (1 - qaytish / 100)


def qqs_tahlil(t, cfg):
    """Транш бўйича: Gavhar қўйган ҚҚС ↔ қонун бўйича ҚҚС (ҳар тур учун, $/гр)."""
    q = qqs_cfg(cfg)
    kG, kL = qqs_k(q["stavka"], q["gavhar_qaytish"]), qqs_k(q["stavka"], q["qonun_qaytish"])
    rows = narx_tarkibi(t, cfg)
    C, manba = t.get("metall_585"), "kiritilgan"
    if not C:
        cs = [x["jami"] - x["qqs"] / kG for x in rows if x["qqs"] > 0.02] if kG > 0 else []
        C, manba = (round(sum(cs) / len(cs), 2), "Gavhar narxidan") if cs else (None, None)
    for x in rows:
        base = x["lom"] + x["ishchilik"]
        if C is None:
            x.update(qqs_qonun=None, jami_qonun=None, ortiqcha=None, qqs_toliq=None)
            continue
        P = (base - kL * C) / (1 - kL) if kL < 1 else base
        if P < C:  # ҚҚС базаси манфий — ҚҚС йўқ
            P = base
        x["qqs_toliq"] = round(max(x["jami"] - C, 0) * (q["stavka"] / 100) / (1 + q["stavka"] / 100), 3)
        x["qqs_qonun"] = round(P - base, 3)
        x["jami_qonun"] = round(P, 2)
        x["ortiqcha"] = round(x["jami"] - P, 3)
    return dict(metall_585=C, metall_999=round(C * 999 / 585, 2) if C else None, manba=manba, rows=rows,
                stavka=q["stavka"], qonun_qaytish=q["qonun_qaytish"], gavhar_qaytish=q["gavhar_qaytish"], kG=kG, kL=kL)


def qqs_ortiqcha_jami(ledger_rows, t, tahlil):
    """Шу пулдан олинган товарлар бўйича ортиқча ҚҚС ($): ҳар қаторнинг шу пулдан сарфи / 1 гр нарх × ортиқча/гр."""
    ort = {x["tur"]: x.get("ortiqcha") for x in tahlil["rows"]}
    jami, gramm = 0.0, 0.0
    for r in ledger_rows:
        c = (r.get("sarf") or {}).get(t["no"])
        if not c:
            continue
        tur_ = "Вакс" if r.get("tur") == "Кузсиз" else r.get("tur")
        p = t["narxlar"].get(tur_) or t["narxlar"].get("Дорика")
        o = ort.get(tur_) if ort.get(tur_) is not None else ort.get("Дорика")
        if not p or o is None:
            continue
        g = c / p
        gramm += g
        jami += g * o
    return round(jami, 2), round(gramm, 2)


def qqs_yangi(p, cfg):
    """Янги пул (транш) нархларидаги ҚҚС ни қабул қилишдан ОЛДИН текшириш. holat: ok | xato | kam | yoq."""
    q = qqs_cfg(cfg)
    t = {"no": 0, "sana": p.get("sana"), "birja_999": p["birja"], "narxlar": dict(p["narxlar"])}
    a = qqs_tahlil(t, cfg)
    L = [f"🧾 НДС текшируви (қайтиш: қонун {q['qonun_qaytish']:g}%, Gavhar {q['gavhar_qaytish']:g}%):"]
    if a["metall_585"] is None:
        return dict(holat="yoq", lines=L + ["   Нархда НДС йўқ (1 гр = лом + ишчилик) — текширадиган нарса йўқ."])
    kG = a["kG"]
    seen, cs = set(), []
    for x in a["rows"]:
        if x["qqs"] > 0.02 and x["jami"] not in seen:
            seen.add(x["jami"])
            cs.append(x["jami"] - x["qqs"] / kG)
    for x in a["rows"]:
        o = x.get("ortiqcha")
        if o is None:
            continue
        belgi = "✅" if abs(o) <= 0.015 else ("❌" if o > 0 else "ℹ")
        L.append(f"   {belgi} {x['tur']}: НДС {x['qqs']:.2f} · қонун бўйича {x['qqs_qonun']:.2f}" + (f" · ортиқча {o:+.2f} $/гр" if abs(o) > 0.015 else ""))
    L.append(f"   Олтин қиймати (Gavhar нархидан): {a['metall_585']:.2f} $/гр 585 (999: {a['metall_999']:.2f})")
    ors = [x["ortiqcha"] for x in a["rows"] if x.get("ortiqcha") is not None]
    holat = "ok"
    if cs and max(cs) - min(cs) > 0.3:
        holat = "xato"
        L.append(f"   ⚠ Турлар бўйича олтин қиймати ҳар хил чиқди ({min(cs):.2f}–{max(cs):.2f}) — НДС бир хил қоида билан ҳисобланмаган, нархлардан бири нотўғри бўлиши мумкин.")
    if ors and max(ors) > 0.015:
        holat = "xato"
        L.append(f"❌ НДС НОТЎҒРИ — ҳар граммга +{max(min(ors), 0):.2f} … +{max(ors):.2f} $ ортиқча.")
    elif holat == "xato":
        L.append("❌ НДС ни текширинг.")
    elif ors and min(ors) < -0.015:
        holat = "kam"
        L.append("ℹ НДС қонундагидан кам — бизга фойда.")
    else:
        L.append("✅ НДС ТЎҒРИ — қонун бўйича ҳисобга мос.")
    return dict(holat=holat, lines=L)


def qqs_gramm_tur(ledger_rows, t):
    """Шу пулдан олинган грамм — нарх тури бўйича (mini app жонли ҳисоби учун)."""
    out = defaultdict(float)
    for r in ledger_rows:
        c = (r.get("sarf") or {}).get(t["no"])
        if not c:
            continue
        tur_ = "Вакс" if r.get("tur") == "Кузсиз" else r.get("tur")
        p = t["narxlar"].get(tur_) or t["narxlar"].get("Дорика")
        if p:
            out[tur_] += c / p
    return {k: round(v, 3) for k, v in out.items()}


def qqs_matn(cfg, ledger_rows):
    q = qqs_cfg(cfg)
    L = ["🧾 ҚҚС (НДС) ТЕКШИРУВИ", "",
         f"Қонун (ПҚ-207, 01.01.2026–01.01.2028): ҚҚС {q['stavka']:g}% фақат устамага (нарх − металл қиймати),",
         f"тўланган ҚҚС нинг {q['qonun_qaytish']:g}% и қайтарилади → нархга {100 - q['qonun_qaytish']:g}% қўшилади.",
         f"Gavhar ҳисоби: {q['gavhar_qaytish']:g}% қайтади деб → нархга {100 - q['gavhar_qaytish']:g}% қўшган.", ""]
    jami_all = 0.0
    for t in sorted(cfg.get("transhlar", []), key=lambda t: t["sana"]):
        a = qqs_tahlil(t, cfg)
        L.append(f"💳 №{t['no']} · {_dmy(t['sana'])} · биржа {t['birja_999']:g}")
        if a["metall_585"] is None:
            L.append("   Gavhar нархида ҚҚС йўқ (ёки металл қиймати киритилмаган)")
            continue
        L.append(f"   Металл қиймати (585): {a['metall_585']:.2f} $/гр ({'киритилган' if a['manba'] == 'kiritilgan' else 'Gavhar нархидан аниқланди'})")
        for x in a["rows"]:
            if x.get("qqs_qonun") is None:
                continue
            L.append(f"   {x['tur']}: Gavhar {x['qqs']:.2f} · қонун бўйича {x['qqs_qonun']:.2f} · ортиқча {x['ortiqcha']:+.2f} $/гр")
        ors = [x["ortiqcha"] for x in a["rows"] if x.get("ortiqcha") is not None]
        if ors and max(ors) > 0.015:
            L.append(f"   ❌ НОТЎҒРИ — ҚҚС ортиқча олинган: ҳар граммга +{max(min(ors), 0):.2f} … +{max(ors):.2f} $")
        elif ors and min(ors) < -0.015:
            L.append("   ℹ ҚҚС қонундагидан кам олинган (бизга фойда)")
        elif ors:
            L.append("   ✅ ТЎҒРИ — қонун бўйича ҳисобга мос")
        s_, g_ = qqs_ortiqcha_jami(ledger_rows, t, a)
        if ors and max(ors) <= 0.015 and min(ors) >= -0.015:
            s_ = 0.0  # тўғри — фарқ фақат яхлитлашдан
        if g_:
            L.append(f"   Шу пулдан олинган {g_:,.0f} гр бўйича ортиқча ҚҚС: {s_:,.2f} $")
            jami_all += s_
    L += ["", f"Жами ортиқча ҚҚС: {jami_all:,.2f} $", "Созлама: mini app → Admin → «QQS (NDS)»."]
    return "\n".join(L)

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


# Gavhar эмас — ўтказиб юборилган қаторлар (бот ва ҳисобот хабарида кўрсатилади)
SKIP = {"boshqa": 0, "nomlar": set(), "tilla_emas": 0}
GAVHAR_MODEL = re.compile(r"\b[A-Z]{3}\d{6}\b")


def skip_reset():
    SKIP.update(boshqa=0, nomlar=set(), tilla_emas=0)


def skip_matn():
    L = []
    if SKIP["boshqa"]:
        nm = ", ".join(sorted(SKIP["nomlar"])[:5]) or "номаълум"
        L.append(f"🚫 Gavhar эмас — {SKIP['boshqa']} та товар ўтказиб юборилди (таъминотчи: {nm}).")
    if SKIP["tilla_emas"]:
        L.append(f"🚫 Тилла товари эмас (номида «Proba 585» йўқ) — {SKIP['tilla_emas']} та қатор ўтказиб юборилди.")
    return L


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
    begona = None
    if "gavhar" not in head.lower() and "гавҳар" not in head.lower() and "гавхар" not in head.lower():
        m = re.search(r"(ООО|OOO|МЧЖ|MChJ|ЧП|ЯТТ)\s*[«\"']?([^»\"'.]{2,40})", head)
        begona = (m.group(2).strip() if m else ws.title.strip())[:40]
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
    # Gavhar фактура рақами доим GVR-… (GVR-27102026). Бошқа рақам (мас. DG-000000795) ва «Gavhar Jewel» ёзуви йўқ бўлса — бегона
    if not begona and faktura and not faktura.upper().startswith("GVR") and "gavhar jewel" not in head.lower():
        begona = f"фактура № {faktura}"
    if begona:  # Gavhar эмас — бошқа таъминотчи фактураси: ўқимаймиз
        SKIP["boshqa"] += len(out)
        SKIP["nomlar"].add(begona)
        return []
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
                SKIP["boshqa"] += 1
                SKIP["nomlar"].add(tam[:40])
                continue
            if not ("proba" in nomi.lower() or "585" in nomi or "проба" in nomi.lower()):
                SKIP["tilla_emas"] += 1  # тилла товар эмас (номида проба йўқ) — ўтказамиз
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
    """Такрор юклашни аниқлаш калити. Артикул (GV…) битта товарники — у возвратдан кейин БОШҚА фактурада қайта
    берилиши мумкин, шунинг учун фактура ва дўкон ҳам калитга киради (05.11.2026: Бахмал мини USD002046 тушиб қолган эди)."""
    if r.get("id"):
        return ("id", str(r["id"]), str(r.get("faktura") or ""), str(r.get("filial") or ""))
    return ("x", r.get("faktura"), r.get("kod"), str(r.get("gramm")), str(r.get("summa")))


# ---------------- возврат ----------------
RET_COLS = ["id", "sana", "filial", "kod", "model", "nomi", "gramm", "summa", "hujjat", "chiqdi_sana", "chiqdi_qator"]
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
        cq = str(r.get("chiqdi_qator") or "").strip()
        r["chiqdi_qator"] = int(float(cq)) if cq else None
        r["chiqdi_sana"] = r.get("chiqdi_sana") or None
    return rows


def ret_key(r):
    return (str(r.get("kod") or ""), str(r.get("model") or ""), str(r.get("sana") or ""),
            round(float(r["gramm"]), 2), round(float(r["summa"]), 2))


def ret_faol(x, i=None, r=None):
    """Возврат товар шу қатор учун рўйхатда борми (кейинги файлда чиқариб ташланмаганми)."""
    cq, cs = x.get("chiqdi_qator"), x.get("chiqdi_sana")
    if cq is not None:
        return i is None or i < cq
    if cs:
        return not (r and r.get("sana") and r["sana"] >= cs)
    return True


def sync_returns(path, new_rows, ledger_len, faqat_qosh=False):
    """Янги возврат файли = тўлиқ жорий рўйхат: янгилари қўшилади (+), файлда йўқлари чиқарилади (−).
    Чиқарилганлар олдин юкланган фактуралар учун сақланади. faqat_qosh=True — фақат қўшиш."""
    old = load_returns(path)
    nk = {}
    for r in new_rows:
        nk.setdefault(ret_key(r), r)
    ok_ = {ret_key(r): r for r in old}
    today = dt.date.today().isoformat()
    plus, minus, back = [], [], []
    for k, r in nk.items():
        if k not in ok_:
            old.append(r)
            plus.append(r)
        elif ok_[k].get("chiqdi_qator") is not None or ok_[k].get("chiqdi_sana"):
            ok_[k]["chiqdi_qator"], ok_[k]["chiqdi_sana"] = None, None
            back.append(ok_[k])
    if not faqat_qosh:
        for k, r in ok_.items():
            if k not in nk and r.get("chiqdi_qator") is None and not r.get("chiqdi_sana"):
                r["chiqdi_qator"], r["chiqdi_sana"] = ledger_len, today
                minus.append(r)
    save_returns(path, old)
    faol = [r for r in old if r.get("chiqdi_qator") is None and not r.get("chiqdi_sana")]
    sm = lambda rs: sum(float(r["summa"]) for r in rs)
    return dict(plus=len(plus), plus_s=sm(plus), minus=len(minus), minus_s=sm(minus), back=len(back),
                n=len(faol), summa=sm(faol))


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


def find_return(r, g, returns, f1=None, tol=0.3, i=None):
    """Возврат рўйхатидан мос товар: (топилган, нархи мосми). Аввал артикул, кейин модель+грамм;
    бир нечта номзод бўлса — нархи фактурага мос келгани олинади."""
    kod, model = str(r.get("kod") or "").strip(), model_of(r.get("nomi"))
    sana_r = r.get("sana") or ""
    returns = [x for x in returns if ret_faol(x, i, r) and not (sana_r and x.get("sana") and str(x["sana"])[:10] > sana_r)]
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


def q_yopiq(q, i, r):
    """Возврат қисми якунланганми (шу қатор учун)."""
    if q.get("_ob_yop"):  # ой боши қолдиғи (Gavhar билан келишилган) дан кейин бу возврат қисми ишлатилмайди
        return True
    yq, ys = q.get("yopilgan_qator"), q.get("yopilgan_sana")
    if yq is not None:
        return i >= yq
    return bool(ys and r.get("sana") and r["sana"] >= ys)


# ---------------- возврат ФИКСАЦИЯСИ (ой бўйича: тош тури → қиймат) ----------------
FOYLAR = ["январ", "феврал", "март", "апрел", "май", "июн", "июл", "август", "сентябр", "октябр", "ноябр", "декабр"]


def oy_nom(oy):
    try:
        return f"{FOYLAR[int(oy[5:7]) - 1]} {oy[:4]}"
    except Exception:
        return oy or ""


def qism_nom(qn):
    """Возврат манбаи номи: 1 → «1-қисм», «Ф2026-09» → «фиксация сентябр 2026»."""
    if isinstance(qn, str) and qn.startswith("Ф"):
        return f"фиксация {oy_nom(qn[1:])}"
    return f"{qn}-қисм"


def fiks_list(cfg):
    return sorted(cfg.get("vozvrat_fiks") or [], key=lambda f: f["oy"])


def fiks_reset(cfg):
    for f in fiks_list(cfg):
        for tn, x in f["turlar"].items():
            x["_used"], x["_n"], x["_g"] = 0.0, 0, 0.0
            x["_nom"] = "Браслет/Занжир" if tn == "Браслет" else tn


def fiks_max_narx(cfg, tur_):
    """Шу турнинг энг юқори нархи (барча пуллар, ўчирилганлари ҳам) — возврат нархи бундан баланд бўлолмайди."""
    vs = [t["narxlar"].get(tur_) for t in cfg.get("transhlar", [])] + [q.get("narxlar", {}).get(tur_) for q in cfg.get("vozvrat_qismlar") or []]
    for f in fiks_list(cfg):
        x = f["turlar"].get(tur_) or {}
        if x.get("gramm"):
            vs.append(x["summa"] / x["gramm"])
    vs = [v for v in vs if v]
    return max(vs) if vs else None


def fiks_narx_matn(cfg, tur_):
    """Фиксациядаги шу тур 1 гр нархлари: «сентябр 2026 — 95.30, октябр 2026 — 90.50»."""
    out = []
    for f in fiks_list(cfg):
        for c in FIKS_MAP.get(tur_, [tur_]):
            x = f["turlar"].get(c)
            if x and x.get("gramm"):
                out.append(f"{oy_nom(f['oy'])} — {x['summa'] / x['gramm']:.2f} $/гр")
                break
    return ", ".join(out) or "бу тур йўқ"


def vozvrat_narx_matn(cfg, tur_):
    """Шу турнинг ёқилган возвратлардаги 1 гр нархлари — хато товар изоҳи учун."""
    out = []
    for q in cfg.get("vozvrat_qismlar") or []:
        if q.get("faol") is False or q.get("_ob_yop"):
            continue
        p = q.get("narxlar", {}).get(tur_) or q.get("narxlar", {}).get("Дорика")
        if p:
            out.append(f"{q['no']}-қисм {p:.2f}")
    fm = fiks_narx_matn(cfg, tur_)
    if fm != "бу тур йўқ":
        out.append("фиксация " + fm.replace(" $/гр", ""))
    return "; ".join(out)


def fiks_try(cfg, r, g, s, chegara):
    """Пул нархига тушмаган товарни возврат фиксациясидан ёпиш. True — ёпилди; ("narx", max) — нарх жуда баланд; None — йўқ."""
    sana = r.get("sana") or ""
    f1 = s / g
    for f in fiks_list(cfg):
        if f.get("faol") is False or (sana and sana[:7] < f["oy"]):
            continue
        # фиксацияда грамм бор бўлса — возврат 1 гр нархи = сумма ÷ грамм; товар шу нархда бўлиши шарт (±чегара)
        x = next((f["turlar"][c] for c in FIKS_MAP.get(r["tur"], [r["tur"]])
                  if c in f["turlar"] and f["turlar"][c]["summa"] - f["turlar"][c].get("_used", 0.0) >= s - 1
                  and (not f["turlar"][c].get("gramm") or (f["turlar"][c]["gramm"] - f["turlar"][c].get("_g", 0.0) >= g - 0.01
                                                          and abs(f1 - f["turlar"][c]["summa"] / f["turlar"][c]["gramm"]) <= chegara))), None)
        if not x:
            continue
        mx = fiks_max_narx(cfg, r["tur"])
        if mx and f1 > mx + chegara:
            return ("narx", mx)
        x["_used"] = x.get("_used", 0.0) + s; x["_n"] = x.get("_n", 0) + 1; x["_g"] = x.get("_g", 0.0) + g
        nom = oy_nom(f["oy"])
        r.update(transh="возврат", vozvrat_qism="Ф" + f["oy"], hisob=s, hisob_1gr=f1, fakt_1gr=f1, farq_gr=0.0, farq=0.0,
                 holat="✓ возврат товар",
                 vozvrat_manba=(f"фиксация {nom}: «{x.get('_nom', r['tur'])}» ({x['summa']:,.0f} $ дан, қолди {x['summa'] - x['_used']:,.0f} $"
                               + (f", {x['gramm'] - x['_g']:,.2f} гр" if x.get("gramm") else "") + f"), {g:.2f} гр × {f1:.2f} $/гр"),
                 izoh=f"Возврат товар — {nom} фиксациясидаги «{x.get('_nom', r['tur'])}» қийматидан айирилди ({f1:.2f} $/гр).")
        return True
    return None


def fiks_stats(cfg):
    out = []
    for f in fiks_list(cfg):
        T = []
        for tn, x in f["turlar"].items():
            u, g_ = x.get("_used", 0.0), x.get("_g", 0.0)
            T.append(dict(tur="Браслет/Занжир" if tn == "Браслет" else tn, summa=x["summa"], gramm=x.get("gramm"), used=u, left=x["summa"] - u, n=x.get("_n", 0), g=g_,
                          ort=(u / g_ if g_ else None), fiks_1gr=(x["summa"] / x["gramm"] if x.get("gramm") else None)))
        sm, u = sum(t["summa"] for t in T), sum(t["used"] for t in T)
        out.append(dict(oy=f["oy"], nom=oy_nom(f["oy"]), faol=f.get("faol", True), summa=sm, used=u, left=sm - u,
                        n=sum(t["n"] for t in T), turlar=T))
    return out


def fiks_tafsil(fs, chegara=0.3):
    L = []
    for t in fs["turlar"]:
        x = f"{t['tur']}: {t['summa']:,.0f} $ → ишлатилди {t['used']:,.0f} ({t['n']} та) · қолди {t['left']:,.0f} $"
        if t.get("gramm"):
            x += f" · грамм: {t['g']:,.2f} / {t['gramm']:,.2f}, қолди {t['gramm'] - t['g']:,.2f} гр"
        if t["ort"]:
            x += f" · ўртача {t['ort']:.2f} $/гр"
            if t["fiks_1gr"] and t["ort"] > t["fiks_1gr"] + chegara:
                x += f" ⚠ фиксацияда {t['fiks_1gr']:.2f}"
        L.append(x)
    return L


FIKS_MAP = {"Кузсиз": ["Вакс"], "Дорика": ["Дорика", "Браслет"], "Ван Клиф": ["Дорика", "Браслет"]}  # Ван Клиф возвратдан ўрнига берилса — Дорика нархида (17.11.2026)  # фиксациядаги «Браслет Занжир» — билакузук ва занжир
_NUM = r"\d{1,3}(?:[ \u00a0\u202f']\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?"


def _fnum(x):
    x = re.sub(r"[ \u00a0\u202f']", "", x)
    return float(x.replace(",", "."))


def fiks_matn_parse(text, aliases):
    """Gavhar возврат фиксацияси матни (Excel дан нусха ҳам бўлади):
    «Сентябрь\nБриллиант 3 310,22 732 $ 315 471,77 $ 95,30\n...\nОктябрь\n...» → ({oy: {тур: {summa, gramm, dona}}}, хатолар).
    Сатрда сонлар: грамм, штук, сумма, $/грамм (ёки фақат сумма)."""
    yil = dt.date.today().year
    res, errs, oy, prev = {}, [], None, None
    for l in (text or "").splitlines():
        ll = l.lower().replace("ё", "е")
        m = re.search(r"\b(\d{1,2})[./-](20\d\d)\b", ll)
        k = next((k for k, nm in enumerate(FOYLAR) if nm[:5] in ll), None)
        if k is None:
            k = {"yanvar": 0, "fevral": 1, "mart": 2, "aprel": 3, "may ": 4, "iyun": 5, "iyul": 6, "avgust": 7,
                 "sentyabr": 8, "oktyabr": 9, "noyabr": 10, "dekabr": 11}.get(next((w for w in ("sentyabr", "oktyabr", "noyabr", "dekabr", "avgust", "iyul", "iyun", "aprel", "mart", "fevral", "yanvar") if w in ll), ""))
        if m or k is not None:
            y = re.search(r"20\d\d", ll)
            oy = f"{m.group(2)}-{int(m.group(1)):02d}" if m else f"{y.group(0) if y else yil}-{k + 1:02d}"
            res.setdefault(oy, {})
        if "жами" in ll or "итого" in ll or "jami" in ll:
            continue
        if "браслет" in ll or "braslet" in ll or ("занжир" in ll and prev == "Браслет"):
            name = "Браслет"
        else:
            name = next((n for n, keys in aliases if any(kk in ll for kk in keys)), None)
        nums = [_fnum(x) for x in re.findall(_NUM, re.sub(r"\b(\d{1,2})[./-](20\d\d)\b|20\d\d", " ", l))]
        if name and not nums:
            prev = name
            continue
        if not name and prev == "Браслет" and nums and "занжир" in ll:
            name = "Браслет"
        prev = name
        if not name or not nums:
            continue
        if oy is None:
            errs.append("ой топилмади (мас. «Сентябрь» ёки «09.2026» деб бошланг)")
            break
        gm = re.search(r"(" + _NUM + r")\s*(гр|gr)\b", ll)
        if gm:
            gramm, summa = _fnum(gm.group(1)), max(nums)
            dona = None
        elif len(nums) >= 3:
            gramm, dona, summa = nums[0], nums[1], nums[2]
        elif len(nums) == 2:
            gramm, summa, dona = min(nums), max(nums), None
        else:
            gramm, summa, dona = None, nums[0], None
        names = ["Вакс", "Кузсиз"] if ("вакс" in ll and ("кузсиз" in ll or "кўзсиз" in ll)) else [name]
        for n in names:
            res[oy][n] = dict(summa=round(summa, 2), gramm=gramm, dona=dona)
    res = {k: v for k, v in res.items() if v}
    if not res and not errs:
        errs.append("тош тури ва қиймати топилмади (мас. «Сентябрь» ва остида «Вакс 2 348,83 577 $ 219 071,37»)")
    return res, errs

def qism_stats(cfg):
    """Суммали возврат қисмлари: берилган / қолган."""
    return [dict(no=q["no"], sana=q.get("sana"), birja=q.get("birja"), summa=q["summa_usd"], berilgan=q.get("_used", 0.0),
                 n=q.get("_n", 0), qolgan=q["summa_usd"] - q.get("_used", 0.0),
                 yopilgan=q.get("yopilgan_sana"), faol=q.get("faol", True), narxlar=q.get("narxlar", {}))
            for q in cfg.get("vozvrat_qismlar") or []]


def gavhar_pullar(rows, cfg):
    """Gavhar ҳар пул нархини қачондан-қачонгача ёзган (фактура нархидан): {пул №: {birinchi, oxirgi, n, summa}}.
    Gavhar пулни қаерда «ёпгани» — охирги шу нархда ёзган санаси, кейин бошқа пул нархига ўтган бўлса."""
    ch = cfg.get("chegara_usd_gr", 0.3)
    trs = sorted([t for t in cfg.get("transhlar", []) if t.get("faol", True)], key=lambda t: (t["sana"], t["no"]))
    out = {}
    for r in rows:
        f, sana = r.get("fakt_1gr"), r.get("sana")
        if f is None or not sana or r.get("transh") == "возврат":
            continue
        tur_ = "Вакс" if r.get("tur") == "Кузсиз" else r.get("tur")
        tt = tosh_tur(r.get("nomi")) if tur_ in ("Браслет", "Дорика") else None  # тошли билак узук — Gavhar тош нархини қўйиши мумкин
        hit = []
        for t in trs:
            if t["sana"] > sana:
                continue
            ps = [t["narxlar"].get(tur_) or t["narxlar"].get("Дорика")] + ([t["narxlar"].get(tt)] if tt else [])
            if any(p and abs(f - p) <= ch for p in ps):
                hit.append(t["no"])
        if len(hit) != 1:  # бир нечта пул нархи бир хил — аниқ эмас
            continue
        x = out.setdefault(hit[0], dict(birinchi=sana, oxirgi=sana, n=0, summa=0.0))
        x["birinchi"] = min(x["birinchi"], sana); x["oxirgi"] = max(x["oxirgi"], sana)
        x["n"] += 1; x["summa"] += r.get("summa") or 0
    sana_max = max((r.get("sana") or "" for r in rows), default="")
    tsana = {t["no"]: (t["sana"], t["no"]) for t in trs}
    for no, x in out.items():
        # кейинги (санаси кейин) пул нархида шу пулдан кейин ҳам ёзилган бўлса — Gavhar бу пулни охирги ишлатган кунида ёпган
        keyin = [n for n, y in out.items() if n != no and tsana.get(n, ("",)) > tsana.get(no, ("",)) and y["oxirgi"] > x["oxirgi"]]
        x["yopdi"] = bool(keyin) and x["oxirgi"] < sana_max
        x["keyingi"] = min(keyin, key=lambda n: tsana[n]) if keyin else None
    return out


def holat_royxat(cfg, trs, v1=None, rows=None):
    """Ҳамма пуллар ва возвратлар — ҳолати аниқ: 🟢 очиқ / 🔒 ёпилган / ✅ тугаган / ⏸ ўчирилган / 📅 фақат эски приход."""
    st = {t["no"]: t for t in trs}
    gp = gavhar_pullar(rows, cfg) if rows else {}
    out = []
    for t in sorted(cfg.get("transhlar", []), key=lambda t: (t["sana"], t["no"])):
        x = st.get(t["no"]) or {}
        used, left = x.get("used"), x.get("left")
        if not t.get("faol", True):
            k, m = "off", "⏸ Ўчирилган — текширувга кирмайди" + (f" (олдин {_dmy(t['yopilgan_sana'])} да ёпилган эди)" if t.get("yopilgan_sana") else " (ёпилмаган — фақат ўчирилган)")
        elif t.get("faqat_tekshiruv"):
            k, m = "eski", "📅 Фақат эски приход текшируви учун"
        elif t.get("yopilgan_sana"):
            k, m = "yopiq", f"🔒 Ёпилган {_dmy(t['yopilgan_sana'])} — кейинги фактураларга ишлатилмайди"
        elif t.get("_avto_yop"):
            k, m = "yopiq", (f"🔒 Gavhar ёпган — {_dmy(t['_avto_yop'])} дан кейин №{t.get('_avto_keyingi')} нархига ўтган, бизда ҳам ёпилди"
                             + (f". ⚠ Шу пулда {left:,.0f} $ қолган — Gavhar билан солиштиринг" if left and left > 1 else ""))
        elif left is not None and left <= 1:
            k, m = "tugagan", "✅ Тугаган"
        else:
            k, m = "ochiq", "🟢 Очиқ — ишлатиляпти" if (used or 0) > 0 else "🟢 Очиқ — навбатда"
        g = gp.get(t["no"])
        gm = None
        if g:
            gm = f"🏭 Gavhar шу пул нархида: {_dmy(g['birinchi'])} – {_dmy(g['oxirgi'])}, {g['n']} та товар"
            if g["yopdi"]:
                gm += f" → {_dmy(g['oxirgi'])} да ёпиб, №{g['keyingi']} нархига ўтган"
                if left is not None and left > max(1000, 0.01 * t["summa_usd"]) and k not in ("tugagan",) and not t.get("_avto_yop"):
                    gm += f". ⚠ Бизнинг ҳисобда шу пулда {left:,.0f} $ қолган эди — Gavhar билан солиштиринг"
            elif k == "tugagan":
                gm += ". ⚠ Бизда пул тугаган, Gavhar ҳали шу нархда ёзяпти"
        elif t.get("faol", True) and not t.get("faqat_tekshiruv") and rows is not None:
            gm = "🏭 Gavhar ҳали шу пул нархида фактура ёзмаган"
        out.append(dict(tur="pul", no=t["no"], sana=t["sana"], nom=f"Пул №{t['no']} · {_dmy(t['sana'])} · биржа {t['birja_999']:g}",
                        summa=t["summa_usd"], used=used, left=left, kod=k, matn=m, gavhar=gm,
                        g_yopdi=bool(g and g["yopdi"]), g_oxirgi=g["oxirgi"] if g else None))
    v = cfg.get("vozvrat_1") or {}
    if v1 and (v1.get("n") or v.get("yopilgan_sana") or v.get("faol") is False):
        if v.get("faol") is False:
            k, m = "off", "⏸ Ўчирилган — текширувга кирмайди" + (f" (олдин {_dmy(v['yopilgan_sana'])} да якунланган эди)" if v.get("yopilgan_sana") else " (якунланмаган — фақат ўчирилган)")
        elif v.get("yopilgan_sana"):
            k, m = "yopiq", f"🔒 Якунланган {_dmy(v['yopilgan_sana'])}"
        elif (v1.get("qolgan_summa") or 0) <= 1:
            k, m = "tugagan", "✅ Тугаган"
        else:
            k, m = "ochiq", "🟢 Очиқ — ишлатиляпти"
        out.append(dict(tur="vozvrat", no=1, sana=None, nom="Возврат 1-қисм · рўйхат (Excel)", summa=v1.get("summa"),
                        used=v1.get("berilgan_summa"), left=v1.get("qolgan_summa"), kod=k, matn=m))
    for q in qism_stats(cfg):
        if q["faol"] is False:
            k, m = "off", "⏸ Ўчирилган — текширувга кирмайди" + (f" (олдин {_dmy(q['yopilgan'])} да якунланган эди)" if q["yopilgan"] else " (якунланмаган — фақат ўчирилган)")
        elif q["yopilgan"]:
            k, m = "yopiq", f"🔒 Якунланган {_dmy(q['yopilgan'])}"
        elif (cfg.get("oy_boshi") or {}).get("sana"):
            k, m = "yopiq", f"🔒 Ой бошидан ({_dmy(cfg['oy_boshi']['sana'])}) ишлатилмайди — возврат фақат фиксация бўйича"
        elif q["qolgan"] <= 1:
            k, m = "tugagan", "✅ Тугаган"
        else:
            k, m = "ochiq", "🟢 Очиқ — ишлатиляпти" if q["berilgan"] > 0 else "🟢 Очиқ — навбатда"
        out.append(dict(tur="vozvrat", no=q["no"], sana=q.get("sana"), nom=f"Возврат {q['no']}-қисм · сумма · биржа {q['birja']}",
                        summa=q["summa"], used=q["berilgan"], left=q["qolgan"], kod=k, matn=m))
    out = [x for x in out if not (x["tur"] == "vozvrat" and x["kod"] == "off")]  # ўчирилган эски возвратлар кўрсатилмайди
    ob = cfg.get("oy_boshi") or {}
    if ob.get("sana"):
        bz = cfg.get("_ob_bizda") or {}
        farq = (bz.get("jami") or 0) - (ob.get("pul") or 0)
        out.insert(0, dict(tur="ob", no=0, sana=ob["sana"], nom=f"Ой боши · {_dmy(ob['sana'])} · Gavhar да ҳақдорлик {ob.get('jami', 0):,.2f} $",
                           summa=ob.get("jami"), used=None, left=None, kod="ochiq",
                           matn=(f"📌 Пул {ob.get('pul', 0):,.2f} $ (№{bz.get('pul_no')} пул нархида) + возврат {ob.get('vozvrat', 0):,.2f} $ "
                                 f"({ob.get('vozvrat_matn', '')}). Шу санадан ҳисоб шу рақамлардан бошланди")
                                + (f". Бизнинг олдинги ҳисобда пулларда {bz.get('jami', 0):,.0f} $ қолган эди — фарқ {farq:+,.0f} $" if bz and abs(farq) > 1 else "")))
    if ob.get("sana"):  # ой бошидан кейин ёпиқ (0 қолган) эски пуллар — рўйхатда кўрсатилмайди
        out = [x for x in out if not (x["tur"] == "pul" and x["sana"] < ob["sana"] and (x.get("left") or 0) <= 1
                                      and x["no"] != (cfg.get("_ob_bizda") or {}).get("pul_no"))]
    for f in fiks_stats(cfg):
        if not f["faol"]:
            k, m = "off", "⏸ Ўчирилган — текширувга кирмайди"
        elif f["left"] <= 1:
            k, m = "tugagan", "✅ Тугаган"
        else:
            k, m = "ochiq", "🟢 Очиқ — ишлатиляпти" if f["used"] > 0 else "🟢 Очиқ — навбатда"
        out.append(dict(tur="fiks", no=f["oy"], sana=f["oy"] + "-01", nom=f"Возврат фиксация · {f['nom']}",
                        summa=f["summa"], used=f["used"], left=f["left"], kod=k, matn=m,
                        tafsil=fiks_tafsil(f, cfg.get("chegara_usd_gr", 0.3))))
    return out


def manba_qatorlar(rs, tmap, cfg):
    """Фактура товарлари қайси пул ва қайси возврат бўйича текширилди: ҳар манба — битта қатор (сони, сумма, тўғри/хато)."""
    g = OrderedDict()
    for r in rs:
        if r.get("transh") == "возврат":
            qn = r.get("vozvrat_qism") or 1
            if isinstance(qn, str):
                k = (1, qn, f"↩️ Возврат {qism_nom(qn)} (тош тури 1 гр нархи)")
            elif qn == 1:
                k = (1, "1", "↩️ Возврат 1-қисм (рўйхат)")
            else:
                q0 = next((q for q in cfg.get("vozvrat_qismlar", []) if q["no"] == qn), {})
                k = (1, str(qn), f"↩️ Возврат {qn}-қисм (биржа {q0.get('birja')} нархи)")
        else:
            ns = list(dict.fromkeys(int(x) for x in re.findall(r"№(\d+)", str(r.get("transh") or ""))))
            if ns:
                k = (0, "+".join(map(str, ns)), " + ".join(f"💳 №{n} пул ({_dmy(tmap[n]['sana'])}, биржа {tmap[n].get('birja_999')})" if n in tmap else f"№{n}" for n in ns))
            else:
                k = (2, "", "❓ пулга боғланмади")
        x = g.setdefault(k, [0, 0.0, 0, 0.0])
        x[0] += 1; x[1] += r["summa"] or 0
        if str(r.get("holat", "")).startswith("⚠"):
            x[2] += 1; x[3] += r.get("farq") or 0
    out = []
    for (_, _, nom), (n, sm, nb_, fq) in sorted(g.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        out.append(f"   {nom}: {n} та · {sm:,.2f} $ — " + ("✅ тўғри" if not nb_ else f"❌ {nb_} та хато ({fq:+,.2f} $)"))
    return out


def holat_qisqa(items):
    """Фактура жавоби учун қисқа ҳолат: ҳар пул/возврат — битта қатор. Тўлиқи — /holat."""
    L = []
    for x in items:
        if x["kod"] in ("off", "eski"):
            continue
        if x["tur"] == "ob":
            L.append(f"📌 {x['nom']}")
            continue
        nom = re.sub(r"^Пул ", "", x["nom"].split(" · биржа")[0])
        if x["tur"] == "pul":
            if x["kod"] == "ochiq":
                L.append(f"🟢 {nom} — қолди {x['left'] or 0:,.0f} $")
            elif x["kod"] == "tugagan":
                L.append(f"✅ {nom} — тугаган")
            else:
                L.append(f"🔒 {nom} — ёпилган" + (f", ⚠ {x['left']:,.0f} $ қолган (Gavhar билан солиштиринг)" if (x.get("left") or 0) > 1 else ""))
        elif x["kod"] == "ochiq":
            L.append(f"{'📑' if x['tur'] == 'fiks' else '↩️'} {nom} — қолди {x['left'] or 0:,.0f} $")
    return L


def holat_royxat_matn(items):
    L = ["📋 Пуллар ва возвратлар ҳолати:"]
    for x in items:
        q = (f" · ишлатилди {x['used']:,.0f} · қолди {x['left']:,.0f} $" if x.get("used") is not None and x.get("left") is not None
             else f" · {x['summa']:,.0f} $" if x.get("summa") else "")
        L.append(f"{'💳' if x['tur'] == 'pul' else '📑' if x['tur'] == 'fiks' else '↩️'} {x['nom']}{q}\n      {x['matn']}"
                 + (f"\n      {x['gavhar']}" if x.get("gavhar") else "")
                 + "".join(f"\n      • {t}" for t in x.get("tafsil") or []))
    return L


def returns_stats(returns):
    returns = [r for r in returns if r.get("_used") or (r.get("chiqdi_qator") is None and not r.get("chiqdi_sana"))]
    used = [r for r in returns if r.get("_used")]
    rest = [r for r in returns if not r.get("_used")]
    sm = lambda rs: sum(float(r["summa"]) for r in rs)
    return dict(n=len(returns), summa=sm(returns), berilgan_n=len(used), berilgan_summa=sm(used),
                qolgan_n=len(rest), qolgan_summa=sm(rest))


# ---------------- тур аниқлаш ----------------
# Gavhar фактураларидан аниқланди: шу серия (модел олд қўшимчаси) билак узук/занжирлар ТОШ НАРХИДА (вакс → Вакс нархи),
# қолганлари (QQG, ZQG …) — Браслет/Дорика нархида. Admin да ўзгартирса бўлади (cfg["seriya_tosh"]).
SERIYA_TOSH_DEFAULT = ["ZSG", "QSG", "QSW", "QSF", "QQF", "QQW"]  # 05.11.2026 фойдаланувчи: ZSG/QSG (вакс тошли занжир/билак узук) — тош (Вакс) нархида; 001491–001496 — Дорика (MODEL_TUR)
SERIYA_TOSH = set(SERIYA_TOSH_DEFAULT)
SERIYA_RE = re.compile(r"\b([A-Z]{2,4})\d{3,}")


def seriya_of(nomi):
    m = SERIYA_RE.search(str(nomi or "").upper())
    return m.group(1) if m else ""


# Модел бўйича нарх тури (энг устун қоида). 04.11.2026 фойдаланувчи: 001491–001496 моделлар (занжир ва билак узук,
# ZSG/QSG/ZQG/QQG …) — Дорика нархида. Калит: тўлиқ модел (ZSG001491) ёки фақат рақам (001491 — занжир ва билак узукка).
MODEL_TUR_DEFAULT = {f"{s}{n:06d}": "Дорика" for s in ("ZSG", "QSG") for n in range(1491, 1497)}
MODEL_TUR = dict(MODEL_TUR_DEFAULT)


LAST_CFG = {}


def seriya_sozla(cfg):
    global SERIYA_TOSH, MODEL_TUR, LAST_CFG
    LAST_CFG = cfg
    SERIYA_TOSH = {x.strip().upper() for x in (cfg.get("seriya_tosh") if cfg.get("seriya_tosh") is not None else SERIYA_TOSH_DEFAULT) if x.strip()}
    MODEL_TUR = {k.upper(): v for k, v in (cfg.get("model_tur") if cfg.get("model_tur") is not None else MODEL_TUR_DEFAULT).items()}


BRASLET_RE = re.compile(r"билак|билаг|браслет|bilak|bilag|braslet|bracelet|б/узук")
ZANJIR_RE = re.compile(r"занжир|zanjir|цеп|chain")
VANCLIF = ("van cl", "vancl", "van-cl", "ван клиф", "ванклиф", "клиф", "klif", "cleef", "clef")


def tur(nomi, kat):
    """Тур аниқлаш. Билак узук (браслет) ва занжир — ТОШ ТУРИДАН ҚАТЪИ НАЗАР категория бўйича нархланади."""
    n, k = (nomi or "").lower(), (kat or "").lower()
    first = n.split()[0] if n.split() else ""
    nom = n.split(" proba")[0]
    md = model_of(nomi)
    mt = MODEL_TUR.get(md)
    if not mt and md and (BRASLET_RE.search(k) or BRASLET_RE.search(nom) or ZANJIR_RE.search(k) or ZANJIR_RE.match(first)):
        mt = MODEL_TUR.get(md[-6:])  # рақам бўйича — фақат занжир ва билак узукка
    if mt:  # модел бўйича белгиланган нарх тури (Admin → «Model bo'yicha narx»)
        return mt
    nom = n.split(" proba")[0]  # товар номи (тош/вазн қисмисиз)
    # 1) категория бўйича (тошга қарамайди): билак узук / браслет → Браслет нархи, занжир → Дорика нархи
    #    истисно: «тош нархидаги» сериялар (QSG, QSW, ZSG …) — тош тури бўйича (пастга ўтади)
    tosh_seriya = seriya_of(nomi) in SERIYA_TOSH
    if (BRASLET_RE.search(k) or BRASLET_RE.search(nom)) and not tosh_seriya:
        return "Браслет"
    if (ZANJIR_RE.search(k) or ZANJIR_RE.match(first)) and not tosh_seriya:
        return "Дорика"
    if "dorika" in n or "дорика" in n:  # Tosh turi - Dorika → Дорика нархи
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


def tosh_tur(nomi):
    """Номдаги тош тури (категорияга қарамай): Бриллиант / Фианит / Вакс / Ван Клиф ёки None."""
    n = (nomi or "").lower()
    if any(v in n for v in VANCLIF):
        return "Ван Клиф"
    if "brilliant" in n or "brilyant" in n or "бриллиант" in n:
        return "Бриллиант"
    if "fianit" in n or "фианит" in n:
        return "Фианит"
    if "waks" in n or "vaks" in n or "вакс" in n:
        return "Вакс"
    return None


def vq_narx(q, r, f1, chegara):
    """Возврат қисмида шу товарнинг нархи: аввал ўз тури (Браслет/Дорика…), мос келмаса — тошли билак узук/занжир учун
    тош тури нархи (Gavhar возвратни шу нархда қабул қилган; бу нарх паст — бизнинг фойдамизга). (нарх, қайси тур) ёки (None, None)."""
    nx = q.get("narxlar", {})
    p = nx.get(r["tur"]) or nx.get("Дорика")
    if p and abs(f1 - p) <= chegara:
        return p, r["tur"]
    if r["tur"] in ("Браслет", "Дорика"):
        tt = tosh_tur(r.get("nomi"))
        p2 = nx.get(tt) if tt else None
        if p2 and abs(f1 - p2) <= chegara and (not p or p2 <= p + chegara):
            return p2, tt
    return None, None


# ---------------- асосий ҳисоб ----------------
def compute(cfg, ledger, returns=None):
    """Ҳисоб. Gavhar пулни ўзи «ёпиб» кейинги пул нархига ўтган бўлса (фактура нархидан аниқланади) —
    бизда ҳам шу пул Gavhar охирги шу нархда ёзган кундан кейин ёпилган ҳисобланади (қолган пули солиштириш учун кўрсатилади)."""
    for t in cfg.get("transhlar", []):
        t.pop("_avto_yop", None)
    rows, trs = _compute(cfg, ledger, returns)
    if cfg.get("avto_yopish", True):
        gp = gavhar_pullar(rows, cfg)
        n = 0
        for t in cfg.get("transhlar", []):
            g = gp.get(t["no"])
            if g and g["yopdi"] and not t.get("yopilgan_sana") and t.get("faol", True) and not t.get("faqat_tekshiruv"):
                t["_avto_yop"] = g["oxirgi"]
                t["_avto_keyingi"] = g["keyingi"]
                n += 1
        if n:
            rows, trs = _compute(cfg, ledger, returns)
    return rows, trs


def _compute(cfg, ledger, returns=None):
    seriya_sozla(cfg)
    returns = returns or []
    for ret in returns:
        ret.pop("_used", None)
    chegara = cfg.get("chegara_usd_gr", 0.3)
    # навбат: биринчи киритилган (тепадаги) пул биринчи тугайди; ⏸ ўчирилган транш ҳисобга кирмайди
    # навбат: энг эски пул биринчи ишлатилади (сана бўйича); ⏸ ўчирилган ва «фақат текширув» пуллари кирмайди
    tr = sorted([t for t in cfg["transhlar"] if t.get("faol", True) and not t.get("faqat_tekshiruv")], key=lambda t: (t["sana"], t["no"]))
    vq = cfg.get("vozvrat_qismlar") or []  # суммаси ва нарх жадвали билан берилган возврат (рўйхатсиз)
    for q in vq:
        q["_used"], q["_n"] = 0.0, 0
    fiks_reset(cfg)
    fk = bool(cfg.get("vozvrat_fiks"))
    left = [float(t["summa_usd"]) for t in tr]
    used = [0.0] * len(tr)
    first_date = [None] * len(tr)
    last_date = [None] * len(tr)

    order = sorted(range(len(ledger)), key=lambda i: (ledger[i]["sana"] or "9999", ledger[i]["faktura"], i))
    out = []
    # ОЙ БОШИ ҚОЛДИҒИ (Gavhar билан келишилган ҳақдорлик): шу санадан бошлаб пул қолдиғи ва возвратлар шу рақамлардан бошланади
    ob = cfg.get("oy_boshi") or {}
    for q in vq + [cfg.get("vozvrat_1") or {}]:
        q.pop("_ob_yop", None)
    cfg.pop("_ob_bizda", None)
    ob_done = [not ob.get("sana")]

    def ob_qoy():
        D = ob["sana"]
        pno = ob.get("pul_no") or max((t for t in tr if t["sana"] <= D), key=lambda t: (t["sana"], t["no"]), default={}).get("no")
        bizda = {}
        for ti, t in enumerate(tr):
            if t["sana"] <= D:
                bizda[t["no"]] = left[ti]
                left[ti] = float(ob.get("pul") or 0) if t["no"] == pno else 0.0
                used[ti] = float(t["summa_usd"]) - left[ti]  # Gavhar билан келишилган: шу пулдан шунча ишлатилган
        cfg["_ob_bizda"] = dict(pullar=bizda, jami=sum(bizda.values()), pul_no=pno)
        for q in vq + ([cfg["vozvrat_1"]] if cfg.get("vozvrat_1") else []):
            q["_ob_yop"] = True
        fiks_reset(cfg)
        ob_done[0] = True

    for i in order:
        if not ob_done[0] and (ledger[i]["sana"] or "9999") >= ob["sana"]:
            ob_qoy()
        r = dict(ledger[i])
        r["_i"] = i  # реестрдаги ўрни — юкланиш тартиби (охирги юкланган — энг катта)
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
        def alloc(commit, start=0):
            lf = left if commit else list(left)
            rem_g, cost, parts = g, 0.0, []
            for ti, t in enumerate(tr):
                if rem_g <= 1e-9:
                    break
                if ti < start:
                    continue
                if r["sana"] and t["sana"] > r["sana"]:  # бу пул фактура санасида ҳали берилмаган эди
                    continue
                if lf[ti] <= 1e-6:
                    continue
                ys, yq = t.get("yopilgan_sana"), t.get("yopilgan_qator")
                if yq is not None:  # ёпилгандан кейин юкланган фактуралар фақат жорий транш нархида
                    if i >= yq:
                        continue
                elif ys and r["sana"] and r["sana"] >= ys:  # эски усул: шу санадан ёпилган
                    continue
                av = t.get("_avto_yop")
                if av and r["sana"] and r["sana"] > av:  # Gavhar бу пулни ёпиб, кейинги пул нархига ўтган
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
                    r.setdefault("sarf", {})[t["no"]] = r.get("sarf", {}).get(t["no"], 0) + c  # шу пулдан айирилган $
            if rem_g > 1e-9:  # пул етмади — фактура санасигача берилган охирги пул нархида
                t = next((x for x in reversed(tr) if not r["sana"] or x["sana"] <= r["sana"]), None)
                sabab_ = "tugagan" if t else "berilmagan"
                t = t or tr[0]  # фактура санасида ҳали пул бўлмаса — энг биринчи пул нархи
                p = t["narxlar"].get(r["tur"]) or t["narxlar"].get("Дорика")
                cost += rem_g * p
                parts.append(("пул етмади", rem_g, p))
                r["pul_yetmadi"] = dict(no=t["no"], sana=t["sana"], sabab=sabab_, p=p)
            elif commit:
                r.pop("pul_yetmadi", None)
            return cost, parts

        cost, parts = alloc(False)
        # нарх транш бўйича тўғри келмаса — возврат рўйхатидан қидирамиз
        if s is not None and (returns or vq or fk) and abs(s / g - cost / g) > chegara:
            v1 = cfg.get("vozvrat_1") or {}
            v1_yopiq = (q_yopiq(v1, i, r) if v1 else False) or v1.get("faol") is False
            ret, price_ok = (None, False) if v1_yopiq else find_return(r, g, returns, s / g, chegara, i)
            if ret is not None and price_ok:
                ret["_used"] = True
                ret["_faktura"] = r.get("faktura")
                ret["_filial2"] = r.get("filial")
                ret["_sana2"] = r.get("sana")
                r1 = float(ret["summa"]) / float(ret["gramm"])
                exp = r1 * g
                r.update(transh="возврат", hisob=exp, hisob_1gr=r1, fakt_1gr=s / g, farq_gr=s / g - r1, farq=s - exp,
                         holat="✓ возврат товар", vozvrat_qism=1,
                         vozvrat_manba=(f"1-қисм (рўйхат): {ret.get('filial') or '—'} дўкони {_dmy(ret.get('sana'))} да қайтарган"
                                        + (f", арт. {ret.get('kod')}" if ret.get('kod') else "") + f", {r1:.2f} $/гр"),
                         izoh=(f"Возврат товар: {ret.get('filial') or '—'} дан {_dmy(ret.get('sana'))} да {r1:.2f} $ дан қайтарилган — "
                               f"шу нархда берилди, тўғри."))
                out.append(r)
                continue
            done = False
            for q in vq:
                if q_yopiq(q, i, r) or q.get("faol") is False:
                    continue
                pq, tq = vq_narx(q, r, s / g, chegara)
                if pq and q["summa_usd"] - q["_used"] >= s - 1:
                    q["_used"] += s; q["_n"] += 1
                    exp = pq * g
                    qosh = "" if tq == r["tur"] else f" (билак узук/занжир — возвратда тош «{tq}» нархида олинган, «{r['tur']}» нархидан арзон)"
                    r.update(transh="возврат", vozvrat_qism=q["no"],
                             vozvrat_manba=f"{q['no']}-қисм (сумма {q['summa_usd']:,.0f} $, биржа {q['birja']}): «{tq}» нархи {pq:.2f} $/гр" + qosh,
                             hisob=exp, hisob_1gr=pq, fakt_1gr=s / g, farq_gr=s / g - pq,
                             farq=s - exp, holat="✓ возврат товар",
                             izoh=f"Возврат товар ({q['no']}-қисм, биржа {q['birja']} нархида): 1 гр {pq:.2f} $ — тўғри." + qosh)
                    out.append(r)
                    done = True
                    break
            if done:
                continue
            for q in vq:  # ўчирилган ёки тугаган возврат қисми нархига мос бўлса — шуни айтамиз
                pq0 = q["narxlar"].get(r["tur"]) or q["narxlar"].get("Дорика")
                if q.get("_ob_yop") and q.get("faol") is not False and vq_narx(q, r, s / g, chegara)[0]:
                    ob_ = cfg.get("oy_boshi") or {}
                    r["_pool_hint"] = (f"Нарх возврат {q['no']}-қисм (биржа {q['birja']}) нархига мос ({s / g:.2f} $/гр), лекин {_dmy(ob_.get('sana'))} "
                                       f"ой бошидан {q['no']}-қисм йўқ — Gavhar билан келишилган возврат фақат {ob_.get('vozvrat_matn') or 'фиксация'}. "
                                       f"Фиксация нархи: {fiks_narx_matn(cfg, r['tur'])}. Янги товар бўлса — {cost / g:.2f} $ (пул нархи).")
                    break
                if q.get("faol") is not False and not q_yopiq(q, i, r) and pq0 and abs(s / g - pq0) <= chegara:
                    r["_pool_hint"] = (f"Нарх возврат {q['no']}-қисм нархига мос ({pq0:.2f} $/гр), лекин {q['no']}-қисм суммаси ТУГАГАН "
                                       f"(қолди {q['summa_usd'] - q['_used']:,.2f} $, бу товар {s:,.2f} $). 2-қисм суммасини Gavhar билан солиштиринг. "
                                       f"Янги товар бўлса — {cost / g:.2f} $ (пул нархи).")
                    break
                if q.get("faol") is False and pq0 and abs(s / g - pq0) <= chegara:
                    r["_pool_hint"] = (f"Бу нарх ({s / g:.2f} $/гр) — возврат {q['no']}-қисм (биржа {q['birja']}) нархи, лекин {q['no']}-қисм "
                                       f"ТЎХТАТИЛГАН (возврат фақат фиксация бўйича: {fiks_narx_matn(cfg, r['tur'])}). "
                                       f"Янги товар бўлса — {cost / g:.2f} $ (пул нархи).")
                    break
            for q in vq:
                if q_yopiq(q, i, r) or q.get("faol") is False:
                    continue
                hit = next((tn for tn, pv in q["narxlar"].items() if abs(s / g - pv) <= chegara), None)
                if hit and hit != r["tur"]:
                    pq = q["narxlar"].get(r["tur"]) or q["narxlar"].get("Дорика")
                    r["_pool_hint"] = (f"Gavhar «{hit}» нархини қўйган (возврат {q['no']}-қисм, {s / g:.2f} $). "
                                       f"Бу товар «{r['tur']}» нархида бўлиши керак: возврат бўлса {pq:.2f} $, янги товар бўлса {cost / g:.2f} $ (транш).")
                    break
            if ret is not None:
                r["_ret_hint"] = (f"Возвратда ўхшаш товар бор ({ret.get('filial')}, {_dmy(ret.get('sana'))}), лекин у "
                                  f"{float(ret['summa']) / float(ret['gramm']):.2f} $ дан қайтарилган — бу ҳам мос эмас.")
        # 1-пул ва возвратга мос келмади — кейинги очиқ пул(лар) нархини текширамиз
        sakrash = 0
        if s is not None and abs(s / g - cost / g) > chegara and parts and parts[0][0] != "пул етмади":
            i0 = next((k for k, t in enumerate(tr) if t["no"] == parts[0][0]), 0)
            for tj in range(i0 + 1, len(tr)):
                c2, p2 = alloc(False, start=tj)
                if p2 and p2[0][0] == tr[tj]["no"] and abs(s / g - c2 / g) <= chegara:
                    sakrash = tj
                    r["_sakrash_dan"] = parts[0][0]
                    break
            if not sakrash and r["tur"] in ("Браслет", "Дорика"):
                # тошли билак узук/занжир: Gavhar тош нархини қўйган — қайси пулнинг тош нархи эканини топамиз (пуллар алмашиб кетмасин)
                tt = tosh_tur(r.get("nomi"))
                p0 = tr[i0]["narxlar"].get(tt) if tt else None
                if tt and not (p0 and abs(s / g - p0) <= chegara):
                    for tj in range(i0 + 1, len(tr)):
                        pj = tr[tj]["narxlar"].get(tt)
                        if pj and abs(s / g - pj) <= chegara and not (r["sana"] and tr[tj]["sana"] > r["sana"]):
                            c2, p2 = alloc(False, start=tj)
                            if p2 and p2[0][0] == tr[tj]["no"]:
                                sakrash = tj
                                r["_tosh_pul"] = (tr[tj]["no"], tt, pj)
                            break
        # пулларнинг ҳеч бирига тушмади — возврат фиксациясидан (тош тури қиймати)
        if fk and s is not None and not sakrash and abs(s / g - cost / g) > chegara:
            fr = fiks_try(cfg, r, g, s, chegara)
            if fr is True:
                out.append(r)
                continue
            if fr:
                r["_pool_hint"] = (f"Возврат фиксациясида «{r['tur']}» бор, лекин 1 гр {s / g:.2f} $ — бу турнинг энг юқори нархидан "
                                   f"({fr[1]:.2f} $) баланд, возврат бўлолмайди.")
        keyingi = []
        if s is not None and parts and parts[0][0] != "пул етмади" and not sakrash:
            i0 = next((k for k, t in enumerate(tr) if t["no"] == parts[0][0]), 0)
            for tj in range(i0 + 1, len(tr)):
                if left[tj] > 1 and not tr[tj].get("yopilgan_sana") and not (r["sana"] and tr[tj]["sana"] > r["sana"]):
                    pj = tr[tj]["narxlar"].get(r["tur"]) or tr[tj]["narxlar"].get("Дорика")
                    keyingi.append(f"{tr[tj]['no']}-пул нархига ({pj:.2f})")
        cost, parts = alloc(True, start=sakrash)
        # Gavhar АРЗОН ёзган бўлса — пулдан фактура суммаси айирилади (Gavhar ҳам шунча олган); арзонлик фарқи хато бўлиб кўринади.
        # ҚИММАТ ёзганда — пулдан тўғри нарх айирилади (ортиқчаси Gavhar дан талаб қилинади).
        if s is not None and cfg.get("arzon_fakt", True) and (cost - s) / g > chegara:
            qayt = cost - s
            for no_, take_, p_ in reversed(parts):
                if no_ == "пул етмади" or qayt <= 1e-9:
                    continue
                ti_ = next(k for k, t in enumerate(tr) if t["no"] == no_)
                back = min(qayt, take_ * p_)
                left[ti_] += back; used[ti_] -= back
                if r.get("sarf", {}).get(no_) is not None:
                    r["sarf"][no_] -= back
                qayt -= back
            r["_arzon_fakt"] = True
        fb = r.get("pul_yetmadi")
        _nom = lambda n: f"№{fb['no']}" if n == "пул етмади" and fb else f"№{n}"
        r["transh"] = " + ".join(f"{_nom(n)}: {gg:.2f} гр × {p:.2f}" for n, gg, p in parts) if len(parts) > 1 else _nom(parts[0][0])
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
                r["izoh"] = togri_izoh(parts, tr, fb, r.get("sana"))
                if r.get("_sakrash_dan") is not None:
                    r["sakrash"] = True
                    r["izoh"] = f"{r.pop('_sakrash_dan')}-пул нархига мос келмади, возвратда ҳам йўқ → " + r["izoh"]
            else:
                r["holat"] = "⚠ ҚИММАТ" if r["farq_gr"] > 0 else "⚠ арзон"
                r["izoh"] = sabab(r, tr, parts)
                if r.pop("_arzon_fakt", None):
                    r["izoh"] = "Пулдан фактура суммаси айирилди (Gavhar шунча олган). " + r["izoh"]
                tp = r.pop("_tosh_pul", None)
                if tp:
                    r["izoh"] = (f"Gavhar №{tp[0]} пулнинг «{tp[1]}» нархини қўйган ({tp[2]:.2f}) — товар шу пулдан ҳисобланди, лекин «{r['tur']}» "
                                 f"нархида бўлиши керак ({r['hisob_1gr']:.2f}). " + r["izoh"])
                if keyingi and not r.get("_pool_hint"):
                    r["izoh"] += " " + ", ".join(keyingi) + " ва возвратга ҳам мос эмас."
                pp = [x for x in parts if x[0] != "пул етмади"]
                if len(pp) > 1:
                    r["izoh"] = (f"{pp[0][0]}-пул шу товарда тугаб, {pp[-1][0]}-пулга ўтди — тўғри нарх аралаш: "
                                 + " + ".join(f"{gg:.2f} гр × {p_:.2f}" for _, gg, p_ in pp) + f" = {cost / g:.2f} $/гр.")
                if fb:
                    r["izoh"] += " " + pul_yetmadi_matn(fb, r.get("sana"))
                if r.get("_pool_hint"):
                    r["izoh"] = r.pop("_pool_hint")
                if r.get("_ret_hint"):
                    r["izoh"] += " " + r.pop("_ret_hint")
                vm = vozvrat_narx_matn(cfg, r["tur"])
                if vm and "возврат" not in r["izoh"].lower():
                    r["izoh"] += f" Возврат нархларига ҳам мос эмас ({vm})."
        if any(n == "пул етмади" for n, _, _ in parts):
            r["holat"] = r.get("holat", "") + " | пул етмади"
        out.append(r)

    if not ob_done[0]:
        ob_qoy()
    trs = []
    for ti, t in enumerate(tr):
        if left[ti] <= 1:
            holat = "Тугаган"
        elif t.get("yopilgan_sana"):
            holat = f"Ёпилган {_dmy(t['yopilgan_sana'])}"
        elif t.get("_avto_yop"):
            holat = f"Gavhar ёпган {_dmy(t['_avto_yop'])}"
        else:
            holat = "Фаол" if used[ti] > 0 else "Ишлатилмаган"
        trs.append(dict(no=t["no"], sana=t["sana"], summa=t["summa_usd"], birja=t["birja_999"],
                        used=used[ti], left=left[ti], holat=holat, birinchi=first_date[ti], oxirgi=last_date[ti]))
    for r in out:  # возврат товар пулдан айирилмайди — «пул етмади» белгиси унга тегишли эмас
        if r.get("transh") == "возврат":
            r.pop("pul_yetmadi", None)
    return belgila(out, cfg), trs


def pul_yetmadi_matn(fb, sana):
    if not fb:
        return ""
    if fb["sabab"] == "berilmagan":
        return (f"Фактура санасида ({_dmy(sana)}) ҳали пул берилмаган эди — энг яқин {fb['no']}-пул ({_dmy(fb['sana'])}) "
                f"нархи ({fb['p']:.2f} $/гр) билан текширилди. Бу эски фактура бўлиши мумкин — «Фақат текшириш» да ташланг.")
    return (f"Шу санагача берилган пуллар тугаган — {fb['no']}-пул ({_dmy(fb['sana'])}) нархи ({fb['p']:.2f} $/гр) билан текширилди. "
            "Янги пул киритилмаган бўлса — киритинг.")


def togri_izoh(parts, tr, fb=None, sana=None):
    """Тўғри қатор учун изоҳ: қайси пул (транш) нархида тўғри."""
    dm = {t["no"]: _dmy(t["sana"]) for t in tr}
    pp = [x for x in parts if x[0] != "пул етмади"]
    if not pp:
        return "Тўғри. " + pul_yetmadi_matn(fb, sana) if fb else "Пул етмади — охирги транш нархида ҳисобланди."
    if len(pp) == 1:
        n, _, p = pp[0]
        return f"{n}-пул ({dm.get(n, '')}) нархида тўғри: {p:.2f} $/гр."
    a, b = pp[0], pp[-1]
    return (f"{a[0]}-пул тугади, {b[0]}-пулга ўтди: " + " + ".join(f"{gg:.2f} гр × {p:.2f} ({n}-пул)" for n, gg, p in pp)
            + " — тўғри.")


def faktura_xulosa(rs, tmap):
    """Фактура бўйича қисқа текширув: қайси пулда тўғри, пул алмашдими, возврат берилдими."""
    L = []
    rv = [r for r in rs if r.get("transh") == "возврат"]
    by_tr, sak = OrderedDict(), OrderedDict()
    for r in rs:
        if r.get("transh") == "возврат":
            continue
        ns = tuple(int(x) for x in re.findall(r"№(\d+)", str(r.get("transh") or "")))
        (sak if r.get("sakrash") else by_tr).setdefault(ns, []).append(r)
    prev_n = None
    for ns, grp_ in sorted(by_tr.items(), key=lambda kv: kv[0]):
        if not ns:
            continue
        okn = sum(1 for r in grp_ if r["holat"].startswith("✓"))
        badn = len(grp_) - okn
        if prev_n is not None and ns[0] != prev_n:
            L.append(f"🔁 {prev_n}-пул тугади → {ns[0]}-пулга ўтди")
        if len(ns) > 1:
            L.append(f"🔁 {ns[0]}-пул тугади → {ns[-1]}-пулга ўтди ({len(grp_)} та товар иккала пул нархида)")
        t0 = tmap.get(ns[-1], {})
        nom = f"{ns[0]}+{ns[-1]}-пул" if len(ns) > 1 else f"{ns[0]}-пул ({_dmy(t0.get('sana'))}, биржа {t0.get('birja_999')})"
        if okn:
            L.append(f"✅ {okn} та товар — {nom} нархида тўғри")
        if badn:
            L.append(f"❌ {badn} та товар — {nom} нархига тўғри келмади")
        prev_n = ns[-1]
    for ns, grp_ in sak.items():
        t0 = tmap.get(ns[0], {})
        L.append(f"➡️ {len(grp_)} та товар — олдинги пулга ва возвратга мос келмади, "
                 f"{ns[0]}-пул ({_dmy(t0.get('sana'))}, биржа {t0.get('birja_999')}) нархида тўғри")
    if rv:
        r1n = sum(1 for r in rv if (r.get("vozvrat_qism") or 1) == 1)
        L.append(f"↩️ {len(rv)} та товар — возврат товар берилди, возврат нархида тўғри"
                 + (f" (рўйхатдан {r1n}, {len(rv) - r1n} таси сумма бўйича)" if 0 < r1n < len(rv) else ""))
    return L


def _dmy(iso):
    return f"{iso[8:10]}.{iso[5:7]}.{iso[:4]}" if iso and len(iso) >= 10 else (iso or "")


def L_xulosa_jami(rows, tmap):
    """Кўп фактура учун умумий текширув: қайси пулда / возвратда нечта товар."""
    L, cnt = ["", "Текширув:"], OrderedDict()
    for r in rows:
        k = r.get("asos") or r.get("xato_turi") or r.get("holat")
        if r["holat"].startswith("⚠"):
            k = "❌ " + (r.get("xato_turi") or "хато")
        else:
            k = ("↩️ " if r.get("transh") == "возврат" else "✅ ") + (r.get("asos") or "тўғри")
        c = cnt.setdefault(k, [0, 0.0])
        c[0] += 1; c[1] += r["summa"] or 0
    for k, (n_, s_) in sorted(cnt.items(), key=lambda kv: (kv[0][0] == "❌", -kv[1][1])):
        L.append(f"{k}: {n_} та · {s_:,.2f} $")
    return L


def write_toplam_report(path, inv):
    """Кўп фактура бир Excel да: фактуралар рўйхати, ҳамма SKU (натижа/асос/изоҳ), хатоликлар."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    F = "Arial"
    hf = PatternFill("solid", fgColor="1F3864")
    red, green = PatternFill("solid", fgColor="F8CBAD"), PatternFill("solid", fgColor="E2EFDA")
    wb = openpyxl.Workbook()

    def head(ws, cols, widths):
        for j, (h, w) in enumerate(zip(cols, widths), 1):
            c = ws.cell(1, j, h)
            c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = hf; c.alignment = Alignment(wrap_text=True)
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
        ws.freeze_panes = "A2"
    ws = wb.active; ws.title = "Фактуралар"
    head(ws, ["№", "Фактура", "Дўкон", "Сана", "SKU", "Грамм", "Фактура $", "Тўғри $", "Фарқ $", "Хато SKU", "Возврат SKU", "Натижа"],
         [5, 16, 22, 11, 7, 10, 13, 13, 11, 9, 10, 26])
    for i, ((fno, fil, sana), rs) in enumerate(inv.items(), 1):
        nb = [r for r in rs if r["holat"].startswith("⚠")]
        fk = sum(r["summa"] or 0 for r in rs); hs = sum(r["hisob"] or 0 for r in rs if r.get("hisob") is not None)
        vals = [i, fno, fil, _dmy(sana), len(rs), round(sum(r["gramm"] for r in rs), 2), round(fk, 2), round(hs, 2), round(fk - hs, 2),
                len(nb), sum(1 for r in rs if r.get("transh") == "возврат"), "❌ Хато бор" if nb else "✅ Тўғри"]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i + 1, j, v); c.font = Font(name=F)
            if j in (7, 8, 9): c.number_format = "#,##0.00"
        ws.cell(i + 1, 12).fill = red if nb else green
    ws.auto_filter.ref = f"A1:L{len(inv) + 1}"
    cols = ["Фактура", "Дўкон", "Сана", "Артикул", "Товар номи", "Категория", "Нарх тури", "Грамм", "Фактура $/гр", "Тўғри $/гр",
            "Фарқ $/гр", "Фактура $", "Тўғри $", "Фарқ $", "Натижа", "Нима бўйича тўғри / Хато тури", "Изоҳ", "Қайси возврат"]
    widths = [15, 20, 11, 11, 42, 13, 11, 8, 11, 11, 10, 11, 11, 10, 10, 34, 60, 45]
    for title, only_bad in (("Ҳамма SKU", False), ("Хатоликлар", True)):
        ws = wb.create_sheet(title); head(ws, cols, widths)
        n = 1
        for (fno, fil, sana), rs in inv.items():
            for r in sorted(rs, key=lambda r: not r["holat"].startswith("⚠")):
                bad = r["holat"].startswith("⚠")
                if only_bad and not bad:
                    continue
                n += 1
                vals = [fno, fil, _dmy(sana), r["kod"], r["nomi"], r.get("kategoriya"), r["tur"], r["gramm"], r.get("fakt_1gr"),
                        r.get("hisob_1gr"), r.get("farq_gr"), r["summa"], r.get("hisob"), r.get("farq"), r.get("natija") or r["holat"],
                        r.get("xato_turi") or r.get("asos") or "", r.get("izoh"), r.get("vozvrat_manba") or ""]
                for j, v in enumerate(vals, 1):
                    c = ws.cell(n, j, v); c.font = Font(name=F)
                    if j in (8, 9, 10, 11): c.number_format = "0.00"
                    if j in (12, 13, 14): c.number_format = "#,##0.00"
                ws.cell(n, 15).fill = red if bad else green
        ws.auto_filter.ref = f"A1:R{n}"
    wb.save(path)


def oy_pullar(rows, cfg, oy):
    """Ой якуни: ҳар пул учун — ой бошида қолдиқ, шу ойда ишлатилди, ой охирида қолдиқ (кейинги ойга ўтди)."""
    oy0 = oy + "-01"
    y, m = int(oy[:4]), int(oy[5:7])
    oy1 = f"{y + (m == 12):04d}-{m % 12 + 1:02d}-01"
    out = []
    for t in sorted([t for t in cfg["transhlar"] if t.get("faol", True) and not t.get("faqat_tekshiruv")], key=lambda t: (t["sana"], t["no"])):
        if t["sana"] >= oy1:
            continue
        ichida = sum((r.get("sarf") or {}).get(t["no"], 0) for r in rows if oy0 <= (r.get("sana") or "") < oy1)
        ob = cfg.get("oy_boshi") or {}
        if ob.get("sana") and oy0 >= ob["sana"] and t["sana"] <= ob["sana"]:  # ой боши қолдиғидан кейин — Gavhar билан келишилган рақамдан
            if t["no"] != (cfg.get("_ob_bizda") or {}).get("pul_no"):
                continue
            oldin = sum((r.get("sarf") or {}).get(t["no"], 0) for r in rows if ob["sana"] <= (r.get("sana") or "") < oy0)
            bosh = float(ob.get("pul") or 0) - oldin
        else:
            oldin = sum((r.get("sarf") or {}).get(t["no"], 0) for r in rows if (r.get("sana") or "") < oy0)
            bosh = t["summa_usd"] - oldin
        out.append(dict(no=t["no"], sana=t["sana"], summa=t["summa_usd"], birja=t["birja_999"], yangi=t["sana"] >= oy0,
                        boshida=bosh, ishlatildi=ichida, oxirida=bosh - ichida, yopilgan=t.get("yopilgan_sana")))
    yetmadi = [r for r in rows if oy0 <= (r.get("sana") or "") < oy1 and r.get("pul_yetmadi")]
    return out, dict(n=len(yetmadi), summa=sum(r["summa"] or 0 for r in yetmadi))


def oy_tovar(rows, oy):
    """Шу ойда Gavhar берган товар: фактуралар, SKU, грамм, сумма — пулдан / возвратдан / пул етмаган."""
    rs = [r for r in rows if (r.get("sana") or "")[:7] == oy]
    sm = lambda xs: sum(r.get("summa") or 0 for r in xs)
    rv = [r for r in rs if r.get("transh") == "возврат"]
    yt = [r for r in rs if r.get("pul_yetmadi") and r.get("transh") != "возврат"]
    return dict(faktura=len({(r.get("faktura"), r.get("filial"), r.get("sana")) for r in rs}), dokon=len({r.get("filial") for r in rs}),
                n=len(rs), gramm=sum(r.get("gramm") or 0 for r in rs), summa=sm(rs),
                puldan=sum(sum((r.get("sarf") or {}).values()) for r in rs), vozvrat=sm(rv), yetmadi=sm(yt),
                xato_n=sum(1 for r in rs if str(r.get("holat", "")).startswith("⚠")),
                ortiqcha=sum(r.get("farq") or 0 for r in rs if str(r.get("holat", "")).startswith("⚠") and (r.get("farq") or 0) > 0),
                kam=sum(r.get("farq") or 0 for r in rs if str(r.get("holat", "")).startswith("⚠") and (r.get("farq") or 0) < 0))


def oy_pullar_matn(rows, cfg, oy):
    pl, yt = oy_pullar(rows, cfg, oy)
    tv = oy_tovar(rows, oy)
    L = [f"📅 {OYLAR[int(oy[5:7])].upper()} {oy[:4]} — ой якуни", "",
         f"📦 Шу ойда Gavhar берган товар: {tv['summa']:,.2f} $",
         f"   {tv['faktura']} та фактура · {tv['dokon']} та дўкон · {tv['n']} та товар · {tv['gramm']:,.2f} гр",
         f"   💳 пулдан (траншлардан): {tv['puldan']:,.2f} $",
         f"   ↩️ возврат товар: {tv['vozvrat']:,.2f} $"]
    if tv["yetmadi"] > 0.5:
        L.append(f"   ⚠ пул етмаган: {tv['yetmadi']:,.2f} $")
    if tv["xato_n"]:
        L.append(f"   ❌ нархи хато: {tv['xato_n']} та товар" + (f", ортиқча ёзилган {tv['ortiqcha']:,.2f} $" if tv["ortiqcha"] else ""))
    L += ["", "💳 Пуллар:"]
    if not pl:
        return "\n".join(L + ["Бу ой учун пул йўқ."])
    for p in pl:
        L.append(f"💳 {p['no']}-пул ({_dmy(p['sana'])}, биржа {p['birja']}, {p['summa']:,.0f} $)")
        L.append(f"   {'шу ойда берилди' if p['yangi'] else 'ой бошида қолдиқ'}: {p['boshida']:,.0f} $")
        L.append(f"   шу ойда ишлатилди: {p['ishlatildi']:,.0f} $")
        if p["oxirida"] <= 1:
            L.append("   ой охирида: тугади ✅")
        else:
            L.append(f"   ой охирида қолдиқ: {p['oxirida']:,.0f} $ → кейинги ойга ўтди" + (f" (ёпилган {_dmy(p['yopilgan'])})" if p["yopilgan"] else ""))
    L += ["", f"Жами шу ойда ишлатилди: {sum(p['ishlatildi'] for p in pl):,.0f} $",
          f"Жами кейинги ойга ўтди: {sum(max(p['oxirida'], 0) for p in pl):,.0f} $"]
    if yt["n"]:
        L.append(f"⚠ {yt['n']} та товар ({yt['summa']:,.0f} $) учун пул етмади — пул киритилмаганми ёки суммаси камми, текширинг.")
    return "\n".join(L)


def pul_ogoh(rows, cfg):
    """Пул етмаган товарлар бўлса ва ботда ишлатилмаётган (ёпилган / ўчирилган / «фақат текширув») пул бор бўлса — огоҳлантириш."""
    yt = [r for r in rows if r.get("pul_yetmadi")]
    if not yt:
        return []
    oxir = max((r.get("sana") or "") for r in yt)
    chetda = []
    for t in sorted(cfg.get("transhlar", []), key=lambda t: t["sana"]):
        if t["sana"] > oxir:
            continue
        sab = ("«📅 фақат эски приход текшируви учун» турибди" if t.get("faqat_tekshiruv") else
               "текширувдан ўчирилган (⏸)" if not t.get("faol", True) else
               f"ёпилган ({_dmy(t['yopilgan_sana'])})" if t.get("yopilgan_sana") else "")
        if sab:
            chetda.append(f"Транш №{t['no']} ({_dmy(t['sana'])}, {t['summa_usd']:,.0f} $) — {sab}")
    L = [f"⚠️ {len(yt)} та товар ({sum(r['summa'] or 0 for r in yt):,.0f} $) учун пул етмади — улар пулдан айирилмади."]
    if chetda:
        L += ["   Ботда ишлатилмаётган пул бор:"] + [f"   • {x}" for x in chetda]
        L.append("   Агар у шу товарлар учун бўлса: Admin → Tekshiruv sozlamasi да «kundalik navbatga qo'shish» / белгини ёқинг "
                 "ёки Transhlar да «Qayta ochish» ни босинг — ҳаммаси ўзи қайта ҳисобланади.")
    else:
        L.append("   Ўша санага пул киритилмаган ёки суммаси кам — пулларни текширинг.")
    return L


def belgila(rows, cfg):
    """Ҳар қаторга қисқа «натижа», «асос» (нима бўйича тўғри) ва «хато тури» ёзади — Excel ва mini app учун."""
    dm = {t["no"]: _dmy(t["sana"]) for t in cfg.get("transhlar", [])}
    for r in rows:
        h = str(r.get("holat", ""))
        tr_ = str(r.get("transh") or "")
        if h.startswith("✓"):
            r["natija"] = "✅ Тўғри"
            if tr_ == "возврат":
                q = r.get("vozvrat_qism") or 1
                r["asos"] = "Возврат рўйхати (1-қисм)" if q == 1 else (f"Возврат {qism_nom(q)}" if isinstance(q, str) else f"Возврат {q}-қисм (сумма бўйича)")
            else:
                ns = list(dict.fromkeys(int(x) for x in re.findall(r"№(\d+)", tr_)))
                fb = r.get("pul_yetmadi")
                if fb and len(ns) <= 1:
                    r["asos"] = (f"{fb['no']}-пул нархи ({dm.get(fb['no'], '')}) — " +
                                 ("фактура санасида ҳали пул берилмаган" if fb["sabab"] == "berilmagan" else "шу санагача пуллар тугаган"))
                elif len(ns) > 1:
                    r["asos"] = f"{ns[0]}+{ns[-1]}-пул (пул алмашган, аралаш нарх)"
                elif ns:
                    r["asos"] = f"{ns[0]}-пул нархи ({dm.get(ns[0], '')})" + (" — олдинги пул ва возвратга мос келмаган" if r.get("sakrash") else "")
                else:
                    r["asos"] = "Тўғри"
            r["xato_turi"] = ""
        elif h.startswith("⚠"):
            r["natija"] = "❌ Хато"
            z = r.get("izoh") or ""
            m = re.match(r"Gavhar «([^»]+)» нархини қўйган", z)
            if m:
                t_ = f"Бошқа тур нархи қўйилган («{m.group(1)}» нархи)"
            elif z.startswith("Бошқа пул"):
                t_ = "Бошқа пул нархи қўйилган"
            elif "биржа нархи" in z:
                t_ = "Бошқа биржа нархида ҳисобланган"
            elif "пул шу товарда тугаб" in z:
                t_ = "Пул алмашган жойда нарх хато"
            elif "Тур аниқланмади" in z:
                t_ = "Товар тури аниқланмади"
            elif "грамм" in h or "нарх йўқ" in h or "транш йўқ" in h:
                t_ = h.replace("⚠", "").strip().capitalize()
            else:
                t_ = "Нарх тўғри келмади"
            yo = "Қиммат" if "ҚИММАТ" in h else "Арзон" if "арзон" in h else ""
            r["xato_turi"] = (f"{yo}: " if yo else "") + t_
            r["asos"] = ""
        else:
            r["natija"], r["asos"], r["xato_turi"] = h, "", ""
    return rows


def sabab(r, tr, parts):
    """Хатонинг содда ва тушунарли сабаби (бир жумла)."""
    f, kerak = r["fakt_1gr"], r["hisob_1gr"]
    tno = parts[0][0]
    if tno == "пул етмади" and r.get("pul_yetmadi"):
        tno = r["pul_yetmadi"]["no"]
    t = next((x for x in tr if x["no"] == tno), tr[-1])
    for name, p in t["narxlar"].items():  # бошқа тур нархи билан ёзилганми
        if name != r["tur"] and name != "Кузсиз" and abs(f - p) <= 0.05:
            return f"Gavhar «{name}» нархини қўйган. Бу товар «{r['tur']}» — нархи {kerak:.2f} $ бўлиши керак."
    for x in tr:  # бошқа транш нархими
        if x["no"] != tno:
            p = x["narxlar"].get(r["tur"])
            if p and abs(f - p) <= 0.05:
                m = (f"Бошқа пул ({_dmy(x['sana'])}) нархи қўйилган. Ҳозир {_dmy(t['sana'])} даги пул ишлатиляпти — "
                     f"нархи {kerak:.2f} $.")
                if x.get("yopilgan_sana") and x["sana"] <= (r.get("sana") or "9"):
                    m += (f" ⚠ {_dmy(x['sana'])} пули ботда ЁПИЛГАН — агар у аслида тугамаган бўлса, "
                          "Transhlar да «Qayta ochish» ни босинг, шунда бу товарлар тўғри чиқади.")
                elif x["sana"] <= (r.get("sana") or "9"):
                    m += (f" ⚠ Ботда {_dmy(x['sana'])} пули тугаган — суммаси тўғри киритилганини текширинг.")
                return m
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
    left_total = sum(t["left"] for t in trs if not str(t["holat"]).startswith(("Ёпилган", "Gavhar ёпган")))
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
    manba = lambda r: ("↩ Возврат " + qism_nom(r.get("vozvrat_qism") or 1)) if r.get("transh") == "возврат" else r.get("transh")
    table(ws, 1, ["Сана", "Фактура №", "Филиал", "Товар коди", "Товар номи", "Категория", "Тур", "Грамм", "Фактура сумма, $", "Фактура $/гр",
                  "Қайси пул / возврат", "Ҳисоб сумма, $", "Ҳисоб $/гр", "Фарқ $/гр", "Фарқ $", "Ҳолати", "Изоҳ"],
          [[r["sana"], r["faktura"], r["filial"], r["kod"], r["nomi"], r["kategoriya"], r["tur"], r["gramm"], r["summa"], r.get("fakt_1gr"),
            manba(r), r.get("hisob"), r.get("hisob_1gr"), r.get("farq_gr"), r.get("farq"), r.get("holat"), r.get("izoh")] for r in rows],
          [11, 11, 20, 14, 55, 14, 11, 8, 12, 10, 10, 12, 10, 9, 9, 14, 50],
          {8: "0.00", 9: "#,##0.00", 10: "0.00", 12: "#,##0.00", 13: "0.00", 14: "0.00", 15: "#,##0.00"})
    for i, r in enumerate(rows, 2):
        ws.cell(i, 16).fill = green if str(r.get("holat")).startswith("✓") else red
    ws.auto_filter.ref = f"A1:Q{len(rows)+1}"

    # возвратдан берилган товарлар — ҳар манба (2-қисм, сентябрь/октябрь фиксацияси) бўйича
    VT = vozvrat_tovarlar(rows)
    if VT:
        ws = wb.create_sheet("Возвратдан берилган")
        data = []
        for q, rs_ in VT.items():
            for r in rs_:
                data.append([qism_nom(q), r["sana"], r["faktura"], r["filial"], r["kod"], model_of(r["nomi"]), r["tur"], r["nomi"], r["gramm"],
                             r.get("fakt_1gr"), r["summa"]])
            data.append([f"ЖАМИ {qism_nom(q)}", None, None, None, None, None, None, f"{len(rs_)} та товар",
                         round(sum(r["gramm"] or 0 for r in rs_), 2), None, round(sum(r["summa"] or 0 for r in rs_), 2)])
        table(ws, 1, ["Қайси возврат", "Сана", "Фактура №", "Филиал", "Товар коди", "Модель", "Тур", "Товар номи", "Грамм", "1 гр, $", "Сумма, $"],
              data, [24, 11, 15, 18, 12, 12, 11, 50, 8, 9, 12], {9: "0.00", 10: "0.00", 11: "#,##0.00"})
        for i, row in enumerate(data, 2):
            if str(row[0]).startswith("ЖАМИ"):
                for j in range(1, 12):
                    ws.cell(i, j).font = Font(name=F, bold=True)
        ws.auto_filter.ref = f"A1:K{len(data) + 1}"

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
    cols = ["№", "Сана", "Фактура №", "Дўкон", "Артикул", "Модель", "Категория", "Нарх тури", "Грамм", "Фактура $/гр", "Тўғри $/гр",
            "Фарқ $/гр", "Фактура $", "Тўғри $", "Фарқ $", "Хато тури", "Сабаб (тушунтириш)", "Возврат рўйхатида", "Возврат нархида фарқ $"]
    header(ws, 1, cols, [5, 11, 15, 18, 11, 12, 14, 11, 8, 11, 11, 10, 11, 11, 10, 34, 70, 45, 12])
    red, amb = PatternFill("solid", fgColor="F8CBAD"), PatternFill("solid", fgColor="FCE4D6")
    for n, r in enumerate(bad, 1):
        vh, vp = vozvrat_holati(r, bym)
        vfarq = (r["summa"] - vp * r["gramm"]) if (vp and r.get("summa") is not None) else None
        vals = [n, r.get("sana"), r.get("faktura"), r.get("filial"), r.get("kod"), model_of(r.get("nomi")), r.get("kategoriya"), r.get("tur"),
                r.get("gramm"), r.get("fakt_1gr"), r.get("hisob_1gr"), r.get("farq_gr"), r.get("summa"), r.get("hisob"), r.get("farq"),
                r.get("xato_turi") or r.get("holat"), r.get("izoh"), vh, vfarq]
        for j, v in enumerate(vals, 1):
            c = ws.cell(n + 1, j, v)
            c.font = Font(name=F)
            if j in (9, 10, 11, 12):
                c.number_format = "0.00"
            if j in (13, 14, 15, 19):
                c.number_format = "#,##0.00"
        ws.cell(n + 1, 16).fill = red if (r.get("farq") or 0) > 0 else amb
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:S{len(bad) + 1}"
    # 3) хато турлари ва категориялар бўйича жамлама
    ws = wb.create_sheet("Хато турлари")
    agg = defaultdict(lambda: [0, 0.0, 0.0])
    for r in bad:
        a = agg[(r.get("kategoriya") or "—", r.get("xato_turi") or r.get("holat"))]
        a[0] += 1; a[1] += r.get("gramm") or 0; a[2] += r.get("farq") or 0
    header(ws, 1, ["Категория", "Хато тури", "Товар сони", "Грамм", "Фарқ $"], [18, 52, 12, 12, 14])
    for n, ((k, t_), (c_, g_, f_)) in enumerate(sorted(agg.items(), key=lambda kv: -abs(kv[1][2])), 2):
        for j, v in enumerate([k, t_, c_, round(g_, 2), round(f_, 2)], 1):
            ws.cell(n, j, v).font = Font(name=F)
        ws.cell(n, 5).number_format = "#,##0.00"
    # 4) Gavhar қайси тур нархини қўйган (Дорика / Браслет / Вакс …) — гуруҳ ва товарма-товар
    if LAST_CFG.get("transhlar"):
        ws = wb.create_sheet("Қайси нарх қўйилган", 1)
        T = xato_turlari(bad, LAST_CFG)
        header(ws, 1, ["Хато (Gavhar нима қилган)", "Товар сони", "Қиммат", "Арзон", "Грамм", "Фарқ $", "Тушунтириш (мисол)", "Мисол товарлар"],
               [34, 11, 9, 9, 11, 13, 70, 60])
        for n, x in enumerate(T, 2):
            for j, v in enumerate([x["tur"], x["n"], x["qimmat"], x["arzon"], round(x["gramm"], 2), round(x["farq"], 2), x["izoh"], "; ".join(x["misol"])], 1):
                ws.cell(n, j, v).font = Font(name=F, bold=(j == 1))
            ws.cell(n, 6).number_format = "#,##0.00"
        r0 = len(T) + 4
        header(ws, r0, ["Хато", "Сана", "Фактура №", "Дўкон", "Модель", "Тур (бўлиши керак)", "Грамм", "Фактура $/гр", "Тўғри $/гр", "Фарқ $", "Тушунтириш"],
               [34, 11, 15, 18, 12, 14, 8, 11, 11, 11, 70])
        for n, r in enumerate(sorted(bad, key=lambda r: (xato_turi_aniq(r, LAST_CFG)[0], r.get("sana") or "")), r0 + 1):
            k, iz = xato_turi_aniq(r, LAST_CFG)
            for j, v in enumerate([k, r.get("sana"), r.get("faktura"), r.get("filial"), model_of(r.get("nomi")), r.get("tur"), r.get("gramm"),
                                   r.get("fakt_1gr"), r.get("hisob_1gr"), r.get("farq"), iz], 1):
                ws.cell(n, j, v).font = Font(name=F)
        ws.freeze_panes = "A2"
    wb.save(path)
    return dict(n=len(bad), farq=sum(r.get("farq") or 0 for r in bad), faktura=len(inv))


def write_invoice_report(path, fno, fil, sana_t, rs, trn, tmap=None):
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
            ("Возврат билан тўғриланди", "; ".join(
                f"{qism_nom(qn)}: {sum(1 for r in rs if r.get('transh') == 'возврат' and (r.get('vozvrat_qism') or 1) == qn)} та" for qn in
                sorted({r.get("vozvrat_qism") or 1 for r in rs if r.get("transh") == "возврат"}, key=str)) or "—"),
            ("ХУЛОСА", "✅ ТЎҒРИ" if not nb else f"❌ ХАТО БОР — {len(nb)} та позиция, {farq:+,.2f} $")]
    for i, (k, v) in enumerate(head, 1):
        ws.cell(i, 1, k).font = Font(name=F, bold=True)
        c = ws.cell(i, 2, v)
        c.font = Font(name=F, bold=k == "ХУЛОСА", color=("C00000" if (k == "ХУЛОСА" and nb) else "1E7B3C" if k == "ХУЛОСА" else "000000"))
    xl = faktura_xulosa(rs, tmap or {})
    if xl:
        r0 = len(head) + 2
        ws.cell(r0, 1, "ТЕКШИРУВ").font = Font(name=F, bold=True)
        for k, line in enumerate(xl):
            c = ws.cell(r0 + 1 + k, 1, line)
            c.font = Font(name=F, color=("C00000" if line.startswith("❌") else "1E7B3C" if line[:1] in "✅↩" else "7F6000"))
        head = head + [None] * (len(xl) + 2)
    start = len(head) + 2
    hdr = ["№", "Артикул / код", "Товар номи", "Категория", "Нарх тури", "Грамм", "Фактура $/гр", "Тўғри $/гр", "Фарқ $/гр",
           "Фактура $", "Тўғри $", "Фарқ $", "Натижа", "Нима бўйича тўғри / Хато тури", "Изоҳ (тушунтириш)", "Қайси возврат"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(start, j, h)
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(wrap_text=True)
    red, green = PatternFill("solid", fgColor="F8CBAD"), PatternFill("solid", fgColor="E2EFDA")
    for i, r in enumerate(sorted(rs, key=lambda r: (not r["holat"].startswith("⚠"), -abs(r.get("farq") or 0))), 1):
        vals = [i, r["kod"], r["nomi"], r.get("kategoriya"), r["tur"], r["gramm"], r.get("fakt_1gr"), r.get("hisob_1gr"), r.get("farq_gr"),
                r["summa"], r.get("hisob"), r.get("farq"), r.get("natija") or r["holat"], r.get("xato_turi") or r.get("asos") or "",
                r.get("izoh"), r.get("vozvrat_manba") or ""]
        for j, v in enumerate(vals, 1):
            c = ws.cell(start + i, j, v)
            c.font = Font(name=F)
            if j in (6, 7, 8, 9):
                c.number_format = "0.00"
            if j in (10, 11, 12):
                c.number_format = "#,##0.00"
            if j in (14, 15, 16):
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(start + i, 13).fill = red if r["holat"].startswith("⚠") else green
        ws.cell(start + i, 14).fill = red if r["holat"].startswith("⚠") else green
    ws.auto_filter.ref = f"A{start}:P{start + len(rs)}"
    for j, w in enumerate([5, 12, 45, 14, 11, 8, 11, 11, 10, 11, 11, 10, 11, 36, 70, 55], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.column_dimensions["A"].width = 18
    ws.freeze_panes = ws.cell(start + 1, 1)
    wb.save(path)


def run(transh_path, ledger_path, out_path, files=(), returns_path=None):
    """Фактураларни реестрга қўшади, ҳамма ҳисобни қайта қилади, ҳисобот ёзади. Қисқа матн қайтаради."""
    cfg = json.load(open(transh_path, encoding="utf-8"))
    ledger = load_ledger(ledger_path)
    seen = {key(r) for r in ledger}
    oldin_id = {str(r["id"]): r for r in ledger if r.get("id")}
    added, dup, new_rows = 0, 0, []
    warns, file_keys, qayta = [], set(), []
    skip_reset()
    for f in files:
        for r in read_invoice(f, warns):
            k = key(r)
            file_keys.add(k)
            if k in seen:
                dup += 1
                continue
            seen.add(k)
            o = oldin_id.get(str(r.get("id") or ""))
            if o:  # шу артикул олдин бошқа фактурада берилган — возвратдан кейин қайта берилган товар
                qayta.append((r, o))
            ledger.append(r)
            new_rows.append(r)
            added += 1
    returns = load_returns(returns_path)
    trs_old, q_old = None, {}
    if added:
        _, trs_old = compute(cfg, ledger[:len(ledger) - added], returns)
        q_old = {q["no"]: q["qolgan"] for q in qism_stats(cfg)}
        save_ledger(ledger_path, ledger)
    rows, trs = compute(cfg, ledger, returns)
    s = write_report(out_path, cfg, rows, trs, returns)

    L = []
    # файлдаги ҲАММА қатор (олдин юкланганлари ҳам) — текширув натижаси кўрсатилади
    new_res = [r for r in rows if key(r) in file_keys]
    if files and dup and not added:
        L.append("ℹ️ Бу фактура олдин юкланган — реестрга қайта қўшилмади (пул икки марта айирилмайди). Текширув натижаси:")
    elif files and dup:
        L.append(f"ℹ️ {dup} та товар олдин юкланган — улар қайта қўшилмади. Ҳаммасининг текшируви:")
    if qayta:
        L.append(f"↩️ {len(qayta)} та товар артикули олдин бошқа фактурада бўлган (возвратдан кейин қайта берилган) — реестрга қўшилди: "
                 + "; ".join(f"{model_of(r['nomi']) or r['kod']} ({o.get('filial')} {_dmy(o.get('sana'))} → {r.get('filial')})" for r, o in qayta[:5])
                 + (" …" if len(qayta) > 5 else ""))
    if new_res and not cfg["transhlar"]:
        L.append(f"🧾 Янги фактура: {len(new_res)} позиция, {sum(r['gramm'] for r in new_res):,.2f} гр, "
                 f"{sum((r['summa'] or 0) for r in new_res):,.2f} $ — реестрга қўшилди.\n"
                 "⚠ Ҳали транш (пул ва нархлар) киритилмаган, шунинг учун нарх текширилмади.\n"
                 "Gavhar нарх хабарини шу ерга forward қилинг ва суммани ёзинг — кейин фактура автомат текширилади.")
        new_res = []
    inv = OrderedDict()
    for r in new_res:
        inv.setdefault((r["faktura"], r["filial"], r["sana"]), []).append(r)
    if new_res and len(inv) > 5:  # кўп фактура бирдан — қисқа жамланма + битта Excel
        tmap = {t["no"]: t for t in cfg["transhlar"]}
        jf = sum(r["summa"] or 0 for r in new_res)
        jh = sum(r["hisob"] or 0 for r in new_res if r.get("hisob") is not None)
        badinv = [(k, rs) for k, rs in inv.items() if any(r["holat"].startswith("⚠") for r in rs)]
        L.append(f"📥 {len(inv)} та фактура текширилди · {len(new_res)} та товар · {jf:,.2f} $")
        L.append(f"✅ Тўғри: {len(inv) - len(badinv)} та фактура   ❌ Хато бор: {len(badinv)} та")
        if jf - jh > 0.5:
            L.append(f"👉 Жами ОРТИҚЧА ёзилган: {jf - jh:,.2f} $")
        elif jf - jh < -0.5:
            L.append(f"👉 Жами КАМ ёзилган: {abs(jf - jh):,.2f} $")
        L += L_xulosa_jami(new_res, tmap)
        L += xato_turlari_matn(new_res, cfg)
        yop_n = sum(1 for r in new_res if "ЁПИЛГАН" in (r.get("izoh") or ""))
        tug_n = sum(1 for r in new_res if "пули тугаган" in (r.get("izoh") or ""))
        if yop_n:
            L.append(f"\n⚠ {yop_n} та товарда Gavhar ЁПИЛГАН пул нархини қўйган. Пул аслида тугамаган бўлса — Transhlar да «Qayta ochish» ни босинг.")
        if tug_n:
            L.append(f"\n⚠ {tug_n} та товарда Gavhar ботда тугаган пул нархини қўйган — ўша пул суммаси тўғрилигини текширинг.")
        if badinv:
            L.append("\n❌ Хато бор фактуралар:")
            for (fno, fil, sana), rs in sorted(badinv, key=lambda kv: -sum(r["farq"] or 0 for r in kv[1] if r["holat"].startswith("⚠")))[:25]:
                f_ = sum(r["farq"] or 0 for r in rs if r["holat"].startswith("⚠"))
                L.append(f"• {fno} · {fil} · {_dmy(sana)} — {sum(1 for r in rs if r['holat'].startswith('⚠'))} та хато, {f_:+,.2f} $")
            if len(badinv) > 25:
                L.append(f"• ... яна {len(badinv) - 25} та (Excel да)")
        L.append("📑 Ҳамма фактура ва SKU — битта Excel файлда (Фактуралар / Ҳамма SKU / Хатоликлар)")
        try:
            write_toplam_report(str(Path(out_path).with_name("Fakturalar_jami.xlsx")), inv)
        except Exception as e:
            L.append(f"   (Excel ёзилмади: {e})")
    elif new_res:
        if len(inv) > 1:  # бир нечта фактура — тепада битта умумий қатор
            nb_all = [k for k, rs in inv.items() if any(r["holat"].startswith("⚠") for r in rs)]
            L.append(f"📥 {len(inv)} та фактура · {len(new_res)} та товар · {sum(r['summa'] or 0 for r in new_res):,.2f} $ — "
                     + ("ҳаммаси ✅ тўғри" if not nb_all else f"✅ {len(inv) - len(nb_all)} та тўғри, ❌ {len(nb_all)} тасида хато"))
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
            if not nb:  # тўғри фактура — сарлавҳа + ҳар манба (пул / возврат) бўйича битта қатор
                L.append(f"✅ {fno or '—'} · {fil or '—'} · {sana_t} · {len(rs)} та · {fakt:,.2f} $")
                L += manba_qatorlar(rs, tmap, cfg)
                try:
                    safe = re.sub(r"[^\w\-]+", "_", f"{fno}_{fil}").strip("_")[:60] or "faktura"
                    write_invoice_report(str(Path(out_path).with_name(f"Faktura_{safe}.xlsx")), fno, fil, sana_t, rs, trn, tmap)
                except Exception:
                    pass
                continue
            L.append("━━━━━━━━━━━━━━━")
            L.append(f"🧾 {fno or '—'} · {fil or '—'} · {sana_t}")
            L.append(f"   {len(rs)} та товар · {sum(r['gramm'] for r in rs):,.2f} гр")
            L += manba_qatorlar(rs, tmap, cfg)
            L.append("Текширув:")
            L += faktura_xulosa(rs, tmap)
            if rv:
                L.append(f"\n↩️ Возврат тафсилоти:")
                for qn in sorted({r.get("vozvrat_qism") or 1 for r in rv}, key=str):
                    qr = [r for r in rv if (r.get("vozvrat_qism") or 1) == qn]
                    if isinstance(qn, str):
                        L.append(f"   {qism_nom(qn)} — {len(qr)} та, {sum(r['summa'] or 0 for r in qr):,.2f} $:")
                        bt = OrderedDict()
                        for r in qr:
                            b_ = bt.setdefault(r["tur"], [0, 0.0, 0.0]); b_[0] += 1; b_[1] += r["gramm"] or 0; b_[2] += r["summa"] or 0
                        for tn, (cnt, gr, sm_) in sorted(bt.items(), key=lambda kv: -kv[1][2]):
                            L.append(f"     • {tn}: {cnt} та, {gr:.2f} гр, {sm_:,.2f} $ — фиксация қийматидан айирилди")
                    elif qn == 1:
                        L.append(f"   1-қисм (возврат рўйхати) — {len(qr)} та, {sum(r['summa'] or 0 for r in qr):,.2f} $:")
                        for r in qr[:6]:
                            L.append(f"     • {r['tur']} {r['gramm']:.2f} гр — {r.get('vozvrat_manba', '').split(': ', 1)[-1]}")
                        if len(qr) > 6:
                            L.append(f"     • ... яна {len(qr) - 6} та (Excel да)")
                    else:
                        q0 = next((q for q in cfg.get("vozvrat_qismlar", []) if q["no"] == qn), {})
                        L.append(f"   {qn}-қисм (сумма бўйича, биржа {q0.get('birja')}) — {len(qr)} та, {sum(r['summa'] or 0 for r in qr):,.2f} $:")
                        bt = OrderedDict()
                        for r in qr:
                            b_ = bt.setdefault(r["tur"], [0, 0.0, r.get("hisob_1gr") or 0])
                            b_[0] += 1; b_[1] += r["gramm"] or 0
                        for tn, (cnt, gr, pr) in sorted(bt.items(), key=lambda kv: -kv[1][1]):
                            L.append(f"     • {tn}: {cnt} та, {gr:.2f} гр × {pr:.2f} $ — возврат нархига мос")
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
                write_invoice_report(str(Path(out_path).with_name(f"Faktura_{safe}.xlsx")), fno, fil, sana_t, rs, trn, tmap)
            except Exception as e:  # ҳисобот ёзилмаса ҳам матн чиқсин
                L.append(f"   (фактура Excel ёзилмади: {e})")
    L += skip_matn()
    if warns:
        L.append("🧮 Фактура арифметикасида хато:\n" + "\n".join("   • " + w for w in warns[:15]))
    if new_res:
        act = next((t for t in trs if t["holat"] in ("Фаол", "Ишлатилмаган")), None)
        L.append("━━━━━━━━━━━━━━━")
        if trs_old:
            old_left = {t["no"]: t["left"] for t in trs_old}
            for t in trs:
                ol = old_left.get(t["no"], t["summa"])
                if ol > 1 and t["left"] <= 1:
                    nxt = next((x for x in trs[trs.index(t) + 1:] if x["left"] > 1), None)
                    L.append(f"🔔 Транш №{t['no']} ({_dmy(t['sana'])}) пули шу фактурада тугади."
                             + (f" Қолган товарлар Транш №{nxt['no']} ({_dmy(nxt['sana'])}) нархида текширилди." if nxt else
                                " Бошқа транш йўқ — янги пул киритинг!"))
                elif t["summa"] and ol / t["summa"] > 0.1 >= t["left"] / t["summa"] and t["left"] > 1:
                    L.append(f"⚠ Транш №{t['no']} да фақат {t['left']:,.0f} $ қолди ({t['left'] / t['summa'] * 100:.0f}%).")
        for q in qism_stats(cfg):
            if q_old.get(q["no"], 0) > 100 and q["qolgan"] <= 100:
                L.append(f"🔔 Возврат {q['no']}-қисм ({q['summa']:,.0f} $) тугади. Якунлаш: /vozvrat_yop {q['no']}")
        L.append(f"💰 Очиқ пуллар қолдиғи: {s['left']:,.0f} $")
        L += holat_qisqa(holat_royxat(cfg, trs, returns_stats(returns) if returns else None, rows))
        og = pul_ogoh(new_res, cfg)
        if og:
            L += [""] + og[:3]
        L.append("Батафсил: /holat · /fiks · Excel · mini app")
        return "\n".join(L)
    s["qismlar"] = qism_stats(cfg)
    s["royxat"] = holat_royxat(cfg, trs, returns_stats(returns) if returns else None, rows)
    s["off"] = [f"Транш №{t['no']} ({_dmy(t['sana'])})" for t in cfg["transhlar"] if not t.get("faol", True)]
    s["pul_ogoh"] = pul_ogoh(rows, cfg)
    s["xato_turlari"] = xato_turlari_matn(rows, cfg)
    s["eski"] = [f"Транш №{t['no']} ({_dmy(t['sana'])})" for t in cfg["transhlar"] if t.get("faol", True) and t.get("faqat_tekshiruv")]
    if returns:
        s["v1"] = dict(returns_stats(returns), yopilgan=(cfg.get("vozvrat_1") or {}).get("yopilgan_sana"))
    import os
    if (os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("RAILWAY_PROJECT_ID")) and not os.environ.get("RAILWAY_VOLUME_MOUNT_PATH"):
        s["volume_ogoh"] = ("⚠️ Railway да Volume уланмаган — янги файл юклаганда маълумотлар ўчиб кетади! "
                            "Railway → сервис → Attach volume → /data.")
    L += holat_matn(rows, trs, s)
    return "\n".join(L)


def holat_matn(rows, trs, s):
    """Содда ҳолат матни (бот учун)."""
    L = [f"📊 ҲОЛАТ · {dt.date.today():%d.%m.%Y}", ""]
    if s.get("volume_ogoh"):
        L = [s["volume_ogoh"], ""] + L
    if not trs:
        return L + ["Ҳали транш (пул) киритилмаган.", "Gavhar нарх хабарини шу ерга forward қилинг — бот суммани сўрайди."]
    L.append(f"💰 Gavhar да пулингиз (очиқ пуллар): {s['left']:,.0f} $")
    if s.get("royxat"):
        L += [""] + holat_royxat_matn(s["royxat"])
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
    if s.get("pul_ogoh"):
        L += [""] + s["pul_ogoh"]
    L += s.get("xato_turlari") or []
    return L


def otgan_oy(rows, bugun=None):
    """Фактура(лар) ўтган ой(лар)га тегишли бўлса — ойини қайтаради (жорий ой ҳисобига қўшилмайди)."""
    joriy = (bugun or dt.date.today()).strftime("%Y-%m")
    oylar = {(r.get("sana") or "")[:7] for r in rows if r.get("sana")}
    return max(oylar) if oylar and max(oylar) < joriy else None


def eski_faktura(cfg, rows):
    """Фактура кундалик навбатдаги биринчи пулдан ОЛДИНГИ санада бўлса — эски фактура (санаси қайтарилади), акс ҳолда None."""
    q = [t["sana"] for t in cfg.get("transhlar", []) if t.get("faol", True) and not t.get("faqat_tekshiruv")]
    ds = sorted(r["sana"] for r in rows if r.get("sana"))
    if not q or not ds:
        return None
    return ds[-1] if ds[-1] < min(q) else None


def sana_tekshir(cfg, rows, returns=None):
    """АВТО режим (эски фактура / ойлик приход): пул айирилмайди, ҳар товар учун бот ўзи топади:
    ① фактура санасидан олдинги энг яқин ёқилган пул нархи → ② возвратлар (санаси олдин бўлгани) →
    ③ бошқа ёқилган пуллар → ④ ҳеч бирига тушмаса ❌ (① даги пул нархи бўйича)."""
    seriya_sozla(cfg)
    returns = returns or []
    for x in returns:
        x.pop("_used", None)
    chegara = cfg.get("chegara_usd_gr", 0.3)
    tr = sorted([t for t in cfg["transhlar"] if t.get("faol", True)], key=lambda t: t["sana"])
    vq = cfg.get("vozvrat_qismlar") or []
    for q in vq:
        q["_used"], q["_n"] = 0.0, 0
    v1 = cfg.get("vozvrat_1") or {}
    fiks_reset(cfg)
    narx = lambda t, tur_: t["narxlar"].get(tur_) or t["narxlar"].get("Дорика")
    out, stat = [], defaultdict(lambda: [0, 0.0])
    for r0 in sorted(rows, key=lambda r: (r.get("sana") or "", r.get("faktura") or "")):
        r = dict(r0)
        r["tur"] = tur(r["nomi"], r["kategoriya"])
        g, sm = to_num(r["gramm"]) or 0.0, to_num(r["summa"])
        r.update(gramm=g, summa=sm, izoh="", transh="")
        if g <= 0 or sm is None:
            r.update(holat="⚠ грамм ёки нарх йўқ", hisob=None, hisob_1gr=None, fakt_1gr=None, farq_gr=None, farq=None)
            out.append(r)
            continue
        f1 = sm / g
        r["fakt_1gr"] = f1
        if not tr:
            r.update(holat="⚠ транш йўқ", hisob=None, hisob_1gr=None, farq_gr=None, farq=None,
                     izoh="Бу санага пул (транш) ёқилмаган ёки киритилмаган")
            out.append(r)
            continue
        sana = r.get("sana") or ""
        oldin = [t for t in tr if t["sana"] <= sana]
        ochiq = [t for t in oldin if not (t.get("yopilgan_sana") and t["yopilgan_sana"] <= sana)]
        # асос — навбатдаги (энг эски, ёпилмаган) пул; ⏸ билан ўчирилганлари ҳисобга кирмайди
        asos = ochiq[0] if ochiq else (oldin[-1] if oldin else tr[0])

        def ok(t, izoh, sak=False):
            p = narx(t, r["tur"])
            r.update(transh=f"№{t['no']}", hisob=p * g, hisob_1gr=p, farq_gr=f1 - p, farq=sm - p * g, holat="✓ тўғри", izoh=izoh)
            if sak:
                r["sakrash"] = True
            stat[t["no"]][0] += 1; stat[t["no"]][1] += sm
        p0 = narx(asos, r["tur"])
        if abs(f1 - p0) <= chegara:
            ok(asos, f"{asos['no']}-пул ({_dmy(asos['sana'])}) нархида тўғри: {p0:.2f} $/гр.")
            out.append(r)
            continue
        # возврат 1-қисм (рўйхат)
        ret, pok = (None, False) if (v1.get("faol") is False or (v1.get("yopilgan_sana") and sana >= v1["yopilgan_sana"])) else \
            find_return(r, g, returns, f1, chegara)
        if ret is not None and pok:
            ret["_used"] = True
            rp = float(ret["summa"]) / float(ret["gramm"])
            r.update(transh="возврат", vozvrat_qism=1, hisob=rp * g, hisob_1gr=rp, farq_gr=f1 - rp, farq=sm - rp * g,
                     holat="✓ возврат товар",
                     vozvrat_manba=f"1-қисм (рўйхат): {ret.get('filial') or '—'} дўкони {_dmy(ret.get('sana'))} да қайтарган"
                                   + (f", арт. {ret.get('kod')}" if ret.get("kod") else "") + f", {rp:.2f} $/гр",
                     izoh=f"Возврат товар: {ret.get('filial') or '—'} дан {_dmy(ret.get('sana'))} да {rp:.2f} $ дан қайтарилган — тўғри.")
            out.append(r)
            continue
        hit = None
        for q in vq:  # возврат 2-қисм (сумма)
            if q.get("faol") is False or (q.get("yopilgan_sana") and sana >= q["yopilgan_sana"]):
                continue
            pq, tq = vq_narx(q, r, f1, chegara)
            if pq and q["summa_usd"] - q["_used"] >= sm - 1:
                q["_used"] += sm; q["_n"] += 1
                r.update(transh="возврат", vozvrat_qism=q["no"], hisob=pq * g, hisob_1gr=pq, farq_gr=f1 - pq, farq=sm - pq * g,
                         holat="✓ возврат товар",
                         vozvrat_manba=f"{q['no']}-қисм (сумма {q['summa_usd']:,.0f} $, биржа {q['birja']}): «{r['tur']}» нархи {pq:.2f} $/гр",
                         izoh=f"Возврат товар ({q['no']}-қисм, биржа {q['birja']} нархида): 1 гр {pq:.2f} $ — тўғри.")
                hit = q
                break
        if hit:
            out.append(r)
            continue
        # бошқа ёқилган пуллар (санаси яқинидан бошлаб)
        boshqa = [t for t in reversed(oldin) if t is not asos] + [t for t in tr if t["sana"] > sana]
        t2 = next((t for t in boshqa if abs(f1 - narx(t, r["tur"])) <= chegara), None)
        if t2:
            ok(t2, f"{asos['no']}-пул ({_dmy(asos['sana'])}) нархига мос эмас, возвратда йўқ → "
                   f"{t2['no']}-пул ({_dmy(t2['sana'])}) нархида тўғри: {narx(t2, r['tur']):.2f} $/гр.", sak=True)
            out.append(r)
            continue
        if cfg.get("vozvrat_fiks") and fiks_try(cfg, r, g, sm, chegara) is True:  # возврат фиксацияси (тош тури қиймати)
            out.append(r)
            continue
        # ❌ хато — асос пул нархи бўйича
        r.update(transh=f"№{asos['no']}", hisob=p0 * g, hisob_1gr=p0, farq_gr=f1 - p0, farq=sm - p0 * g,
                 holat="⚠ ҚИММАТ" if f1 > p0 else "⚠ арзон")
        r["izoh"] = sabab(r, tr, [(asos["no"], g, p0)])
        if len(tr) > 1:
            r["izoh"] += " Бошқа пуллар ва возвратга ҳам мос эмас."
        out.append(r)
    trs = [dict(no=t["no"], sana=t["sana"], summa=t["summa_usd"], birja=t["birja_999"], used=stat[t["no"]][1],
                left=t["summa_usd"] - stat[t["no"]][1], holat="Текширувда", birinchi=None, oxirgi=None, n=stat[t["no"]][0]) for t in tr]
    return belgila(out, cfg), trs


OYLAR = ["", "Январ", "Феврал", "Март", "Апрел", "Май", "Июн", "Июл", "Август", "Сентябр", "Октябр", "Ноябр", "Декабр"]


# ---------------- ХАТО ТУРЛАРИ: Gavhar қайси тур нархини қўйган (Дорика / Браслет / Вакс …) ----------------
def xato_turi_aniq(r, cfg, ch=None):
    """Хато товар: Gavhar қўйган нарх қайси тур/пулга тўғри келади. (қисқа ном, тушунтириш)."""
    ch = ch if ch is not None else cfg.get("chegara_usd_gr", 0.3)
    f1 = r.get("fakt_1gr")
    if f1 is None:
        return "Нарх йўқ", "Фактурада нарх ёзилмаган"
    nos = [int(x) for x in re.findall(r"№(\d+)", str(r.get("transh") or ""))]
    trs = sorted(cfg.get("transhlar", []), key=lambda t: (t["no"] not in nos, t["sana"]))
    tt = tosh_tur(r.get("nomi"))
    # ҳамма пулларнинг бошқа тур нархлари ичидан энг мосини танлаймиз: аввал товарнинг ўз тоши, кейин энг яқин нарх, кейин шу пул
    cand = [(tn != tt, round(abs(f1 - p), 3), t["no"] not in nos, t, tn, p) for t in trs for tn, p in t["narxlar"].items()
            if tn != r["tur"] and tn != "Кузсиз" and abs(f1 - p) <= ch and abs(f1 - (t["narxlar"].get(r["tur"]) or 0)) > ch]
    if cand:
        _, _, boshqa, t, tn, p = min(cand, key=lambda c: c[:3])
        return (f"{r['tur']} ўрнига {tn} нархи",
                f"Gavhar «{tn}» нархини қўйган ({p:.2f}, №{t['no']} пул {_dmy(t['sana'])}), «{r['tur']}» бўлиши керак ({r.get('hisob_1gr') or 0:.2f})")
    for t in trs:
        p = t["narxlar"].get(r["tur"]) or t["narxlar"].get("Дорика")
        if p and abs(f1 - p) <= ch and t["no"] not in nos:
            return "Бошқа пул нархи", f"№{t['no']} пул ({_dmy(t['sana'])}) нархи қўйилган ({p:.2f}), бизда №{', №'.join(map(str, nos)) or '?'} пул нархи керак"
    for q in cfg.get("vozvrat_qismlar") or []:
        for tn, p in (q.get("narxlar") or {}).items():
            if abs(f1 - p) <= ch:
                return "Возврат нархи (мос эмас)", f"Возврат {q['no']}-қисм «{tn}» нархи ({p:.2f}) — бу товарга тўғри келмади"
    return ("Нарх бошқа биржада" if "бошқа биржа" in (r.get("izoh") or "") else "Бошқа нарх"), (r.get("izoh") or "").split(". ")[0]


def xato_turlari(rows, cfg):
    """Хатолар гуруҳи: [{tur: «Вакс ўрнига Браслет нархи», n, gramm, farq, qimmat, arzon, misol}] — фарқ бўйича катта биринчи."""
    G = OrderedDict()
    for r in rows:
        if not str(r.get("holat", "")).startswith("⚠"):
            continue
        k, izoh = xato_turi_aniq(r, cfg)
        x = G.setdefault(k, dict(tur=k, n=0, gramm=0.0, farq=0.0, qimmat=0, arzon=0, izoh=izoh, misol=[]))
        x["n"] += 1; x["gramm"] += r.get("gramm") or 0; x["farq"] += r.get("farq") or 0
        x["qimmat" if (r.get("farq") or 0) > 0 else "arzon"] += 1
        if len(x["misol"]) < 3:
            x["misol"].append(f"{model_of(r.get('nomi')) or r.get('kod')} · {r.get('filial')} · {_dmy(r.get('sana'))}")
    return sorted(G.values(), key=lambda x: -abs(x["farq"]))


def xato_turlari_matn(rows, cfg, n=6):
    T = xato_turlari(rows, cfg)
    if not T:
        return []
    L = ["", "🔍 Хато турлари (Gavhar қайси нархни қўйган):"]
    for x in T[:n]:
        L.append(f"• {x['tur']}: {x['n']} та, {x['gramm']:,.2f} гр → {x['farq']:+,.2f} $ ({'қиммат' if x['farq'] > 0 else 'арзон'})")
    if len(T) > n:
        L.append(f"• … яна {len(T) - n} хил (Excel да)")
    return L


# ---------------- ВОЗВРАТДАН БЕРИЛГАН ТОВАРЛАР РЎЙХАТИ ----------------
def vozvrat_tovarlar(rows, qn=None):
    """Возврат ҳисобидан берилган товарлар: {манба: [қаторлар]}. qn — фақат шу манба (1, 2, «Ф2026-09»)."""
    out = OrderedDict()
    for r in sorted([r for r in rows if r.get("transh") == "возврат"], key=lambda r: (str(r.get("vozvrat_qism") or 1), r.get("sana") or "", r.get("filial") or "")):
        q = r.get("vozvrat_qism") or 1
        if qn is not None and str(q) != str(qn):
            continue
        out.setdefault(q, []).append(r)
    return out


def write_vozvrat_tovarlar(path, rows, cfg):
    """Excel: ҳар возврат манбаи (2-қисм, сентябрь/октябрь фиксацияси …) — алоҳида варақ, товарма-товар."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Хулоса"
    hd, fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3864")
    for j, c in enumerate(["Возврат манбаи", "Товар сони", "Грамм", "Сумма $", "Ўртача $/гр"], 1):
        x = ws.cell(1, j, c); x.font = hd; x.fill = fill
    T = vozvrat_tovarlar(rows)
    for i, (q, rs) in enumerate(T.items(), 2):
        gr, sm_ = sum(r["gramm"] or 0 for r in rs), sum(r["summa"] or 0 for r in rs)
        for j, v in enumerate([("Возврат " + qism_nom(q)), len(rs), round(gr, 2), round(sm_, 2), round(sm_ / gr, 2) if gr else None], 1):
            ws.cell(i, j, v)
    ws.column_dimensions["A"].width = 34
    for q, rs in T.items():
        w = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "", qism_nom(q))[:30])
        cols = ["Сана", "Фактура", "Филиал", "Артикул", "Модель", "Тур", "Товар номи", "Грамм", "1 гр $", "Сумма $", "Изоҳ"]
        for j, c in enumerate(cols, 1):
            x = w.cell(1, j, c); x.font = hd; x.fill = fill
        for i, r in enumerate(rs, 2):
            for j, v in enumerate([r.get("sana"), r.get("faktura"), r.get("filial"), r.get("kod"), model_of(r.get("nomi")), r.get("tur"), r.get("nomi"),
                                   round(r["gramm"] or 0, 2), round(r.get("fakt_1gr") or 0, 2), round(r["summa"] or 0, 2), r.get("vozvrat_manba")], 1):
                w.cell(i, j, v)
        n = len(rs) + 2
        w.cell(n, 1, "ЖАМИ").font = Font(bold=True)
        w.cell(n, 8, round(sum(r["gramm"] or 0 for r in rs), 2)).font = Font(bold=True)
        w.cell(n, 10, round(sum(r["summa"] or 0 for r in rs), 2)).font = Font(bold=True)
        for j, wd in enumerate([11, 15, 20, 11, 12, 11, 50, 8, 8, 11, 60], 1):
            w.column_dimensions[openpyxl.utils.get_column_letter(j)].width = wd
        w.freeze_panes = "A2"; w.auto_filter.ref = f"A1:K{len(rs) + 1}"
    wb.save(path)
    return {q: len(rs) for q, rs in T.items()}


# ---------------- ДАСТУР КИРИМИ ↔ БОТ (Gavhar фактуралари) СОЛИШТИРИШ ----------------
def _kirim_expand(prog):
    """Дастурда бир қаторда 2 дона (сони=2) бўлса — 2 та товарга ажратилади (Gavhar ҳар донани алоҳида ёзади)."""
    out = []
    for p in prog:
        n = int(round(to_num(p.get("soni")) or 1)) or 1
        for _ in range(n):
            q = dict(p)
            q["soni"], q["gramm"], q["summa"] = 1, (to_num(p.get("gramm")) or 0) / n, (to_num(p.get("summa")) or 0) / n
            out.append(q)
    return out


def kirim_solishtir(ledger, prog):
    """ledger — ботдаги Gavhar фактуралари, prog — дастурдаги кирим (экспорт). Товарма-товар: модел + грамм + нарх,
    аввал бир хил сана ва филиал, кейин санаси бошқа. Натижа: togri, sana (санаси бошқа), gramm, narx, model (номи бошқа),
    botda (Gavhar ёзган — дастурда йўқ), dasturda (дастурда бор — ботда Gavhar фактураси йўқ)."""
    prog = _kirim_expand(prog)
    g = lambda r: round(to_num(r.get("gramm")) or 0, 2)
    sm = lambda r: to_num(r.get("summa")) or 0
    ds = lambda a, b: abs((dt.date.fromisoformat(a) - dt.date.fromisoformat(b)).days) if a and b else 99
    idx = defaultdict(list)
    for j, p in enumerate(prog):
        idx[model_of(p["nomi"])].append(j)
    fm = defaultdict(Counter)
    for r in ledger:
        for j in idx.get(model_of(r["nomi"]), []):
            if prog[j]["sana"] == r["sana"] and g(prog[j]) == g(r):
                fm[r["filial"]][prog[j]["filial"]] += 1
    fmap = {b: c.most_common(1)[0][0] for b, c in fm.items()}
    F = lambda r, p: p["filial"] == fmap.get(r["filial"])
    res = dict(togri=[], sana=[], gramm=[], narx=[], model=[], botda=[], dasturda=[], fmap=fmap)
    passes = [("togri", lambda r, p: p["sana"] == r["sana"] and F(r, p) and g(p) == g(r) and abs(sm(p) - sm(r)) <= 0.05, False),
              ("narx", lambda r, p: p["sana"] == r["sana"] and F(r, p) and g(p) == g(r), False),
              ("gramm", lambda r, p: p["sana"] == r["sana"] and F(r, p) and abs(g(p) - g(r)) <= 0.3 and abs(sm(p) - sm(r)) <= 0.05, False),
              ("sana", lambda r, p: F(r, p) and g(p) == g(r) and ds(r["sana"], p["sana"]) <= 10, False),
              ("sana", lambda r, p: g(p) == g(r) and abs(sm(p) - sm(r)) <= 0.05 and ds(r["sana"], p["sana"]) <= 10, False),
              ("gramm", lambda r, p: abs(sm(p) - sm(r)) <= 0.05 and abs(g(p) - g(r)) <= 0.3 and ds(r["sana"], p["sana"]) <= 10, False),
              ("model", lambda r, p: F(r, p) and abs(sm(p) - sm(r)) <= 0.05 and ds(r["sana"], p["sana"]) <= 10, True),
              ("model", lambda r, p: abs(sm(p) - sm(r)) <= 0.05 and abs(g(p) - g(r)) <= 0.3 and ds(r["sana"], p["sana"]) <= 10, True)]
    used, left = set(), list(ledger)
    for nom, cond, any_model in passes:
        rest = []
        for r in left:
            cand = range(len(prog)) if any_model else idx.get(model_of(r["nomi"]), [])
            j = next((j for j in cand if j not in used and cond(r, prog[j])), None)
            if j is None:
                rest.append(r)
            else:
                used.add(j); res[nom].append((r, prog[j]))
        left = rest
    res["dasturda"] = [p for j, p in enumerate(prog) if j not in used]
    # ойдан ойга ўтган: Gavhar ўтган ой берган, дўкон шу ой кирим қилган — файлнинг асосий ойи бўйича ажратамиз
    oylar = Counter((p.get("sana") or "")[:7] for p in prog)
    asosiy = oylar.most_common(1)[0][0] if oylar else ""
    res["asosiy_oy"] = asosiy
    res["otgan_oy"] = [(a, b) for k in ("togri", "sana", "gramm", "narx", "model") for a, b in res[k] if (a.get("sana") or "")[:7] < asosiy]
    res["botda"] = [r for r in left if (r.get("sana") or "")[:7] >= asosiy]  # ўтган ой Gavhar фактураси бу файлда бўлмаслиги нормал
    res["botda_otgan"] = [r for r in left if (r.get("sana") or "")[:7] < asosiy]
    return res


def kirim_matn(res):
    sm = lambda xs: sum(to_num(x.get("summa")) or 0 for x in xs)
    jf = lambda ps: sm([a for a, _ in ps])
    L = ["", "🔄 ДАСТУР КИРИМИ ↔ БОТДАГИ Gavhar ФАКТУРАЛАРИ:",
         f"✅ Мос: {len(res['togri'])} та · {jf(res['togri']):,.2f} $"]
    if res["sana"]:
        L.append(f"📅 Санаси бошқа (кейинроқ кирим қилинган): {len(res['sana'])} та · {jf(res['sana']):,.2f} $ — нормал")
    if res["gramm"]:
        L.append(f"⚖️ Грамми фарқ (нархи бир хил): {len(res['gramm'])} та")
    if res["narx"] or res["model"]:
        L.append(f"🔀 Нархи ёки модели бошқа ёзилган: {len(res['narx']) + len(res['model'])} та")
    if res.get("otgan_oy"):
        oo = res["otgan_oy"]
        L.append(f"↪️ Ойдан ойга ўтган: {len(oo)} та товар ({jf(oo):,.2f} $) — Gavhar ўтган ой берган, дўкон {oy_nom(res.get('asosiy_oy', ''))} кирим қилган. Ботда ўтган ой ҳисобида, тўғри.")
    b = res["botda"]
    L.append(f"❗ Gavhar фактурасида бор — бу файлда КИРИМ ЙЎҚ: {len(b)} та · {sm(b):,.2f} $ "
             f"(дўкон кейинги ойда кирим қилган бўлса — кейинги ой файлида «↪️ ойдан ойга ўтган» бўлиб чиқади)" if b
             else "✅ Gavhar ёзган ҳамма товар дастурга кирим қилинган")
    if b:
        g = OrderedDict()
        for r in sorted(b, key=lambda r: (r["sana"], r["filial"])):
            x = g.setdefault((r["sana"], r["filial"], r.get("faktura")), [0, 0.0]); x[0] += 1; x[1] += to_num(r.get("summa")) or 0
        for (sa, fi, fk), (n, s_) in list(g.items())[:15]:
            L.append(f"   • {_dmy(sa)} · {fi} · {fk} — {n} та, {s_:,.2f} $")
        if len(g) > 15:
            L.append(f"   • … яна {len(g) - 15} та фактура (Excel да)")
    d = res["dasturda"]
    if d:
        L.append(f"📥 Дастурда кирим бор — ботда Gavhar фактураси ЙЎҚ: {len(d)} та · {sm(d):,.2f} $")
        g = OrderedDict()
        for p in sorted(d, key=lambda p: (p["sana"], p["filial"])):
            x = g.setdefault((p["sana"], p["filial"]), [0, 0.0]); x[0] += 1; x[1] += to_num(p.get("summa")) or 0
        for (sa, fi), (n, s_) in list(g.items())[:15]:
            L.append(f"   • {_dmy(sa)} · {fi} — {n} та, {s_:,.2f} $")
        if len(g) > 15:
            L.append(f"   • … яна {len(g) - 15} та (Excel да)")
        L.append("   Бу фактураларни Gavhar дан олиб, ботга ташланг — шунда пулдан айирилади ва нархи текширилади.")
    return L


def write_kirim_report(path, res):
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Хулоса"
    hd = Font(bold=True, color="FFFFFF"); fill = PatternFill("solid", fgColor="1F3864")
    for k, line in enumerate([l for l in kirim_matn(res) if l], 1):
        ws.cell(k, 1, line)
    ws.column_dimensions["A"].width = 110

    def sheet(nom, cols, data):
        w = wb.create_sheet(nom)
        for j, c in enumerate(cols, 1):
            x = w.cell(1, j, c); x.font = hd; x.fill = fill
        for i, row in enumerate(data, 2):
            for j, v in enumerate(row, 1):
                w.cell(i, j, v)
        for j in range(1, len(cols) + 1):
            w.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 16 if j != 6 else 60
        w.freeze_panes = "A2"; w.auto_filter.ref = w.dimensions
    gf = lambda x: round(to_num(x) or 0, 2)
    sheet("Дастурда йўқ (Gavhar ёзган)", ["Сана", "Gavhar фактура", "Филиал", "Артикул", "Модель", "Товар номи", "Грамм", "Сумма $"],
          [[r["sana"], r.get("faktura"), r["filial"], r.get("kod"), model_of(r["nomi"]), r["nomi"], gf(r["gramm"]), gf(r["summa"])] for r in res["botda"]])
    sheet("Ботда йўқ (дастурда бор)", ["Сана", "Фактура №", "Филиал", "Товар коди", "Модель", "Товар номи", "Грамм", "Сумма $"],
          [[p["sana"], p.get("faktura"), p["filial"], p.get("kod"), model_of(p["nomi"]), p["nomi"], gf(p["gramm"]), gf(p["summa"])] for p in res["dasturda"]])
    fr = [(k, a, b) for k in ("gramm", "narx", "model", "sana") for a, b in res[k]]
    sheet("Фарқли (грамм, нарх, сана)", ["Фарқ", "Gavhar сана", "Дастур сана", "Филиал", "Модель (Gavhar)", "Gavhar товар", "Gavhar гр", "Дастур гр",
                                         "Gavhar $", "Дастур $", "Дастур товар"],
          [[{"gramm": "грамм", "narx": "нарх", "model": "модель", "sana": "сана"}[k], a["sana"], b["sana"], b["filial"], model_of(a["nomi"]), a["nomi"],
            gf(a["gramm"]), gf(b["gramm"]), gf(a["summa"]), gf(b["summa"]), b["nomi"]] for k, a, b in fr])
    wb.save(path)


def oylik_tekshir(cfg, ledger, returns, files, out_path):
    """Ойлик приход файли бўйича тўлиқ текширув. Реестр ўзгармайди.
    Ой бошигача юкланган фактуралар + шу ойнинг приходи (файлдан) → 1, 2-пул ва возвратлар бўйича қайта ҳисоб."""
    import copy
    warns, frows = [], []
    skip_reset()
    for f in files:
        frows += read_invoice(f, warns)
    seen, uniq = set(), []
    for r in frows:
        k = key(r)
        if k not in seen:
            seen.add(k)
            uniq.append(r)
    frows = uniq
    if not frows:
        return None, "❌ Файлда приход қаторлари топилмади."
    months = sorted({(r["sana"] or "")[:7] for r in frows if r.get("sana")})
    m0 = min((r["sana"] for r in frows if r.get("sana")), default="0000")  # 15 кунлик ёки ойлик — файлдаги биринчи сана
    cfg = copy.deepcopy(cfg)
    for t in cfg["transhlar"]:  # қатор рақами бўйича ёпиш бу ерда маъносиз — санаси қолади
        t.pop("yopilgan_qator", None)
    for q in cfg.get("vozvrat_qismlar") or []:
        q.pop("yopilgan_qator", None)
    (cfg.get("vozvrat_1") or {}).pop("yopilgan_qator", None)
    returns = [dict(x, chiqdi_qator=None) for x in returns]
    mr, trs = sana_tekshir(cfg, frows, returns)
    s = write_report(out_path, cfg, mr, trs, returns)
    oy = ", ".join(f"{OYLAR[int(m[5:7])]} {m[:4]}" for m in months) or "—"
    fakt = sum(r["summa"] or 0 for r in mr)
    inv = {(r["faktura"], r["filial"]) for r in mr}
    t0 = min((t["sana"] for t in cfg["transhlar"] if t.get("faol", True)), default=None)
    oldin = [r for r in mr if t0 and (r.get("sana") or "") < t0]
    L0 = []
    if not cfg["transhlar"]:
        L0 = ["⚠️ Ботда пул (транш) киритилмаган — нархлар текширилмади. Аввал шу давр пулларини киритинг.", ""]
    elif oldin:
        L0 = [f"⚠️ {len(oldin)} та товар ({sum(r['summa'] or 0 for r in oldin):,.0f} $) биринчи пулдан ({_dmy(t0)}) олдинги санада.",
              "   Бу давр пули ботда йўқ ёки ўчирилган — ўша пулни киритиб «Текширувда» қилинг, кейин қайта ташланг.", ""]
    L = L0 + [f"🔎 АВТО ТЕКШИРУВ — санага ва нархга қараб (реестрга қўшилмайди)", f"📅 {oy} ({_dmy(min(r['sana'] for r in mr))} – {_dmy(max(r['sana'] for r in mr))})", "",
         f"📦 Приход: {len(mr)} та товар · {sum(r['gramm'] for r in mr):,.2f} гр · {fakt:,.2f} $",
         f"   {len(inv)} та фактура · {len({r['filial'] for r in mr})} та дўкон", "", "💳 Текширувда ёқилган пуллар:"]
    for t in trs:
        L.append(f"   {t['no']}-пул ({_dmy(t['sana'])}, биржа {t['birja']}): шу нархда {t['n']} та товар · {t['used']:,.0f} $")
    off = [t for t in cfg["transhlar"] if not t.get("faol", True)]
    if off:
        L.append("   ⏸ ўчирилган: " + ", ".join(f"{t['no']}-пул ({_dmy(t['sana'])})" for t in off))
    rv = [r for r in mr if r.get("transh") == "возврат"]
    ok = [r for r in mr if r["holat"].startswith("✓") and r.get("transh") != "возврат"]
    bad = [r for r in mr if r["holat"].startswith("⚠")]
    L += ["", "Текширув:"]
    by = OrderedDict()
    for r in ok:
        ns = re.findall(r"№(\d+)", str(r.get("transh") or ""))
        by.setdefault("+".join(dict.fromkeys(ns)) or "?", []).append(r)
    for k, rs_ in sorted(by.items()):
        L.append(f"✅ {k}-пул нархида тўғри: {len(rs_)} та · {sum(r['summa'] or 0 for r in rs_):,.2f} $")
    for qn in sorted({r.get("vozvrat_qism") or 1 for r in rv}, key=str):
        qr = [r for r in rv if (r.get("vozvrat_qism") or 1) == qn]
        L.append(f"↩️ Возврат {qism_nom(qn)}{' (рўйхат)' if qn == 1 else ''}: {len(qr)} та · {sum(r['summa'] or 0 for r in qr):,.2f} $")
    qim = sum(r["farq"] or 0 for r in bad if (r["farq"] or 0) > 0)
    arz = sum(r["farq"] or 0 for r in bad if (r["farq"] or 0) < 0)
    if bad:
        L.append(f"❌ Хато: {len(bad)} та товар · {len({(r['faktura'], r['filial']) for r in bad})} та фактурада")
        if qim:
            L.append(f"   👉 ОРТИҚЧА ёзилган: {qim:,.2f} $")
        if arz:
            L.append(f"   👉 КАМ ёзилган: {abs(arz):,.2f} $")
        byf = defaultdict(float)
        for r in bad:
            byf[r["filial"]] += r["farq"] or 0
        L.append("   Энг кўп фарқ:")
        for f_, v in sorted(byf.items(), key=lambda kv: -abs(kv[1]))[:5]:
            L.append(f"   • {f_}: {v:+,.2f} $")
    else:
        L.append("✅ Хато йўқ — ҳамма нархлар тўғри")
    for q in qism_stats(cfg):
        if q.get("faol") is not False:
            L.append(f"↩️ Возврат {q['no']}-қисм: берилди {q['berilgan']:,.0f} $, қолди {q['qolgan']:,.0f} $")
    for f in fiks_stats(cfg):
        if f["faol"]:
            L.append(f"↩️ Возврат фиксация {f['nom']}: ишлатилди {f['used']:,.0f} $, қолди {f['left']:,.0f} $")
    L += skip_matn()
    if warns:
        L.append(f"🧮 Файлда арифметика хатоси: {len(warns)} та")
    oxir = max((r["sana"] for r in mr if r.get("sana")), default="")
    eslat = [f"возврат {q['no']}-қисм ({q['summa_usd']:,.0f} $, {_dmy(q.get('sana'))} да киритилган)"
             for q in cfg.get("vozvrat_qismlar") or [] if q.get("faol", True) and q.get("sana") and oxir and q["sana"][:7] > oxir[:7]]
    if eslat:
        L += ["", "🔔 Эслатма: " + ", ".join(eslat) + " ҳам ҳисобга олинди. Бу давр учун тегишли бўлмаса — "
              "Admin → Tekshiruv sozlamasi да белгисини олиб, файлни қайта ташланг."]
    L += xato_turlari_matn(mr, cfg)
    L += ["", "ℹ️ Бу текширув реестрни ўзгартирмайди. Батафсил — Excel файлда (Хатоликлар варағи)."]
    return mr, "\n".join(L)


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
