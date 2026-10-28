> **Ҳамма файл битта жойда (папкасиз)** — GitHub га «Add file → Upload files» орқали ҳаммасини бирданига ташланг.

# Gavhar — тилло нарх назорати: Telegram Mini App (Railway)

Бот Gavhar Gold zavod фактураларини текширади:
- ҳар позицияни **транш нарх жадвали** билан солиштиради (лом 585 + ишчилик + НДС), фарқ ±0,3 $/гр дан ошса ⚠;
- пулни **FIFO** бўйича траншларга тақсимлайди: олдинги транш тугамагунча товар унинг нархида;
- ҳар юкламадан кейин қисқа хулоса + Excel ҳисобот (Назорат, Транш, Хатоликлар, Приход, Кунлик, Нархлар) юборади;
- такрорий фактурани иккинчи марта қўшмайди, `/bekor` билан охирги юкламани қайтаради.

Траншлар (берилган пуллар) бошланғич файлда ЙЎҚ — уларни фақат фойдаланувчи ўзи киритади (ботга нарх хабари + сумма ёки mini app → Transhlar). Бот пулни ўзи қўшмайди ва ўзгартирмайди.

## Railway га ўрнатиш (15 дақиқа)

1. **Бот яратиш.** Telegram да @BotFather → `/newbot` → **токен**ни нусхалаб олинг.
2. **Кодни GitHub га юкланг** (private репозиторий, шу папкадаги ҳамма файл).
3. **Railway** → *New Project* → *Deploy from GitHub repo* → репозиторийни танланг (`Dockerfile` ни ўзи топади).
4. **Volume:** сервис → *Attach volume* → **Mount path: `/data`** (маълумот сақланиши учун мажбурий).
5. **Домен:** сервис → *Settings → Networking* → **Generate Domain** (порт сўраса: 8080).
   Шу домен mini app манзили бўлади — бот уни ўзи олади (`RAILWAY_PUBLIC_DOMAIN`).
6. **Variables:** `BOT_TOKEN` = токен, `ALLOWED_USERS` = Telegram ID ингиз (бир нечта бўлса вергул билан),
   `PORT` = 8080. ID ни билмасангиз: ботга `/myid` ёзинг.
7. *Deploy* → Logs да «Бот ишга тушди» чиқса тайёр. Ботга `/start` ёзинг → «🟡 Назоратни очиш» ёки пастдаги «Назорат» менюси.

> Сервисни фақат **1 нусха** да ишлатинг. Mini app ни фақат `ALLOWED_USERS` даги одамлар очади — Telegram имзоси серверда текширилади.

## Mini app bo'limlari (dizayn C — Telegram minimal, lotin yozuvi)
- **Holat** — ta'minotchidagi qoldiq, faol transh va ishlatilgan %, oxirgi faktura ogohlantirishi (qimmat bo'lsa qizil),
  jami berilgan / olingan, fakturalar soni, jami farq, kunlik sarf, taxminiy tugash, kunlik grafik, tur bo'yicha jadval,
  «Excel hisobotni chatga yuborish».
- **Faktura** — transh tanlanadi (Transh №2 · 13.10 / Transh №1 · 15.07): pul berilgandan keyin kelgan **fakturalar soni,
  jami qiymati**, hisob bo'yicha qiymati, farq, ishlatilgan %. Pastda fakturalar ro'yxati (raqam, sana, filial, pozitsiya,
  gramm, summa, farq). Faktura bosilsa — tur bo'yicha tafsilot (faktura narxi / kerakli narx / farq) va «xatoliklar» tugmasi.
- **Xatoliklar** — farqli pozitsiyalar: Qimmat / Arzon / Boshqa, filial va tur filtri, kod yoki faktura № bo'yicha qidiruv.
- **Transhlar** — transhlar va narxlari, yangi transh formasi (Tekshirish → Saqlash), transhni sanadan yopish / qayta ochish,
  ruxsat etilgan farq.
- **Faktura yuklash** — Excel faylni bot chatiga tashlang (Gavhar INVOICE formati ham, tizim eksporti ham o'qiladi).

## Фойдаланиш

| Нима қилиш керак | Қандай |
|---|---|
| Фактурани текшириш | Тизим экспортини (.xlsx) ботга ташланг |
| Ҳозирги ҳолат | `/holat` |
| Янги пул берилди | `/transh` ва кейинги қаторларда: сана, «Биржа нархи …», ҳар тур нархи, «Сумма …» (намуна `/start` да). Бот лом ва ишчилик+НДС ни олдинги транш билан солиштириб кўрсатади, «✅ Сақлаш» босилгач сақлайди |
| Транш тугади / ёпиш | `/yop 1 2026-10-13` (қайта очиш: `/och 1`) |
| Нархларни кўриш | `/narxlar` |
| Чегарани ўзгартириш | `/chegara 0.3` |
| Охирги юкламани бекор қилиш | `/bekor` |
| Реестрни юклаб олиш | `/reestr` |

**Фактура файли** — шу устунлар бўлиши керак: «Товар номи» (ичида `Vazni - X`, `Tosh turi - …`),
«Товар категорияси», «Сони», «Нарх» ёки «Қиймати», «Ташкилот», «Таъминотчи», «Таъминотчи фактура рақами»,
«Таъминотчи фактура санаси».

**Тур аниқлаш:** Билак узуклар → Браслет, Занжирлар → Дорика (тош турига қарамай). Истисно: QSG, QSW, QSF, QQF, ZSG сериялари тош нархида (вакс → Вакс) — рўйхат Admin → Tekshiruv sozlamasi да; Van Cleef/Клиф → Ван Клиф; brilliant → Бриллиант;
fianit → Фианит; Занжирлар → Дорика; waks → Вакс; Tosh turi йўқ → Детский.

## Файллар
- `app.py` — асосий файл: mini app сервери (API) + бот
- `index.html` — mini app интерфейси, `chart.umd.min.js` — график кутубхонаси
- `bot.py` — бот ёрдамчилари (фактура қабул қилиш, /bekor)
- `nazorat.py` — ҳисоб (FIFO, нарх текшириш, Excel ҳисобот); алоҳида ҳам ишлайди:
  `python nazorat.py --transh seed_transhlar.json --ledger seed_ledger.csv --out hisobot.xlsx faktura.xlsx`
- `seed_transhlar.json`, `seed_ledger.csv` — бошланғич транш ва реестр (биринчи ишга тушганда `/data` га кўчирилади)
- Volume `/data` да: `transhlar.json`, `ledger.csv`, `backup/` (охирги 30 нусха)
