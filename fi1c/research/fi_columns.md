# Семантика колонок XML-формата Finansijski iskazi (fi.xml/fi.xsd) — портал eprijava.tax.gov.me

Дата рисёрча: 2026-08-08.

## Итог / TL;DR

| Отчёт | Iznos-колонки | Уверенность |
|---|---|---|
| **BilanStanja** | Iznos1 = **tekuća godina** (конец текущего периода); Iznos2 = **prethodna godina — krajnje stanje** (конец предыдущего периода, обычное сравнительное значение); Iznos3 = **prethodna godina — početno stanje** (начало предыдущего периода) — заполняется **только** при ретроспективном пересчёте/переклассификации, в обычной сдаче = 0 | **Высокая** (первичный источник — текст Правилника, ст. 4) |
| **BilanUspjeha** | Iznos1 = tekuća godina, Iznos2 = prethodna godina | **Высокая** (ст. 5 Правилника + реальный поданный отчёт) |
| **IskazOTokovimaGotovine** | = Прилог 3, **direktna metoda**. Iznos1 = tekuća godina (колона 3), Iznos2 = prethodna godina (колона 4) | **Высокая** (ст. 6, 8 Правилника) |
| **TokoviGotovine3a** | = Прилог 3a, **indirektna metoda**. Те же колонки, что и выше | **Высокая** (ст. 9 Правилника — прямая цитата) |
| **IskazOPromenamaKapitala** | RedniBr2..11/Iznos2..11 = 10 колонок капитала, см. таблицу ниже | **Высокая** (ст. 18 Правилника — построчная цитата + реальный поданный отчёт, номера red.br. совпадают 1:1) |
| **StatistickiAneks** | Iznos1 = tekuća godina, Iznos2 = prethodna godina | **Высокая** (структура идентична BilanUspjeha, подтверждено реальным поданным отчётом) |
| **ObracunAmortizacije** | Iznos2..8 = Početni saldo / Kupovina / Prodaja / Neotpisana vrijednost / Stopa % / Amortizacija / Neotpisana vrijednost na kraju — см. таблицу ниже | **Высокая** (реальный поданный «Obrazac OA» + Zakon o porezu na dobit) |

---

## 1. BilanStanja (Iskaz o finansijskoj poziciji / Bilans stanja)

### Источник истины
Официальный **Pravilnik o sadržini i formi obrazaca finansijskih iskaza za privredna društva i druga pravna lica** («Sl. list CG», br. 011/20 od 06.03.2020, sa izmjenama 139/21, 013/22, 139/22) — прочищенный текст, член 4:

> «U koloni 5 iskazuju se podaci na kraju tekućeg izvještajnog perioda na osnovu stanja na računima glavnih knjiga finansijskog i pogonskog knjigovodstva tekućeg izvještajnog perioda.
> ...
> **U koloni 6 unose se podaci sa kraja prethodnog izvještajnog perioda.**
> **U koloni 7 unose se podaci sa početka prethodnog izvještajnog perioda. Podaci se unose u kolonu 7 samo u sledećim situacijama: retrospektivne primjene računovodstvene politike ili retrospektivnog preračunavanja stavki u finansijskim iskazima, ili kada se reklasifikuju stavke u finansijskim iskazima. Ako se mijenja makar jedan podatak u koloni 7 popunjava se cijela kolona 7 i u aktivi i u pasivi.**»

В печатной форме обрасца (Прилог 1) колонки называются буквально:
`Grupa računa | POZICIJA | Redni broj | Napomena-broj | Tekuća godina | Prethodna godina - krajnje stanje | Prethodna godina - početno stanje`
— подтверждено реально поданным отчётом (см. Источники, п. B): «ISKAZ O FINANSIJSKOJ POZICIJI /BILANS STANJA/» DOO «Strelkomerc», Podgorica, broj iskaza 53984/2020 — колонки 5/6/7 названы дословно так, значения по строкам (RedniBroj 001, GrupaRacuna «00», Pozicija «A. NEUPLAĆENI UPISANI KAPITAL» и т.д.) **пословно совпадают** со строками в вашем `fi.xml`, включая экзотические позиции типа «058340/59371» и коды групп счетов «030,039(dio)».

### Отображение на XSD-поля
- **Iznos1 → колона 5 → Tekuća godina** (баланс на конец текущего отчётного периода — стандартное значение «на 31.12 текущего года»).
- **Iznos2 → колона 6 → Prethodna godina - krajnje stanje** (баланс на конец предыдущего периода — стандартная сравнительная колонка, «на 31.12 прошлого года»). Это тот столбец, который в обычной практике называют просто «prethodna godina» / это НЕ ispravka vrijednosti и НЕ bruto.
- **Iznos3 → колона 7 → Prethodna godina - početno stanje** (баланс на начало предыдущего периода, т.е. «на 01.01 прошлого года» = «на 31.12 позапрошлого года»). Заполняется **только** в трёх случаях, прямо перечисленных в законе: (а) ретроспективное применение учётной политики, (б) ретроспективный пересчёт статей, (в) реклассификация статей — это стандартное IAS 1 требование о «третьем балансе» при ретроспективном пересчёте. При обычной сдаче отчётности действующей компании без пересчётов эта колонка **пустая/0**.

### ⚠️ Важно для реализации в fi1c
Первоначальная гипотеза research-задания «Bruto/Ispravka vrijednosti/Neto» **не подтвердилась** — это НЕ то, что означают Iznos1/2/3. Колонки означают **три временные точки одного и того же нетто-баланса** (текущий год, конец пред. года, начало пред. года), а не бухгалтерское разложение бруто-баланса на составляющие. Для расчёта в fi1c:
- Iznos1 = остатки на конец `Год` (как и предполагает SPEC.md).
- Iznos2 = остатки на конец `Год-1` (= то, что было Iznos1 в отчёте за `Год-1`).
- Iznos3 = 0 в общем случае; заполняется вручную только если в текущем году делается ретроспективная корректировка/переклассификация сравнительных данных (редкий кейс, `Ручное`/`Корректировки`).

Отдельно ст. 3 Правилника уточняет: при статусной промене/ликвидации/стечаju текущий год показывается объединённо (net, после взаимозачёта результатов до/после события) с пометкой «statusna promjena»/«stečaj»/«likvidacija» на всех формах — это НЕ про количество колонок, а про то, как заполняется одна колонка «tekuća godina».

Строки RedniBroj 001–144 (актива 001–046, пасива 101–144) заполняются из соответствующих групп счетов (GrupaRacuna), формулы-суммы (например «B. STALNA IMOVINA (003+008+016)») — как и предполагает mapping.json.

---

## 2. BilanUspjeha (Iskaz o ukupnom rezultatu / Bilans uspjeha)

Правилник, ст. 5:
> «U obrazac Iskaz o ukupnom rezultatu /bilans uspjeha/ pravno lice unosi podatke u kolone 5 i 6... Podaci se sa odgovarajućih računa... unose tako što se u kolonu 6 (prethodna godina) unose podaci obračuna odgovarajućeg perioda prethodne godine, reklasifikovani u skladu sa strukturom podataka za tekući period, dok se za tekuću godinu, u koloni 5, podaci unose s računa na kraju perioda za koji se obračun sastavlja.»

- **Iznos1 = kolona 5 = Tekuća godina** (текущий отчётный период).
- **Iznos2 = kolona 6 = Prethodna godina** (предыдущий период, реклассифицирован под структуру текущего периода).

Подтверждено и реально поданным отчётом (заголовки «Tekuća godina»/«Prethodna godina», колонки 5/6), и IFRS-отчётностью ЦБ Черногории (Bilans uspjeha «2022. godina»/«2021. godina» — тот же принцип tekuća/prethodna, хотя это другой формат — не портальный AOP-обрасац, а полный IFRS-отчёт).

Никакой третьей колонки (в отличие от BilanStanja) нет и не может быть в XSD (`Iznos1`, `Iznos2` — всё, что описывает схема) — потому что Bilans uspjeha описывает поток за период, а не остаток на дату, и требование о «третьем состоянии» из IAS 1 к нему не относится.

Знаки: ряд позиций вносится со знаком минус в специальных случаях (202 — уменьшение запасов, 243 — убыток от прекращённой деятельности, 247/258 — отложенный налоговый доход) — это относится к содержанию суммы, а не к семантике колонки.

---

## 3. IskazOTokovimaGotovine vs TokoviGotovine3a

Правилник, ст. 2: формы утверждены как **Прилог 1, 1a, 2, 2a, 3, 3a и 4**. Cт. 6-9 (дословно):

> «Ст. 6: Pri sastavljanju obrasca Iskaz o tokovima gotovine, u koloni 4 unose se podaci iz prethodnog izvještajnog perioda, a za tekući izvještajni period podaci se unose u koloni 3.
> Ст. 7: Pravno lice samostalno odlučuje da li će iskaz o tokovima gotovine sastavljati primjenom **direktne metode** ili primjenom **indirektne metode**.
> Ст. 8: U slučaju izbora direktne metode... pravno lice popunjava obrazac dat u **prilogu 3**.
> Ст. 9: U slučaju izbora indirektne metode... pravno lice popunjava obrazac dat u **prilogu 3a**.»

Значит:
- **`IskazOTokovimaGotovine`** (в XSD) = **Prilog 3** = **direktna metoda** (прямой метод). Строки начинаются с «I. Prilivi gotovine iz poslovnih aktivnosti» → «1. Prodaja i primljeni avansi», «2. Primljene kamate...» и т.д. — классическая структура прямого метода (наличные поступления/выплаты по видам операций).
- **`TokoviGotovine3a`** = **Prilog 3a** = **indirektna metoda** (косвенный метод). Строки начинаются с «1. Rezultat prije oporezivanja» → «2. Amortizacija», «3. Promjena zaliha», «4. Promjena potraživanja»... — классическая реконciliация чистой прибыли до налога с денежным потоком (косвенный метод). Подтверждается и напрямую текстом ст. 9 Правилника, построчно объясняющим позиции 301-310 Приложения 3a (Rezultat prije oporezivanja / Amortizacija / Promjena zaliha / Promjena potraživanja / Promjena obaveza prema dobavljačima / Promjena rezervisanja / Plaćene kamate / Porez na dobit / Plaćanja po osnovu ostalih javnih prihoda / Promjena odloženih poreza).

Оба отчёта — это **один и тот же логический отчёт (Iskaz o tokovima gotovine), только по выбору юрлица заполняется ЛИБО Прилог 3, ЛИБО Прилог 3a** (метод выбирается самостоятельно предприятием, ст. 7). Секции B (investicione aktivnosti) и C (finansiranje) в обоих приложениях идентичны — различается только раздел A (poslovne aktivnosti).

Колонки в обоих: **Iznos1 = tekuća godina (kolona 3), Iznos2 = prethodna godina (kolona 4)** — подтверждено и текстом закона, и реальным поданным отчётом («ISKAZ O TOKOVIMA GOTOVINE - direktna metoda», колонки «Tekuća godina»/«Prethodna godina»).

Для fi1c: скорее всего организация выбирает ОДИН метод (обычно прямой, как самый распространённый в малых/средних юрлицах ЦГ) — при выгрузке заполняется соответствующая секция, вторая остаётся нулевой/пустой (обе секции обязательны в XSD как элементы, но по смыслу заполняется только выбранная).

---

## 4. IskazOPromenamaKapitala (Iskaz o promjenama na kapitalu)

Правилник, раздел V, ст. 18 (Позиция 1, «Stanje na dan 01.01 prethodne godine») дословно перечисляет все 10 колонок с привязкой к группам счетов и red.br.:

| RedniBrN / IznosN | Наименование колонки (из Правилника, ст. 18) | Группа/счёт плана счетов | red.br. позиции 1 (образец) |
|---|---|---|---|
| RedniBr2 / Iznos2 | **Osnovni kapital** (grupa 30 bez računa 309) | grupa 30, без 309 | 401 |
| RedniBr3 / Iznos3 | **Ostali kapital** | račun 309 | 410 |
| RedniBr4 / Iznos4 | **Neuplaćeni upisani kapital** | grupa 31 | 419 |
| RedniBr5 / Iznos5 | **Emisiona premija** | račun 320 | 428 |
| RedniBr6 / Iznos6 | **Rezerve** (zakonske + statutarne/druge) | računi 321, 322 | 437 |
| RedniBr7 / Iznos7 | **Revalorizacione rezerve** (i nerealizovani dobici/gubici) | grupa 33 | 446 |
| RedniBr8 / Iznos8 | **Neraspoređena dobit** | grupa 34 | 455 |
| RedniBr9 / Iznos9 | **Gubitak** | grupa 35 | 464 |
| RedniBr10 / Iznos10 | **Otkupljene sopstvene akcije i udjeli** | račun 237 | 473 |
| RedniBr11 / Iznos11 | **Ukupno** (kol. 2+3+4+5+6+7+8-9-10) | — (сумма) | 482 |

9 строк (Pozicija 1–9 в XSD/XML) — это НЕ 9 разных лет, а 9 фиксированных стадий одного «моста» между началом и концом ДВУХ периодов (текущего и предыдущего года):
1. Stanje na dan 01.01 prethodne godine (начало пред. года)
2. Efekti retroaktivne ispravke grešaka i promjena rač. politika (корректировки к нач. пред. года)
3. Korigovano početno stanje 01.01 prethodne godine (r.br. 1+2)
4. Neto promjene u prethodnoj godini
5. Stanje na dan 31.12 prethodne godine (r.br. 3+4) = начало текущего года
6. Efekti retroaktivne ispravke grešaka (к нач. тек. года)
7. Korigovano početno stanje (r.br. 5+6)
8. Neto promjene u tekućoj godini
9. Stanje na dan 31.12 tekuće godine (r.br. 7+8)

Полностью подтверждено реально поданным отчётом (тот же файл, что и для BilanStanja): red.br. 401/410/419/…/482 для позиции 1, значения по колонкам совпадают с суммами в Bilans stanja того же юрлица (Osnovni kapital = сумма счёта 302 из Bilans stanja и т.д.) — 100% совпадение номеров и структуры с вашим fi.xml.

Знак: убытки и уменьшения капитала вносятся отрицательными числами (стандартное правило, следует из формулы Ukupno = 2+3+4+5+6+7+8-9-10, где Gubitak (кол.9) и Otkupljene akcije (кол.10) вычитаются).

---

## 5. StatistickiAneks

Не описан отдельно в Pravilniku o finansijskim iskazima (это отдельная форма, обязательная по Zakonu o računovodstvu, используется в т.ч. для нужд MONSTAT). Структура в fi.xml идентична BilanUspjeha/BilanStanja по форме строки (`RedniBroj`, `Iznos1`, `Iznos2`, `NapomenaBroj`, `Pozicija`, `GrupaRacuna`).

Подтверждено реальным поданным отчётом: печатная форма «STATISTIČKI ANEKS» имеет колонки `Grupa računa, računa | POZICIJA | Redni broj | Napomena - broj | Tekuća godina | Prethodna godina` (колонки 5 и 6) — т.е. структурно это точная копия таблицы Bilans uspjeha, просто с другим набором строк (более детальная разбивка по счетам для статистики — «Prosječan broj zaposlenih», «Prihodi od prodaje robe», «Troškovi zarada i naknada zarada (bruto)» и т.д., плюс отдельный аналитический блок по нематериальным активам).

- **Iznos1 = Tekuća godina**
- **Iznos2 = Prethodna godina**

Уверенность высокая — подтверждено прямым наблюдением реального поданного документа с этими же заголовками колонок и той же структурой строк, что и в вашем fi.xml (Prosječan broj zaposlenih 001, Prihodi od prodaje robe 002 = grupa 60, и т.д. — построчное совпадение).

---

## 6. ObracunAmortizacije (Obrazac OA — Obračun amortizacije osnovnih sredstava)

Это отдельная форма («Obrazac OA», надпись в правом верхнем углу бланка), прилагаемая к налоговой декларации по налогу на прибыль (Prijava poreza na dobit), а не к Pravilniku o finansijskim iskazima — регулируется **Zakonom o porezu na dobit pravnih lica** (амортизационные группы, ст. о амортизации осн. средств) и соответствующим подзаконным Pravilnikom o razvrstavanju osnovnih sredstava po grupama.

### Колонки (подтверждено реальным поданным «Obrazac OA», DOO Strelkomerc, 2020 год)

| Поле XSD | Заголовок колонки (обрасц) | Формула/смысл |
|---|---|---|
| `Grupa` | Broj grupe | I / II / III / IV / V |
| Iznos2 | **Početni saldo** (кол. 2) | неотписана вредность на начало года = неотписана вредность на конец пред. года |
| Iznos3 | **Kupovina sredstava koja se stavljaju u upotrebu** (кол. 3) | приобретения в течение года |
| Iznos4 | **Prodaja sredstava tokom godine** (кол. 4) | выбытия в течение года |
| Iznos5 | **Neotpisana vrijednost** (кол. 5) | = 2+3-4 (база для начисления амортизации) |
| Iznos6 | **Stopa %** (кол. 6) | налоговая ставка амортизации по группе |
| Iznos7 | **Amortizacija** (кол. 7) | = 5 × 6 (начисленная амортизация за год) |
| Iznos8 | **Neotpisana vrijednost na kraju godine** (кол. 8) | = 5 − 7 (переносится как Iznos2 в отчёт следующего года) |

Примечание из формы (дословно): «NAPOMENA: Početni saldo na početku godine jednak je neotpisanoj vrijednosti na kraju prethodne godine.»

### Ставки по группам (Iznos6) — ⚠️менялись, важно для года отчёта

Метод: группа I амортизируется **линейным (proporcionalna) методом отдельно по каждому средству**; группы II–V — **дегрессивным (declining balance) методом, по группам целиком** (пул).

- **До 01.01.2024** (закон в редакции до Sl. list CG 125/23): **I = 5%, II = 15%, III = 20%, IV = 25%, V = 30%**. Подтверждено реально поданным Obrazac OA за 2020 год (ставки в бланке: I=5, II=15, III=20, IV=25, V=30) и независимым источником (unija.com, до-2024 описание).
- **С 01.01.2024** (после изменений Zakona o porezu na dobit, «Sl. list CG», br. 125/23 od 31.12.2023, применяются с 01.01.2024): ставки снижены до **I = 2.5%, II = 10%, III = 15%, IV = 20%, V = 30%**. Это **ровно те значения**, что стоят по умолчанию в вашем fi.xml (Iznos6: 2.5 / 10 / 15 / 20 / 30) — значит ваш образец fi.xml/fi.xsd относится к текущей (пост-2024) версии портала/схемы, и правильные ставки для новых деклараций — **2.5/10/15/20/30**, а не 5/15/20/25/30.

Для fi1c: при заполнении Obrazac OA за год ≥2024 использовать ставки 2.5/10/15/20/30; за годы 2020-2023 — 5/15/20/25/30 (если когда-нибудь понадобится ретро-формирование). Точный текст изменения закона (какая именно статья/пункт правит ставки) не процитирован дословно — рекомендуется при первом реальном использовании сверить с действующим текстом Zakona o porezu na dobit pravnih lica (актуальная консолидированная версия на gov.me/wapi.gov.me) перед боевой эксплуатацией, но связь «fi.xml default = 2.5/10/15/20/30 = ставки с 2024» подтверждена независимо двумя источниками (сам файл + unija.com).

---

## Источники

### Первичные (наивысшая надёжность)
- **A. Prečišćeni tekst Pravilnika o sadržini i formi obrazaca finansijskih iskaza za privredna društva i druga pravna lica** («Sl. list CG», br. 011/20, 139/21, 013/22, 139/22), составитель Pravni ekspert doo, Podgorica — полный текст статей 1-21 прочитан целиком (члены 1-9 — Bilans stanja/Bilans uspjeha/Tokovi gotovine дословно процитированы выше; член 18-21 — Iskaz o promjenama na kapitalu). PDF: `https://cdn.prod.website-files.com/6458e788814cab18be394a81/6537cc6d5cc53b4a96d45faa_3.%20Pravilnik%20o%20sadrzini%20i%20formi%20obrazaca.pdf` (внутренне это тот же «Katalog propisa 2023»).
- **B. Реально поданный финансовый отчёт** DOO «Strelkomerc» (bivši «Streljački savez Crne Gore» hosting), broj iskaza 53984/2020, na dan 31.12.2020 — содержит ВСЕ формы (Bilans stanja, Bilans uspjeha, Iskaz o tokovima gotovine — direktna metoda, Iskaz o promjenama na kapitalu, Statistički aneks, Obrazac OA, Prijava poreza na dobit) с реальными заголовками колонок и заполненными суммами; номера red.br. и коды GrupaRacuna совпадают 1:1 со строками в `xsd/fi.xml`. URL: `https://streljackisavez-crnegore.me/wp-content/uploads/strelkom-zav-racun-2020.pdf`.
- **C. `xsd/fi.xml` и `xsd/fi.xsd`** (сам образец с портала) — использован для сверки структуры и значений по умолчанию (в частности, ставок амортизации).

### Вторичные / контекст
- Финансовые отчёты Centralne banke Crne Gore (IFRS, не портальный AOP-обрасц, но подтверждает общий принцип tekuća/prethodna godina): `https://www.cbcg.me/slike_i_fajlovi/fajlovi/fajlovi_publikacije/god_izv_finansijski_izv/fs_cbcg_2022.pdf`.
- Zakon o porezu na dobit pravnih lica Crne Gore (амортизационные группы) — обзор от unija.com (аккаунтинг-фирма, ЦГ): `https://unija.com/cg/obracun-poreske-amortizacije-crna-gora/` — подтверждает действующие (после 2024) ставки 2.5/10/15/20/30 и метод (I — proporcionalna po sredstvu, II-V — degresivna po grupama).
- Изменения ставок амортизации с 01.01.2024 (Sl. list CG 125/23 od 31.12.2023) — упомянуты в нескольких вторичных источниках (найдено через WebSearch, включая unija.com «Izmijenjen je i dopunjen Zakon o porezu na dobit pravnih lica»), прямой текст поправки не процитирован дословно.
- Официальная страница Poreske uprave CG с Pravilnikom: `https://eprijava.tax.gov.me/TaxisPortal/FinancialStatement/Index`; портал izrsg.org (Institut sertifikovanih računovođa Crne Gore): `https://www.isrcg.org/pravilnik`.

## Открытые вопросы / что не 100% подтверждено
1. Дословный текст изменения ставок амортизации (Sl. list CG 125/23) не найден и не процитирован — сама смена ставок 5/15/20/25/30 → 2.5/10/15/20/30 подтверждена косвенно (совпадение fi.xml-дефолтов с текущим описанием у unija.com), но не первоисточником закона.
2. Правовая база `StatistickiAneks` (какая именно норма Zakona o računovodstvu / решение MONSTAT обязывает его сдавать и точный состав строк) не процитирована дословно — структура колонок подтверждена практически (реальный поданный отчёт), но не текстом нормы.
3. Не проверено на реальном примере поведение колонки Iznos3 (Prethodna godina - početno stanje) в BilanStanja при РЕАЛЬНОМ ретроспективном пересчёте (только текст закона + косвенно подозрительный пример из п. B, где сама подача, наоборот, показывает данные ИСКЛЮЧИТЕЛЬНО в колонках 6/7 при пустой колонке 5 — похоже на специфический edge-case подачи данных о статусной промене/окончании деятельности, требует дополнительной проверки при появлении второго живого примера).
