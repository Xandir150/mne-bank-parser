# Obrazac PD — построчная семантика (Crna Gora, porez na dobit pravnih lica)

Исследование выполнено 2026‑08‑08. Задача: сопоставить поля `Iznos1..Iznos59` (+ подпункты
`15a/15b, 18a‑c, 21a/21b, 49a/49b, 51a‑c, 55a‑c, 56a/56b`, приложения `IznosPG1_*`, `IznosPG2_*`)
из `/Users/alexander.shekhovtsov/work/1с/izvod/xsd/pd.xml` и `pd.xsd` с официальным содержанием
декларации по налогу на прибыль Черногории (Obrazac PD), которую сдают через
eprijava.tax.gov.me / IRMS.

**TL;DR**: официальный первоисточник найден и прочитан целиком — это действующий (с 2024 г.)
консолидированный Pravilnik с полным текстом формы и построчной инструкцией (Uputstvo).
99% полей `Iznos1..Iznos58` и приложений PG1/PG2 сопоставлены с высокой уверенностью
дословно по официальной инструкции. Есть три поля в XSD, которых **нет** в найденном мною
официальном PDF формы (`Iznos18c`, `Iznos56a`, `Iznos56b`, `Iznos59`) — по ним даю обоснованные
гипотезы, но это **не подтверждено** первоисточником (см. раздел «Открытые вопросы»).

---

## 0. Источники (первичные, скачаны и прочитаны целиком)

Все файлы сохранены локально в
`/private/tmp/claude-501/-Users-alexander-shekhovtsov-work-1--izvod/72788ddf-300d-45f0-81c2-01c07a94c0a8/scratchpad/pd_research/`
на случай, если понадобится повторная сверка.

1. **Действующий Pravilnik (актуальная редакция, прочитан целиком) — ГЛАВНЫЙ ИСТОЧНИК**
   Прочищени текст на gov.me:
   `https://www.gov.me/dokumenta/2ae946d5-418e-41a1-912b-99c6369dc1e7` →
   основной текст: `https://wapi.gov.me/download/618575b3-37ff-415a-8d2d-7c5524dd459c?version=1.0`
   (RTF, содержит перечень всех редакций Pravilnika и текст последнего Člana 1–5) —
   и приложение (сама форма + инструкция):
   `https://wapi.gov.me/download/69fb7c10-ee3f-4e44-a6ee-10c0e86609d7?version=1.0` (PDF, 204 KB).
   Это **консолидированный текст**, включающий редакцию
   *"Pravilnik o izmjeni Pravilnika o obliku i sadržini poreske prijave za utvrđivanje poreza
   na dobit pravnih lica"* ("Službeni list Crne Gore", br. 114/24 od 29.11.2024), которая
   **применяется начиная с decl. за 2024 год** и, следовательно, актуальна и для 2025/2026.
   Содержит: полный бланк Obrazac PD (строки 1–58), Dodatak PG1, Dodatak PG2 и
   **UPUTSTVO ZA POPUNJAVANJE** — официальную построчную инструкцию по каждой строке 1–58.
   → это основа всей таблицы ниже.

2. **Более ранняя (промежуточная) редакция формы**, отражающая правки 2022 г.
   (115/22 от 14.10.2022, 127/22 от 21.11.2022), но ещё без правок 114/24:
   `https://wapi.gov.me/download/2ae946d5-418e-41a1-912b-99c6369dc1e7?version=1.0` (PDF, 821 KB,
   опубликован на gov.me 09.02.2023). Содержит форму со строками 1–62 (включая ещё не удалённые
   на тот момент строки "Iznos poreza koji se duguje / Preplaćeni porez / Iznos mjesečne
   akontacije" — их в редакции 2024 г. уже нет) и **отдельную такую же построчную инструкцию**.
   Использована для сверки историчности нумерации (см. §4).

3. **Zakon o porezu na dobit pravnih lica** — консолидированный текст (по состоянию на правки
   028/23 от 10.03.2023, издание "Pravni ekspert doo, Katalog propisa 2023"):
   `https://www.investinkotor.me/files/documents/1706539733-Zakon%20o%20porezu%20na%20dobit%20pravnih%20lica.pdf`
   — прочитан целиком (16 страниц). Источник ставок налога (čl. 28), правил переноса убытков
   (čl. 25, čl. 22 st. 5), капитальных доходов/убытков (čl. 21‑24), трансфертных цен (čl. 38,
   38a‑38c), налогового кредита за рубежом (čl. 33), сроков подачи (čl. 40).

4. **KPMG Montenegro** — подтверждение хронологии и содержания изменений:
   - "Progresivno oporezivanje dobiti, transferne cijene i PDV niže stope od 2022. godine"
     (kpmg.com/me, январь 2022) — подтверждает, что прогрессивная шкала 9/12/15% введена
     **с 2022 налогового года**.
   - "Izmjene crnogorskih poreskih propisa" (kpmg.com/me, октябрь 2024) — по краткому
     содержанию из поиска: Skupština Crne Gore приняла 06 и 26.09.2024 пакет налоговых
     поправок (Sl. list CG br. 88/2024 od 13.09.2024 i 94/2024 od 30.09.2024); в т.ч. по
     Zakon o porezu na dobit — новая допустимая статья расходов "izdaci za nacionalne
     sportske saveze" до 5% ukupnog prihoda, применяется **с 1 января 2025 года**. Прямой
     доступ к полному тексту PDF по адресу `assets.kpmg.com/.../Izmjene-crnogorskih-poreskih-propisa.pdf`
     на момент исследования вернул 404 — содержание восстановлено только по сниппету поиска
     (см. пометку уверенности в §2, строка 18).

5. **Устаревшая (историческая) форма PD** для справки о происхождении нумерации (2009–2011 гг.,
   ещё до прогрессивной шкалы, с плоской ставкой): скачан `.doc` с
   `https://posao.crna.gora.me/zakoni_i_ugovori/prijava_poreza_na_dobit_pravnih_lica_-_obrazac_pd.doc`
   — не образует часть основной таблицы, использован только для понимания эволюции.

6. Общий контекст по срокам/порталу: подача — только электронно, через **eprijava.tax.gov.me**;
   c 12.01.2026 налоговая администрация Черногории перевела приём деклараций на новую
   интегрированную систему **IRMS** (irms.tax.gov.me) — это может объяснять часть расхождений
   между печатной формой Pravilnika и фактической XML‑схемой (см. §5).

---

## 1. Общая структура декларации

Титул: **"PRIJAVA ZA UTVRĐIVANJE POREZA NA DOBIT PRAVNIH LICA"** (Obrazac PD).
Подаётся резидентными и нерезидентными юр. лицами — плательщиками налога на прибыль,
за финансовый период (обычно календарный год), не позднее 3 месяцев после его окончания
(čl. 40 Zakona), строго в электронном виде.

Титульная часть (не входит в `Iznos*`, соответствует полям `Zaglavlje` в XSD и разделу
"PODACI O PORESKOM OBVEZNIKU" формы):

| Поле формы | XML (`Zaglavlje`) | Описание |
|---|---|---|
| 0.1 / 0.2 | — (обычно отдельный флаг типа декларации, в XSD не выделен явно) | Osnovna / Izmijenjena prijava (первичная / уточнённая) |
| 0.3 / 0.4 | — | Rezident / Nerezident |
| 1 | `Naziv`, `Pib` | Naziv poreskog obveznika, PIB (8 цифр) |
| 2 | — | Adresa glavnog mjesta poslovanja (opština, ulica, telefon) |
| 3 | `PoreskiPeriodOd` / `PoreskiPeriodDo` | Poreski period (od–do) |
| 4 | — | Bankovni računi |
| 5 | `FinansijskaKonsolidacija` | Odobrena finansijska konsolidacija (znak "x") |
| 6 | `BrojRjesenja` | Broj rješenja kojim je odobrena finansijska konsolidacija |
| 7 | `DatumRjesenja` | Datum rješenja sa r.br. 6 |
| 8 | `OdgovornoLice_Naziv`(?) | Naziv poreskog obveznika na kojeg se prenosi poreska obaveza — **см. примечание ниже** |
| 9 | `OdgovornoLice_JMBG`(?) | PIB poreskog obveznika na kojeg se prenosi obaveza — **см. примечание ниже** |

> ⚠️ **Уверенность: НИЗКАЯ** для строк 8–9 ↔ `OdgovornoLice_Naziv`/`OdgovornoLice_JMBG`.
> По официальной форме поля 8–9 — это **наименование и PIB другого юр. лица** (получателя
> перенесённой налоговой обязанности при консолидации), а не ФИО и ЈМБГ ответственного лица
> подписанта (которое по идее и должно называться `OdgovornoLice_*`, судя по подпункту
> "Potpis ovlašćenog lica ... PIB" в конце формы, но там указывается PIB **самого
> налогоплательщика**, а не физлица; ЈМБГ ответственного лица в печатной форме вообще не
> запрашивается явно). Похоже, `OdgovornoLice_Naziv`/`OdgovornoLice_JMBG` в XSD — это
> отдельная пара полей для лица, подписывающего декларацию (директора), которых на печатной
> форме нет как отдельных вводимых полей в этом виде (только подпись). Требует проверки
> в реальном интерфейсе IRMS.

---

## 2. Построчная таблица (строки 1–58, по официальной инструкции Pravilnika, редакция 114/24)

Обозначения в столбце «Тип»:
- **[БУ]** — берётся из финансовой отчётности (bilans uspjeha/stanja)
- **[Ф]** — вычисляется по формуле от других строк (формула приведена)
- **[РК]** — ручная налоговая корректировка (usklađivanje rashoda/prihoda)
- **[СТ]** — расчёт налога по ставке / шкале
- **[ЛГ]** — льгота / налоговый кредит (umanjenje poreza)

| № | XML‑поле | Название (черногор.) | Перевод | Тип | Как заполняется / формула |
|---|---|---|---|---|---|
| **A. Poslovna dobit i gubici — I. Finansijski rezultat iz bilansa uspjeha** |
| 1 | `Iznos1` | Dobit poslovne godine | Прибыль отчётного года | **[БУ]** | Прибыль по bilans uspjeha (МСФО/МСБУ), до вычета налога на прибыль |
| 2 | `Iznos2` | Gubitak poslovne godine | Убыток отчётного года | **[БУ]** | Убыток по bilans uspjeha, до вычета налога. Взаимоисключающе со стр.1 |
| **II. Kapitalni dobici i gubici** |
| 3 | `Iznos3` | Kapitalni dobici | Капитальные доходы (текущего периода) | **[БУ/РК]** | По ст. 21, 22, 24 Zakona — из данных bilans uspjeha в целях определения суммы капитального дохода |
| 4 | `Iznos4` | Kapitalni gubici | Капитальные убытки (текущего периода) | **[БУ/РК]** | Аналогично стр. 3, по ст. 21,22,24 Zakona |
| **III. Usklađivanje rashoda (корректировка расходов — надбавки к базе)** |
| 5 | `Iznos5` | Troškovi amortizacije iskazani u bilansu uspjeha | Амортизация по бухучёту | **[БУ]** | Сумма амортизации ОС, как в bilans uspjeha |
| 6 | `Iznos6` | Troškovi amortizacije koji se priznaju u poreske svrhe | Амортизация, признаваемая для целей налога | **[РК]** | Расчёт по Pravilniku o razvrstavanju osnovnih sredstava (Sl. list RCG 28/02); группы ОС: I‑5%, II‑15%, III‑20%, IV‑25%, V‑30% (čl.13 Zakona); гр.I — линейный метод, гр.II‑V — регрессивный (declining balance) |
| 7 | `Iznos7` | Troškovi koji nijesu nastali u svrhu obavljanja poslovnih aktivnosti | Расходы не для целей деятельности | **[РК]** | Напр. личные расходы владельца/сотрудников (čl.11 t.1 Zakona) |
| 8 | `Iznos8` | Troškovi koji se ne mogu dokumentovati | Недокументированные расходы | **[РК]** | Нет достоверных документов о сумме/времени/назначении (čl.11 t.2) |
| 9 | `Iznos9` | Kamata za neblagovremeno plaćene poreze i doprinose | Проценты за несвоевременную уплату налогов/взносов | **[РК]** | čl.11 t.3 Zakona |
| 10 | `Iznos10` | Kamata plaćena nerezidentima po stopi većoj od uobičajene komercijalne stope | Проценты нерезидентам сверх рыночной ставки | **[РК]** | čl.11 t.4 |
| 11 | `Iznos11` | Administrativni troškovi plaćeni od strane stalne poslovne jedinice nerezidentnoj centrali | Админ. расходы постоянного представительства в пользу нерезидентного головного офиса | **[РК]** | čl.11 t.5 |
| 12 | `Iznos12` | Primanja zaposlenih ili drugih lica po osnovu raspodjele dobiti | Выплаты сотрудникам/иным лицам из распределения прибыли | **[РК]** | čl.11 t.6 |
| 13 | `Iznos13` | Novčane kazne i penali | Штрафы и пени | **[РК]** | čl.11 t.7 |
| 14 | `Iznos14` | Ispravka vrijednosti pojedinačnih potraživanja od lica kojima se istovremeno duguje, do iznosa obaveze prema tom licu | Обесценение встречной дебиторки в пределах встречной кредиторки | **[РК]** | čl.11 t.8 |
| 15 | `Iznos15` | **Ukupno (a+b)** | Итого по строке 15 | **[Ф]** | `Iznos15 = Iznos15a + Iznos15b` |
| 15a | `Iznos15a` | Prilozi dati političkim organizacijama | Взносы политическим организациям | **[РК]** | čl.11 t.9 |
| 15b | `Iznos15b` | Pokloni i drugi prenosi bez naknade | Подарки и иные безвозмездные передачи | **[РК]** | Не изъятые из налогообложения по Zakonu (новая позиция ред. 114/24) |
| 16 | `Iznos16` | Troškovi zarada, autorskih naknada, naknada po ugovoru o djelu ... koji nijesu isplaćeni u poreskom periodu | Начисленные, но не выплаченные в периоде зарплаты/авт.вознагр./выходные пособия/ и т.п. | **[РК]** | čl.11a Zakona — признаются расходом для налога **в периоде фактической выплаты** (кассовый принцип для этой категории) |
| 17 | `Iznos17` | Troškovi materijala i nabavna vrijednost prodate robe iznad iznosa obračunatog metodom prosječne cijene ili FIFO | Превышение стоимости материалов/товаров сверх среднеlevel./FIFO | **[РК]** | čl.12 Zakona |
| 18 | `Iznos18` | **Ukupno (a+b [+c?])** | Итого по строке 18 | **[Ф]** | `Iznos18 = Iznos18a + Iznos18b (+ Iznos18c?)` — см. §5 по 18c |
| 18a | `Iznos18a` | Izdaci isplaćeni pravnim licima za zdravstvene, socijalne, obrazovne ... svrhe ... iznad 3,5% ukupnog prihoda, uvećan za iznos ovih izdataka isplaćenih fizičkim licima | Пожертвования сверх 3,5% выручки (широкий список гуманитарных/социальных целей) + всё, выплаченное физлицам по этой статье | **[РК]** | čl.14 Zakona; лимит 3,5% ukupnog prihoda; выплаты физлицам по этим целям — не признаются вовсе (добавляются целиком) |
| 18b | `Iznos18b` | Izdaci po osnovu donacija u naučnoistraživačke projekte ili inovacionu infrastrukturu po osnovu kojih je obveznik stekao status korisnika podsticajnih mjera | Донации на научно‑исследоват./инновационную инфраструктуру (по которым получен статус льготника) | **[РК]** | Новая позиция ред. 114/24. **Важно**: во избежание двойного уменьшения налога, сумма этой донации одновременно отражается как льгота в стр. 55b |
| 19 | `Iznos19` | Troškovi reprezentacije iznad 1% ukupnog prihoda | Представительские расходы свыше 1% выручки | **[РК]** | čl.15 Zakona |
| 20 | `Iznos20` | Članarine komorama, savezima i udruženjima iznad 0,1% ukupnog prihoda | Членские взносы палатам/союзам сверх 0,1% выручки | **[РК]** | čl.16 Zakona |
| 21 | `Iznos21` | **Ukupno (a+b)** | Итого по строке 21 | **[Ф]** | `Iznos21 = Iznos21a + Iznos21b` |
| 21a | `Iznos21a` | Otpisana vrijednost sumnjivih potraživanja koja se ne priznaju u poreske svrhe | Списанная сомнительная дебиторка, не признаваемая для налога | **[РК]** | Не выполнены условия čl.17 Zakona (нет доказательств иска/претензии, долг < 365 дней и т.п.) |
| 21b | `Iznos21b` | Ispravka vrijednosti sumnjivih potraživanja koja se ne priznaju u poreske svrhe | Обесценение (резерв) сомнительной дебиторки, не признаваемое для налога | **[РК]** | То же основание, но резерв, а не списание |
| 22 | `Iznos22` | Rezervisanja koja se ne priznaju u poreske svrhe | Резервы, не признаваемые для налога | **[РК]** | Сверх лимитов čl.18 Zakona (для банков/страховщиков — особые правила) |
| 23 | `Iznos23` | Obračunate kamate i pripadajući troškovi prema povezanim licima koje prelaze troškove kamata "van dohvata ruke"; zatezne kamate između povezanih lica | Проценты связанным лицам сверх принципа "вытянутой руки"; пени между связанными лицами | **[РК]** | čl.19 Zakona |
| 24 | `Iznos24` | Rashodi po osnovu obezvređivanja imovine koji se ne priznaju u poreske svrhe | Обесценение активов, не признаваемое для налога | **[РК]** | čl.18a Zakona — разница между балансовой (МСФО) и возмещаемой стоимостью, до момента фактического выбытия/повреждения |
| 25 | `Iznos25` | Rashodi po osnovu efekata promjene računovodstvene politike nastali usled prve primjene MRS/MSFI | Расходы от эффекта смены учётной политики при первом применении МСФО | **[РК]** ⚠ | čl.7 Zakona. **Формально в разделе "Usklađivanje rashoda", но в итоговой формуле (стр.38) ВЫЧИТАЕТСЯ**, т.к. это разовая переходная сумма, прошедшая через капитал/нераспределённую прибыль (минуя bilans uspjeha), но признаваемая вычетом для налога — см. пояснение в §3 |
| **IV. Korekcija rashoda po osnovu transfernih cijena između povezanih lica** |
| 26 | `Iznos26` | Obračunati troškovi primjenom transfernih cijena | Расходы, рассчитанные по трансфертным ценам | **[БУ/РК]** | čl.38 Zakona |
| 27 | `Iznos27` | Obračunati troškovi primjenom cijena "van dohvata ruke" | Расходы по рыночным ценам ("вытянутая рука") | **[РК]** | čl.38b st.1 Zakona |
| 28 | `Iznos28` | Razlika obračunatih troškova (26‑27)>0 | Превышение трансфертных расходов над рыночными | **[Ф]** | `Iznos28 = max(Iznos26 - Iznos27, 0)` |
| **V. Usklađivanje prihoda (корректировка доходов)** |
| 29 | `Iznos29` | Iznos poreza po odbitku na dividende i udjele u dobiti koji je platila nerezidentna filijala rezidentnog poreskog obveznika | Налог у источника на дивиденды, уплаченный нерезидентным филиалом | **[РК]** | Также отражается в стр.57 в пределах налогового кредита по čl.33 st.2 Zakona |
| 30 | `Iznos30` | Obračunate kamate ... prema povezanim licima koje su manje od prihoda po kamati "van dohvata ruke" | Заниженные (относительно рыночных) проценты, начисленные связанным лицам | **[РК]** | čl.20 Zakona — доначисление процентного дохода до рыночного уровня |
| 31 | `Iznos31` | Prihodi po osnovu efekata promjene računovodstvene politike nastali usled prve primjene MRS/MSFI | Доходы от эффекта смены учётной политики при первом применении МСФО | **[РК]** | čl.7 Zakona; зеркально к стр.25 — переходный доход, минующий bilans uspjeha, добавляется в базу |
| **VI. Korekcija prihoda po osnovu transfernih cijena između povezanih lica** |
| 32 | `Iznos32` | Obračunati prihodi primjenom transfernih cijena | Доходы по трансфертным ценам | **[БУ/РК]** | čl.38 Zakona |
| 33 | `Iznos33` | Obračunati prihodi primjenom cijena "van dohvata ruke" | Доходы по рыночным ценам | **[РК]** | čl.38b st.1 |
| 34 | `Iznos34` | Razlika obračunatih prihoda (33‑32)>0 | Занижение дохода относительно рыночного уровня | **[Ф]** | `Iznos34 = max(Iznos33 - Iznos32, 0)` |
| **VII. Ostale korekcije** |
| 35 | `Iznos35` | Prihod od naplaćenih sumnjivih potraživanja ostvaren u periodu, a koji u ranijem periodu nijesu priznata kao rashod | Возврат/погашение ранее списанной сомнительной дебиторки, не признанной расходом ранее | **[РК −]** | Реверс — вычитается, во избежание двойного налогообложения того же дохода |
| 36 | `Iznos36` | Rashodi po impairment-у, ranije ne priznati, za imovinu koja je otuđena/oštećena u periodu | Обесценение, ранее не признанное, теперь (при выбытии/утрате) признаётся | **[РК −]** | čl.18a Zakona — вычет в периоде фактического выбытия |
| 37 | `Iznos37` | Troškovi zarada i sl. koji nijesu bili priznati u prethodnim periodima, a sada isplaćeni | Ранее не признанные зарплатные/подобные расходы, теперь фактически выплаченные | **[РК −]** | Зеркально к стр.16 — вычет в периоде выплаты |
| **VIII. Oporeziva dobit** |
| 38 | `Iznos38` | Oporeziva dobit | Налогооблагаемая прибыль | **[Ф]** | `= 1 − 3 + 4 + 5 − 6 + (7..24 суммарно) − 25 + 28 + 29 + 30 + 31 + 34 − 35 − 36 − 37`, если результат > 0. Иначе результат — убыток, отражается в стр.39 |
| 39 | `Iznos39` | Gubitak | Убыток | **[Ф]** | Зеркальная (отрицательная) формула к стр.38, если итог <0 |
| 40 | `Iznos40` | Iznos gubitka u poreskoj prijavi iz prethodnih godina, do visine oporezive dobiti | Перенесённый убыток прошлых лет (в пределах текущей налог. прибыли) | **[Ф/РК]** | `= min(перенос за 5 лет, Iznos38)`. Источник суммы — **стр.3 приложения PG1** (`IznosPG1_3`) |
| 41 | `Iznos41` | Ostatak oporezive dobiti (38‑40)>0 | Остаток налогооблагаемой прибыли | **[Ф]** | `Iznos41 = Iznos38 − Iznos40` |
| **B. Kapitalni dobici i gubici** |
| 42 | `Iznos42` | Ukupni kapitalni dobici tekuće godine | Итого капитальные доходы текущего года | **[Ф]** | `= Iznos3` (переносится) |
| 43 | `Iznos43` | Ukupni kapitalni gubici tekuće godine | Итого капитальные убытки текущего года | **[Ф]** | `= Iznos4` (переносится) |
| 44 | `Iznos44` | Kapitalni dobici (42‑43)>0 | Чистый капитальный доход | **[Ф]** | `Iznos44 = max(Iznos42 − Iznos43, 0)` |
| 45 | `Iznos45` | Kapitalni gubici (43‑42)>0 | Чистый капитальный убыток | **[Ф]** | `Iznos45 = max(Iznos43 − Iznos42, 0)` |
| 46 | `Iznos46` | Prenijeti kapitalni gubici iz ranijih godina do visine iznosa pod r.br. 44 | Перенесённые капитальные убытки прошлых лет (в пределах стр.44) | **[Ф/РК]** | čl.22 st.5 Zakona — перенос **до 5 лет**. Источник — **стр.3 приложения PG2** (`IznosPG2_3`) |
| 47 | `Iznos47` | Ostatak kapitalnog dobitka (44‑46)≥0 | Остаток капитального дохода | **[Ф]** | `Iznos47 = Iznos44 − Iznos46` |
| **C. Poreska osnovica** |
| 48 | `Iznos48` | Poreska osnovica (41 + 100% od 47)>0 | Налоговая база (до льгот) | **[Ф]** | `Iznos48 = Iznos41 + Iznos47` (капитальный доход включается на 100%, льготной ставки для капитальных доходов сейчас нет — в отличие от некоторых старых редакций) |
| Умен. базы | | **Umanjenje poreske osnovice** | | | |
| 49 | `Iznos49` | **Ukupno (a+b)** | Итого по строке 49 (необлагаемые доходы) | **[Ф]** | `Iznos49 = Iznos49a + Iznos49b` |
| 49a | `Iznos49a` | Prihodi po osnovu dividendi i udjela u dobiti rezidentnih pravnih lica | Дивиденды и доли в прибыли от резидентных юр. лиц | **[БУ/РК −]** | čl.9 Zakona — исключаются из базы (уже обложены на уровне выплачивающей стороны) |
| 49b | `Iznos49b` | Prihodi po osnovu likvidacionog ostatka | Доходы от ликвидационного остатка | **[БУ/РК −]** | čl.9 Zakona, при условии что плательщик уже удержал и уплатил налог у источника |
| 50 | `Iznos50` | Poreska osnovica — oporeziva dobit (48‑49)>0 | Итоговая налоговая база | **[Ф]** | `Iznos50 = Iznos48 − Iznos49` |
| 51 | `Iznos51` | **a) 9% за oporezivu dobit do 100.000€; b) 12% за dio koji prelazi 100.000€; c) 15% за dio koji prelazi 1.500.000€** | Расчёт налога по прогрессивной шкале | **[СТ]** | См. §3 — формула шкалы |
| 51a | `Iznos51a` | (9% часть) | Налог по ставке 9% | **[СТ]** | `= min(Iznos50, 100000) × 9%` |
| 51b | `Iznos51b` | (12% часть) | Налог по ставке 12% | **[СТ]** | `= max(min(Iznos50,1500000) − 100000, 0) × 12%` |
| 51c | `Iznos51c` | (15% часть) | Налог по ставке 15% | **[СТ]** | `= max(Iznos50 − 1500000, 0) × 15%` |
| 52 | `Iznos52` | Ukupan iznos poreza (51a+51b+51c) | Итого начисленный налог | **[Ф]** | `Iznos52 = Iznos51a + Iznos51b + Iznos51c` |
| Умен. налога | | **Umanjenje poreza** | | | |
| 53 | `Iznos53` | Iznos poreza na dobit ostvarenog u nedovoljno razvijenim opštinama | Льгота: налог на прибыль от производственной деятельности в слаборазвитых муниципалитетах | **[ЛГ]** | čl.31 Zakona — для новых предприятий: полное освобождение первые 8 лет (максимум 200 000 € суммарно, пропорционально доле такой прибыли) |
| 54 | `Iznos54` | Iznos poreza na dobit koju pravno lice reinvestira u naučnoistraživačke/inovativne projekte ... Fonda za inovacije CG | Льгота: налог на реинвестированную в НИОКР/инновации прибыль | **[ЛГ]** | čl.23 st.2 Zakona o podsticajnim mjerama za razvoj istraživanja i inovacija (Sl.list CG 82/20). Основание — Rješenje о статусе получателя мер поддержки |
| 55 | `Iznos55` | **Ukupno (a+b+c)** | Итого по строке 55 | **[Ф/ЛГ]** | `Iznos55 = Iznos55a + Iznos55b + Iznos55c` |
| 55a | `Iznos55a` | Ulaganja u udjele ili akcije startapova i spinofova | Вложения в доли/акции стартапов и спин‑оффов | **[ЛГ]** | čl.23 st.3 того же закона |
| 55b | `Iznos55b` | Donacije naučnoistraživačkim ustanovama i subjektima inovacione infrastrukture | Донации научно‑исслед. учреждениям/инновационной инфраструктуре | **[ЛГ]** | Сюда же дублируется сумма из стр.18b — чтобы избежать двойной льготы |
| 55c | `Iznos55c` | Ulaganja u Fond za inovacije CG i/ili druge investicione fondove u inovacionu djelatnost | Вложения в Фонд инноваций ЧГ / иные инвестфонды в инновационную деятельность | **[ЛГ]** | Там же |
| 56 | `Iznos56` | Utvrđena poreska obaveza (52‑53‑54‑55) | Установленное налоговое обязательство | **[Ф]** | `Iznos56 = Iznos52 − Iznos53 − Iznos54 − Iznos55`. ⚠ по печатной форме подпунктов a/b тут **нет** — см. §5 |
| 57 | `Iznos57` | Iznos poreza plaćenog u drugoj državi | Налог, уплаченный в другом государстве | **[ЛГ]** | čl.33 st.2 Zakona — налоговый кредит, не более суммы, которая была бы уплачена в ЧГ на такой доход |
| 58 | `Iznos58` | Ukupna poreska obaveza (56‑57)>0 | Итоговое налоговое обязательство к уплате | **[Ф]** | `Iznos58 = max(Iznos56 − Iznos57, 0)` — это **финальная сумма налога к уплате за период** |

---

## 3. Прогрессивная шкала налога на прибыль (čl. 28 Zakona o porezu na dobit pravnih lica)

Дословно из первоисточника (Katalog propisa 2023, консолидированный текст закона):

> **Član 28**
> (1) Stope poreza na dobit pravnih lica su progresivne.
> (2) Stope poreza na iznos oporezive dobiti iznose:
> 1) do 100.000,00 eura — **9%**;
> 2) od 100.000,01 do 1.500.000,00 eura — **9.000,00 eura + 12% na iznos preko 100.000,01 eura**;
> 3) preko 1.500.000,01 eura — **177.000,00 eura + 15% na iznos preko 1.500.000,01 eura**.

Это подтверждает формулировку из задания **дословно**: 9% до 100k, 9000+12% на диапазон
100k–1.5M, 177000+15% свыше 1.5M.

**Актуальность на 2025/2026**: подтверждена. Прогрессивная шкала введена законом
Zakon o izmjenama i dopunama Zakona o porezu na dobit pravnih lica ("Sl. list CG", 146/21 od
31.12.2021) и **действует с налогового 2022 года** (см. заголовок статьи KPMG,
январь 2022: «Progresivno oporezivanje dobiti ... od 2022. godine»). Последующие правки
Zakona (152/22, 028/23, а также пакет от сентября 2024 г. — Sl.list CG 88/2024, 94/2024)
**ставок налога не меняли** — правки сентября 2024 касались новой статьи допустимых
расходов (спортивные федерации, лимит 5% выручки, с 1.01.2025) и ряда норм по НДС/другим
налогам. Веб‑источник (ekonomik.co, статья про налоги 2026 года) подтверждает те же
9/12/15% как действующие в 2026 г. — **Уверенность: ВЫСОКАЯ**.

Формула построчно (см. также таблицу выше, стр. 51a‑c):

```
Iznos51a = min(Iznos50, 100000) * 0.09
Iznos51b = max(min(Iznos50, 1500000) - 100000, 0) * 0.12
Iznos51c = max(Iznos50 - 1500000, 0) * 0.15
Iznos52  = Iznos51a + Iznos51b + Iznos51c
```

### Пояснение к знаку строки 25/31 в формуле стр.38/39

Строка 25 формально находится в разделе III «Usklađivanje rashoda» (расходные
корректировки — обычно надбавки к базе), однако в итоговой формуле стр.38
(`Oporeziva dobit`) она **вычитается**, а не прибавляется — и это не опечатка: и в
формуле стр.38 (`...-25+28...`), и в зеркальной формуле стр.39 (`...+25-28...`) знак
у 25 последовательно противоположен знакам остальных элементов раздела III. Причина —
семантика самой строки: «Rashodi po osnovu efekata promjene računovodstvene politike
nastali usled prve primjene MRS/MSFI» (čl. 7 Zakona) — это **разовый переходный расход**,
который прошёл через капитал/нераспределённую прибыль при первом переходе на МСФО, **минуя
bilans uspjeha** (т.е. не сидит внутри стр.1/2), но который закон признаёт вычитаемым для
налоговой базы — отсюда вычитание. Зеркальная позиция — строка 31 («Prihodi po osnovu...»,
раздел V «Usklađivanje prihoda») — доход того же происхождения, соответственно **прибавляется**
к базе. — **Уверенность: СРЕДНЯЯ-ВЫСОКАЯ** (сама формула процитирована из первоисточника
дословно; логическое объяснение знака — моя интерпретация на основе текста строк 25/31 и
čl. 7 Zakona, не встречал отдельного explicit комментария у налоговой к этому нюансу).

---

## 4. Перенос убытков — Dodatak PG1 и Dodatak PG2

**PG1 = "Prenos poslovnih gubitaka"** (перенос операционных убытков), **PG2 = "Prenos
kapitalnih gubitaka"** (перенос капитальных убытков). Это **не два периода**, а два
**разных вида** убытка (операционный / капитальный), каждый — с одинаковой внутренней
структурой "расчёт остатка убытка, доступного к переносу на текущий период, из **пяти**
предыдущих лет".

Правовое основание: **čl. 25 Zakona** — операционные убытки переносятся на будущие периоды,
но **не более чем на 5 лет** («ne duže od pet godina»); **čl. 22 st.5 Zakona** — аналогичное
правило для капитальных убытков («u narednih pet godina»).

### Структура Dodatak PG1 (дословно из формы, 3 строки на бланке):

| Redni broj (форма) | Описание (форма) | XML‑поля |
|---|---|---|
| 1 | Iznos dobiti iskazane pod rednim brojem **38.** poreske prijave | `IznosPG1_1` = переносится значение `Iznos38` |
| 2 | Iznos gubitka iz prethodnih **pet godina**: `___god. ___€` × 5 строк, с итоговой суммой справа | `IznosPG1_21..IznosPG1_25` = сумма убытка за каждый из 5 предыдущих лет (год N‑1 .. N‑5, годы вписываются вручную в пустые поля `___god.`); `IznosPG1_2` = сумма этих пяти сумм (итог, напечатанный чертой под перечнем годов на бланке) |
| 3 | Iznos gubitka za pokriće (upisuje se **manji** od iznosa pod rednim brojem 1 ili 2) | `IznosPG1_3` = `min(IznosPG1_1, IznosPG1_2)` — это значение затем переносится в **стр.40** основной декларации |

### Структура Dodatak PG2 (аналогично, но по капитальным убыткам):

| Redni broj (форма) | Описание (форма) | XML‑поля |
|---|---|---|
| 1 | Iznos kapitalnog dobitka iskazan pod rednim brojem **44.** poreske prijave | `IznosPG2_1` = переносится значение `Iznos44` |
| 2 | Iznos kapitalnog gubitka iz prethodnih pet godina: `___god. ___€` × 5 | `IznosPG2_21..IznosPG2_25` = убыток по годам N‑1..N‑5; `IznosPG2_2` = сумма пяти |
| 3 | Iznos gubitka za pokriće (manji od 1 ili 2) | `IznosPG2_3` = `min(IznosPG2_1, IznosPG2_2)` — переносится в **стр.46** основной декларации |

Сопоставление `_21..._25` с «пятью годами» **выводится напрямую из визуальной структуры
бланка** (пять одинаковых пустых строк `___ god. ___ €` под пунктом 2 в PDF‑форме,
см. цитату ниже) и является наиболее логичным объяснением наименования полей —
**Уверенность: ВЫСОКАЯ** (структурно обоснованное соответствие «номер строки формы (2) +
порядковый номер года (1‑5)» → `_21..._25`), хотя дословного текста «поле _21 = первый
год» в самой инструкции, разумеется, нет (в бланке это просто повторяющиеся пустые строки,
не пронумерованные официально глубже, чем "2").

Цитата из формы (PG1, дословно):
```
2. Iznos gubitka iz prethodnih pet godina:
   __________ god. ___________€
   __________ god. ___________€
   __________ god. ___________€
   __________ god. ___________€
   __________ god. ___________€
                                    ____________  ← сумма (итог) сюда
3. Iznos gubitka za pokriće (upisuje se manji od iznosa pod rednim brojem 1 ili 2)
```

---

## 5. Открытые вопросы / расхождения между официальным бланком и XSD

Официальный бланк (редакция 114/24, самая свежая найденная) заканчивается строкой **58**
и не содержит подпунктов **18c**, **56a**, **56b**, строки **59** вообще нет. При этом
`pd.xsd` в репозитории явно содержит `Iznos18c`, `Iznos56a`, `Iznos56b`, `Iznos59`.
Это значит одно из двух:
(a) существует ещё более новая, не найденная мной в открытом доступе, редакция Pravilnika
(возможно, синхронизированная с запуском новой системы **IRMS**, которая заменила портал
eprijava и была переведена в промышленную эксплуатацию **12.01.2026** — т.е. уже после
редакции 114/24 от конца 2024 г.); либо
(b) реальная XML‑схема системы приёма деклараций **исторически шире**, чем то, что печатается
на бумажном бланке (в электронных системах это обычное дело — доп. поля для валидации/расчёта,
не выводимые как отдельные визуальные ячейки бланка).

Официального текста, подтверждающего смысл именно этих четырёх полей, я не нашёл. Ниже —
**обоснованные гипотезы**, не более того:

- **`Iznos18c`** — Уверенность: **СРЕДНЯЯ**. Наиболее вероятная гипотеза: новая статья
  допустимых (с лимитом) расходов — *«Izdaci za nacionalne sportske saveze»* (расходы на
  национальные спортивные федерации), введённая пакетом поправок к Zakonu от сентября
  2024 г. (Sl. list CG 88/2024, 94/2024), с лимитом **5% ukupnog prihoda** и вступлением в
  силу **с 1 января 2025 года** (т.е. actually применяется к декларации за 2025 год — что
  по срокам совпадает с переходом на IRMS в январе 2026 и вполне могло войти в форму позже
  114/24). По аналогии с 18a (лимит 3,5%) и 18b (донации на науку/инновации) вероятная
  формулировка: *«Iznos izdataka za nacionalne sportske saveze iznad 5% ukupnog prihoda»*
  (превышение лимита добавляется к налоговой базе). Это моя реконструкция по косвенным
  данным (сниппет поиска KPMG) — **сам текст статьи не читал напрямую** (assets.kpmg.com
  вернул 404 при попытке скачать полный PDF).

- **`Iznos56a` / `Iznos56b`** — Уверенность: **НИЗКАЯ**. Строка 56 (`Utvrđena poreska
  obaveza`) на бланке — одно значение без подпунктов. Гипотеза (не подтверждена):
  разбивка может быть связана с блоком «Podaci o poreskom konsolidovanju» (поля 5–9
  титульной части, где часть налоговой обязанности **передаётся** другому члену группы при
  finansijskoj konsolidaciji) — т.е. `56a` = «сопственная» обязанность до консолидации,
  `56b` = сумма, переданная/принятая в рамках консолидации. Альтернативная гипотеза:
  технический артефакт схемы IRMS без видимого отражения на бланке. **Требует проверки**.

- **`Iznos59`** — Уверенность: **НИЗКАЯ**. У бланка 2024 г. нет аналога. У **более старой**
  редакции формы (та, что действовала примерно до 2017 г., где я читал полную инструкцию
  1‑го скачанного PDF) далее шли строки: 59 «Iznos poreza koji se duguje», 60 «Preplaćeni
  porez», 61 «Iznos preplaćenog poreza za koji se traži povraćaj», 62 «Iznos mjesečne
  akontacije». В редакции 114/24 они **отсутствуют на бланке** (декларация заканчивается
  Ukupna poreska obaveza, стр.58) — но `Iznos59` может быть служебным полем IRMS,
  унаследованным от этой более старой логики (напр. для сверки с уже уплаченными
  авансовыми платежами по данным лицевого счёта налогоплательщика — и потому не выводимым
  как ручное поле бланка, а вычисляемым системой).

**Рекомендация для верификации** (не проверял в рамках этого исследования — нет доступа к
залогиненной сессии IRMS/eprijava): либо (1) выполнить тестовое заполнение декларации PD
через реальный интерфейс irms.tax.gov.me / eprijava.tax.gov.me и посмотреть подписи полей
19c/56a/56b/59 на экране, либо (2) перехватить фактический XML/JSON payload формы через
сетевые запросы браузера при тестовой отправке, либо (3) напрямую запросить свежую
редакцию Pravilnika/Obrasca PD у eprijava@tax.gov.me (адрес техподдержки, указан в
новостях gov.me).

---

## 6. Прочее (справочно)

- **Um. poreske osnovice за капитальный доход**: льготной пониженной ставки для капитального
  дохода в действующей редакции нет — стр.47 включается в базу **на 100%** (формула стр.48:
  `41 + 100% od 47`). В более старых редакциях закона такая формулировка (явно «100% od»)
  уже присутствовала, т.е. это не новация 2024 г., а устоявшаяся норма.
- **Финансовая консолидация (grupno oporezivanje)** — čl. 35‑37 Zakona: материнская и
  дочерние компании образуют группу, если материнская прямо/косвенно контролирует ≥75%
  акций/долей дочерней; одобренная консолидация действует минимум 5 лет; убытки одной
  компании группы могут перекрывать прибыль других в рамках консолидированной декларации.
  Именно это отражают поля 5‑9 титульной части.
- **Трансфертное ценообразование** — čl. 38, 38a, 38b, 38c Zakona: принцип "van dohvata
  ruke" (arm's length), 6 методов определения рыночной цены (метод сопоставимой
  неконтролируемой цены, цена+наценка, перепродажная цена, чистая маржа, распределение
  прибыли, комбинированный/иной метод); порог документации — 75 000 € в год по сделкам со
  связанным лицом (для «крупных» плательщиков — без порога).
- **Срок подачи и уплаты**: не позднее 3 месяцев после окончания периода (обычно —
  31 марта); только электронно (čl. 40 st. 6 Zakona).

---

## Сводка по уверенности

| Блок | Уверенность |
|---|---|
| Титул/шапка формы (поля 1‑9) | Высокая, кроме сопоставления полей 8‑9 vs `OdgovornoLice_*` — низкая |
| Строки 1‑58 основной таблицы (описание, тип, формулы) | **Высокая** — дословно по официальному Uputstvu (редакция 114/24) |
| Прогрессивная шкала налога (9/12/15%, пороги 100k/1.5M) | **Высокая** — дословная цитата čl. 28 Zakona + подтверждение актуальности на 2025/2026 из независимых источников |
| Перенос убытков (5 лет, PG1/PG2, поля `_1/_2/_21‑25/_3`) | **Высокая** для сути (5 лет, čl.25 и čl.22 st.5 Zakona); **высокая** для структурной интерпретации `_21..._25` как пяти лет, хоть и не названа так дословно в первоисточнике |
| `Iznos18c`, `Iznos56a/56b`, `Iznos59` | **Низкая/средняя** — не подтверждено официальным текстом, только обоснованные гипотезы (см. §5) |
| Знак строки 25/31 в формуле стр.38/39 | Средняя‑высокая (формула — точная цитата; толкование знака — моя интерпретация) |

## Источники (сводный список ссылок)

- Pravilnik (действующая редакция, консолид. текст + приложение с формой и Uputstvom):
  https://www.gov.me/dokumenta/2ae946d5-418e-41a1-912b-99c6369dc1e7 ;
  https://wapi.gov.me/download/618575b3-37ff-415a-8d2d-7c5524dd459c?version=1.0 (текст правок) ;
  https://wapi.gov.me/download/69fb7c10-ee3f-4e44-a6ee-10c0e86609d7?version=1.0 (форма+инструкция, PDF)
- Более ранняя редакция формы (для сверки эволюции нумерации):
  https://wapi.gov.me/download/2ae946d5-418e-41a1-912b-99c6369dc1e7?version=1.0
- Zakon o porezu na dobit pravnih lica (консолид. текст, Katalog propisa 2023):
  https://www.investinkotor.me/files/documents/1706539733-Zakon%20o%20porezu%20na%20dobit%20pravnih%20lica.pdf
- KPMG Montenegro — прогрессивная шкала с 2022 г.:
  https://kpmg.com/me/en/home/insights/2022/01/progresivno-oporezivanje-dobiti-transferne-cijene-i-pdv-niza-stopa-od-2022-godine.html
- KPMG Montenegro — поправки сентября 2024 (спортивные федерации, 5% лимит с 1.01.2025):
  https://kpmg.com/me/en/home/insights/2024/10/izmjene-crnogorskih-poreskih-propisa.html
  (полный PDF `assets.kpmg.com/.../Izmjene-crnogorskih-poreskih-propisa.pdf` вернул 404 —
  содержание восстановлено только по сниппету поиска)
- gov.me — новость об изменении Pravilnika (114/24):
  https://www.gov.me/clanak/pravilnik-o-izmjeni-pravilnika-o-obliku-i-sadrzini-poreske-prijave-za-utvrdivanje-poreza-na-dobit-pravnih-lica
- Портал подачи / переход на новую систему IRMS:
  https://eprijava.tax.gov.me ; https://irms.tax.gov.me/public/ ;
  запуск IRMS в промышленную эксплуатацию — 12.01.2026 (по сниппету поиска, не проверено
  напрямую документом с этой датой)
- Обзор текущих ставок на 2026 г. (вторичный источник, для перекрёстной проверки):
  https://www.ekonomik.co/me/news/taxes-in-Montenegro-in-2026
