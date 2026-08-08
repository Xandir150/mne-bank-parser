# mapping.json — notes, caveats, open issues

Generated from:

- `xsd/fi.xml` — official FI template (7 sections, 315 `Stavka` rows)
- `xsd/fi.xsd` — schema (column names for the two non-tabular sections)
- `fi1c/reference/mne_chart.tsv` — 1C chart of accounts (1529 accounts, 1277 on-balance / 252 off-balance)

Every row of every section of the template is present in `mapping.json`, in template order.
`redniBroj` / `pozicija` / `grupaRacuna` are copied verbatim from the template so the file can be
diffed against a newer `fi.xml`.

## 1. Row counts by kind

| Section | header | account | formula | manual | total |
|---|---|---|---|---|---|
| BilanStanja | 2 | 68 | 21 | 1 | 92 |
| BilanUspjeha | 0 | 41 | 19 | 6 | 66 |
| IskazOTokovimaGotovine | 3 | 0 | 11 | 26 | 40 |
| TokoviGotovine3a | 3 | 0 | 9 | 28 | 40 |
| IskazOPromenamaKapitala | 0 | 0 | 4 | 5 | 9 |
| StatistickiAneks | 1 | 61 | 0 | 1 | 63 |
| ObracunAmortizacije | 0 | 0 | 0 | 5 | 5 |
| **total** | **9** | **170** | **64** | **72** | **315** |

`balance` distribution over the 170 account rows:

- BilanStanja: 37 `debitNet`, 29 `creditNet`, 1 `creditOnly` (109), 1 `debitOnly` (110)
- BilanUspjeha: 25 `turnoverCreditNet`, 16 `turnoverDebitNet`
- StatistickiAneks: 30 `debitNet`, 8 `creditNet`, 9 `turnoverCreditNet`, 14 `turnoverDebitNet`

## 2. Validation results

1. **JSON validity** — `json.load()` OK.
2. **Prefix existence** — 11 prefixes do not exist in `mne_chart.tsv` (see §3).
3. **Formula references** — 0 broken references; every token of every `formula` resolves to an
   existing `redniBroj` of the same section. No `formula` on a non-formula row, no duplicate
   `redniBroj` inside a section, no `accounts`/`balance` on non-account rows.
4. **Coverage** — 4 synthetic on-balance accounts of classes 0–6 (plus 2 of their sub-accounts)
   are not reachable from any account row (see §4).

## 3. Prefixes referenced by the template but absent from the 1C chart

These rows will always evaluate to 0 until the chart is extended (or the row is remapped).

| Section | Row | Prefix | Comment |
|---|---|---|---|
| BilanStanja | 109 | 335, 336, 337 | chart has only 330–334 under class 33 |
| BilanStanja | 110 | 335, 336, 337 | same |
| BilanStanja | 128 | 467 | "Kratkoročna rezervisanja"; chart class 46 has 460–465, 469 only. **Consequence:** short-term provisions currently land in row 139 (`45 i 46`), and row 128 stays 0. |
| BilanUspjeha | 220 | 592 | chart class 59 has 590, 591, 599 only |
| BilanUspjeha | 255 | 335 | see 109 |
| BilanUspjeha | 256 | 336 | see 109 |
| BilanUspjeha | 257 | 337 | see 109 |

The prefixes are deliberately **kept** in the mapping: they are what the official form prescribes,
and the calculation engine must simply return 0 for a prefix with no matching accounts.

## 4. On-balance accounts of classes 0–6 not covered by any account row

| Account | Name | Assessment |
|---|---|---|
| 565 | Rashodi od učešća u gubitku zavisnih pravnih lica i zajedničkih ulaganja (metod udjela) | **Real gap.** Belongs economically to BilanUspjeha 237/240, but the form enumerates only `562, 563, 564, 569`. Recommend adding `565` to row 240 `include` after confirming with the accountant. |
| 665 (+ 6650, 6651) | Prihodi po osnovu učešća u dobitku zavisnih pravnih lica i zajedničkih ulaganja | **Real gap.** Mirror of 565; belongs to BilanUspjeha 222 (rows 223–225 enumerate only `660/661/669`). Recommend adding `665` to row 223 or 225. |
| 599 | Prenos rashoda | **Expected.** Clearing/transfer account, nets to zero at period end; intentionally outside the form. |
| 699 | Prenos prihoda | **Expected.** Same as 599. |

Everything else is reachable: **1242 of the 1248** on-balance accounts of classes 0–6 (classes 0–4
via BilanStanja, classes 5–6 via BilanUspjeha) — 99.5 % coverage. Coverage rule used by the validator: account `A` is covered by an
account row if some `include` prefix `P` satisfies `A.startswith(P)` (and `A` is not caught by an
`exclude` prefix) **or** `P.startswith(A)` (i.e. `A` is a synthetic parent whose children are mapped).

Class 7 (721, 722 …) is used only by BilanUspjeha rows 246/247 and is outside the coverage
requirement; classes 8–9 are off-balance and out of scope.

## 5. Semantic decisions and deviations

### 5.1 Deliberate deviation: BilanStanja 114 / 115 use `debitNet`, not `creditNet`

The blanket rule "PASIVA → creditNet" breaks for the two loss rows:

- 114 = account 350 "Gubitak ranijih godina"
- 115 = account 351 "Gubitak tekuće godine"

Both carry **debit** balances, and row 111 is `112+113-114-115`. With `creditNet` they would come
out negative and the subtraction would *add* the loss back, inverting the sign of accumulated
result and thereby of total capital (101) and total liabilities (144). They are therefore mapped
as `debitNet`, i.e. reported as positive amounts that the formula subtracts — which is how the
official form is filled in.

If the calculation engine prefers the literal spec, flip these two to `creditNet` **and**
simultaneously change formula 111 to `112+113+114+115`.

### 5.2 BilanStanja 109 / 110 — split of accounts 330–337 by side

- 109 `creditOnly` over `330, 331, 332, 333, 334, 335, 336, 337`
- 110 `debitOnly` over `331 … 337`

Per template, 330 participates only in row 109. Since 330 ("Revalorizacione rezerve") is a pure
passive account, `creditOnly` and `creditNet` give the same figure for it; the `creditOnly`
semantics matter only for the dual-sided 331–337 group. The two rows split the *same* accounts by
side, which is why 331–334 appear in both rows — expected, not a double count (105 subtracts 110).

### 5.3 BilanUspjeha 243 — `690-590` expressed as one account row

The template puts an expression, not an account group, into `GrupaRacuna`. It is encoded as
`kind: "account"`, `include: ["690", "590"]`, `balance: "turnoverCreditNet"`, because
credit-net over the pair equals `(credit−debit of 690) + (credit−debit of 590)` = income of
discontinued operations − expense of discontinued operations = exactly `690 − 590`.

### 5.4 BilanUspjeha 208 — bogus `GrupaRacuna`

Row 208 "Troškovi poslovanja (209+210+210a)" carries `GrupaRacuna = 208`, which is not an account
(class 2 = receivables). Treated as a template artifact: `kind: "formula"`, `formula: "209+210+210a"`.
The raw string is preserved in `grupaRacuna` for traceability.

### 5.5 Formula rows that also carry a `GrupaRacuna` (control totals)

| Section | Row | GrupaRacuna | Formula |
|---|---|---|---|
| BilanStanja | 003 | 01 | 004+005+006+007 |
| BilanStanja | 122 | 41 | 123+124 |
| BilanUspjeha | 208 | 208 (bogus) | 209+210+210a |

`kind` is `formula` (the sub-rows are authoritative). The group is a useful **reconciliation
check**: `debitNet(01)` must equal row 003, `creditNet(41)` must equal row 122. Worth wiring into
the engine as a warning.

### 5.6 BilanUspjeha 260 — sign kept as printed

`IX. NETO SVEOBUHVATNI REZULTAT ( 248-259)` is encoded literally as `248-259`, although
comprehensive income is normally `248+259` (net result + other comprehensive result). The template
text is followed; verify against a filed report before trusting row 260.

### 5.7 BilanUspjeha 250–257 use `turnoverCreditNet`

Rows 250–257 report the *change* in each other-comprehensive-income component (330–337) during the
period, so period turnover (credit − debit) is used, not the closing balance. A gain is positive, a
loss negative.

### 5.8 BilanUspjeha 247 — account 722

722 "Odloženi poreski rashodi i prihodi perioda" is two-sided. Mapped `turnoverDebitNet` (a
deferred tax *expense* is positive, deferred tax income negative), consistent with row 245 being
subtracted in 248.

### 5.9 StatistickiAneks — valuation-allowance rows use `creditNet`

Rows 034, 035, 042, 043, 047, 053, 054, 059 (`0108, 0109, 0118, 0119, 0129, 0148, 0149, 0159`) are
contra-asset accounts with credit balances. They are mapped `creditNet` so that the annex shows a
positive allowance; all other class-0 annex rows use `debitNet`.

Also note: rows **026, 027 and 028** (patents / copyright / licence income) all point at the same
account **652**. The template gives no analytical split — filling all three from 652 would triple
count. Recommendation: compute 026 from 652 and leave 027/028 to the accountant, or drive them from
1C sub-conto. Currently all three are `account` rows over 652 as printed.

### 5.10 Range expansion `440-449`

BilanStanja 133 (`433, 434, 440-449`) is expanded to accounts that actually exist in the chart:
**440, 441, 442, 449** (443–448 do not exist). Re-run the generator if the chart gains new 44x
accounts.

### 5.11 Sections without account rules

- **IskazOTokovimaGotovine / TokoviGotovine3a** — the template has no `GrupaRacuna` at all. All
  detail rows are `manual`; only the subtotals are `formula`. Note that `(1 do 3)` / `(1 do 5)` /
  `(I-II)` in `Pozicija` refer to *item numbers within the block*, and were resolved to the actual
  `redniBroj` values (e.g. ITG 301 → `302+303+304`, 311 → `301-305`; TG3a 311 → `301+…+310`).
- **IskazOPromenamaKapitala** — a 9×10 matrix, not a list. The template has no `RedniBroj`; the
  row number lives in `Pozicija` (`1.` … `9.`) and the label in `Opis`. In the mapping,
  `redniBroj` = `1`…`9` and `pozicija` = the `Opis` text. Rows 3, 5, 7, 9 are `formula`
  (`1+2`, `3+4`, `5+6`, `7+8`); the rest are `manual`.
  Cell ids per column (`RedniBr2`…`RedniBr11`, rows 1→9):
  col2 401–409, col3 410–418, col4 419–427, col5 428–436, col6 437–445, col7 446–454,
  col8 455–463, col9 464–472, col10 473–481, col11 482–490.
- **ObracunAmortizacije** — 5 rows keyed by `Grupa` I…V with columns `Iznos2`…`Iznos8`; `Iznos6`
  holds the statutory rate (2.5 / 10 / 15 / 20 / 30 %). No positions, no accounts in the template —
  all `manual`. `redniBroj` is set to the roman group number and `pozicija` synthesised as
  "Amortizaciona grupa N (stopa X%)".

### 5.12 Manual rows in BilanStanja / BilanUspjeha

BilanStanja 116 (učešće koje ne obezbjeđuje kontrolu — consolidation only);
BilanUspjeha 258, 261, 262, 263, 264, 265 (deferred tax on OCI, earnings per share,
result attributable to owners / NCI). No account basis in the chart.

## 6. `(dio)` rows — accounts split across several positions

38 rows include at least one account marked `(dio)` in the template, i.e. only *part* of the
account belongs to that position. These **cannot be computed from the synthetic account alone** —
they need analytics (sub-account, sub-conto, or a manual split). Every such account is listed in
`accounts.dio` while also staying in `accounts.include`, so an engine can (a) compute a provisional
figure and (b) flag the row as requiring review.

### BilanStanja

| Row | GrupaRacuna | dio accounts | Split needed between |
|---|---|---|---|
| 017 | `030, 039(dio)` | 039 | 017–023 all share 039 |
| 018 | `033(dio), 039(dio)` | 033, 039 | 018 vs 020 share 033 |
| 019 | `031(dio), 032(dio), 039(dio)` | 031, 032, 039 | 019 vs 021 share 031/032 |
| 020 | `033(dio), 039(dio)` | 033, 039 | see 018 |
| 021 | `031(dio), 032(dio)` | 031, 032 | see 019 |
| 022 | `032(dio), 034, 035, 036, 039(dio)` | 032, 039 | 032 also in 019/021 |
| 023 | `038, 039(dio)` | 039 | — |
| 032 | `202, 203, 209(dio)` | 209 | 209 (allowance) split over 032/033/034 |
| 033 | `200, 209(dio)` | 209 | same |
| 034 | `201, 209(dio)` | 209 | same |
| 040 | `236(dio)` | 236 | 040 vs 042 |
| 042 | `23 osim 236(dio) i osim 237` | 236 | the *remainder* of 236 |
| 107 | `322(dio)` | 322 | statutarne vs druge rezerve (107 vs 108) |
| 108 | `322(dio)` | 322 | same |
| 119 | `404(dio)` | 404 | part of 404 stays in 121? (121 excludes 404 entirely) |
| 120 | `400(dio)` | 400 | part of 400 stays in 121? (121 excludes 400 entirely) |
| 126 | `495(dio)` | 495 | 126 (long-term) vs 143 (short-term) |
| 130 | `422–425(dio), 426, 429(dio)` | 422, 423, 424, 425, 429 | non-bank lenders vs banks (130 vs 131) |
| 131 | `422–425(dio), 429(dio)` | 422, 423, 424, 425, 429 | same |
| 134 | `439(dio)` | 439 | bills payable vs other (134 vs 138) |
| 138 | `439(dio)` | 439 | same |
| 143 | `490, 491, 494, 495(dio), 496, 497, 499` | 495 | see 126 |

**Note on 119/120/121:** the template says 119 = `404(dio)`, 120 = `400(dio)`, but 121 =
`40, sem 400 i 404` — i.e. 121 excludes 400 and 404 *entirely*. So the "(dio)" on 119/120 has no
counterpart, and if 404/400 are not fully consumed by 119/120 the difference disappears from the
balance sheet. Practical reading: treat 119 = whole 404 and 120 = whole 400 unless analytics say
otherwise (that makes 118 = 119+120+121 = whole class 40, which reconciles).

### BilanUspjeha

| Row | GrupaRacuna | dio accounts | Split needed between |
|---|---|---|---|
| 210 | `53, 54 (dio) i 55` | 54 | 54 minus 540 (540 is row 210a) — practical rule: `include 54, exclude 540` |
| 212 | `52 (dio)` | 52 | net wages |
| 214 | `52 (dio)` | 52 | wage taxes |
| 215 | `52 (dio)` | 52 | pension contributions |
| 216 | `52 (dio)` | 52 | other contributions |
| 218 | `580, 581, 582, 589 (dio)` | 589 | fixed-asset part of 589 |
| 219 | `584, 589 (dio)` | 589 | current-asset part of 589 |
| 223 | `660(dio)` | 660 | 660 shared by 223 / 227 / 231 |
| 224 | `661(dio)` | 661 | 661 shared by 224 / 228 / 232 |
| 225 | `669(dio)` | 669 | 669 shared by 225 / 229 / 233 |
| 227 | `660(dio)` | 660 | see 223 |
| 228 | `661(dio)` | 661 | see 224 |
| 229 | `662(dio), 663(dio), 664(dio), 669(dio)` | 662, 663, 664, 669 | shared with 233 |
| 231 | `660(dio)` | 660 | see 223 |
| 232 | `661(dio)` | 661 | see 224 |
| 233 | `662(dio), 663(dio), 664(dio), 669(dio)` | 662, 663, 664, 669 | shared with 229 |

**Suggested defaults** (until analytics exist), all of which keep the section totals correct
because the parent subtotal rows sum the split rows:

- 210: `54 osim 540` (mechanical, safe)
- 212/214/215/216: split 52 by sub-account — 520/521 → 212, 522…526 → 214–216 per the 1C chart
  (`52` has 520, 521, 522, 523, 524, 525, 526, 529); needs an accountant's sign-off
- 223/227/231 (660), 224/228/232 (661), 225/229/233 (669): put the whole account into the *first*
  row of the group (223 / 224 / 225) and 0 into the rest — total 241 stays correct
- 218/219 (589): put all of 589 into 219 (current assets) by default

## 7. Rows where the same account appears in more than one position of the same section

25 accounts are referenced by two or more account rows of the same section. **All 25 are exactly
the `(dio)` cases from §6 plus the 109/110 side split** — there is no accidental double counting.
Full list: BS 031, 032, 033, 039, 209, 322, 331, 332, 333, 334, 422, 423, 424, 425, 429, 439, 495;
BU 52, 589, 660, 661, 662, 663, 664, 669.

## 8. Regeneration

```bash
cd fi1c/mapping
python3 build_mapping.py       # regenerates mapping.json from ../../xsd/fi.xml + ../reference/mne_chart.tsv
python3 validate_mapping.py    # prints counts, missing prefixes, broken refs, coverage, dio rows
```

The mapping is produced from the template by a deterministic script; if `fi.xml` or the chart
changes, regenerate rather than hand-edit. The rule tables in `build_mapping.py` (formula
expressions, balance overrides, the three special `accounts` overrides for BS 109/110 and BU 243)
are the only hand-written parts — everything else is parsed from `GrupaRacuna`.

`GrupaRacuna` grammar handled by the parser:

- separators: `,`, ` i `, ` и ` (cyrillic)
- exclusions: `osim`, `sem`, `bez` (everything after the keyword is excluded)
- partial marker: `NNN(dio)` / `NNN (dio)` → code goes into `include` **and** `dio`
  (in an exclusion clause it goes into `exclude` **and** `dio`)
- ranges: `440-449` → expanded over codes that exist in the chart
- cyrillic phrases `potr. saldo rač. …` / `dugov. saldo rač. …` → hand-coded overrides (BS 109/110)
- the pseudo-expression `690-590` → hand-coded override (BU 243)
