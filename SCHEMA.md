# Workbook schema

The dashboard reads a standard `.xlsx` workbook. The **only fixed contract** is the
sheet names, the section headers on `Accounts`, and the column positions listed
below. Everything else — your categories, subcategories, account names, colours —
is derived from your data.

The simplest way to conform is to start from `finance-tracker-TEMPLATE.xlsx`
(blank, headers in place) or `finance-tracker-SAMPLE.xlsx` (the same structure
filled with synthetic data) and replace the contents.

## Sheets

| Sheet | Required | Drives |
|-------|----------|--------|
| `Accounts` | **Yes** | Balance sheet, net worth, wealth donut, liabilities, net liquid assets |
| `Transactions` | **Yes** | Every spending view, budget actuals, the auto-detected insights |
| `Historical` | Recommended | Month-end cash, savings rate, net savings |
| `Expenses` | Recommended | Budget-vs-actual pacing and safe-to-spend |
| `Bills` | Recommended | Bills due this month, the cash-flow forecast |
| `Dashboard` | Optional | Net pay, payday, cash-reserve goal, target card spend, wealth rates, footer note |
| `Daily Log` | Optional | Month-on-month deltas on the hero tiles |
| `Wish List` | Optional | Affordability panel |
| `Config` | Optional | Currency, locale, colours, groupings, thresholds |

**Graceful degradation.** Missing optional sheets or values don't error — the
affected panel is simply skipped or shows a placeholder. The minimum to render
something useful is `Accounts` plus `Transactions`.

## `Accounts` — required

Section headers in **column A**, one per block, spelled exactly:
`CASH`, `SAVINGS`, `REFUNDS DUE`, `INVESTMENTS`, `LIABILITIES`.

Under each header, one row per line item: **name in column A, value in column D.**
Column D rather than B so the workbook can keep a native-currency amount in B and
a converted value in D; if you hold everything in one currency, put the same
number in both.

**Column E holds a liability's due date.** This is how short-term debt is told
from long-term: anything due within twelve months is deducted from net liquid
assets, anything later (a mortgage) is not. A legacy `S`/`L` text flag still
works as a fallback. Rows whose name contains "subtotal" are skipped.

Optional summary rows, placed under a `SUMMARY` header so they aren't parsed as
line items — label in column A, value in column D (column B for
`Company valuation`): `Net worth`, `Gross assets`, `Total liquid`,
`Total investments`, `Total liabilities`, `Company valuation`.

## `Transactions` — required

The ledger. Header in row 1, then one row per transaction:

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| Date | Month (`YYYY-MM`) | Account | Description | Amount | List 1 (category) | List 2 (subcategory) | Bucket | Confidence |

**Amount positive = money out**, negative = refund or credit.

**Bucket** is one of `Recurring`, `One-off`, `Income`, `Savings`, `REVIEW`.
`Recurring` drives the spending views; `One-off` is excluded from them so a single
capital outlay doesn't distort your averages; `Income` and `Savings` are flows
between your own accounts.

Categories, subcategories and colours are all derived from `List 1` / `List 2`, so
recategorising in the spreadsheet flows through on the next load with no code
change.

## `Historical`

One row per month, oldest first. `A` = month (`YYYY-MM` as text), `B` = cash
balance, `C` = pay/income, `D` = expenses, `E` = net savings, `F` = savings rate
as a fraction (`0.25` = 25%), `G` = pension contribution.

## `Expenses`

A `Reporting month` label in column A with the month in column B. Then a block
beginning with a `Category` header in column A and `Budget` in column B, one row
per category, ended by a `Total` row. Actuals are recomputed from the ledger, so
the `Actual` column is for your own reference only.

## `Bills`

A header row with `Bill` in column A. Then `A` = bill name, `B` = amount,
`C` = direct-debit day of month. Optional column H flags a bill you can't easily
cancel (`F`, `Y` or `1`).

## `Dashboard`

Label/value pairs — label in column A, value in column B. Matched by substring, so
the exact wording is flexible: `Net pay`, `Payday`, `Month opening net liquid`,
`Cash reserve goal`, `Cash reserve monthly target`, `Cash savings rate`,
`All-in wealth rate`, `Target card spend per month`, `Footer note`.

`Target card spend per month` is the forward parameter for the 12-month cash-flow
forecast. Paying a credit-card bill moves money between your own cash and your own
liability, so it leaves net liquidity unchanged and the forecast does not charge it;
what reduces net liquidity is the spending, charged as a smooth daily accrual from
this figure. Set it by hand — it is used verbatim, including zero. Omit the row and
the dashboard falls back to a run-rate averaged over the last three complete months
in `Transactions`.

## `Daily Log`

A header row with `Date` in column A and a column headed `Net Liquid` (and,
optionally, one headed `Net Worth`). Dated snapshots below drive the
month-on-month deltas on the hero tiles. One row per month-end is plenty. Rows
whose date cell is a formula are skipped, so a live-connection row won't be read
as a snapshot.

## `Wish List`

Label/value pairs in columns A/B: `Net liquidity`, `Provisions`, `Free cash`,
`Affordability ratio`. Then an item table starting at a header row with `ID` in
column A: `A` = id, `B` = item, `C` = category, `D` = cost, `T` = verdict.

## `Config` — optional

Entirely optional. Delete the sheet and the dashboard derives what it can from
your data: categories and colours from the ledger, wealth groups from generic
keywords, `£` and your browser's locale.

Each block is introduced by a header in column A. Unknown rows are ignored, so you
can leave notes in the spare columns.

### `SETTINGS`

| A (name) | B (value) | |
|---|---|---|
| `Currency symbol` | `£` | Used everywhere, privacy mode included |
| `Locale` | `en-GB` | Date and number formatting — `en-US`, `it-IT`, `de-DE` … Blank = follow the UI language |
| `Language` | `it` | UI language: `en`, `it`, `es`, `fr`, `de`. Blank = follow the browser. A pick made in the header overrides this |
| `Owner` | `Sample User` | Shown in the header; leave blank to hide |
| `History from` | `2023-01` | `YYYY-MM`; trims the long-run charts. Blank = all data |

Changing `Currency symbol` re-denominates the whole dashboard — there is no second
place to edit, and privacy-mode blurring follows it automatically.

`Language` sets which language the interface opens in; it never touches your own
labels. `Locale` and `Language` are independent: leave `Locale` blank and number and
date formatting follow the chosen language, or set it to pin formatting regardless
of language.

### `WEALTH GROUPS`

Controls the *Where the wealth sits* donut.

| A (Match) | B (Label) | C (Colour) |
|-----------|-----------|-----------|
| `home,property` | Home | `#5b9cf6` |
| `pension,sipp` | Pension | `#f59e0b` |
| `company,stake,equity` | Company / business | `#a78bfa` |

`Match` is a comma-separated keyword list; any `INVESTMENTS` line whose name
**contains** one of them joins that group, first match wins. Everything unmatched
becomes *Other investments*, plus a *Liquid* segment (cash + savings).

### `CATEGORIES`

Pins the colour and display order of your spending categories, consistently across
every view.

| A (Category) | B (Colour) | C (Order) |
|--------------|-----------|-----------|
| Home | `#5b9cf6` | 1 |
| Transport | `#2dd4bf` | 2 |

`Category` must match a `List 1` value in your ledger. Any category you omit still
appears — it just sorts after the pinned ones and gets an automatic palette colour.

### `ACCOUNT GROUPS`

| A (Match) | B (Role) |
|-----------|----------|
| `card,credit` | `card repayment` |

Bills matching a `card repayment` row are treated as moving money between your own
accounts rather than buying anything.

### `THRESHOLDS`

Tune when the dashboard warns you.

| A (Name) | B (Value) | |
|---|---|---|
| `Runway warning months` | `2` | Fewer months of liquid runway than this → amber verdict |
| `Pace warning pct` | `100` | Spend vs day-proportional budget → amber |
| `Pace alert pct` | `110` | → red |
| `Insight min spend` | `40` | Ignore categories smaller than this when detecting changes |
| `Insight min change pct` | `15` | Only report a change bigger than this |
