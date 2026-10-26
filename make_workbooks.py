#!/usr/bin/env python3
"""
make_workbooks.py
Generates the two workbooks that ship with this repo:

  finance-tracker-SAMPLE.xlsx    every sheet, filled with synthetic but
                                 internally consistent data — open this first to
                                 see a fully-populated dashboard.
  finance-tracker-TEMPLATE.xlsx  the same structure, blank: every sheet, section
                                 header and column header in place, with one
                                 clearly-marked example row per section. Delete
                                 the example rows, keep the headers, add yours.

Both carry a filled-in `Config` sheet, so they double as the reference for the
schema, and a `Start here` sheet explaining the colour code. Nothing here is real
data.

Totals, subtotals, budget actuals, bill dates, net liquidity and wish-list
verdicts are live Excel formulas, so the workbook keeps itself consistent as you
type. Colour code: soft-yellow cells with blue text are inputs; grey cells are
calculated (don't type over them).

The dashboard reads the *cached* result of each formula, and openpyxl cannot
calculate. So after saving, this script asks LibreOffice (if installed) to
calculate the workbook and copies the results back into the file. Without
LibreOffice the files are still correct, but open each one in Excel and save
it once before loading it into the dashboard.

Requires openpyxl.  pip install openpyxl
Optional: LibreOffice (`soffice` on PATH, or the standard macOS app location).

Run:  python3 make_workbooks.py            # both
      python3 make_workbooks.py sample
      python3 make_workbooks.py template
"""
import sys, os, re, random, shutil, subprocess, tempfile, zipfile, datetime as dt
from xml.sax.saxutils import escape
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

WHICH = sys.argv[1].lower() if len(sys.argv) > 1 else "both"
if WHICH not in ("both", "sample", "template"):
    sys.exit("usage: make_workbooks.py [both|sample|template]")

SAMPLE_OUT = "finance-tracker-SAMPLE.xlsx"
TEMPLATE_OUT = "finance-tracker-TEMPLATE.xlsx"

TODAY = dt.date.today()

# ----------------------------------------------------------------- styling
# Accents are the dashboard's own palette, so the workbook and the page read as
# one product. Fills are light tints of the same hues so text stays legible.
FONT = "Arial"
INK, MUTED, WHITE = "1F2430", "6B7280", "FFFFFF"
HEAD_BG = "2A2D3E"          # dashboard panel colour
INPUT_BG, INPUT_FG = "FFF7D6", "1F4FD8"
CALC_BG = "EEF0F4"
LINK_FG = "137333"          # formulas that pull from another sheet
LINE = "C9CDD6"

SECTION = {                 # fill tint, accent text
    "CASH":        ("E3F6EA", "1E7B45"),
    "SAVINGS":     ("DDF5F2", "0F766E"),
    "REFUNDS DUE": ("FDF3D7", "8A5A00"),
    "INVESTMENTS": ("EEE9FD", "5B3FC4"),
    "LIABILITIES": ("FDE5E5", "B42318"),
    "SUMMARY":     ("E3EDFD", "1D4ED8"),
}
TAB = {"start": "2A2D3E", "required": "5B9CF6", "recommended": "2DD4BF",
       "optional": "A0A4B0"}

MONEY = '#,##0.00;[Red]-#,##0.00;"–"'
MONEY0 = '#,##0;[Red]-#,##0;"–"'
PCT = '0.0%;[Red]-0.0%;"–"'
DATE = "dd mmm yyyy"
INT = "0"

THIN = Side(style="thin", color=LINE)
MED = Side(style="medium", color=INK)


def font(**kw):
    kw.setdefault("name", FONT)
    kw.setdefault("size", 10)
    kw.setdefault("color", INK)
    return Font(**kw)


def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


BOLD = font(bold=True)
NOTE = font(italic=True, color=MUTED, size=9)


def title(ws, text, sub=None, span=8, sub_row=None):
    """Row 1 banner. The text stays in A1 (no merged cells — the dashboard and
    anyone sorting the sheet are both happier without them)."""
    ws["A1"] = text
    ws["A1"].font = font(bold=True, size=13, color=WHITE)
    for c in range(1, span + 1):
        ws.cell(1, c).fill = fill(HEAD_BG)
    ws.row_dimensions[1].height = 26
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    if sub:
        note(ws, "A2" if sub_row is None else f"A{sub_row}", sub)


def note(ws, ref, text):
    ws[ref] = text
    ws[ref].font = NOTE
    ws[ref].alignment = Alignment(vertical="center")


def header(ws, row, labels, start=1, right=(), center=(), height=20):
    """Dark header row. Alignment follows the column's content: numbers right,
    flags centred, text left."""
    for j, h in enumerate(labels, start):
        if h is None:
            continue
        c = ws.cell(row, j, h)
        c.font = font(bold=True, color=WHITE)
        c.fill = fill(HEAD_BG)
        c.alignment = Alignment(vertical="center", wrap_text=height > 20,
                                horizontal="right" if j in right else
                                "center" if j in center else "left")
    ws.row_dimensions[row].height = height


def inp(c, fmt=None):
    """An input cell: the user types here."""
    c.fill = fill(INPUT_BG)
    c.font = font(color=INPUT_FG)
    c.border = Border(bottom=THIN)
    if fmt:
        c.number_format = fmt
    return c


def calc(c, fmt=None, link=False, bold=False):
    """A calculated cell: formula, don't type over it."""
    c.fill = fill(CALC_BG)
    c.font = font(color=LINK_FG if link else INK, bold=bold)
    c.border = Border(bottom=THIN)
    if fmt:
        c.number_format = fmt
    return c


def label(c, bold=False, indent=0):
    c.font = font(bold=bold)
    if indent:
        c.alignment = Alignment(indent=indent)
    return c


def band(ws, row, name, span):
    """Coloured section-header band (CASH, SAVINGS, …)."""
    bg, fg = SECTION[name]
    for c in range(1, span + 1):
        ws.cell(row, c).fill = fill(bg)
    ws.cell(row, 1).font = font(bold=True, color=fg)
    ws.row_dimensions[row].height = 18


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def print_setup(wb):
    """Landscape, one page wide — so a printout or PDF export is readable."""
    from openpyxl.worksheet.properties import PageSetupProperties
    for ws in wb.worksheets:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = "1:1" if ws.title == "Transactions" else None


def arial_everywhere(wb):
    """Any cell styled elsewhere without an explicit font still gets Arial."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.font is None or c.font.name != FONT:
                    f = c.font
                    c.font = Font(name=FONT, size=f.size or 10, bold=f.bold,
                                  italic=f.italic, underline=f.underline,
                                  color=f.color if f.color is not None else INK)

# --------------------------------------------------------------- taxonomy
CATS = {
    "Home":            ["Rent", "Groceries", "Utilities", "Internet/TV", "Maintenance"],
    "Transport":       ["Fuel", "Insurance", "Service", "Public transport"],
    "Me":              ["Subscriptions", "Health", "Clothes"],
    "Fun":             ["Travel", "Concerts", "Leisure"],
    "Food drinks out": ["Eating out", "Drinks", "Take away"],
    "Family":          ["Gifts", "Support"],
    "Work":            ["Equipment", "Fees"],
}
CATCOL = {
    "Home": "#5b9cf6", "Transport": "#2dd4bf", "Me": "#4ade80", "Fun": "#a78bfa",
    "Food drinks out": "#f472b6", "Family": "#f87171", "Work": "#fbbf24",
}
MERCHANTS = {
    "Rent": ["RIVERSIDE LETTINGS"],
    "Groceries": ["FRESHMART", "GREENGROCER CO", "COSTSAVE"],
    "Utilities": ["BRIGHT ENERGY", "AQUAFLOW WATER"],
    "Internet/TV": ["STREAMNET", "FIBREONE"],
    "Maintenance": ["FIXIT HARDWARE", "HOMESTORE"],
    "Fuel": ["FUELSTOP", "ECOCHARGE"],
    "Insurance": ["SAFEDRIVE INS"],
    "Service": ["AUTOCARE GARAGE"],
    "Public transport": ["METRO TRAVEL", "RAILCONNECT"],
    "Subscriptions": ["STREAMFLIX", "CLOUDTUNES"],
    "Health": ["WELLNESS CLINIC", "PHARMA PLUS"],
    "Clothes": ["URBAN THREADS", "SHOE DEPOT"],
    "Travel": ["SKYJET AIR", "CITY HOTELS"],
    "Concerts": ["TICKETLINE", "ARENA LIVE"],
    "Leisure": ["CINEPLEX", "BOULDER GYM"],
    "Eating out": ["THE OLD KITCHEN", "BISTRO 21"],
    "Drinks": ["THE ANCHOR PUB", "CRAFT TAPROOM"],
    "Take away": ["NOODLE BAR", "PIZZA CO"],
    "Gifts": ["GIFTBOX CO"],
    "Support": ["A. RELATIVE"],
    "Equipment": ["TECHWORLD", "OFFICE DEPOT CO"],
    "Fees": ["PRO BODY FEES"],
}
TYP = {
    "Rent": 950, "Groceries": 380, "Utilities": 140, "Internet/TV": 70, "Maintenance": 90,
    "Fuel": 70, "Insurance": 45, "Service": 60, "Public transport": 85,
    "Subscriptions": 55, "Health": 90, "Clothes": 110,
    "Travel": 180, "Concerts": 90, "Leisure": 80,
    "Eating out": 220, "Drinks": 160, "Take away": 70,
    "Gifts": 40, "Support": 60,
    "Equipment": 50, "Fees": 30,
}
CONF = ["existing", "learned", "learned", "existing", "low"]


def month_range(start, end):
    y, m = start
    ey, em = end
    out = []
    while (y, m) <= (ey, em):
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return out


# 36 months ending with the current month, so there is always a "this month".
MONTHS = month_range((TODAY.year - 3, TODAY.month), (TODAY.year, TODAY.month))
CUR_M = MONTHS[-1]


def build_transactions():
    random.seed(42)
    tx = []
    accounts = ["Sample Bank", "Everyday Card", "Rewards Card"]
    for ym in MONTHS:
        y, m = int(ym[:4]), int(ym[5:])
        maxday = 28 if ym != CUR_M else max(1, min(28, TODAY.day))
        for cat, subs in CATS.items():
            for sub in subs:
                if random.random() < 0.15:
                    continue
                tx.append([dt.datetime(y, m, random.randint(1, maxday)), ym,
                           random.choice(accounts), random.choice(MERCHANTS[sub]),
                           round(TYP[sub] * random.uniform(0.6, 1.5), 2),
                           cat, sub, "Recurring", random.choice(CONF)])
        tx.append([dt.datetime(y, m, min(25, maxday)), ym, "Sample Bank",
                   "EMPLOYER PAYROLL", -round(4200 + random.uniform(-150, 400), 2),
                   "Income", "Salary", "Income", "existing"])
        tx.append([dt.datetime(y, m, min(26, maxday)), ym, "Sample Bank",
                   "TO SAVINGS", round(max(0, random.uniform(-200, 1100)), 2),
                   "Savings", "Transfer", "Savings", "existing"])
        if random.random() < 0.25:
            tx.append([dt.datetime(y, m, random.randint(1, maxday)), ym, "Sample Bank",
                       "REFUND CO", -round(random.uniform(20, 120), 2),
                       "Home", "Maintenance", "Recurring", "existing"])
        if random.random() < 0.3:
            tx.append([dt.datetime(y, m, random.randint(1, maxday)), ym, "Sample Bank",
                       "J. FRIEND", round(random.uniform(40, 300), 2),
                       "Me", "Health", "One-off", "review"])
    for off, label, cat, sub, amt in [
        (6, "FURNISH CO", "Home", "Maintenance", 12500),
        (5, "INTERIORS LTD", "Home", "Maintenance", 6400),
        (4, "TAX AUTHORITY", "Work", "Fees", 4200),
    ]:
        ym = MONTHS[max(0, len(MONTHS) - 1 - off)]
        y, m = int(ym[:4]), int(ym[5:])
        tx.append([dt.datetime(y, m, 12), ym, "Sample Bank", label, amt, cat, sub,
                   "One-off", "existing"])
    tx.sort(key=lambda t: t[0])
    return tx


def rollups(tx):
    R = {ym: {"rec": 0.0, "oneoff": 0.0, "income": 0.0, "tosav": 0.0,
              "cats": {c: 0.0 for c in CATS}} for ym in MONTHS}
    for t in tx:
        ym, amt, cat, bucket = t[1], t[4], t[5], t[7]
        r = R[ym]
        if bucket == "Recurring":
            r["rec"] += amt
            if cat in r["cats"]:
                r["cats"][cat] += amt
        elif bucket == "One-off":
            r["oneoff"] += amt
        elif bucket == "Income":
            r["income"] += -amt
        elif bucket == "Savings":
            r["tosav"] += amt
    return R


# Accounts: (section, [(name, value, due_in_days)]) — value goes in column D.
SAMPLE_ACCOUNTS = [
    ("CASH", [("Sample Bank", 6200, None), ("Everyday Card", 0, None),
              ("Rewards Card", 120, None)]),
    ("SAVINGS", [("Instant saver", 16000, None), ("Premium bonds", 5000, None)]),
    ("REFUNDS DUE", [("Deposit refund", 300, None), ("Expense claim", 165, None)]),
    ("INVESTMENTS", [("Stocks & shares ISA", 24000, None), ("Pension pot", 52000, None),
                     ("Index funds", 18000, None), ("Crypto", 3500, None),
                     ("Home (property)", 420000, None),
                     ("Company shareholding", 45000, None)]),
    ("LIABILITIES", [("Credit card A", 3200, 25), ("Credit card B", 2400, 40),
                     ("Personal loan", 5000, 200),
                     ("Mortgage (outstanding)", 260000, 365 * 18)]),
]
TEMPLATE_ACCOUNTS = [
    ("CASH", [("Example current account (replace)", 0, None)]),
    ("SAVINGS", [("Example savings (replace)", 0, None)]),
    ("REFUNDS DUE", [("Example refund owed to you (replace)", 0, None)]),
    ("INVESTMENTS", [("Example home/property (replace)", 0, None),
                     ("Example pension (replace)", 0, None),
                     ("Example other investment (replace)", 0, None)]),
    ("LIABILITIES", [("Example mortgage (replace)", 0, 365 * 18),
                     ("Example credit card (replace)", 0, 25)]),
]
SAMPLE_BILLS = [("Water", 34, 1, "F"), ("Council tax", 188, 1, "F"),
                ("Life cover", 82, 1, "F"), ("Gas & electric", 135, 1, "F"),
                ("Health plan", 123, 4, "F"), ("Mortgage", 1150, 4, "F"),
                ("Gym membership", 45, 9, ""), ("TV & broadband", 67, 17, "F"),
                ("Mobile", 30, 25, "F"), ("Cloud storage", 10, 29, ""),
                ("Personal loan", 73, 29, "F"), ("Credit card A", 250, 20, "")]
TEMPLATE_BILLS = [("Example bill — direct debit (replace)", 0, 1, "F"),
                  ("Example bill — variable (replace)", 0, 15, "")]


def build(blank):
    """blank=True -> the empty TEMPLATE; blank=False -> the filled SAMPLE."""
    tag = "TEMPLATE" if blank else "SAMPLE"
    tx = [] if blank else build_transactions()
    R = rollups(tx) if tx else {}
    accounts = TEMPLATE_ACCOUNTS if blank else SAMPLE_ACCOUNTS
    bills = TEMPLATE_BILLS if blank else SAMPLE_BILLS
    wb = openpyxl.Workbook()

    # --------------------------------------------------------- Start here
    # Not read by the dashboard (it ignores sheets it doesn't know), so it can
    # say whatever helps a first-time user.
    ws = wb.active
    ws.title = "Start here"
    ws.sheet_properties.tabColor = TAB["start"]
    ws.sheet_view.showGridLines = False
    title(ws, "Personal Finance Tracker — " +
          ("blank template" if blank else "sample workbook (synthetic data)"), span=4)
    note(ws, "A2", "Load this file into the dashboard with 'Load workbook'. "
                   "It is read in your browser and never uploaded.")
    r = 4
    ws.cell(r, 1, "HOW TO USE").font = font(bold=True, color=SECTION["SUMMARY"][1])
    r += 1
    steps = ([
        "1.  Replace the example rows (marked 'replace') with your own — keep every header.",
        "2.  Accounts and Transactions are the only required sheets; the rest add panels.",
        "3.  Type only in the yellow cells. Grey cells are formulas and update themselves.",
        "4.  Save, then load (or reload) the file in the dashboard.",
    ] if blank else [
        "1.  Every figure here is synthetic — load it to see a fully populated dashboard.",
        "2.  Accounts and Transactions are the only required sheets; the rest add panels.",
        "3.  Yellow cells are inputs; grey cells are formulas and update themselves.",
        "4.  To start on your own numbers, use finance-tracker-TEMPLATE.xlsx.",
    ])
    for s in steps:
        ws.cell(r, 1, s).font = font()
        r += 1
    r += 1
    ws.cell(r, 1, "COLOUR CODE").font = font(bold=True, color=SECTION["SUMMARY"][1])
    r += 1
    for sample, what, styler in [
        ("1,234.00", "Input — type here", lambda c: inp(c, MONEY)),
        ("1,234.00", "Calculated — formula, leave it alone", lambda c: calc(c, MONEY)),
        ("1,234.00", "Calculated from another sheet (green text)",
         lambda c: calc(c, MONEY, link=True)),
    ]:
        c = ws.cell(r, 1, sample)
        styler(c)
        c.alignment = Alignment(horizontal="right")
        ws.cell(r, 2, what).font = font()
        r += 1
    r += 1
    ws.cell(r, 1, "SHEETS").font = font(bold=True, color=SECTION["SUMMARY"][1])
    r += 1
    header(ws, r, ["Sheet (click to open)", "Status", "What it drives"])
    r += 1
    for name, status, what in [
        ("Accounts", "Required", "Balance sheet, net worth, net liquid assets, wealth donut"),
        ("Transactions", "Required", "Every spending view, budget actuals, insights"),
        ("Historical", "Recommended", "Month-end cash, savings rate, net savings"),
        ("Expenses", "Recommended", "Budget vs actual, safe-to-spend"),
        ("Bills", "Recommended", "Bills due this month, cash-flow forecast"),
        ("Dashboard", "Optional", "Net pay, payday, cash-reserve goal, forecast spend"),
        ("Daily Log", "Optional", "Month-on-month deltas on the headline tiles"),
        ("Wish List", "Optional", "Affordability panel"),
        ("Config", "Optional", "Currency, language, colours, groupings, thresholds"),
    ]:
        c = ws.cell(r, 1, name)
        c.hyperlink = f"#'{name}'!A1"
        c.font = font(color=INPUT_FG, underline="single")
        st = ws.cell(r, 2, status)
        st.font = font(bold=status == "Required",
                       color={"Required": "1D4ED8", "Recommended": "0F766E"}.get(status, MUTED))
        ws.cell(r, 3, what).font = font()
        for cc in range(1, 4):
            ws.cell(r, cc).border = Border(bottom=THIN)
        r += 1
    r += 1
    ws.cell(r, 1, "GROUND RULES").font = font(bold=True, color=SECTION["SUMMARY"][1])
    r += 1
    for s in [
        "Keep sheet names, section headers (CASH, SAVINGS…) and column positions as they are — "
        "that is what the dashboard reads. Full detail in SCHEMA.md.",
        "Add rows anywhere inside a block (e.g. just above a subtotal or total); "
        "the formulas pick them up.",
        "Transactions: amount positive = money out, negative = refund/credit. "
        "Month as YYYY-MM text.",
        "Bucket: Recurring (counts as spending) · One-off (excluded from averages) · "
        "Income · Savings · REVIEW.",
    ]:
        ws.cell(r, 1, "•  " + s).font = font()
        r += 1
    widths(ws, {"A": 26, "B": 46, "C": 58, "D": 4})

    # ---------------------------------------------------------- Dashboard
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.tabColor = TAB["optional"]
    ws.sheet_view.showGridLines = False
    title(ws, f"DASHBOARD — optional  ({tag})", span=4)
    header(ws, 2, ["Setting", "Value", None, "What it does"])
    # "Target card spend per month" is the forward parameter for the 12-month cash-flow
    # forecast: set it by hand and the forecast uses it verbatim. Leave the row out entirely
    # and the dashboard falls back to a run-rate computed from the Transactions ledger.
    hints = {
        "Net pay": "monthly take-home pay",
        "Payday": "day of the month you are paid",
        "Cash reserve goal": "the net-liquidity level you are rebuilding towards",
        "Cash reserve monthly target": "how much a month you aim to add to it",
        "Month opening net liquid": "net liquid assets on the 1st of this month",
        "Target card spend per month": "drives the 12-month cash-flow forecast; "
                                       "delete the row to use your recent run-rate",
        "All-in wealth rate": "share of pay that becomes wealth (incl. pension, equity)",
        "Cash savings rate": "share of pay left in the bank",
        "Footer note": "shown at the foot of the dashboard",
    }
    fmts = {"Net pay": MONEY0, "Payday": INT, "Cash reserve goal": MONEY0,
            "Cash reserve monthly target": MONEY0, "Month opening net liquid": MONEY0,
            "Target card spend per month": MONEY0, "All-in wealth rate": PCT,
            "Cash savings rate": PCT}
    rows = ([("Net pay", 0), ("Payday", 25), ("Cash reserve goal", 0),
             ("Cash reserve monthly target", 0), ("Month opening net liquid", 0),
             ("Target card spend per month", 0),
             ("All-in wealth rate", 0), ("Cash savings rate", 0),
             ("Footer note", "Replace with your own note, or delete this row.")]
            if blank else
            [("Net pay", 4200), ("Payday", 25), ("Cash reserve goal", 20000),
             ("Cash reserve monthly target", 800), ("Month opening net liquid", 26000),
             ("Target card spend per month", 1600),
             ("All-in wealth rate", 0.28), ("Cash savings rate", 0.14),
             ("Footer note", "All figures in this workbook are synthetic sample data.")])
    for i, (k, v) in enumerate(rows, start=3):
        label(ws.cell(i, 1, k))
        c = inp(ws.cell(i, 2, v), fmts.get(k))
        if k == "Footer note":
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[i].height = 40
        ws.cell(i, 4, hints[k]).font = NOTE
    widths(ws, {"A": 32, "B": 24, "C": 2, "D": 70})
    ws.freeze_panes = "A3"

    # ------------------------------------------------------------- Config
    ws = wb.create_sheet("Config")
    ws.sheet_properties.tabColor = TAB["optional"]
    title(ws, f"CONFIG — optional  ({tag})",
          "Delete the whole sheet and the dashboard derives everything from your data. "
          "Sections are matched by the headers below.", span=4)
    lang_dv = DataValidation(type="list", formula1='"en,it,es,fr,de"', allow_blank=True)
    ws.add_data_validation(lang_dv)

    def block(r, name, hint):
        c = ws.cell(r, 1, name)
        c.font = font(bold=True, color=SECTION["SUMMARY"][1])
        for cc in range(1, 5):
            ws.cell(r, cc).fill = fill(SECTION["SUMMARY"][0])
        if hint:
            ws.cell(r, 4, hint).font = NOTE
        ws.row_dimensions[r].height = 18

    def swatch(c):
        """Colour cells show their own colour — the dashboard reads the text."""
        inp(c)
        try:
            c.fill = fill(str(c.value).lstrip("#").upper())
            c.font = font(color=WHITE, bold=True)
            c.alignment = Alignment(horizontal="center")
        except Exception:
            pass

    r = 3
    block(r, "SETTINGS", "")
    r += 1
    for k, v, hint in [
        ("Currency symbol", "£", "used everywhere, privacy mode included"),
        ("Locale", "en-GB", "date and number formatting, e.g. en-US, it-IT, de-DE"),
        ("Language", "", "UI language: en, it, es, fr or de. Blank = follow the browser"),
        ("Owner", "" if blank else "Sample User", "shown in the header; blank to hide"),
        ("History from", "", "YYYY-MM — trims the long-run charts; blank = all data"),
    ]:
        label(ws.cell(r, 1, k)); inp(ws.cell(r, 2, v), "@"); note(ws, f"D{r}", hint)
        if k == "Language":
            lang_dv.add(f"B{r}")
        r += 1

    r += 1
    block(r, "WEALTH GROUPS", "groups INVESTMENTS lines into the wealth donut")
    r += 1
    header(ws, r, ["Match (investment name contains; comma-separated)", "Label", "Colour"],
           center=(3,))
    r += 1
    for match, lab, colour in [("home,property", "Home", "#5b9cf6"),
                               ("pension,sipp", "Pension", "#f59e0b"),
                               ("company,stake,equity", "Company / business", "#a78bfa")]:
        inp(ws.cell(r, 1, match)); inp(ws.cell(r, 2, lab)); swatch(ws.cell(r, 3, colour))
        r += 1

    r += 1
    block(r, "CATEGORIES", "pins colour and display order; omitted categories still appear")
    r += 1
    header(ws, r, ["Category", "Colour", "Order"], center=(2,), right=(3,))
    r += 1
    if blank:
        inp(ws.cell(r, 1, "Example category"))
        note(ws, f"D{r}", "replace with your List 1 categories (must match the ledger)")
        swatch(ws.cell(r, 2, "#5b9cf6")); inp(ws.cell(r, 3, 1), INT); r += 1
    else:
        for i, c in enumerate(CATS, start=1):
            inp(ws.cell(r, 1, c)); swatch(ws.cell(r, 2, CATCOL[c])); inp(ws.cell(r, 3, i), INT)
            r += 1

    r += 1
    block(r, "ACCOUNT GROUPS", "role 'card repayment' = a bill that clears a card, not a purchase")
    r += 1
    header(ws, r, ["Match (bill name contains; comma-separated)", "Role"])
    r += 1
    inp(ws.cell(r, 1, "card,credit")); inp(ws.cell(r, 2, "card repayment")); r += 1

    r += 1
    block(r, "THRESHOLDS", "tune when the dashboard warns you")
    r += 1
    header(ws, r, ["Name", "Value"], right=(2,))
    r += 1
    for k, v, hint in [
        ("Runway warning months", 2, "fewer months of liquid runway than this -> amber"),
        ("Pace warning pct", 100, "spend vs day-proportional budget -> amber"),
        ("Pace alert pct", 110, "-> red"),
        ("Insight min spend", 40, "ignore categories smaller than this when detecting changes"),
        ("Insight min change pct", 15, "only report a change bigger than this"),
    ]:
        label(ws.cell(r, 1, k)); inp(ws.cell(r, 2, v), INT); note(ws, f"D{r}", hint); r += 1
    widths(ws, {"A": 46, "B": 20, "C": 12, "D": 68})

    # ----------------------------------------------------------- Accounts
    ws = wb.create_sheet("Accounts")
    ws.sheet_properties.tabColor = TAB["required"]
    title(ws, f"ACCOUNTS — required  ({tag})",
          "Section header in column A, then one row per line item: name in A, value in D.",
          span=5, sub_row=3)
    label(ws["A2"])
    inp(ws.cell(2, 2, dt.datetime(TODAY.year, TODAY.month, TODAY.day)), DATE)
    ws["A2"] = "Last updated"
    header(ws, 4, ["Account", "Amount", "Currency", "GBP value", "Due (liabilities)"],
           right=(2, 4, 5), center=(3,))
    note(ws, "G4", "column D is what the dashboard reads (here it simply mirrors B); "
                   "column E splits short- from long-term debt")
    note(ws, "G5", "subtotal rows are skipped by the dashboard; add rows anywhere "
                   "between a section header and its subtotal")
    r = 6
    sub_row = {}
    sub_name = {"CASH": "Cash subtotal", "SAVINGS": "Savings subtotal",
                "REFUNDS DUE": "Refunds subtotal", "INVESTMENTS": "Investments subtotal",
                "LIABILITIES": "Liabilities subtotal"}
    for name, items in accounts:
        ws.cell(r, 1, name)
        band(ws, r, name, 5)
        hdr = r
        r += 1
        for n, v, due in items:
            inp(ws.cell(r, 1, n))
            inp(ws.cell(r, 2, v), MONEY)
            inp(ws.cell(r, 3, "GBP")).alignment = Alignment(horizontal="center")
            calc(ws.cell(r, 4, f"=B{r}"), MONEY)
            if name == "LIABILITIES":
                d = TODAY + dt.timedelta(days=due) if due is not None else None
                inp(ws.cell(r, 5, dt.datetime(d.year, d.month, d.day) if d else None), DATE)
            r += 1
        # SUM from the header row to the row just above this one, so rows inserted
        # anywhere in the block — including directly above the subtotal — are counted.
        ws.cell(r, 1, sub_name[name]).font = font(bold=True, color=SECTION[name][1])
        ws.cell(r, 1).alignment = Alignment(indent=1)
        c = calc(ws.cell(r, 4, f"=SUM(D{hdr}:INDEX(D:D,ROW()-1))"), MONEY, bold=True)
        c.border = Border(top=Side(style="thin", color=SECTION[name][1]), bottom=THIN)
        sub_row[name] = r
        r += 2
    ws.cell(r, 1, "SUMMARY")
    band(ws, r, "SUMMARY", 5)
    note(ws, f"G{r}", "optional — the dashboard recomputes these from the line items "
                      "when absent")
    r += 1
    label(ws.cell(r, 1, "Company valuation"))
    inp(ws.cell(r, 2, 0 if blank else 500000), MONEY0)
    r += 1
    S = sub_row
    summary = [
        ("Total liquid", f"=D{S['CASH']}+D{S['SAVINGS']}+D{S['REFUNDS DUE']}"),
        ("Total investments", f"=D{S['INVESTMENTS']}"),
        ("Total liabilities", f"=D{S['LIABILITIES']}"),
    ]
    first = r
    for lab, f in summary:
        label(ws.cell(r, 1, lab)); calc(ws.cell(r, 4, f), MONEY); r += 1
    label(ws.cell(r, 1, "Gross assets")); calc(ws.cell(r, 4, f"=D{first}+D{first + 1}"), MONEY)
    r += 1
    label(ws.cell(r, 1, "Net worth"), bold=True)
    c = calc(ws.cell(r, 4, f"=D{r - 1}-D{first + 2}"), MONEY, bold=True)
    c.border = Border(top=MED, bottom=Side(style="double", color=INK))
    widths(ws, {"A": 40, "B": 16, "C": 10, "D": 16, "E": 16, "F": 3, "G": 70})
    ws.freeze_panes = "A5"

    # -------------------------------------------------------------- Bills
    ws = wb.create_sheet("Bills")
    ws.sheet_properties.tabColor = TAB["recommended"]
    title(ws, f"BILLS — optional  ({tag})", "Recurring commitments.", span=8)
    header(ws, 3, ["Bill", "Amount", "Direct debit day", "To pay (1/0)",
                   "Next collection", "Days until", "Status", "Fixed?"],
           right=(2, 3, 5, 6), center=(4, 7, 8), height=30)
    note(ws, "J3", "column H: F/Y/1 marks a bill you cannot easily cancel")
    note(ws, "J4", "E–G work themselves out from the debit day; "
                   "To pay = 0 pauses a bill in the totals below")
    for i, (n, amt, day, fixed) in enumerate(bills):
        row = 4 + i
        inp(ws.cell(row, 1, n))
        inp(ws.cell(row, 2, amt), MONEY)
        inp(ws.cell(row, 3, day), INT)
        inp(ws.cell(row, 4, 1), INT).alignment = Alignment(horizontal="center")
        # Next collection: this month if the day hasn't passed, else next month.
        # MIN(…EOMONTH) keeps a 31st-of-month debit inside short months.
        calc(ws.cell(row, 5,
             f'=IF($C{row}="","",IF($C{row}>=DAY(TODAY()),'
             f'DATE(YEAR(TODAY()),MONTH(TODAY()),MIN($C{row},DAY(EOMONTH(TODAY(),0)))),'
             f'DATE(YEAR(EDATE(TODAY(),1)),MONTH(EDATE(TODAY(),1)),'
             f'MIN($C{row},DAY(EOMONTH(TODAY(),1))))))'), DATE)
        calc(ws.cell(row, 6, f'=IF($E{row}="","",$E{row}-TODAY())'), INT)
        calc(ws.cell(row, 7, f'=IF($C{row}="","",IF($D{row}<>1,"Paused",'
                             f'IF($C{row}<DAY(TODAY()),"Paid","Due")))'))
        inp(ws.cell(row, 8, fixed))
        ws.cell(row, 7).alignment = Alignment(horizontal="center")
        ws.cell(row, 8).alignment = Alignment(horizontal="center")
    last = 4 + len(bills) - 1
    t = last + 2        # one blank spacer row inside the ranges, so a new bill
                        # typed straight under the list is still counted
    tot = [("Total monthly (all active)",
            '=SUMIF($D$3:INDEX($D:$D,ROW()-1),1,$B$3:INDEX($B:$B,ROW()-1))'),
           ("Left to pay this month",
            '=SUMIFS($B$3:INDEX($B:$B,ROW()-2),$C$3:INDEX($C:$C,ROW()-2),'
            '">="&DAY(TODAY()),$D$3:INDEX($D:$D,ROW()-2),1)'),
           ("Paid so far this month",
            '=SUMIFS($B$3:INDEX($B:$B,ROW()-3),$C$3:INDEX($C:$C,ROW()-3),'
            '"<"&DAY(TODAY()),$D$3:INDEX($D:$D,ROW()-3),1)')]
    for k, (lab, f) in enumerate(tot):
        label(ws.cell(t + k, 1, lab), bold=k == 0)
        calc(ws.cell(t + k, 2, f), MONEY, bold=k == 0)
    ws.cell(t, 2).border = Border(top=MED, bottom=THIN)
    rng = f"A4:H{last}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND($G4="Due",$F4<=7)'], fill=fill("FDF3D7")))
    ws.conditional_formatting.add(f"G4:G{last}", CellIsRule(
        operator="equal", formula=['"Paid"'], font=Font(name=FONT, color="1E7B45")))
    widths(ws, {"A": 34, "B": 13, "C": 11, "D": 10, "E": 15, "F": 10, "G": 10,
                "H": 9, "I": 3, "J": 70})
    ws.freeze_panes = "A4"

    # ----------------------------------------------------------- Expenses
    ws = wb.create_sheet("Expenses")
    ws.sheet_properties.tabColor = TAB["recommended"]
    title(ws, f"EXPENSES — optional  ({tag})", "Monthly budget per category.",
          span=5, sub_row=3)
    label(ws["A2"], bold=True)
    ws["A2"] = "Reporting month"
    inp(ws.cell(2, 2, CUR_M), "@")
    note(ws, "C2", "YYYY-MM — the Actual column follows this month")
    header(ws, 4, ["Category", "Budget", "Actual", "Variance", "Share of net pay"],
           right=(2, 3, 4, 5), height=30)
    note(ws, "G4", "actuals are recomputed from the ledger; the block ends at 'Total'")
    note(ws, "G5", "Actual = Recurring spend in Transactions for the reporting month "
                   "(List 1 = category)")
    r = 5

    def exp_row(r, cat, budget):
        inp(ws.cell(r, 1, cat))
        inp(ws.cell(r, 2, budget), MONEY0)
        calc(ws.cell(r, 3, f'=SUMIFS(Transactions!$E:$E,Transactions!$B:$B,$B$2,'
                           f'Transactions!$F:$F,$A{r},Transactions!$H:$H,"Recurring")'),
             MONEY, link=True)
        calc(ws.cell(r, 4, f"=B{r}-C{r}"), MONEY)
        calc(ws.cell(r, 5, f'=IFERROR(C{r}/SUMIF(Dashboard!$A:$A,"Net pay",'
                           f'Dashboard!$B:$B),0)'), PCT, link=True)

    if blank:
        # Same name as the example ledger row, so the Actual formula visibly links up.
        exp_row(r, "Example category", 0)
        note(ws, f"G{r + 1}", "replace 'Example category' with your List 1 categories "
                              "— they must match the ledger exactly")
        r += 1
    else:
        for c in CATS:
            actual = round(R[CUR_M]["cats"][c], 2)
            ref = (sum(R[m]["cats"][c] for m in MONTHS[-13:-1]) / 12
                   if len(MONTHS) > 12 else max(actual, 50))
            budget = round(max(50, ref * random.uniform(0.95, 1.2)))
            exp_row(r, c, budget)
            r += 1
    label(ws.cell(r, 1, "Total"), bold=True)
    for col, fmt in [(2, MONEY0), (3, MONEY), (4, MONEY)]:
        L = get_column_letter(col)
        c = calc(ws.cell(r, col, f"=SUM({L}$4:INDEX({L}:{L},ROW()-1))"), fmt, bold=True)
        c.border = Border(top=MED, bottom=Side(style="double", color=INK))
    c = calc(ws.cell(r, 5, f'=IFERROR(C{r}/SUMIF(Dashboard!$A:$A,"Net pay",'
                           f'Dashboard!$B:$B),0)'), PCT, link=True, bold=True)
    c.border = Border(top=MED, bottom=Side(style="double", color=INK))
    ws.conditional_formatting.add(f"D5:D{r}", CellIsRule(
        operator="lessThan", formula=["0"], fill=fill("FDE5E5")))
    widths(ws, {"A": 32, "B": 13, "C": 13, "D": 13, "E": 13, "F": 3, "G": 72})
    ws.freeze_panes = "A5"

    # --------------------------------------------------------- Historical
    ws = wb.create_sheet("Historical")
    ws.sheet_properties.tabColor = TAB["recommended"]
    title(ws, f"HISTORICAL — optional  ({tag})", "One row per month, oldest first.",
          span=7)
    header(ws, 3, ["Month", "Cash", "Pay", "Expenses", "Net savings",
                   "Savings rate", "Pension contrib"], right=(2, 3, 4, 5, 6, 7))
    note(ws, "I3", "month as YYYY-MM text; savings rate = net savings ÷ pay (calculated)")

    def hist_row(row, ym, vals):
        inp(ws.cell(row, 1, ym), "@")
        for j, v in zip((2, 3, 4, 5, 7), vals):
            inp(ws.cell(row, j, v), MONEY0)
        calc(ws.cell(row, 6, f"=IF(N(C{row})=0,0,E{row}/C{row})"), PCT)

    if blank:
        hist_row(4, CUR_M, (0, 0, 0, 0, 0))
    else:
        cash = 12000.0
        for i, ym in enumerate(MONTHS):
            r0 = R[ym]
            exp = r0["rec"] + r0["oneoff"]
            cash += r0["income"] - exp - r0["tosav"]
            hist_row(4 + i, ym, (round(cash), round(r0["income"]), round(exp),
                                 round(r0["tosav"]), 320))
    widths(ws, {"A": 11, "B": 13, "C": 13, "D": 13, "E": 13, "F": 12, "G": 14,
                "H": 3, "I": 64})
    ws.freeze_panes = "A4"

    # ---------------------------------------------------------- Daily Log
    ws = wb.create_sheet("Daily Log")
    ws.sheet_properties.tabColor = TAB["optional"]
    title(ws, f"DAILY LOG — optional  ({tag})",
          "Dated snapshots; drives the month-on-month deltas.", span=6)
    header(ws, 3, ["Date", "Net Liquid", "Investments", "Liabilities", "Net Worth", "Note"],
           right=(2, 3, 4, 5))
    note(ws, "H3", "one row per month-end is plenty")
    note(ws, "H4", "row 4 is LIVE from Accounts (its date is a formula, so the dashboard "
                   "skips it). To log a snapshot: copy row 4, Paste Special › Values below.")
    # Live connection row — same idea as in a full tracker: today's figures, so a
    # snapshot is one copy/paste-values away.
    LIVE_NL = ('=SUMIF(Accounts!$A:$A,"Total liquid",Accounts!$D:$D)'
               '-SUMIFS(Accounts!$D:$D,Accounts!$E:$E,"<="&(TODAY()+365),Accounts!$E:$E,"<>")')
    calc(ws.cell(4, 1, "=TODAY()"), DATE)
    calc(ws.cell(4, 2, LIVE_NL), MONEY0, link=True)
    calc(ws.cell(4, 3, '=SUMIF(Accounts!$A:$A,"Total investments",Accounts!$D:$D)'),
         MONEY0, link=True)
    calc(ws.cell(4, 4, '=SUMIF(Accounts!$A:$A,"Total liabilities",Accounts!$D:$D)'),
         MONEY0, link=True)
    calc(ws.cell(4, 5, '=SUMIF(Accounts!$A:$A,"Net worth",Accounts!$D:$D)'),
         MONEY0, link=True)
    calc(ws.cell(4, 6, "LIVE — today, from Accounts")).alignment = Alignment(indent=1)
    ws.cell(4, 6).font = font(italic=True, color=LINK_FG)
    for cc in range(1, 7):
        ws.cell(4, cc).border = Border(bottom=MED)
    netliq = (sum(v for _, it in accounts[:3] for _, v, _ in it)
              - (0 if blank else 3200 + 2400 + 5000))
    liquid = sum(v for _, it in accounts[:3] for _, v, _ in it)
    invest = sum(v for _, v, _ in accounts[3][1])
    liab = sum(v for _, v, _ in accounts[4][1])

    def snap(row, day, vals, txt):
        inp(ws.cell(row, 1, dt.datetime(day.year, day.month, day.day)), DATE)
        for j, v in enumerate(vals, 2):
            inp(ws.cell(row, j, v), MONEY0)
        inp(ws.cell(row, 6, txt)).alignment = Alignment(indent=1)

    if blank:
        snap(5, TODAY, (0, 0, 0, 0), "Example snapshot (replace)")
    else:
        d = TODAY.replace(day=1)
        snaps = []
        for _ in range(7):
            d = d - dt.timedelta(days=1)
            snaps.append(d)
            d = d.replace(day=1)
        snaps = list(reversed(snaps)) + [TODAY]
        for i, day in enumerate(snaps):
            drift = (len(snaps) - i) * 900
            snap(5 + i, day, (round(netliq + drift), invest, liab,
                              round(liquid + invest - liab + drift)), "synthetic snapshot")
    widths(ws, {"A": 14, "B": 14, "C": 14, "D": 14, "E": 14, "F": 30, "G": 3, "H": 80})
    ws.freeze_panes = "A4"

    # ---------------------------------------------------------- Wish List
    ws = wb.create_sheet("Wish List")
    ws.sheet_properties.tabColor = TAB["optional"]
    title(ws, f"WISH LIST — optional  ({tag})",
          "Affordability panel. Columns E–S are hidden spare columns; the verdict sits "
          "in column T, which is where the dashboard reads it.", span=4)
    label(ws.cell(3, 1, "Net liquidity"))
    calc(ws.cell(3, 2, LIVE_NL), MONEY0, link=True)
    note(ws, "D3", "liquid assets − debts due within 12 months (from Accounts)")
    label(ws.cell(4, 1, "Provisions"))
    calc(ws.cell(4, 2, '=SUMIF(Bills!$A:$A,"Total monthly (all active)",Bills!$B:$B)'),
         MONEY0, link=True)
    note(ws, "D4", "this month's committed bills (from Bills)")
    label(ws.cell(5, 1, "Free cash"), bold=True)
    calc(ws.cell(5, 2, "=B3-B4"), MONEY0, bold=True)
    label(ws.cell(6, 1, "Affordability ratio"))
    inp(ws.cell(6, 2, 0.25), PCT)
    note(ws, "D6", "share of free cash any single item may consume")
    label(ws.cell(7, 1, "Max single item"))
    calc(ws.cell(7, 2, "=MAX(0,B5*B6)"), MONEY0)
    header(ws, 8, ["ID", "Item", "Category", "Cost"], right=(4,), center=(1,))
    header(ws, 8, ["Verdict"], start=20)
    wish = ([(1, "Example wish-list item (replace)", "Fun", 0)] if blank else
            [(1, "Road bike", "Fun", 1800),
             (2, "Kitchen refit", "Home", 14000),
             (3, "Laptop", "Work", 2200),
             (4, "Holiday", "Fun", 3500)])
    for i, (wid, item, cat, cost) in enumerate(wish):
        row = 9 + i
        inp(ws.cell(row, 1, wid), INT).alignment = Alignment(horizontal="center")
        inp(ws.cell(row, 2, item))
        inp(ws.cell(row, 3, cat)); inp(ws.cell(row, 4, cost), MONEY0)
        calc(ws.cell(row, 20, f'=IF(N($D{row})=0,"",IF($D{row}<=$B$7,"OK to buy now",'
                              f'"Hold — above single-item gate"))'))
    last = 9 + len(wish) - 1
    ws.conditional_formatting.add(f"T9:T{last + 50}", FormulaRule(
        formula=['LEFT($T9,2)="OK"'], font=Font(name=FONT, color="1E7B45", bold=True)))
    ws.conditional_formatting.add(f"T9:T{last + 50}", FormulaRule(
        formula=['LEFT($T9,4)="Hold"'], font=Font(name=FONT, color="1D4ED8", bold=True)))
    for col in range(5, 20):
        ws.column_dimensions[get_column_letter(col)].hidden = True
    ws.cell(8, 20).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for row in range(9, last + 1):
        ws.cell(row, 20).alignment = Alignment(indent=1)
    widths(ws, {"A": 22, "B": 34, "C": 14, "D": 12, "T": 34})

    # ------------------------------------------------------- Transactions
    ws = wb.create_sheet("Transactions")
    ws.sheet_properties.tabColor = TAB["required"]
    header(ws, 1, ["Date", "Month", "Account", "Description", "Amount",
                   "List 1", "List 2", "Bucket", "Confidence"], right=(5,), center=(2,))
    fmts = [DATE, "@", None, None, MONEY, None, None, None, None]
    if blank:
        vals = [dt.datetime(TODAY.year, TODAY.month, 1), CUR_M,
                "Example account", "EXAMPLE MERCHANT (replace)", 0,
                "Example category", "Example subcategory", "Recurring", "existing"]
        for j, v in enumerate(vals, 1):
            c = inp(ws.cell(2, j, v), fmts[j - 1])
            if j == 2:
                c.alignment = Alignment(horizontal="center")
            elif j == 6:
                c.alignment = Alignment(indent=1)
        note(ws, "K1", "Amount positive = money out. Bucket: Recurring / One-off / "
                       "Income / Savings / REVIEW")
        note(ws, "K2", "Month is YYYY-MM text (column B is pre-formatted as text so "
                       "Excel won't turn it into a date)")
        last = 2
    else:
        zebra = fill("F7F8FA")
        for i, t in enumerate(tx):
            row = 2 + i
            for j, v in enumerate(t, 1):
                c = ws.cell(row, j, v)
                if fmts[j - 1]:
                    c.number_format = fmts[j - 1]
                if j == 2:
                    c.alignment = Alignment(horizontal="center")
                elif j == 6:
                    c.alignment = Alignment(indent=1)
                if i % 2:
                    c.fill = zebra
        last = 1 + len(tx)
    # Pre-format column B as text for rows typed later, so "2026-11" stays text.
    ws.column_dimensions["B"].number_format = "@"
    bucket_dv = DataValidation(type="list",
                               formula1='"Recurring,One-off,Income,Savings,REVIEW"',
                               allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(bucket_dv)
    bucket_dv.add("H2:H20000")
    ws.conditional_formatting.add("A2:I20000", FormulaRule(
        formula=['$H2="REVIEW"'], fill=fill("FDF3D7")))
    ws.conditional_formatting.add("E2:E20000", CellIsRule(
        operator="lessThan", formula=["0"], font=Font(name=FONT, color="1E7B45")))
    ws.auto_filter.ref = f"A1:I{last}"
    ws.freeze_panes = "A2"
    widths(ws, {"A": 12, "B": 9, "C": 16, "D": 30, "E": 12, "F": 17, "G": 18,
                "H": 11, "I": 11, "J": 3, "K": 70})

    arial_everywhere(wb)
    print_setup(wb)
    out = TEMPLATE_OUT if blank else SAMPLE_OUT
    wb.save(out)
    print(f"Wrote {out}")
    print(f"  {len(tx)} transactions · sheets: {', '.join(wb.sheetnames)}")
    return out


# ------------------------------------------------------- cached formula values
def find_soffice():
    for cand in (shutil.which("soffice"), shutil.which("libreoffice"),
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if cand and os.path.exists(cand):
            return cand
    return None


def _excel_serial(v):
    if isinstance(v, dt.datetime):
        delta = v - dt.datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400
    if isinstance(v, dt.date):
        return (v - dt.date(1899, 12, 30)).days
    return v


def cache_formula_results(path):
    """openpyxl writes formulas without results, and the dashboard reads results.
    Let LibreOffice calculate a throw-away copy, then write those results into
    our own file as cached values. Our file's styling is left untouched."""
    soffice = find_soffice()
    if not soffice:
        print(f"  ! LibreOffice not found: open {path} in Excel and save it once "
              f"before loading it into the dashboard.")
        return False
    with tempfile.TemporaryDirectory() as tmp:
        prof = "file://" + os.path.join(tmp, "profile")
        subprocess.run([soffice, f"-env:UserInstallation={prof}", "--headless",
                        "--calc", "--convert-to", "xlsx", "--outdir", tmp, path],
                       capture_output=True, timeout=180)
        calc_path = os.path.join(tmp, os.path.basename(path))
        if not os.path.exists(calc_path):
            print(f"  ! LibreOffice could not calculate {path}; open it in Excel and "
                  f"save once before loading it into the dashboard.")
            return False
        wf = openpyxl.load_workbook(path)
        wv = openpyxl.load_workbook(calc_path, data_only=True)
        results, errors = {}, []
        for idx, ws in enumerate(wf.worksheets, 1):
            vals = {}
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        v = wv[ws.title][c.coordinate].value
                        if isinstance(v, str) and v.startswith("#"):
                            errors.append(f"{ws.title}!{c.coordinate} {v}")
                        vals[c.coordinate] = v
            results[f"xl/worksheets/sheet{idx}.xml"] = vals
        if errors:
            sys.exit("  ! formula errors: " + ", ".join(errors[:20]))

    cell_re = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*)><f>(.*?)</f><v\s*/>|'
                         r'<c r="([A-Z]+[0-9]+)"([^>]*)><f>(.*?)</f><v></v>')

    def patch(xml, vals):
        def sub(m):
            ref = m.group(1) or m.group(4)
            attrs = m.group(2) if m.group(1) else m.group(5)
            f = m.group(3) if m.group(1) else m.group(6)
            attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
            v = _excel_serial(vals.get(ref))
            if v is None:
                return f'<c r="{ref}"{attrs} t="str"><f>{f}</f><v></v>'
            if isinstance(v, bool):
                return f'<c r="{ref}"{attrs} t="b"><f>{f}</f><v>{int(v)}</v>'
            if isinstance(v, (int, float)):
                return f'<c r="{ref}"{attrs}><f>{f}</f><v>{repr(float(v)) if isinstance(v, float) else v}</v>'
            return f'<c r="{ref}"{attrs} t="str"><f>{f}</f><v>{escape(str(v))}</v>'
        return cell_re.sub(sub, xml)

    tmp_out = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in results:
                data = patch(data.decode("utf-8"), results[item.filename]).encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp_out, path)
    n = sum(len(v) for v in results.values())
    print(f"  cached {n} formula results (calculated by LibreOffice)")
    return True


outs = []
if WHICH in ("both", "sample"):
    outs.append(build(blank=False))
if WHICH in ("both", "template"):
    outs.append(build(blank=True))
for o in outs:
    cache_formula_results(o)
