"""Correct the Apple 10-year workbook using figures from Apple's 10-K / Q4 8-K filings."""
from copy import copy

import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Font

SRC = "orig.xlsx"
OUT = "Apple_10_Year_Financial_and_Sentiment_Analysis_FIXED.xlsx"
YEARS = [2025, 2024, 2023, 2022, 2021, 2020, 2019, 2018, 2017, 2016]  # columns C..L
BLUE = Font(name="Arial", size=10, color="0000FF")

wb = openpyxl.load_workbook(SRC)

# ---------------------------------------------------------------- Balance sheet
# row -> values FY2025..FY2016 (USD millions), as reported in each year's filing
bs_rows = {
    7: [35934, 29943, 29965, 23646, 34940, 38016, 48844, 25913, 20289, 20484],      # cash
    8: [18763, 35228, 31590, 24658, 27699, 52927, 51713, 40388, 53892, 46671],      # mkt sec (current)
    9: [39777, 33410, 29508, 28184, 26278, 16120, 22926, 23186, 17874, 15754],      # A/R
    10: [5718, 7286, 6331, 4946, 6580, 4061, 4106, 3956, 4855, 2132],               # inventories
    11: [33180, 32833, 31477, 32748, 25228, 21325, 22878, 25809, 17799, 13545],     # vendor NTR
    12: [14585, 14287, 14695, 21223, 14111, 11264, 12352, 12087, 13936, 8283],      # other current
    15: [77723, 91479, 100544, 120805, 127877, 100887, 105341, 170799, 194714, 170430],
    16: [49834, 45680, 43715, 42117, 39440, 36766, 37378, 41304, 33783, 27010],
    17: [83727, 74834, 64758, 54428, 48849, 42522, 32978, 22283, 18177, 17377],
    22: [69860, 68960, 62611, 64115, 54763, 42296, 46236, 55888, 49049, 37294],     # A/P
    23: [66387, 78304, 58829, 60845, 47493, 42684, 37720, 32687, 25744, 22027],     # other CL
    24: [9055, 8249, 8061, 7912, 7612, 6643, 5522, 7543, 7548, 8080],               # deferred rev
    25: [7979, 9967, 5985, 9982, 6000, 4996, 5980, 11964, 11977, 8105],             # comm. paper
    26: [12350, 10912, 9822, 11128, 9613, 8773, 10260, 8784, 6496, 3500],           # term debt cur
    29: [78328, 85750, 95281, 98959, 109106, 98667, 91807, 93735, 97207, 75427],
    30: [41549, 45888, 49848, 49142, 53325, 54490, 50503, 47977, 43251, 39004],
    34: [93568, 83276, 73812, 64849, 57365, 50779, 45174, 40201, 35867, 31251],
    35: [-14264, -19154, -214, -3068, 5562, 14966, 45898, 70400, 98330, 96364],
    36: [-5571, -7172, -11452, -11109, 163, -406, -584, -3454, -150, 634],
}
bs = wb["Balance Sheet"]
for r, vals in bs_rows.items():
    for i, v in enumerate(vals):
        bs.cell(r, 3 + i).value = v
bs["A2"] = "Fiscal Year-End Dates (In Millions USD) - Assets = Liabilities + Equity"
notes = [
    "Notes:",
    "1. Figures as reported in each fiscal year's Form 10-K / Q4 earnings release (Form 8-K Ex. 99.1), SEC EDGAR CIK 0000320193.",
    "2. FY2016-FY2017 'Other non-current assets' includes goodwill and acquired intangibles; "
    "FY2016-FY2018 'Other non-current liabilities' includes non-current deferred revenue.",
    "3. Corrected Oct 31, 2026: component line items for FY2017, FY2018 and FY2023-FY2025 replaced; totals were already correct.",
]
for k, t in enumerate(notes):
    c = bs.cell(41 + k, 1, t)
    c.font = Font(name="Arial", size=9, bold=(k == 0), italic=(k > 0), color="555555")

# ---------------------------------------------------------------- Cash flow
cf = wb["Cash Flow Statement"]
sty_head = [copy(c._style) for c in cf[5][:12]]
sty_line = [copy(c._style) for c in cf[6][:12]]
sty_tot = [copy(c._style) for c in cf[10][:12]]
sty_net = [copy(c._style) for c in cf[20][:12]]
for row in cf.iter_rows(min_row=5, max_row=cf.max_row):
    for c in row:
        c.value = None

# Reported figures FY2025..FY2016. Plug rows ("other") tie to each year's reported subtotal.
ni = [112010, 93736, 96995, 99803, 94680, 57411, 55256, 59531, 48351, 45687]
da = [11698, 11445, 11519, 11104, 11284, 11056, 12547, 10903, 10157, 10505]
sbc = [12863, 11688, 10833, 9038, 7906, 6829, 6068, 5340, 4840, 4210]
cfo = [111482, 118254, 110543, 122151, 104038, 80674, 69391, 77434, 63598, 65824]
capex = [-12715, -9447, -10959, -10708, -11085, -7309, -10495, -13313, -12451, -12734]
mktsec = [29390, 13690, 16001, -9560, -3075, 5453, 57460, 32363, -33147, -30634]
cfi = [15195, 2935, 3705, -22354, -14545, -4289, 45896, 16066, -46446, -45977]
buyback = [-90711, -94949, -77550, -89402, -85971, -72358, -66897, -72738, -32900, -29722]
divs = [-15421, -15234, -15025, -14841, -14467, -14081, -14119, -13712, -12769, -12150]
debt = [-8483, -5998, -9901, -123, 12665, 2499, -7819, 432, 29014, 22057]
cff = [-120686, -121983, -108488, -110749, -93353, -86820, -90976, -87876, -17347, -20483]
# FY2023 marketable-securities net per 8-K lines is 15,999 (rounding vs. 10-K); other investing is the plug.
mktsec[2] = 15999
oth_op = [cfo[i] - ni[i] - da[i] - sbc[i] for i in range(10)]
oth_inv = [cfi[i] - capex[i] - mktsec[i] for i in range(10)]
oth_fin = [cff[i] - buyback[i] - divs[i] - debt[i] for i in range(10)]

layout = [
    ("Operating Activities:", "Section Header", None, "head"),
    ("    Net income", "Cash Flow Reconciliation", ni, "line"),
    ("    Adjustments: Depreciation and amortization", "Non-Cash Expense", da, "line"),
    ("    Adjustments: Share-based compensation expense", "Non-Cash Expense", sbc, "line"),
    ("    Other adjustments & changes in working capital", "Working Capital / Other", oth_op, "line"),
    ("Cash Generated by Operating Activities", "Operating Cash Flow", "=SUM({c}6:{c}9)", "tot"),
    ("Investing Activities:", "Section Header", None, "head"),
    ("    Payments for acquisition of property, plant and equipment", "Capital Expenditures", capex, "line"),
    ("    Marketable securities: purchases, maturities & sales, net", "Securities Portfolio", mktsec, "line"),
    ("    Other investing (acquisitions, non-marketable securities, other)", "Other Investing", oth_inv, "line"),
    ("Cash Generated by / (Used in) Investing Activities", "Investing Cash Flow", "=SUM({c}12:{c}14)", "tot"),
    ("Financing Activities:", "Section Header", None, "head"),
    ("    Repurchases of common stock", "Shareholder Capital Return", buyback, "line"),
    ("    Payments for dividends and dividend equivalents", "Shareholder Capital Return", divs, "line"),
    ("    Term debt & commercial paper, net", "Debt Financing", debt, "line"),
    ("    Other financing (net share settlement taxes, stock issuance, other)", "Other Financing", oth_fin, "line"),
    ("Cash Used in Financing Activities", "Financing Cash Flow", "=SUM({c}17:{c}20)", "tot"),
    ("Net Increase / (Decrease) in Cash, Cash Equivalents & Restricted Cash", "Net Cash Change",
     "={c}10+{c}15+{c}21", "net"),
]
styles = {"head": sty_head, "line": sty_line, "tot": sty_tot, "net": sty_net}
for k, (label, cls, vals, kind) in enumerate(layout):
    r = 5 + k
    for j in range(12):
        cf.cell(r, 1 + j)._style = copy(styles[kind][j])
    cf.cell(r, 1, label)
    cf.cell(r, 2, cls)
    if vals is None:
        continue
    for i in range(10):
        col = openpyxl.utils.get_column_letter(3 + i)
        cf.cell(r, 3 + i).value = vals.format(c=col) if isinstance(vals, str) else vals[i]

cf_notes = [
    "Notes:",
    "1. Figures as reported in each fiscal year's Form 10-K / Q4 earnings release (Form 8-K Ex. 99.1).",
    "2. 'Other' rows combine the remaining reported line items so each section ties to Apple's reported subtotal.",
    "3. FY2016-FY2017 net change excludes restricted cash (Apple adopted the restricted-cash standard in FY2018).",
    "4. Corrected Oct 31, 2026: line items in every year replaced; prior sheet had wrong components and "
    "wrong FY2017 financing / FY2024-FY2025 investing and financing totals.",
]
for k, t in enumerate(cf_notes):
    c = cf.cell(24 + k, 1, t)
    c.font = Font(name="Arial", size=9, bold=(k == 0), italic=(k > 0), color="555555")

# ---------------------------------------------------------------- Ratio inputs -> live links
ra = wb["Financial Ratio Analysis"]
links = {12: ("Income Statement", 8), 13: ("Income Statement", 22), 14: ("Balance Sheet", 13),
         15: ("Balance Sheet", 19), 16: ("Balance Sheet", 27), 17: ("Balance Sheet", 32)}
for r, (sheet, srow) in links.items():
    for i in range(10):
        src_col = openpyxl.utils.get_column_letter(3 + i)
        c = ra.cell(r, 4 + i)
        c.value = f"='{sheet}'!{src_col}{srow}"
        f = copy(c.font)
        f.color = "008000"
        c.font = f
ra["M18"].comment = Comment("FY2015 total assets of $290,479M from Apple's FY2016 Form 10-K.", "Analyst")

# ---------------------------------------------------------------- Split-adjusted prices
hp = wb["Historical Stock Prices"]
for col, head in ((8, "Split Factor (to post-2020 basis)"), (9, "Split-Adjusted Close ($ USD)")):
    c = hp.cell(4, col, head)
    c._style = copy(hp.cell(4, 7)._style)
for r in range(5, 15):
    fy = int(hp.cell(r, 1).value.split()[-1])
    f = hp.cell(r, 8, 4 if fy <= 2019 else 1)
    f._style = copy(hp.cell(r, 4)._style)
    f.font = BLUE
    p = hp.cell(r, 9, f"=D{r}/H{r}")
    p._style = copy(hp.cell(r, 4)._style)
hp.column_dimensions["H"].width = 18
hp.column_dimensions["I"].width = 18
hp["A19"] = ("3. 'Adjusted Close' is Yahoo Finance's dividend- and split-adjusted series. "
             "'Split-Adjusted Close' (column I) = Closing Price / Split Factor and is used in all charts.")
hp["A20"] = "4. Year-ends fall on last Saturday; prices on preceding Friday."
hp["A21"] = "5. Verify FY2025 close (Sep 26, 2025) against Yahoo Finance; one source reports $255.13."

# ---------------------------------------------------------------- Stock Price vs Sentiment
sv = wb["Stock Price vs Sentiment"]
sv["B4"] = "Split-Adjusted Closing Price ($ USD)"
for r in range(5, 15):            # FY2016..FY2025 top to bottom
    hp_row = 14 - (r - 5)         # Historical sheet: FY2025 at row 5 ... FY2016 at row 14
    sa_row = hp_row               # Sentiment sheet uses same order (FY2025 at row 5)
    sv.cell(r, 2).value = f"='Historical Stock Prices'!I{hp_row}"
    sv.cell(r, 3).value = f"='Sentiment Analysis'!M{sa_row}"
    sv.cell(r, 4).value = f"='Sentiment Analysis'!N{sa_row}"
    sv.cell(r, 2).number_format = '$#,##0.00'
    sv.cell(r, 3).number_format = '0.00'
    sv.cell(r, 4).number_format = '0.0000'


def dual_chart(ws, price_col, pn_col, first, last, title, anchor):
    """Line chart: price on the left axis, P/N ratio on a secondary right axis."""
    c1 = LineChart()
    c1.title = title
    c1.y_axis.title = "Split-Adjusted Price ($)"
    c1.add_data(Reference(ws, min_col=price_col, min_row=first - 1, max_row=last), titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=first, max_row=last)
    c1.set_categories(cats)
    c1.series[0].marker.symbol = "circle"
    c2 = LineChart()
    c2.add_data(Reference(ws, min_col=pn_col, min_row=first - 1, max_row=last), titles_from_data=True)
    c2.set_categories(cats)
    c2.y_axis.axId = 200
    c2.y_axis.title = "P/N Ratio"
    c2.y_axis.crosses = "max"
    c2.y_axis.majorGridlines = None
    c2.series[0].marker.symbol = "diamond"
    c2.series[0].graphicalProperties.line.dashStyle = "dash"
    for ax in (c1.x_axis, c1.y_axis, c2.y_axis):
        ax.delete = False
    for ser in c1.series + c2.series:
        ser.smooth = False
    c1 += c2
    c1.legend.position = "b"
    c1.width, c1.height = 18, 9
    ws.add_chart(c1, anchor)


sv._charts = []
dual_chart(sv, 2, 3, 5, 14, "Apple Inc.: Split-Adjusted Stock Price vs. MD&A Sentiment P/N Ratio", "G4")

# ---------------------------------------------------------------- Dashboard -> live links
db = wb["Financial Dashboard"]
db["I8"] = "Split-Adj. Price ($)"
for r in range(9, 19):           # FY2016..FY2025
    stmt_col = openpyxl.utils.get_column_letter(3 + (18 - r))   # FY2025 -> C
    ra_col = openpyxl.utils.get_column_letter(4 + (18 - r))     # FY2025 -> D
    src_row = 5 + (18 - r)                                      # FY2025 -> row 5
    formulas = {
        2: f"='Income Statement'!{stmt_col}8", 3: f"='Income Statement'!{stmt_col}22",
        4: f"='Income Statement'!{stmt_col}18", 5: f"='Financial Ratio Analysis'!{ra_col}5",
        6: f"='Financial Ratio Analysis'!{ra_col}6", 7: f"='Financial Ratio Analysis'!{ra_col}7",
        8: f"='Financial Ratio Analysis'!{ra_col}8", 9: f"='Historical Stock Prices'!I{src_row}",
        10: f"='Sentiment Analysis'!C{src_row}", 11: f"='Sentiment Analysis'!D{src_row}",
        12: f"='Sentiment Analysis'!M{src_row}",
    }
    for col, fml in formulas.items():
        db.cell(r, col).value = fml
db["B4"] = ('="10-Year Cumulative Revenue"&CHAR(10)&TEXT(SUM(B9:B18),"$#,##0")&"M"')
db["E4"] = ('="10-Year Cumulative Net Income"&CHAR(10)&TEXT(SUM(C9:C18),"$#,##0")&"M"')
db["H4"] = ('="10-Year Cumulative Operating Cash Flow"&CHAR(10)&'
            'TEXT(SUM(\'Cash Flow Statement\'!C10:L10),"$#,##0")&"M"')
db["K4"] = ('="FY 2025 Closing Stock Price"&CHAR(10)&TEXT(\'Historical Stock Prices\'!D5,"$0.00")&" per share"')
db["N4"] = ('="FY 2025 Current Ratio"&CHAR(10)&TEXT(E18,"0.00")&"x (Working Capital "&'
            'TEXT((\'Balance Sheet\'!C13-\'Balance Sheet\'!C27)/1000,"$#,##0.00;-$#,##0.00")&"B)"')

# Replace chart 3 (price) and chart 5 (dual axis lost on load) with rebuilt versions.
keep = db._charts[:2] + [db._charts[3]]
price_chart = LineChart()
price_chart.title = "Chart 3: Split-Adjusted Closing Stock Price ($ USD)"
price_chart.add_data(Reference(db, min_col=9, min_row=8, max_row=18), titles_from_data=True)
price_chart.set_categories(Reference(db, min_col=1, min_row=9, max_row=18))
price_chart.x_axis.delete = False
price_chart.y_axis.delete = False
price_chart.width, price_chart.height = 15, 7.5
db._charts = keep
db.add_chart(price_chart, "B37")
dual_chart(db, 9, 12, 9, 18, "Chart 5: Split-Adjusted Stock Price vs. Sentiment P/N Ratio", "F53")

# ---------------------------------------------------------------- Sentiment counts (re-scored)
# Same method as NVIDIA: src/lm_sentiment.py, exact Loughran-McDonald matching on the
# sentiment-ready MD&A text. Replaces the original program's counts, which could not be reproduced.
import json
lm = {r["year"]: r for r in json.load(open("/home/claude/10k-sentiment-analysis/data/compare/AAPL_sentiment_lm.json"))}
sa = wb["Sentiment Analysis"]
sa["A2"] = "Item 7 MD&A word counts from src/lm_sentiment.py (exact Loughran-McDonald matching, sentiment-ready text)"
keys = ["total", "positive", "negative", "uncertainty", "litigious", "strongModal", "weakModal", "constraining"]
for k, y in enumerate(YEARS):
    r = 5 + k
    assert sa.cell(r, 1).value == f"FY {y}"
    for j, key in enumerate(keys):
        sa.cell(r, 2 + j).value = lm[y][key]
sa["A17"] = ("Note: counts re-scored Oct 31, 2026 with the same script used for NVIDIA so the two companies are "
             "comparable. Strong/Weak modal = LM modal codes 1 and 3; no stemming.")
sa["A17"].font = Font(name="Arial", size=9, italic=True)

# ---------------------------------------------------------------- Sources tab text fixes
sr = wb["Sources and Verification"]
sr["B8"] = "Balance Sheet Equation (A = L + E)"
sr["E11"] = "FY 2015 balance sheet ending assets ($290,479M) utilized to establish FY 2016 beginning assets."
sr["E7"] = ("Line items re-verified Oct 31, 2026 against each year's 10-K / Q4 8-K Ex. 99.1; "
            "balance sheet and cash flow components corrected.")
sr["E14"] = ("Exact balance sheet dates identified; weekend dates rolled back to preceding Friday. "
             "Charts use split-adjusted close.")

wb.save(OUT)
print("saved", OUT)
