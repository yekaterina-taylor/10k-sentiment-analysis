"""Build the NVIDIA (NVDA) 10-year workbook, FY2017-FY2026, with live formulas.

All figures in USD millions, from NVIDIA's Q4 earnings releases (Form 8-K, Ex. 99.1),
which present the same condensed statements as the 10-K. None = not yet available.
"""
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

YEARS = list(range(2026, 2016, -1))            # columns C..L, most recent first
BLUE, GOLD = "004B8D", "FFC423"
HDR = Font(name="Arial", bold=True, color="FFFFFF", size=10)
TITLE = Font(name="Arial", bold=True, color=BLUE, size=14)
SUB = Font(name="Arial", italic=True, color="555555", size=10)
BOLD = Font(name="Arial", bold=True, size=10)
NORM = Font(name="Arial", size=10)
INPUT = Font(name="Arial", size=10, color="0000FF")
LINK = Font(name="Arial", size=10, color="008000")
SECTION = Font(name="Arial", bold=True, color=BLUE, size=10)
NOTE = Font(name="Arial", italic=True, color="555555", size=9)
FILL = PatternFill("solid", fgColor=BLUE)
SOFT = PatternFill("solid", fgColor="EEF3F8")
TOP = Border(top=Side(style="thin", color="8899AA"))
NUM = '#,##0;(#,##0);"-"'

SRC = {
    2026: "https://www.sec.gov/Archives/edgar/data/1045810/000104581026000019/q4fy26pr.htm",
    2024: "https://www.sec.gov/Archives/edgar/data/1045810/000104581024000028/q4fy24pr.htm",
    2022: "https://www.sec.gov/Archives/edgar/data/1045810/000104581022000008/q4fy22pr.htm",
    2020: "https://www.sec.gov/Archives/edgar/data/1045810/000104581020000007/q4fy20pr.htm",
    2019: "https://www.sec.gov/Archives/edgar/data/1045810/000104581019000007/q4fy19pr.htm",
    2018: "https://www.sec.gov/Archives/edgar/data/1045810/000104581018000004/q4fy18pr.htm",
    2017: "https://www.sec.gov/Archives/edgar/data/1045810/000104581017000016/q4fy17pr.htm",
}
PERIOD_END = {2026: "2026-01-25", 2025: "2025-01-26", 2024: "2024-01-28", 2023: "2023-01-29",
              2022: "2022-01-30", 2021: "2021-01-31", 2020: "2020-01-26", 2019: "2019-01-27",
              2018: "2018-01-28", 2017: "2017-01-29"}

wb = openpyxl.Workbook()


def sheet(title, heading, sub, first=True):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws["A1"], ws["A2"] = heading, sub
    ws["A1"].font, ws["A2"].font = TITLE, SUB
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 22
    for i in range(len(YEARS)):
        ws.column_dimensions[L(3 + i)].width = 12
    return ws


def header(ws, row, first_label, second="Classification"):
    for j, v in enumerate([first_label, second] + [f"FY {y}" for y in YEARS]):
        c = ws.cell(row, 1 + j, v)
        c.font, c.fill = HDR, FILL
        c.alignment = Alignment(horizontal="left" if j < 2 else "right")
    ws.freeze_panes = ws.cell(row + 1, 3)


def write_rows(ws, start, rows):
    """rows: (label, class, kind, values|formula-template). Returns {key: row}."""
    pos = {}
    for k, (key, label, cls, kind, vals) in enumerate(rows):
        r = start + k
        pos[key] = r
        ws.cell(r, 1, label if kind != "line" else "    " + label)
        ws.cell(r, 2, cls)
        if kind == "head":
            ws.cell(r, 1).font = SECTION
            ws.cell(r, 2).font = Font(name="Arial", size=9, color="666666")
            continue
        ws.cell(r, 1).font = BOLD if kind == "total" else NORM
        ws.cell(r, 2).font = Font(name="Arial", size=9, color="666666")
        for i, y in enumerate(YEARS):
            c = ws.cell(r, 3 + i)
            col = L(3 + i)
            v = vals(col, pos) if callable(vals) else vals[i]
            if callable(v):
                v = v(col, pos)
            c.value = v
            c.number_format = NUM
            c.font = BOLD if kind == "total" else (INPUT if not (isinstance(v, str) and v.startswith("=")) else NORM)
            if kind == "total":
                c.border = TOP
        if kind == "total":
            ws.cell(r, 1).border = ws.cell(r, 2).border = TOP
    return pos


def f(template):
    """Formula builder: template uses {c} for column and {key} for row numbers."""
    return lambda col, pos: "=" + template.format(c=col, **pos)


def notes(ws, row, items):
    for k, t in enumerate(items):
        ws.cell(row + k, 1, t).font = NOTE if k else Font(name="Arial", bold=True, size=9, color="555555")


# ------------------------------------------------------------------ Income statement
ws = sheet("Income Statement", "NVIDIA Corp. - Consolidated Statements of Income",
           "10 Fiscal Years, FY2017-FY2026 (In Millions USD). Fiscal year ends last Sunday of January.")
header(ws, 4, "Income Statement Line Item")
IS = write_rows(ws, 5, [
    ("rev", "Revenue", "Total Revenue", "total",
     [215938, 130497, 60922, 26974, 26914, 16675, 10918, 11716, 9714, 6910]),
    ("cor", "Cost of revenue", "Cost Category", "line",
     [62475, 32639, 16621, 11618, 9439, 6279, 4150, 4545, 3892, 2847]),
    ("gp", "Gross Profit", "Operating Profitability", "total", f("{c}{rev}-{c}{cor}")),
    ("h1", "Operating Expenses:", "Section Header", "head", None),
    ("rd", "Research and development", "Opex", "line",
     [18497, 12914, 8675, 7339, 5268, 3924, 2829, 2376, 1797, 1466]),
    ("sga", "Sales, general and administrative", "Opex", "line",
     [4579, 3491, 2654, 2440, 2166, 1940, 1093, 991, 815, 663]),
    ("atc", "Acquisition termination cost (Arm)", "Opex", "line", [0, 0, 0, 1353, 0, 0, 0, 0, 0, 0]),
    ("opex", "Total Operating Expenses", "Total Opex", "total", f("SUM({c}{rd}:{c}{atc})")),
    ("oi", "Operating Income", "Operating Margin", "total", f("{c}{gp}-{c}{opex}")),
    ("h2", "Other Income / (Expense):", "Section Header", "head", None),
    ("ii", "Interest income", "Non-Operating", "line", [2300, 1786, 866, 267, 29, 57, 178, 136, 69, 54]),
    ("ie", "Interest expense", "Non-Operating", "line", [-259, -247, -257, -262, -236, -184, -52, -58, -61, -58]),
    ("on", "Other, net", "Non-Operating", "line", [9022, 1034, 237, -48, 107, 4, -2, 14, -22, -25]),
    ("toi", "Total Other Income / (Expense), Net", "Non-Operating", "total", f("SUM({c}{ii}:{c}{on})")),
    ("pti", "Income Before Income Tax", "Pretax Earnings", "total", f("{c}{oi}+{c}{toi}")),
    ("tax", "Income tax expense / (benefit)", "Tax Expense", "line",
     [21383, 11146, 4058, -187, 189, 77, 174, -245, 149, 239]),
    ("ni", "Net Income", "Bottom-Line Earnings", "total", f("{c}{pti}-{c}{tax}")),
])
notes(ws, 25, ["Notes:",
               "1. Source: NVIDIA Q4 earnings releases (Form 8-K Ex. 99.1) for FY2026, FY2024, FY2022, FY2020 and FY2018; each presents two fiscal years.",
               "2. Blue = reported input; black = formula."])

# ------------------------------------------------------------------ Balance sheet
bs = sheet("Balance Sheet", "NVIDIA Corp. - Condensed Consolidated Balance Sheets",
           "Fiscal Year-End Dates (In Millions USD) - Assets = Liabilities + Equity", first=False)
header(bs, 4, "Balance Sheet Line Item")
BS = write_rows(bs, 5, [
    ("h1", "Current Assets:", "Sub-Header", "head", None),
    ("cash", "Cash, cash equivalents and marketable securities", "Current Asset", "line",
     [62556, 43210, 25984, 13296, 21208, 11561, 10897, 7422, 7108, 6798]),
    ("ar", "Accounts receivable, net", "Current Asset", "line",
     [38466, 23065, 9999, 3827, 4650, 2429, 1657, 1424, 1265, 826]),
    ("inv", "Inventories", "Current Asset", "line", [21403, 10080, 5282, 5159, 2605, 1826, 979, 1575, 796, 794]),
    ("pre", "Prepaid expenses and other current assets", "Current Asset", "line",
     [3180, 3771, 3080, 791, 366, 239, 157, 136, 86, 118]),
    ("tca", "Total Current Assets", "Subtotal", "total", f("SUM({c}{cash}:{c}{pre})")),
    ("h2", "Non-Current Assets:", "Sub-Header", "head", None),
    ("ppe", "Property and equipment, net", "Long-Term Tangible Asset", "line",
     [10383, 6283, 3914, 3807, 2778, 2149, 1674, 1404, 997, 521]),
    ("ola", "Operating lease assets", "Non-Current Asset", "line", [2867, 1793, 1346, 1038, 829, 707, 618, 0, 0, 0]),
    ("gw", "Goodwill", "Non-Current Asset", "line", [20832, 5188, 4430, 4372, 4349, 4193, 618, 618, 618, 618]),
    ("int", "Intangible assets, net", "Non-Current Asset", "line", [3306, 807, 1112, 1676, 2339, 2737, 49, 45, 52, 104]),
    ("dta", "Deferred income tax assets", "Non-Current Asset", "line", [13258, 10979, 6081, 3396, 1222, 806, 548, 560, 0, 0]),
    ("nmes", "Non-marketable equity securities", "Non-Current Asset", "line", [22251, 3387, 0, 0, 0, 0, 0, 0, 0, 0]),
    ("oa", "Other assets", "Other Non-Current Asset", "line", [8301, 3038, 4500, 3820, 3841, 2144, 118, 108, 319, 62]),
    ("tnca", "Total Non-Current Assets", "Subtotal", "total", f("SUM({c}{ppe}:{c}{oa})")),
    ("ta", "TOTAL ASSETS", "Grand Total Assets", "total", f("{c}{tca}+{c}{tnca}")),
    ("h3", "Current Liabilities:", "Sub-Header", "head", None),
    ("ap", "Accounts payable", "Current Liability", "line", [9812, 6310, 2699, 1193, 1783, 1149, 687, 511, 596, 485]),
    ("acc", "Accrued and other current liabilities", "Current Liability", "line",
     [21352, 11737, 6682, 4120, 2552, 1777, 1097, 818, 542, 507]),
    ("std", "Short-term debt (incl. convertible)", "Short-Term Debt", "line", [999, 0, 1250, 1250, 0, 999, 0, 0, 15, 796]),
    ("tcl", "Total Current Liabilities", "Subtotal", "total", f("SUM({c}{ap}:{c}{std})")),
    ("h4", "Non-Current Liabilities:", "Sub-Header", "head", None),
    ("ltd", "Long-term debt", "Long-Term Debt", "line", [7469, 8463, 8459, 9703, 10946, 5964, 1991, 1988, 1985, 1983]),
    ("oll", "Long-term operating lease liabilities", "Long-Term Liability", "line", [2572, 1519, 1119, 902, 741, 634, 561, 0, 0, 0]),
    ("oltl", "Other long-term liabilities", "Long-Term Liability", "line", [7306, 4245, 2541, 1913, 1553, 1375, 775, 633, 632, 277]),
    ("tncl", "Total Non-Current Liabilities", "Subtotal", "total", f("SUM({c}{ltd}:{c}{oltl})")),
    ("tl", "TOTAL LIABILITIES", "Total Liabilities", "total", f("{c}{tcl}+{c}{tncl}")),
    ("mezz", "Convertible debt conversion obligation", "Temporary Equity", "line", [0, 0, 0, 0, 0, 0, 0, 0, 0, 31]),
    ("eq", "TOTAL SHAREHOLDERS' EQUITY", "Total Equity", "total",
     [157293, 79327, 42978, 22101, 26612, 16893, 12204, 9342, 7471, 5762]),
    ("tle", "TOTAL LIABILITIES AND SHAREHOLDERS' EQUITY", "Grand Total Liab & Eq", "total",
     f("{c}{tl}+{c}{mezz}+{c}{eq}")),
    ("chk", "Accounting Balance Check (Assets - Liab & Eq)", "Audit Verification", "line", f("{c}{ta}-{c}{tle}")),
])
notes(bs, 41, ["Notes:",
               "1. Earnings releases report cash, cash equivalents and marketable securities as one line.",
               "2. Before FY2020 deferred tax assets were in other assets (FY2017-FY2018); before FY2025, non-marketable equity securities were in other assets.",
               "3. FY2017 other long-term liabilities include $6M of long-term capital lease obligations.",
               "4. FY2016 total assets ($7,370M, FY2017 release) used for FY2017 average total assets."])

# ------------------------------------------------------------------ Cash flow
cf = sheet("Cash Flow Statement", "NVIDIA Corp. - Condensed Consolidated Statements of Cash Flows",
           "Operating, Investing, and Financing Flows (In Millions USD)", first=False)
header(cf, 4, "Cash Flow Line Item")
P = None   # pending (FY2017-FY2018 detail not in earnings releases)
cfo = [102718, 64089, 28090, 5641, 9108, 5822, 4761, 3743, 3502, 1672]
da = [2843, 1864, 1508, 1544, 1174, 1098, 381, 262, P, P]
sbc = [6386, 4737, 3549, 2709, 2004, 1397, 844, 557, 391, 248]
ni = [120067, 72880, 29760, 4368, 9752, 4332, 2796, 4141, 3047, 1666]
capex = [-6042, -3236, -1069, -1833, -976, -1128, -489, -600, -593, -176]
mkt = [-14233, -14885, -8429, 9334, -8567, -9989, 6648, -3488, P, P]
cfi = [-52228, -20421, -10566, 7375, -9830, -19675, 6145, -4097, P, P]
buy = [-40086, -33706, -9533, -10039, 0, 0, 0, -1579, P, -739]
div = [-974, -834, -395, -398, -399, -395, -390, -371, P, -261]
debt = [0, -1250, -1250, 0, 3977, 4968, 0, -16, P, P]
cff = [-48474, -42359, -13633, -11617, 1865, 3804, -792, -2866, P, P]
oth_op = [cfo[i] - ni[i] - da[i] - sbc[i] if da[i] is not None else P for i in range(10)]
oth_inv = [cfi[i] - capex[i] - mkt[i] if cfi[i] is not None else P for i in range(10)]
oth_fin = [cff[i] - buy[i] - div[i] - debt[i] if cff[i] is not None else P for i in range(10)]


def total(tmpl, known):
    """Formula for complete years; reported figure (or blank) where detail is pending."""
    out = []
    for i in range(10):
        if known[i] is None:
            out.append(None)
        elif i >= 8:
            out.append(known[i])      # FY2018/FY2017: reported subtotal, components incomplete
        else:
            out.append(lambda col, pos, t=tmpl: "=" + t.format(c=col, **pos))
    return out


CF = write_rows(cf, 5, [
    ("h1", "Operating Activities:", "Section Header", "head", None),
    ("ni", "Net income", "Cash Flow Reconciliation", "line",
     [lambda col, pos: f"='Income Statement'!{col}{IS['ni']}"] * 10),
    ("da", "Depreciation and amortization", "Non-Cash Expense", "line", da),
    ("sbc", "Stock-based compensation expense", "Non-Cash Expense", "line", sbc),
    ("oth", "Other adjustments & changes in working capital", "Working Capital / Other", "line", oth_op),
    ("cfo", "Net Cash Provided by Operating Activities", "Operating Cash Flow", "total",
     total("SUM({c}{ni}:{c}{oth})", cfo)),
    ("h2", "Investing Activities:", "Section Header", "head", None),
    ("capex", "Purchases of property, equipment & intangibles", "Capital Expenditures", "line", capex),
    ("mkt", "Marketable securities: purchases, maturities & sales, net", "Securities Portfolio", "line", mkt),
    ("oinv", "Other investing (acquisitions, equity investments, other)", "Other Investing", "line", oth_inv),
    ("cfi", "Net Cash Provided by / (Used in) Investing Activities", "Investing Cash Flow", "total",
     total("SUM({c}{capex}:{c}{oinv})", cfi)),
    ("h3", "Financing Activities:", "Section Header", "head", None),
    ("buy", "Repurchases of common stock", "Shareholder Capital Return", "line", buy),
    ("div", "Dividends paid", "Shareholder Capital Return", "line", div),
    ("debt", "Debt issuance / (repayment), net", "Debt Financing", "line", debt),
    ("ofin", "Other financing (employee stock plans, RSU taxes, other)", "Other Financing", "line", oth_fin),
    ("cff", "Net Cash Provided by / (Used in) Financing Activities", "Financing Cash Flow", "total",
     total("SUM({c}{buy}:{c}{ofin})", cff)),
    ("net", "Change in Cash and Cash Equivalents", "Net Cash Change", "total",
     [lambda col, pos: f"={col}{pos['cfo']}+{col}{pos['cfi']}+{col}{pos['cff']}"] * 8 + [None, None]),
])
notes(cf, 26, ["Notes:",
               "1. FY2019-FY2026 from the twelve-month cash flow statements in NVIDIA's Q4 earnings releases.",
               "2. FY2017-FY2018: the earnings releases give only operating cash flow, capex, stock-based compensation, "
               "and (FY2017) buybacks and dividends. Blank cells are pending the full 10-K cash flow statements.",
               "3. 'Other' rows combine remaining reported lines so each section ties to NVIDIA's reported subtotal."])

# ------------------------------------------------------------------ Ratios
ra = sheet("Financial Ratio Analysis", "NVIDIA Corp. - 10-Year Financial Ratio Analysis",
           "Liquidity, Solvency, Efficiency, and Profitability", first=False)
ra.column_dimensions["C"].width = 34
for j, v in enumerate(["Financial Ratio / Metric", "Category", "Formula Definition"] + [f"FY {y}" for y in YEARS]):
    c = ra.cell(4, 1 + j, v)
    c.font, c.fill = HDR, FILL
for i in range(10):
    ra.column_dimensions[L(4 + i)].width = 12
inputs = [("Revenue", "Income Statement", IS["rev"]), ("Net Income", "Income Statement", IS["ni"]),
          ("Total Current Assets", "Balance Sheet", BS["tca"]), ("Total Assets (Ending Balance)", "Balance Sheet", BS["ta"]),
          ("Total Current Liabilities", "Balance Sheet", BS["tcl"]), ("Total Liabilities", "Balance Sheet", BS["tl"])]
ra.cell(10, 1, "Underlying Financial Statement Inputs ($ in Millions)").font = SECTION
for k, (label, sh, srow) in enumerate(inputs):
    r = 11 + k
    ra.cell(r, 1, label).font = NORM
    ra.cell(r, 2, sh).font = Font(name="Arial", size=9, color="666666")
    for i in range(10):
        c = ra.cell(r, 4 + i, f"='{sh}'!{L(3 + i)}{srow}")
        c.font, c.number_format = LINK, NUM
ra.cell(17, 1, "Average Total Assets (Beg & End)").font = NORM
ra.cell(17, 3, "AVERAGE(Current Year, Prior Year)")
for i in range(10):
    col, nxt = L(4 + i), L(5 + i)
    ra.cell(17, 4 + i, f"=AVERAGE({col}14,{nxt}14)" if i < 9 else f"=AVERAGE({col}14,7370)").number_format = NUM
ra["M17"].font = INPUT
ratio_defs = [("Current Ratio", "Liquidity", "Current Assets / Current Liabilities", "={c}13/{c}15", "0.00"),
              ("Debt-to-Assets Ratio", "Solvency", "Total Liabilities / Total Assets", "={c}16/{c}14", "0.00"),
              ("Asset Turnover", "Efficiency", "Revenue / Average Total Assets", "={c}11/{c}17", "0.00"),
              ("Net Profit Margin", "Profitability", "Net Income / Revenue", "={c}12/{c}11", "0.0%")]
for k, (n, cat, frm, t, fmt) in enumerate(ratio_defs):
    r = 5 + k
    ra.cell(r, 1, n).font = BOLD
    ra.cell(r, 2, cat)
    ra.cell(r, 3, frm)
    for i in range(10):
        c = ra.cell(r, 4 + i, t.format(c=L(4 + i)))
        c.number_format, c.font = fmt, NORM
notes(ra, 20, ["Notes:", "1. FY2017 average total assets uses FY2016 ending total assets of $7,370M (FY2017 Q4 earnings release)."])

# ------------------------------------------------------------------ Prices & sentiment placeholders
hp = sheet("Historical Stock Prices", "NVIDIA Corp. - Historical Stock Prices at Balance Sheet Dates",
           "Yahoo Finance daily closes via src/get_nvda_stock_prices.py (yfinance)", first=False)
for j, v in enumerate(["Fiscal Year", "Balance Sheet Date", "Actual Trading Date", "Closing Stock Price ($ USD)",
                       "Split Factor (to post-2024 basis)", "Split-Adjusted Close ($ USD)"]):
    c = hp.cell(4, 1 + j, v)
    c.font, c.fill = HDR, FILL
# Actual closing price on the trading date (yfinance split-adjusted Close x split factor)
CLOSE = {2026: 187.67, 2025: 142.62, 2024: 610.31, 2023: 203.65, 2022: 228.40,
         2021: 519.59, 2020: 250.48, 2019: 160.15, 2018: 243.33, 2017: 111.77}
trade = {2026: "2026-01-23", 2025: "2025-01-24", 2024: "2024-01-26", 2023: "2023-01-27", 2022: "2022-01-28",
         2021: "2021-01-29", 2020: "2020-01-24", 2019: "2019-01-25", 2018: "2018-01-26", 2017: "2017-01-27"}
for k, y in enumerate(YEARS):
    r = 5 + k
    factor = 1 if y >= 2025 else (10 if y >= 2022 else 40)
    hp.cell(r, 1, f"FY {y}")
    hp.cell(r, 2, PERIOD_END[y])
    hp.cell(r, 3, trade[y])
    hp.cell(r, 4, CLOSE[y]).font = INPUT
    hp.cell(r, 4).number_format = "$#,##0.00"
    hp.cell(r, 5, factor).font = INPUT
    hp.cell(r, 6, f'=IF(D{r}="","",D{r}/E{r})').number_format = "$#,##0.00"
for c in ("B", "C", "D", "E", "F"):
    hp.column_dimensions[c].width = 22
notes(hp, 16, ["Notes:",
               "1. NVIDIA split 4-for-1 on July 20, 2021 and 10-for-1 on June 10, 2024. Split factor converts older prices to today's share basis.",
               "2. Fiscal years end on the last Sunday of January; prices are from the preceding Friday.",
               "3. Source: Yahoo Finance via yfinance. Its Close is already split-adjusted; the actual close shown is that value x split factor."])

sa = sheet("Sentiment Analysis", "NVIDIA Corp. - Loughran-McDonald MD&A Sentiment Analysis",
           "Item 7 (MD&A) word counts from src/lm_sentiment.py, sentiment-ready text", first=False)
# Counts per fiscal year: total, positive, negative, uncertainty, litigious, strong modal, weak modal, constraining
SENT = {
    2026: [4882, 22, 41, 60, 24, 10, 31, 48],
    2025: [5458, 30, 47, 76, 22, 9, 45, 63],
    2024: [5803, 42, 55, 72, 25, 8, 39, 55],
    2023: [5553, 33, 61, 52, 20, 7, 25, 45],
    2022: [6062, 41, 58, 75, 27, 6, 30, 37],
    2021: [6051, 24, 48, 83, 24, 15, 22, 40],
    2020: [4965, 31, 49, 42, 26, 6, 13, 35],
    2019: [6916, 37, 70, 60, 43, 9, 15, 41],
    2018: [7046, 49, 75, 57, 46, 17, 18, 45],
    2017: [7627, 42, 96, 89, 37, 20, 33, 49],
}
for j, v in enumerate(["Fiscal Year", "Total Words", "Positive", "Negative", "Uncertainty", "Litigious",
                       "Strong Modal", "Weak Modal", "Constraining", "P/N Ratio", "Sentiment Score"]):
    c = sa.cell(4, 1 + j, v)
    c.font, c.fill = HDR, FILL
for k, y in enumerate(YEARS):
    r = 5 + k
    sa.cell(r, 1, f"FY {y}").font = BOLD
    for j, v in enumerate(SENT[y]):
        c = sa.cell(r, 2 + j, v)
        c.font, c.number_format = INPUT, "#,##0"
    c = sa.cell(r, 10, f"=IF(D{r}=0,\"\",C{r}/D{r})")
    c.font, c.number_format = NORM, "0.00"
    c = sa.cell(r, 11, f"=(C{r}-D{r})/B{r}")
    c.font, c.number_format = NORM, "0.0000"
for col in "BCDEFGHIJK":
    sa.column_dimensions[col].width = 14
notes(sa, 17, ["Notes:",
               "1. Counts use exact matching against the Loughran-McDonald Master Dictionary (no stemming); Strong/Weak modal = LM modal codes 1 and 3.",
               "2. P/N Ratio = Positive / Negative. Sentiment Score = (Positive - Negative) / Total Words.",
               "3. Text: NVDA_FY<year>_MDA_SentimentReady.txt (table fragments and page headers removed)."])

so = sheet("Sources", "NVIDIA Corp. - Sources", "SEC EDGAR, CIK 0001045810", first=False)
for j, v in enumerate(["Fiscal years covered", "Document", "URL"]):
    c = so.cell(4, 1 + j, v)
    c.font, c.fill = HDR, FILL
rows = [("FY2026, FY2025", "Q4 FY2026 earnings release (Form 8-K Ex. 99.1)", SRC[2026]),
        ("FY2024, FY2023", "Q4 FY2024 earnings release (Form 8-K Ex. 99.1)", SRC[2024]),
        ("FY2022, FY2021", "Q4 FY2022 earnings release (Form 8-K Ex. 99.1)", SRC[2022]),
        ("FY2020, FY2019", "Q4 FY2020 earnings release (Form 8-K Ex. 99.1)", SRC[2020]),
        ("FY2019, FY2018 (balance sheet)", "Q4 FY2019 earnings release (Form 8-K Ex. 99.1)", SRC[2019]),
        ("FY2018, FY2017", "Q4 FY2018 earnings release (Form 8-K Ex. 99.1)", SRC[2018]),
        ("FY2017, FY2016 (total assets)", "Q4 FY2017 earnings release (Form 8-K Ex. 99.1)", SRC[2017])]
for k, row in enumerate(rows):
    for j, v in enumerate(row):
        so.cell(5 + k, 1 + j, v).font = NORM
so.column_dimensions["C"].width = 95

out = "NVDA_10_Year_Analysis.xlsx"
wb.save(out)
print("saved", out, IS, BS["ta"], CF["cfo"])
