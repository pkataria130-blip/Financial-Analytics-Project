#!/usr/bin/env python3
r"""
===============================================================================
Financial Analytics using Python
Script 09 : Two-Stock Efficient Frontier  (Historical)
Author    : Vipul Mehta
===============================================================================

PURPOSE
-------
Builds the classic two-asset efficient-frontier exhibit for a chosen PAIR of
stocks, using HISTORICAL (not assumed) inputs taken from the price file. The
whole thing is LIVE EXCEL: the expected return, standard deviation and
correlation are formulas that read the daily returns, and every portfolio's
return and risk is a formula of the weights -- change a weight and the table and
chart update.

We use SUNPHARMA.NS + INFY.NS by default because they are the lowest-correlated
pair, so the frontier bends noticeably to the left (the diversification story).

THE TWO-ASSET FORMULAS  (with A = Stock A, B = Stock B)
------------------------------------------------------
    E[r_p] = w_A * E[r_A] + w_B * E[r_B]
    sd_p   = sqrt( w_A^2 sd_A^2 + w_B^2 sd_B^2 + 2 w_A w_B rho sd_A sd_B )
where E[.] and sd are ANNUALISED historical figures and rho is the historical
correlation of daily returns. Everything is built on SIMPLE returns (a portfolio
return is a weighted sum of simple returns; log returns do not aggregate that way).

INPUT  : Yahoo_Daily_Prices_2016_2026.csv
OUTPUT : Efficient_Frontier_TwoStock.xlsx
    Sheet 1 "Efficient Frontier" : inputs, weight/return/risk table, scatter chart
    Sheet 2 "Daily Returns"      : the two daily return series the formulas read
    Sheet 3 "Instructions"       : legend + exercises

COLOUR CONVENTION (same as our other models)
    BLUE  = inputs you may edit (weights, annualisation factor)
    BLACK = formulas on this sheet
    GREEN = formulas that read from the Daily Returns sheet
===============================================================================
"""

# =============================================================================
# IMPORT LIBRARIES
# =============================================================================
import os
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker

# =============================================================================
# USER SETTINGS
# =============================================================================
FOLDER  = r"D:\Trimester 4\Finance\Financial Analytics\Financial Analytics Project\02_Output"
STOCK_A = "GRASIM.NS"           # "Share A"
STOCK_B = "INFY.NS"             # "Share B"
ANNUALIZATION_FACTOR = 250
WEIGHT_STEP = 0.10              # 0.20 -> 6 portfolios (100/0 ... 0/100).
                               # Set 0.10 or 0.05 for a smoother frontier curve.

INPUT_FILE  = os.path.join(FOLDER, "Yahoo_Daily_Prices_2016_2026.csv")
OUTPUT_FILE = os.path.join(FOLDER, "Efficient_Frontier_TwoStock.xlsx")

# =============================================================================
# STEP 1 : READ PRICES, KEEP THE TWO CHOSEN STOCKS' ADJUSTED CLOSE
# =============================================================================
if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(INPUT_FILE)

df = pd.read_csv(INPUT_FILE)
date_col = df.columns[0]

def adj_col_for(ticker):
    hits = [c for c in df.columns if "Adj Close" in c and ticker in c]
    if not hits:
        hits = [c for c in df.columns if ticker in c]
    if not hits:
        raise ValueError(f"Could not find a price column for {ticker}. "
                         f"Is it in the data file? (Re-run Script 01.)")
    return hits[0]

colA, colB = adj_col_for(STOCK_A), adj_col_for(STOCK_B)
prices = df[[date_col, colA, colB]].copy()
prices.columns = ["Date", STOCK_A, STOCK_B]
prices["Date"] = pd.to_datetime(prices["Date"])

# =============================================================================
# STEP 2 : DAILY SIMPLE RETURNS (aligned, no gaps)
# =============================================================================
returns = prices.set_index("Date").pct_change().dropna()
n_days = len(returns)
print(f"Pair: {STOCK_A} (A) + {STOCK_B} (B)   |   {n_days} return observations")

# =============================================================================
# STEP 3 : WEIGHT GRID  (Weight A from 100% down to 0%)
# =============================================================================
n_steps = int(round(1.0 / WEIGHT_STEP))
weights_A = [round(1.0 - i * WEIGHT_STEP, 10) for i in range(n_steps + 1)]
n_port = len(weights_A)

# =============================================================================
# STEP 4 : STYLES
# =============================================================================
BLUE  = Font(name="Arial", color="0000FF")
BLACK = Font(name="Arial", color="000000")
GREEN = Font(name="Arial", color="008000")
BOLD  = Font(name="Arial", bold=True)
WBOLD = Font(name="Arial", bold=True, color="FFFFFF")
TITLE = Font(name="Arial", bold=True, size=14)
HEAD  = PatternFill("solid", fgColor="4F81BD")
YEL   = PatternFill("solid", fgColor="FFF2CC")
thin  = Side(style="thin", color="BFBFBF")
BOX   = Border(left=thin, right=thin, top=thin, bottom=thin)
CTR   = Alignment(horizontal="center")
PCT, PCT2 = "0.0%", "0.00%"

wb = Workbook()

# =============================================================================
# STEP 5 : DAILY RETURNS SHEET  (source the formulas read from)
# =============================================================================
ret_ws = wb.create_sheet("Daily Returns")
ret_ws.append(["Date", STOCK_A, STOCK_B])
for c in ret_ws[1]:
    c.fill = HEAD; c.font = WBOLD
for row in returns.reset_index().itertuples(index=False):
    ret_ws.append([row[0].date(), float(row[1]), float(row[2])])
for r in ret_ws.iter_rows(min_row=2, min_col=2):
    for c in r:
        c.number_format = "0.0000%"
last = n_days + 1
A_RET = f"'Daily Returns'!$B$2:$B${last}"      # Stock A daily returns
B_RET = f"'Daily Returns'!$C$2:$C${last}"      # Stock B daily returns

# =============================================================================
# STEP 6 : EFFICIENT FRONTIER SHEET
# =============================================================================
ws = wb.create_sheet("Efficient Frontier", 0)
ws["A1"] = "Two-Stock Efficient Frontier (Historical)"
ws["A1"].font = TITLE

# ---- 6a. Annualisation factor (blue input) --------------------------------
ws["A3"] = "Annualisation factor (trading days/yr)"; ws["A3"].font = BOLD
ws["B3"] = ANNUALIZATION_FACTOR; ws["B3"].font = BLUE; ws["B3"].border = BOX

# ---- 6b. Historical inputs block (green = read from Daily Returns) ---------
ws["B5"] = "Share A"; ws["C5"] = "Share B"
ws["B5"].value = STOCK_A; ws["C5"].value = STOCK_B
for cell in ("B5", "C5"):
    ws[cell].fill = HEAD; ws[cell].font = WBOLD; ws[cell].alignment = CTR; ws[cell].border = BOX

ws["A6"] = "Expected return (annualised)"; ws["A6"].font = BOLD
ws["B6"] = f"=AVERAGE({A_RET})*$B$3"
ws["C6"] = f"=AVERAGE({B_RET})*$B$3"

ws["A7"] = "Standard deviation (annualised)"; ws["A7"].font = BOLD
ws["B7"] = f"=STDEV({A_RET})*SQRT($B$3)"
ws["C7"] = f"=STDEV({B_RET})*SQRT($B$3)"

ws["A8"] = "Correlation (A,B)"; ws["A8"].font = BOLD
ws["B8"] = f"=CORREL({A_RET},{B_RET})"

for cell, fmt in [("B6", PCT2), ("C6", PCT2), ("B7", PCT2), ("C7", PCT2), ("B8", "0.000")]:
    ws[cell].font = GREEN; ws[cell].number_format = fmt; ws[cell].border = BOX

# Named input cells used by the table below.
EXP_A, EXP_B = "$B$6", "$C$6"
SD_A,  SD_B  = "$B$7", "$C$7"
RHO          = "$B$8"

# ---- 6c. Portfolio table (mirrors the classroom exhibit) ------------------
hdr = 10
for col, text in zip("ABCDE",
                      ["Portfolio", "Weight A", "Weight B",
                       "Expected Return", "Standard Deviation"]):
    c = ws[f"{col}{hdr}"]; c.value = text
    c.fill = HEAD; c.font = WBOLD; c.alignment = CTR; c.border = BOX

first = hdr + 1
for i, wA in enumerate(weights_A):
    r = first + i
    ws.cell(row=r, column=1, value=f"Portfolio {i+1}").font = BLACK
    # Weight A : BLUE editable input
    a = ws.cell(row=r, column=2, value=wA); a.font = BLUE; a.number_format = PCT; a.border = BOX
    # Weight B = 1 - Weight A
    b = ws.cell(row=r, column=3, value=f"=1-B{r}"); b.font = BLACK; b.number_format = PCT; b.border = BOX
    # Expected return = wA*E_A + wB*E_B
    d = ws.cell(row=r, column=4,
                value=f"=B{r}*{EXP_A}+C{r}*{EXP_B}")
    d.font = BLACK; d.number_format = PCT2; d.border = BOX
    # Standard deviation = sqrt( wA^2 sdA^2 + wB^2 sdB^2 + 2 wA wB rho sdA sdB )
    e = ws.cell(row=r, column=5,
                value=(f"=SQRT(B{r}^2*{SD_A}^2 + C{r}^2*{SD_B}^2 "
                       f"+ 2*B{r}*C{r}*{RHO}*{SD_A}*{SD_B})"))
    e.font = BLACK; e.number_format = PCT2; e.border = BOX
last_port = first + n_port - 1

# ---- 6d. Analytical minimum-variance portfolio (the tip of the bullet) ----
mv = last_port + 2
ws.cell(row=mv, column=1, value="Minimum-variance portfolio (analytical)").font = BOLD
ws.cell(row=mv+1, column=1, value="Weight A*").font = BLACK
# w_A* = (sdB^2 - rho sdA sdB) / (sdA^2 + sdB^2 - 2 rho sdA sdB)
ws.cell(row=mv+1, column=2,
        value=(f"=({SD_B}^2-{RHO}*{SD_A}*{SD_B})/"
               f"({SD_A}^2+{SD_B}^2-2*{RHO}*{SD_A}*{SD_B})"))
ws.cell(row=mv+1, column=3, value=f"=1-B{mv+1}")               # Weight B*
ws.cell(row=mv+2, column=1, value="Return* / Risk*").font = BLACK
ws.cell(row=mv+2, column=2,                                     # Return*
        value=f"=B{mv+1}*{EXP_A}+C{mv+1}*{EXP_B}")
ws.cell(row=mv+2, column=3,                                     # Risk*
        value=(f"=SQRT(B{mv+1}^2*{SD_A}^2 + C{mv+1}^2*{SD_B}^2 "
               f"+ 2*B{mv+1}*C{mv+1}*{RHO}*{SD_A}*{SD_B})"))
for (rr, cc, fmt) in [(mv+1,2,PCT),(mv+1,3,PCT),(mv+2,2,PCT2),(mv+2,3,PCT2)]:
    cell = ws.cell(row=rr, column=cc)
    cell.font = BLACK; cell.number_format = fmt; cell.border = BOX; cell.fill = YEL

# ---- 6e. Markowitz scatter chart (X = risk, Y = return) -------------------
chart = ScatterChart()
chart.title = f"Portfolios 1-{n_port}: Markowitz Shape"
chart.x_axis.title = "Standard deviation"
chart.y_axis.title = "Return"
chart.x_axis.delete = False
chart.y_axis.delete = False
chart.x_axis.numFmt = "0.0%"
chart.y_axis.numFmt = "0.0%"
chart.height = 9; chart.width = 15
xref = Reference(ws, min_col=5, min_row=first, max_row=last_port)   # SD column
yref = Reference(ws, min_col=4, min_row=first, max_row=last_port)   # Return column
series = Series(yref, xref, title="Portfolios")
series.marker = Marker(symbol="circle", size=7)
series.graphicalProperties.line.noFill = True                      # markers only
chart.series.append(series)
ws.add_chart(chart, "G5")

# ---- 6f. Column widths ----------------------------------------------------
ws.column_dimensions["A"].width = 34
for col in "BCDE":
    ws.column_dimensions[col].width = 15

# =============================================================================
# STEP 7 : INSTRUCTIONS SHEET
# =============================================================================
ins = wb.create_sheet("Instructions")
lines = [
 ("Two-Stock Efficient Frontier", TITLE),
 ("", BLACK),
 ("Colour legend", BOLD),
 ("  BLUE  = inputs you may edit (Weight A cells, annualisation factor)", BLUE),
 ("  BLACK = formulas on the Efficient Frontier sheet", BLACK),
 ("  GREEN = formulas that read the Daily Returns sheet", GREEN),
 ("", BLACK),
 ("What the sheet does", BOLD),
 ("  1. Expected return, SD and correlation are computed from daily returns.", BLACK),
 ("  2. For each weight pair: E[rp]=wA*EA+wB*EB.", BLACK),
 ("  3. Risk: sd_p = sqrt(wA^2 sdA^2 + wB^2 sdB^2 + 2 wA wB rho sdA sdB).", BLACK),
 ("  4. The scatter (risk on X, return on Y) traces the Markowitz shape.", BLACK),
 ("  5. The yellow block is the analytical minimum-variance portfolio.", BLACK),
 ("", BLACK),
 ("Exercises", BOLD),
 ("  1. Read Weight A down from 100% to 0% and watch risk bend LEFT before", BLACK),
 ("     it rises -- diversification lowers risk below a straight line.", BLACK),
 ("  2. Confirm the lowest-risk table row sits near the Weight A* value.", BLACK),
 ("  3. Set WEIGHT_STEP=0.05 in the script for a smoother curve; re-run.", BLACK),
 ("  4. Swap in a HIGH-correlation pair and see the bend almost disappear.", BLACK),
 ("  5. Mark the efficient half: only points on/above the min-variance point.", BLACK),
]
for i, (t, f) in enumerate(lines, start=1):
    ins.cell(row=i, column=1, value=t).font = f
ins.column_dimensions["A"].width = 72

# =============================================================================
# STEP 8 : SAVE
# =============================================================================
os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
wb.save(OUTPUT_FILE)
print("Efficient-frontier workbook created:", OUTPUT_FILE)
