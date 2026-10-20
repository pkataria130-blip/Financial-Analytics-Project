# =============================================================================
# SCRIPT 02 : Compute Daily Simple and Log Returns
#
# PURPOSE
# -------
# Reads the price file from Script 01, keeps the Adjusted Close prices, and
# computes (1) daily simple returns and (2) daily log returns, saving both to a
# single Excel workbook with two worksheets.
#
#   Simple return_t = (P_t / P_(t-1)) - 1
#   Log return_t    = ln(P_t / P_(t-1))
#
# INPUT  : Yahoo_Daily_Prices_2016_2026.csv
# OUTPUT : Daily_Returns.xlsx   (Sheet1 = Simple Returns, Sheet2 = Log Returns)
#
# TEACHING NOTE
# -------------
# The first row of every return series is blank (NaN): there is no prior price to
# compare against. Simple returns aggregate across STOCKS (portfolios); log
# returns aggregate across TIME (compounding). We use both later for that reason.
# =============================================================================

import pandas as pd
import numpy as np
import os

# =============================================================================
# SETTINGS
# =============================================================================
FOLDER = r"D:\Trimester 4\Finance\Financial Analytics Project\02_Output_Assignment_1"

input_file  = os.path.join(FOLDER, "Yahoo_Daily_Prices_2016_2026.csv")
output_file = os.path.join(FOLDER, "Daily_Returns.xlsx")

# =============================================================================
# READ DATA
# =============================================================================
df = pd.read_csv(input_file)
df["Date"] = pd.to_datetime(df["Date"])

# =============================================================================
# KEEP ONLY DATE + ADJUSTED CLOSE, THEN CLEAN THE COLUMN NAMES
# =============================================================================
adj_close_cols = [c for c in df.columns if "Adj Close" in c]
if not adj_close_cols:
    raise ValueError("No 'Adj Close' columns found -- re-run Script 01.")

prices = df[["Date"] + adj_close_cols].copy()
# 'RELIANCE.NS_Adj Close'  ->  'RELIANCE.NS'
prices.rename(columns={c: c.replace("_Adj Close", "") for c in adj_close_cols},
              inplace=True)

# Quick data sanity check (always look at your data before analysing it).
print("Stocks :", list(prices.columns[1:]))
print("Rows   :", len(prices),
      "| Dates:", prices['Date'].min().date(), "->", prices['Date'].max().date())
late = prices.iloc[:, 1:].apply(lambda s: s.first_valid_index())
if late.nunique() > 1:
    print("NOTE: stocks have different start dates (later listings):")
    for stock, idx in late.items():
        print(f"   {stock:15s} first price on {prices.loc[idx, 'Date'].date()}")

# =============================================================================
# SIMPLE DAILY RETURNS
# =============================================================================
simple_returns = prices.copy()
simple_returns.iloc[:, 1:] = prices.iloc[:, 1:].pct_change()

# =============================================================================
# LOG RETURNS
# =============================================================================
log_returns = prices.copy()
log_returns.iloc[:, 1:] = np.log(prices.iloc[:, 1:] / prices.iloc[:, 1:].shift(1))

# =============================================================================
# SAVE BOTH SHEETS
# =============================================================================
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    simple_returns.to_excel(writer, sheet_name="Simple Returns", index=False)
    log_returns.to_excel(writer, sheet_name="Log Returns", index=False)

print("=" * 60)
print("Daily returns saved to:\n", output_file)
print("=" * 60)
