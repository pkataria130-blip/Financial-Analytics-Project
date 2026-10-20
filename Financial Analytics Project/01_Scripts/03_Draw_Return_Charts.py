# =============================================================================
# SCRIPT 03 : Draw Price and Return Charts
# AUTHOR    : Vipul Mehta
#
# PURPOSE
# -------
# From the Adjusted Close prices, create publication-quality charts:
#   Figure 1 : all stocks' daily adjusted closing prices (one combined chart)
#   Figure 2 : all stocks' SCALED prices, base = 100 on each stock's first day
#              (this is the fair way to compare performance across price levels)
#   Figure 3 : daily log returns, one chart per stock
#
# Comparing Figure 1 with Figure 2 is the lesson: raw prices mislead because a
# Rs 3,000 stock and a Rs 300 stock are not comparable in levels; rebasing to 100
# shows GROWTH, which is what an investor actually earns.
#
# INPUT  : Yahoo_Daily_Prices_2016_2026.csv
# OUTPUT : Financial Analytics/Charts/
#              Combined_Daily_Prices.png
#              Combined_Scaled_Prices.png
#              <TICKER>/<TICKER>_LogReturns.png   (one folder per stock)
# =============================================================================

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# =============================================================================
# SETTINGS
# =============================================================================
FOLDER = r"D:\Trimester 4\Finance\Financial Analytics Project\02_Output_Assignment_1"

input_file   = os.path.join(FOLDER, "Yahoo_Daily_Prices_2016_2026.csv")
chart_folder = os.path.join(FOLDER, "Charts")
os.makedirs(chart_folder, exist_ok=True)

# =============================================================================
# READ PRICES AND KEEP ADJUSTED CLOSE
# =============================================================================
df = pd.read_csv(input_file)
df["Date"] = pd.to_datetime(df["Date"])

adj_cols = [c for c in df.columns if "Adj Close" in c]
prices = df[["Date"] + adj_cols].copy()
prices.rename(columns=lambda x: x.replace("_Adj Close", ""), inplace=True)

stock_list = list(prices.columns[1:])
print(f"Stocks found: {stock_list}\n")

# =============================================================================
# FIGURE 1 : COMBINED DAILY ADJUSTED CLOSING PRICES
# =============================================================================
fig, ax = plt.subplots(figsize=(12, 6))
for stock in stock_list:
    ax.plot(prices["Date"], prices[stock], linewidth=1.3, label=stock)
ax.set_title("Daily Adjusted Closing Prices", fontsize=14, fontweight="bold")
ax.set_xlabel("Date"); ax.set_ylabel("Adjusted Closing Price")
ax.legend(loc="best", fontsize=9); ax.grid(alpha=0.30)
fig.tight_layout()
fig.savefig(os.path.join(chart_folder, "Combined_Daily_Prices.png"),
            dpi=300, bbox_inches="tight")
plt.close(fig)
print("Saved: Combined_Daily_Prices.png")

# =============================================================================
# FIGURE 2 : COMBINED SCALED PRICES  (base = 100 on each stock's first day)
# =============================================================================
fig, ax = plt.subplots(figsize=(12, 6))
for stock in stock_list:
    stock_data = prices[["Date", stock]].dropna()
    scaled = stock_data[stock] / stock_data[stock].iloc[0] * 100
    ax.plot(stock_data["Date"], scaled, linewidth=1.3, label=stock)
ax.set_title("Scaled Price Performance (Base = 100)", fontsize=14, fontweight="bold")
ax.set_xlabel("Date"); ax.set_ylabel("Index (100 = First Trading Day)")
ax.legend(loc="best", fontsize=9); ax.grid(alpha=0.30)
fig.tight_layout()
fig.savefig(os.path.join(chart_folder, "Combined_Scaled_Prices.png"),
            dpi=300, bbox_inches="tight")
plt.close(fig)
print("Saved: Combined_Scaled_Prices.png")

# =============================================================================
# FIGURE 3 : DAILY LOG RETURNS  (one chart per stock)
# =============================================================================
for stock in stock_list:
    stock_folder = os.path.join(chart_folder, stock)
    os.makedirs(stock_folder, exist_ok=True)

    data = prices[["Date", stock]].dropna().reset_index(drop=True)
    log_return = np.log(data[stock] / data[stock].shift(1))

    fig, ax = plt.subplots(figsize=(11, 4.5))
    ax.plot(data["Date"], log_return, linewidth=0.8)
    ax.axhline(y=0, color="red", linestyle="--", linewidth=1)
    ax.set_title(f"{stock} Daily Log Returns", fontsize=13)
    ax.set_xlabel("Date"); ax.set_ylabel("Log Return")
    ax.grid(alpha=0.30)

    # Auto-scale y-axis to this stock's data (don't inherit a previous plot).
    y_min, y_max = log_return.min(), log_return.max()
    margin = (y_max - y_min) * 0.05 if pd.notna(y_max - y_min) else 0.01
    ax.set_ylim(y_min - margin, y_max + margin)

    fig.tight_layout()
    fig.savefig(os.path.join(stock_folder, f"{stock}_LogReturns.png"),
                dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"Saved: {stock}/{stock}_LogReturns.png")

print("\nAll charts created successfully.")
