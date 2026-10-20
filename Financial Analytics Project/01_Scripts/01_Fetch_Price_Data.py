# =============================================================================
# SCRIPT 01 : Fetch Daily Price Data from Yahoo Finance
#
# PURPOSE
# -------
# Downloads daily OHLC + Adjusted Close + Volume for a list of tickers and saves
# them to one flat CSV that every later script reads.
#
# IMPORTANT (yfinance behaviour)
# ------------------------------
# Since yfinance 0.2.28+, yf.download() defaults to auto_adjust=True, which
# REMOVES the 'Adj Close' column (Close becomes the adjusted series). Our whole
# course computes returns from 'Adj Close', so we set auto_adjust=False
# EXPLICITLY below to keep that column. Also keep yfinance up to date
# (pip install -U yfinance) as older versions throw 'rate limit' errors.
#
# Install once if needed:
#   pip install -U yfinance pandas
# =============================================================================

import yfinance as yf
import pandas as pd
import os


# =============================================================================
# SET YOUR FOLDER PATH BELOW

FOLDER = r"D:\Trimester 4\Finance\Financial Analytics Project\02_Output_Assignment_1"
TICKERS = [
    "NVDA",
    "AAPL",
    "GOOGL",
    "MSFT"
]

start_date = "2016-08-01"
end_date   = "2026-08-01"     # end date is EXCLUSIVE, so this includes 31-Jul-2026

output_file = os.path.join(FOLDER, "Yahoo_Daily_Prices_2016_2026.csv")

# =============================================================================
# DOWNLOAD
# =============================================================================
df = yf.download(
    tickers=TICKERS,
    start=start_date,
    end=end_date,
    interval="1d",
    auto_adjust=False,        # <-- keeps the 'Adj Close' column (see note above)
    group_by="ticker",
    progress=True,
)

# Guard: an empty frame usually means a rate-limit or a bad ticker.
if df.empty:
    raise RuntimeError(
        "yfinance returned no data. Try again in a minute (rate limit), "
        "run 'pip install -U yfinance', or check the ticker symbols."
    )

# =============================================================================
# FLATTEN COLUMNS  ->  e.g. RELIANCE.NS_Adj Close
# =============================================================================
# With several tickers, columns are a MultiIndex (ticker, field). With a single
# ticker yfinance returns a flat index, so we handle both cases -- this makes the
# same script reusable for Step 4 (global indices), including a lone index.
if isinstance(df.columns, pd.MultiIndex):
    df.columns = [f"{ticker}_{field}" for ticker, field in df.columns]
else:
    only = tickers[0]
    df.columns = [f"{only}_{field}" for field in df.columns]

# Make Date a normal column.
df.reset_index(inplace=True)

# Sanity check: confirm the Adjusted Close columns we depend on are present.
adj_cols = [c for c in df.columns if "Adj Close" in c]
if not adj_cols:
    raise RuntimeError(
        "No 'Adj Close' columns found. Your yfinance likely ignored "
        "auto_adjust=False. Upgrade with 'pip install -U yfinance' and re-run."
    )

# =============================================================================
# SAVE
# =============================================================================
os.makedirs(FOLDER, exist_ok=True)
df.to_csv(output_file, index=False)

print(f"Data saved to:\n{output_file}")
print(f"Rows: {len(df)}   |   Tickers with Adj Close: {len(adj_cols)}")
print(f"Date range: {df['Date'].min().date()}  ->  {df['Date'].max().date()}")
