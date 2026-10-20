#!/usr/bin/env python3
r"""
===============================================================================
Financial Analytics using Python
Script 10 : Monte Carlo Efficient Frontier  (1000 Random Portfolios)
Author    : Vipul Mehta
===============================================================================

PURPOSE
-------
Step 9 drew the frontier for TWO stocks by hand-picking weights. Here we let the
computer explore the whole opportunity set for ALL stocks in the data file by
generating 1000 RANDOM weight combinations, computing each portfolio's
annualised return and volatility, and scattering them. The upper-left EDGE of
the resulting cloud is the efficient frontier.

METHOD (exactly the loop-and-append recipe)
-----------------------------------------------------------------------
    for each of 1000 portfolios:
        w  = random numbers, one per stock
        w /= sum(w)                     # make the weights add to 100%
        portfolio return     = w . mean_annual_returns
        portfolio volatility = sqrt( w . AnnualCovariance . w )
        append return and volatility to lists
    build a DataFrame of Return / Volatility (+ the weights and Sharpe)
    inspect with head() and tail()
    scatter-plot volatility (x) vs return (y)

(The same result can be produced with one vectorised matrix expression, but the
explicit loop makes each portfolio's construction visible to students.)

INPUT  : Yahoo_Daily_Prices_2016_2026.csv
OUTPUT : Charts/MonteCarlo_Frontier.png
         MonteCarlo_Portfolios.csv   (the simulated cloud, for inspection)
===============================================================================
"""

# =============================================================================
# IMPORT LIBRARIES
# =============================================================================
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

# =============================================================================
# USER SETTINGS
# =============================================================================
FOLDER = r"D:\Trimester 4\Finance\Financial Analytics\Financial Analytics Project\02_Output"
NUM_PORTFOLIOS = 1000
ANNUALIZATION_FACTOR = 250
RISK_FREE = 0.00          # annual risk-free rate for the Sharpe colouring
SEED = 42                 # fixes the random draw so the whole class sees the same cloud

INPUT_FILE  = os.path.join(FOLDER, "Yahoo_Daily_Prices_2016_2026.csv")
chart_dir   = os.path.join(FOLDER, "Charts")
os.makedirs(chart_dir, exist_ok=True)
CHART_FILE  = os.path.join(chart_dir, "MonteCarlo_Frontier.png")
SIM_CSV     = os.path.join(FOLDER, "MonteCarlo_Portfolios.csv")

# =============================================================================
# STEP 1 : READ PRICES AND KEEP ADJUSTED CLOSE FOR ALL STOCKS
# =============================================================================
if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(INPUT_FILE)

df = pd.read_csv(INPUT_FILE)
date_col = df.columns[0]

adj_cols = [c for c in df.columns if "Adj Close" in c]
if not adj_cols:
    raise ValueError("No 'Adj Close' columns found -- re-run Script 01.")

prices = df[[date_col] + adj_cols].copy()
prices.columns = ["Date"] + [c.replace("_Adj Close", "").strip() for c in adj_cols]
prices["Date"] = pd.to_datetime(prices["Date"])

# =============================================================================
# STEP 2 : DAILY SIMPLE RETURNS  ->  ANNUALISED MEAN AND COVARIANCE
# =============================================================================
returns = prices.set_index("Date").pct_change().dropna()
stocks  = list(returns.columns)
n_assets = len(stocks)

mean_annual = returns.mean() * ANNUALIZATION_FACTOR              # vector (per stock)
cov_annual  = returns.cov()  * ANNUALIZATION_FACTOR              # n x n matrix

mean_vec = mean_annual.values
cov_mat  = cov_annual.values

print(f"Stocks ({n_assets}): {stocks}")
print(f"Simulating {NUM_PORTFOLIOS} random portfolios...\n")

# =============================================================================
# STEP 3 : SIMULATE RANDOM PORTFOLIOS  (loop + append)
# =============================================================================
np.random.seed(SEED)

returns_list    = []      # portfolio annual returns
volatility_list = []      # portfolio annual volatilities
weights_list    = []      # the weight vectors themselves

for i in range(NUM_PORTFOLIOS):
    w = np.random.random(n_assets)      # one random number per stock
    w /= np.sum(w)                      # normalise so weights add to 1 (100%)

    port_return     = np.dot(w, mean_vec)                 # w . mean
    port_volatility = np.sqrt(np.dot(w, np.dot(cov_mat, w)))   # sqrt(w . Cov . w)

    returns_list.append(port_return)
    volatility_list.append(port_volatility)
    weights_list.append(w)

# Convert the list of weight arrays into one 2-D NumPy array (rows = portfolios).
weights_arr = np.array(weights_list)

# =============================================================================
# STEP 4 : BUILD THE RESULTS DATAFRAME
# =============================================================================
sim = pd.DataFrame({
    "Return":     returns_list,
    "Volatility": volatility_list,
})
# Add one weight column per stock so each row is a full portfolio.
for j, s in enumerate(stocks):
    sim[s] = weights_arr[:, j]

# Sharpe ratio = (return - risk-free) / volatility  (used for colour + highlights)
sim["Sharpe"] = (sim["Return"] - RISK_FREE) / sim["Volatility"]

# Inspect the top and bottom of the simulation.
print("HEAD of the simulated portfolios:")
print(sim.head(), "\n")
print("TAIL of the simulated portfolios:")
print(sim.tail(), "\n")

# Two portfolios worth naming.
min_vol = sim.loc[sim["Volatility"].idxmin()]      # lowest-risk simulated portfolio
max_shp = sim.loc[sim["Sharpe"].idxmax()]          # best risk-adjusted portfolio
print("Minimum-volatility portfolio:")
print(min_vol[stocks].apply(lambda x: f"{x:.1%}").to_string(), "\n")
print("Maximum-Sharpe portfolio:")
print(max_shp[stocks].apply(lambda x: f"{x:.1%}").to_string(), "\n")

sim.to_csv(SIM_CSV, index=False)

# =============================================================================
# STEP 5 : SCATTER PLOT (the frontier cloud)
# =============================================================================
fig, ax = plt.subplots(figsize=(11, 7))

sc = ax.scatter(sim["Volatility"], sim["Return"],
                c=sim["Sharpe"], cmap="viridis", s=12, alpha=0.75)

cbar = fig.colorbar(sc)
cbar.set_label("Sharpe ratio", fontsize=11)

# Highlight the two special portfolios.
ax.scatter(min_vol["Volatility"], min_vol["Return"],
           marker="*", s=350, color="red", edgecolor="black",
           label="Minimum volatility", zorder=5)
ax.scatter(max_shp["Volatility"], max_shp["Return"],
           marker="*", s=350, color="gold", edgecolor="black",
           label="Maximum Sharpe", zorder=5)

ax.set_title(f"Monte Carlo Efficient Frontier  ({NUM_PORTFOLIOS} random portfolios)",
             fontsize=14, fontweight="bold")
ax.set_xlabel("Volatility  (annualised standard deviation)", fontsize=12)
ax.set_ylabel("Expected Return  (annualised)", fontsize=12)
ax.xaxis.set_major_formatter(mtick.PercentFormatter(xmax=1.0))
ax.yaxis.set_major_formatter(mtick.PercentFormatter(xmax=1.0))
ax.grid(alpha=0.30)
ax.legend(loc="best", fontsize=10)

fig.tight_layout()
fig.savefig(CHART_FILE, dpi=300, bbox_inches="tight")
plt.close(fig)

print("Saved chart :", CHART_FILE)
print("Saved cloud :", SIM_CSV)

# =============================================================================
# MINI TEACHING MODULE
# =============================================================================
# WHAT STUDENTS SHOULD SEE
#   * The cloud has a bullet / umbrella shape.
#   * Its upper-left BOUNDARY is the efficient frontier: for any risk level, the
#     highest-return portfolio; for any return, the lowest-risk portfolio.
#   * The red star (min volatility) sits at the leftmost tip.
#   * The gold star (max Sharpe) is where a line from the risk-free rate is
#     steepest -- the tangency portfolio.
#
# EXERCISES
#   1. Raise NUM_PORTFOLIOS to 10000 -- the frontier edge sharpens.
#   2. Change SEED and confirm the frontier edge barely moves (it is structural).
#   3. Set RISK_FREE to a realistic rate (e.g. 0.065) and watch max-Sharpe shift.
#   4. Compare the two-stock curve from Script 09 with this N-stock cloud.
