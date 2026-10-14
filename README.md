# Iron Ore Price Forecasting: Marginal Cost Curve + VECM

Two-year-ahead iron ore price forecasts combining a seaborne marginal cost curve with vector error correction models (VECMs). NPSC3000 research project.

## Files

| File | What it is |
|---|---|
| `NPSC3000_RawData.xlsx` | Input data: supply, costs, demand, prices, macro and government forecasts. Read only, never changed. |
| `iron_ore_forecasting.py` | The model. One script that runs the whole analysis. |
| `Iron_Ore_Results.docx` | Output: every result table and chart. |
| `Iron_Ore_Dashboard.html` | Output: interactive dashboard. Opens in any browser, works offline. |

## How to run

1. Install Python 3.10 or newer.
2. Install the packages:
   ```
   pip install pandas numpy scipy statsmodels==0.15.0 openpyxl matplotlib python-docx plotly
   ```
3. Put `iron_ore_forecasting.py` and `NPSC3000_RawData.xlsx` in the same folder.
4. Run:
   ```
   python iron_ore_forecasting.py
   ```
   It takes about 15 seconds and writes the `.docx` and `.html` files to the same folder.

## What the model does

1. Builds the annual cost curve (2010–2035) and finds the marginal producer and marginal cost.
2. Tests six variables for stationarity (ADF and KPSS).
3. Runs Johansen cointegration tests, with the Reinsel–Ahn correction, on all 15 VECM specifications.
4. Backtests every model recursively, one and two years ahead, over 2021–2025.
5. Forecasts 2026 (from data to 2024) and 2027 (from data to 2025), plus a demand sensitivity.

## Key settings

- Estimation sample: 2010–2025 (16 annual observations)
- Price: 62% Fe CFR China
- VECM: one lagged difference, rank 1, constant restricted to the long-run relationship
- `statsmodels==0.15.0` is pinned because Model 2's 2021 backtest forecast is numerically sensitive. All other results are the same across versions.

## Using the dashboard

On the **Cost curve** tab:

- Move the year, demand and Tier 1–3 supply sliders. The marginal producer, marginal cost and calibrated price update live.
- Click any block to see that mine's details.

The other tabs cover the balance table, backtest, forecasts, model selection and diagnostics.
