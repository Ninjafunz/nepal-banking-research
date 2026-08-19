import glob as _glob, os as _os
for _p in _glob.glob(_os.path.expanduser(r"~\AppData\Local\Packages\PythonSoftwareFoundation.Python.3*\LocalCache\local-packages\Python3*\site-packages")) + _glob.glob(_os.path.expanduser(r"~\AppData\Roaming\Python\Python3*\site-packages")):
    import sys as _sys
    if _p not in _sys.path: _sys.path.insert(0, _p)

"""
fetch_macro.py
==============
Pulls Nepal macro data from World Bank API and NRB-sourced estimates.
Writes directly into 07_macro_indicators.xlsx.

Usage:
    python fetch_macro.py
"""

import sys, os, json, urllib.request
sys.path.insert(0, os.path.dirname(__file__))

import pandas as pd
import numpy as np

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "DATA")

# ─────────────────────────────────────────────────────────────────────────────
# Nepal fiscal year → calendar year mapping
# Nepal FY ends mid-July. World Bank uses calendar year.
# FY2020 (mid-Jul 2019 – mid-Jul 2020) → WB year 2020
# FY2021 → 2021, FY2022 → 2022, FY2023 → 2023, FY2024 → 2024, FY2025 → 2025
# ─────────────────────────────────────────────────────────────────────────────
FY_MAP = {2020: 2020, 2021: 2021, 2022: 2022,
          2023: 2023, 2024: 2024, 2025: 2025, 2026: 2026}

MACRO_YEARS = list(range(2020, 2027))

WB_INDICATORS = {
    "gdp_growth_pct":             "NY.GDP.MKTP.KD.ZG",
    "inflation_pct":              "FP.CPI.TOTL.ZG",
    "remittance_usd_mn":          "BX.TRF.PWKR.CD.DT",   # current USD → convert to mn
    "avg_lending_rate_pct":       "FR.INR.LEND",
    "avg_deposit_rate_pct":       "FR.INR.DPST",
    "gdp_current_usd_bn":         "NY.GDP.MKTP.CD",       # current USD → convert to bn
    "private_credit_pct_gdp":     "FS.AST.PRVT.GD.ZS",
}

def wb_fetch(indicator, country="NPL", mrv=10):
    url = (f"https://api.worldbank.org/v2/country/{country}/indicator/"
           f"{indicator}?format=json&per_page={mrv}&mrv={mrv}")
    try:
        with urllib.request.urlopen(url, timeout=15) as r:
            data = json.loads(r.read())
        return {int(item["date"]): item["value"] for item in data[1] if item["value"] is not None}
    except Exception as e:
        print(f"  [WARN] Could not fetch {indicator}: {e}")
        return {}

def build_macro():
    print("Fetching World Bank data for Nepal...")
    raw = {}
    for field, indicator in WB_INDICATORS.items():
        series = wb_fetch(indicator)
        raw[field] = series
        n = len([v for v in series.values() if v is not None])
        print(f"  {field}: {n} years fetched")

    rows = []
    for fy in MACRO_YEARS:
        cal_yr = FY_MAP[fy]
        row = {"fy": fy}

        row["gdp_growth_pct"]         = raw["gdp_growth_pct"].get(cal_yr)
        row["inflation_pct"]          = raw["inflation_pct"].get(cal_yr)
        remit_usd                      = raw["remittance_usd_mn"].get(cal_yr)
        row["remittance_usd_mn"]      = round(remit_usd / 1e6, 1) if remit_usd else None
        row["avg_lending_rate_pct"]   = raw["avg_lending_rate_pct"].get(cal_yr)
        row["avg_deposit_rate_pct"]   = raw["avg_deposit_rate_pct"].get(cal_yr)
        row["private_credit_pct_gdp"] = raw["private_credit_pct_gdp"].get(cal_yr)
        gdp_usd                        = raw["gdp_current_usd_bn"].get(cal_yr)
        row["gdp_current_usd_bn"]     = round(gdp_usd / 1e9, 2) if gdp_usd else None
        row["source"]                 = "World Bank API"
        row["notes"]                  = f"WB calendar year {cal_yr} mapped to Nepal FY{fy}"
        rows.append(row)

    df = pd.DataFrame(rows)

    # ── NRB-sourced estimates (from NRB Annual Reports / published data) ──────
    # These are well-documented published figures for Nepal
    # Sources: NRB Annual Reports, Monetary Policy documents
    nrb_data = {
        # NRB policy rate (bank rate / repo rate) — from NRB Monetary Policy
        "policy_rate_pct": {
            2020: 5.00,   # NRB repo rate FY2020
            2021: 3.00,   # Reduced due to COVID
            2022: 5.50,   # Tightening cycle began
            2023: 7.00,   # Peak tightening
            2024: 5.50,   # Easing started
            2025: 5.00,   # Continued easing
            2026: None,
        },
        "bank_rate_pct": {
            2020: 6.00,
            2021: 5.00,
            2022: 7.00,
            2023: 8.50,
            2024: 7.00,
            2025: 6.50,
            2026: None,
        },
        # NRB Banking & Financial Statistics — system-level (approx from annual reports)
        "n_commercial_banks": {
            2020: 27, 2021: 27, 2022: 26, 2023: 22, 2024: 20, 2025: 20, 2026: None
        },
    }
    for col, year_map in nrb_data.items():
        df[col] = df["fy"].map(year_map)

    # Reorder columns to match template
    col_order = [
        "fy", "gdp_growth_pct", "inflation_pct",
        "remittance_usd_mn", "private_credit_pct_gdp",
        "gdp_current_usd_bn",
        "policy_rate_pct", "bank_rate_pct",
        "avg_deposit_rate_pct", "avg_lending_rate_pct",
        "n_commercial_banks",
        "source", "notes"
    ]
    df = df[[c for c in col_order if c in df.columns]]

    out_path = os.path.join(DATA, "07_macro_indicators.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl", mode="a",
                        if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="macro_indicators", index=False)

    print(f"\n  Written: {out_path}")
    print(f"  Rows: {len(df)} | Columns: {len(df.columns)}")
    print("\n  Preview:")
    print(df[["fy","gdp_growth_pct","inflation_pct","policy_rate_pct","n_commercial_banks"]].to_string(index=False))

if __name__ == "__main__":
    print("=== fetch_macro.py ===\n")
    build_macro()
