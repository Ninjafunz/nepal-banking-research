import glob as _glob, os as _os
for _p in _glob.glob(_os.path.expanduser(r"~\AppData\Local\Packages\PythonSoftwareFoundation.Python.3*\LocalCache\local-packages\Python3*\site-packages")) + _glob.glob(_os.path.expanduser(r"~\AppData\Roaming\Python\Python3*\site-packages")):
    import sys as _sys
    if _p not in _sys.path: _sys.path.insert(0, _p)
"""
calculate_market_shares.py
==========================
Reads 01_bank_financials.xlsx.
Calculates per-bank market shares, bank rankings, HHI, CR4 / CR5 / CR10.
Writes results to 03_market_shares.xlsx and ANALYSIS/industry_structure.xlsx.

Usage:
    cd SCRIPTS
    python calculate_market_shares.py
"""

import sys, os
sys.path.insert(0, os.path.dirname(__file__))

import pandas as pd
import numpy as np

BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA     = os.path.join(BASE, "DATA")
ANALYSIS = os.path.join(BASE, "ANALYSIS")


def load_bs():
    path = os.path.join(DATA, "01_bank_financials.xlsx")
    df = pd.read_excel(path, sheet_name="balance_sheet", dtype={"bank_code": str, "fy": int})
    return df


def load_is():
    path = os.path.join(DATA, "01_bank_financials.xlsx")
    df = pd.read_excel(path, sheet_name="income_statement", dtype={"bank_code": str, "fy": int})
    return df


def safe_div(a, b):
    with np.errstate(divide="ignore", invalid="ignore"):
        return np.where((b == 0) | pd.isnull(b), np.nan, a / b)


def calc_shares(bs, inc):
    """Calculate per-bank market shares for each year."""
    # Industry totals per year
    industry = bs.groupby("fy")[["total_assets", "gross_loans", "total_deposits"]].sum().rename(
        columns={"total_assets": "ind_assets", "gross_loans": "ind_loans", "total_deposits": "ind_deposits"}
    )
    # Profit totals
    inc_totals = inc.groupby("fy")[["profit_after_tax"]].sum().rename(
        columns={"profit_after_tax": "ind_profit"}
    )
    industry = industry.join(inc_totals)

    df = bs.merge(inc[["bank_code", "fy", "profit_after_tax"]], on=["bank_code", "fy"], how="left")
    df = df.merge(industry.reset_index(), on="fy", how="left")

    df["asset_share_pct"]   = safe_div(df["total_assets"],      df["ind_assets"]) * 100
    df["loan_share_pct"]    = safe_div(df["gross_loans"],        df["ind_loans"])  * 100
    df["deposit_share_pct"] = safe_div(df["total_deposits"],     df["ind_deposits"]) * 100
    df["profit_share_pct"]  = safe_div(df["profit_after_tax"],   df["ind_profit"])   * 100

    # Rankings (1 = largest)
    df["asset_rank"]   = df.groupby("fy")["total_assets"].rank(ascending=False, method="min").astype("Int64")
    df["loan_rank"]    = df.groupby("fy")["gross_loans"].rank(ascending=False, method="min").astype("Int64")
    df["deposit_rank"] = df.groupby("fy")["total_deposits"].rank(ascending=False, method="min").astype("Int64")
    df["profit_rank"]  = df.groupby("fy")["profit_after_tax"].rank(ascending=False, method="min").astype("Int64")

    out_cols = ["bank_code", "bank_name", "fy",
                "asset_share_pct", "loan_share_pct", "deposit_share_pct", "profit_share_pct",
                "asset_rank", "loan_rank", "deposit_rank", "profit_rank"]
    return df[[c for c in out_cols if c in df.columns]].sort_values(["fy", "asset_rank"])


def calc_concentration(shares):
    """Calculate HHI and CR-N statistics per year."""
    rows = []
    for fy, grp in shares.groupby("fy"):
        row = {"fy": fy}

        for metric, col in [("assets", "asset_share_pct"),
                             ("loans",  "loan_share_pct"),
                             ("deposits","deposit_share_pct")]:
            s = grp[col].dropna().sort_values(ascending=False)
            if s.empty:
                continue
            # HHI: sum of squared percentage shares  (max = 10,000)
            row[f"hhi_{metric}"]  = round((s ** 2).sum(), 1)
            # CR-N
            row[f"cr4_{metric}"]  = round(s.head(4).sum(),  2)
            row[f"cr5_{metric}"]  = round(s.head(5).sum(),  2)
            row[f"cr10_{metric}"] = round(s.head(10).sum(), 2)

        row["n_banks"] = len(grp)

        # Simple interpretation
        hhi_a = row.get("hhi_assets", 0)
        if hhi_a < 1000:
            row["interpretation"] = "Unconcentrated"
        elif hhi_a < 1800:
            row["interpretation"] = "Moderately concentrated"
        else:
            row["interpretation"] = "Highly concentrated"

        rows.append(row)
    return pd.DataFrame(rows).sort_values("fy")


def validate_shares(shares):
    flags = []
    for fy, grp in shares.groupby("fy"):
        for col, label in [("asset_share_pct", "assets"), ("loan_share_pct", "loans"),
                           ("deposit_share_pct", "deposits")]:
            total = grp[col].sum()
            if not (95 <= total <= 105):
                flags.append(f"  [FLAG] FY{fy} {label} market shares sum to {total:.1f}% (expected ~100%)")
    return flags


def run():
    print("Loading balance sheet and income statement...")
    bs  = load_bs()
    inc = load_is()

    print("Calculating market shares...")
    shares = calc_shares(bs, inc)

    print("Calculating concentration metrics (HHI, CR4, CR5, CR10)...")
    conc = calc_concentration(shares)

    # Validation
    flags = validate_shares(shares)

    # Write 03_market_shares.xlsx
    out_path = os.path.join(DATA, "03_market_shares.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        shares.to_excel(writer, sheet_name="market_shares", index=False)
        conc.to_excel(writer, sheet_name="concentration", index=False)
        if flags:
            pd.DataFrame({"validation_flags": flags}).to_excel(
                writer, sheet_name="validation_flags", index=False
            )
    print(f"  Written: {out_path}")

    # Also write to ANALYSIS/industry_structure.xlsx
    os.makedirs(ANALYSIS, exist_ok=True)
    ind_path = os.path.join(ANALYSIS, "industry_structure.xlsx")
    with pd.ExcelWriter(ind_path, engine="openpyxl") as writer:
        conc.to_excel(writer, sheet_name="industry_structure", index=False)
        shares.to_excel(writer, sheet_name="bank_shares_detail", index=False)
    print(f"  Written: {ind_path}")

    print(f"\nConcentration summary:")
    print(conc[["fy", "hhi_assets", "cr4_assets", "cr5_assets", "cr10_assets", "interpretation"]].to_string(index=False))

    if flags:
        print("\nValidation flags:")
        for f in flags:
            print(f)
    else:
        print("\n  All validation checks passed.")


if __name__ == "__main__":
    print("=== calculate_market_shares.py ===\n")
    run()

