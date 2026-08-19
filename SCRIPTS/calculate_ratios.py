import glob as _glob, os as _os
for _p in _glob.glob(_os.path.expanduser(r"~\AppData\Local\Packages\PythonSoftwareFoundation.Python.3*\LocalCache\local-packages\Python3*\site-packages")) + _glob.glob(_os.path.expanduser(r"~\AppData\Roaming\Python\Python3*\site-packages")):
    import sys as _sys
    if _p not in _sys.path: _sys.path.insert(0, _p)
"""
calculate_ratios.py
===================
Reads 01_bank_financials.xlsx + 04_operating_metrics.xlsx.
Calculates all Dataset 02 ratios and writes 02_bank_ratios.xlsx.

Usage:
    cd SCRIPTS
    python calculate_ratios.py
"""

import sys, os
sys.path.insert(0, os.path.dirname(__file__))

import pandas as pd
import numpy as np
from banks import BANK_CODES, FISCAL_YEARS

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "DATA")


def load_financials():
    path = os.path.join(DATA, "01_bank_financials.xlsx")
    bs   = pd.read_excel(path, sheet_name="balance_sheet",   dtype={"bank_code": str, "fy": int})
    inc  = pd.read_excel(path, sheet_name="income_statement", dtype={"bank_code": str, "fy": int})
    return bs, inc


def load_operating():
    path = os.path.join(DATA, "04_operating_metrics.xlsx")
    if not os.path.exists(path):
        print("  [WARN] 04_operating_metrics.xlsx not found — skipping operational ratios.")
        return pd.DataFrame()
    op = pd.read_excel(path, sheet_name="operating_metrics", dtype={"bank_code": str, "fy": int})
    return op


def avg_balance(df, col, key="bank_code"):
    """Compute (t + t-1) / 2 average balance, NaN for first observation."""
    df = df.sort_values([key, "fy"])
    df[f"avg_{col}"] = df.groupby(key)[col].transform(
        lambda x: (x + x.shift(1)) / 2
    )
    return df


def safe_div(a, b):
    """Element-wise division returning NaN where denominator is 0 or NaN."""
    with np.errstate(divide="ignore", invalid="ignore"):
        result = np.where((b == 0) | pd.isnull(b), np.nan, a / b)
    return pd.Series(result, index=a.index if hasattr(a, "index") else None)


def yoy_growth(df, col, key="bank_code"):
    df = df.sort_values([key, "fy"])
    df[f"{col}_growth"] = df.groupby(key)[col].pct_change() * 100
    return df


def calculate():
    print("Loading source data...")
    bs, inc = load_financials()
    op      = load_operating()

    # Merge balance sheet + income statement on (bank_code, fy)
    df = pd.merge(bs, inc, on=["bank_code", "bank_name", "fy"], how="outer", suffixes=("_bs", "_is"))

    # -----------------------------------------------------------------------
    # Average balances (for ROA, ROE, NIM)
    # -----------------------------------------------------------------------
    for col in ["total_assets", "shareholders_equity", "gross_loans", "net_loans",
                "total_deposits"]:
        if col in df.columns:
            df = avg_balance(df, col)

    # -----------------------------------------------------------------------
    # Profitability
    # -----------------------------------------------------------------------
    if "profit_after_tax" in df.columns:
        df["roa"] = safe_div(df["profit_after_tax"], df.get("avg_total_assets", df.get("total_assets"))) * 100
        df["roe"] = safe_div(df["profit_after_tax"], df.get("avg_shareholders_equity", df.get("shareholders_equity"))) * 100

    if "net_interest_income" in df.columns:
        avg_earning = df.get("avg_gross_loans", df.get("gross_loans"))  # proxy; refine if investment data added
        df["nim"] = safe_div(df["net_interest_income"], avg_earning) * 100

    if all(c in df.columns for c in ["profit_after_tax", "operating_income"]):
        df["profit_margin"] = safe_div(df["profit_after_tax"], df["operating_income"]) * 100

    # -----------------------------------------------------------------------
    # Efficiency
    # -----------------------------------------------------------------------
    if all(c in df.columns for c in ["operating_expenses", "operating_income"]):
        df["cost_income"] = safe_div(df["operating_expenses"], df["operating_income"]) * 100

    # -----------------------------------------------------------------------
    # Growth (YoY %)
    # -----------------------------------------------------------------------
    for col in ["total_assets", "gross_loans", "total_deposits", "shareholders_equity", "profit_after_tax"]:
        if col in df.columns:
            df = yoy_growth(df, col)
    # Rename growth columns
    rename_map = {
        "total_assets_growth":      "asset_growth",
        "gross_loans_growth":       "loan_growth",
        "total_deposits_growth":    "deposit_growth",
        "shareholders_equity_growth":"equity_growth",
        "profit_after_tax_growth":  "profit_growth",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # -----------------------------------------------------------------------
    # Liquidity / Funding
    # -----------------------------------------------------------------------
    if all(c in df.columns for c in ["gross_loans", "total_deposits"]):
        df["loan_deposit_ratio"] = safe_div(df["gross_loans"], df["total_deposits"]) * 100

    # -----------------------------------------------------------------------
    # Operational ratios — merge from 04
    # -----------------------------------------------------------------------
    if not op.empty and all(c in op.columns for c in ["employees", "branches"]):
        op_sub = op[["bank_code", "fy", "employees", "branches",
                      "gross_npl_pct", "net_npl_pct", "provision_coverage_pct",
                      "car_pct", "cet1_pct"]].copy()
        df = pd.merge(df, op_sub, on=["bank_code", "fy"], how="left")

        df["assets_per_employee"]    = safe_div(df["total_assets"],    df["employees"])
        df["loans_per_employee"]     = safe_div(df["gross_loans"],     df["employees"])
        df["deposits_per_employee"]  = safe_div(df["total_deposits"],  df["employees"])
        df["profit_per_employee"]    = safe_div(df["profit_after_tax"],df["employees"])
        df["assets_per_branch"]      = safe_div(df["total_assets"],    df["branches"])

    # -----------------------------------------------------------------------
    # Build output columns
    # -----------------------------------------------------------------------
    ratio_cols = [
        "bank_code", "bank_name", "fy",
        "roa", "roe", "nim", "profit_margin",
        "asset_growth", "loan_growth", "deposit_growth", "equity_growth", "profit_growth",
        "cost_income",
        "assets_per_employee", "loans_per_employee", "deposits_per_employee",
        "profit_per_employee", "assets_per_branch",
        "gross_npl_pct", "net_npl_pct", "provision_coverage_pct",
        "car_pct", "cet1_pct",
        "loan_deposit_ratio",
    ]
    out_cols = [c for c in ratio_cols if c in df.columns]
    out = df[out_cols].copy()
    out = out.sort_values(["bank_code", "fy"]).reset_index(drop=True)

    # -----------------------------------------------------------------------
    # Validation flags
    # -----------------------------------------------------------------------
    flags = []
    if "roa" in out.columns:
        bad_roa = out[(out["roa"].notna()) & ((out["roa"] < -5) | (out["roa"] > 15))]
        if not bad_roa.empty:
            flags.append(f"  [FLAG] ROA out of range [-5%, 15%] for {len(bad_roa)} rows")
    if "roe" in out.columns:
        bad_roe = out[(out["roe"].notna()) & ((out["roe"] < -30) | (out["roe"] > 60))]
        if not bad_roe.empty:
            flags.append(f"  [FLAG] ROE out of range [-30%, 60%] for {len(bad_roe)} rows")
    if "cost_income" in out.columns:
        bad_ci = out[(out["cost_income"].notna()) & ((out["cost_income"] < 10) | (out["cost_income"] > 120))]
        if not bad_ci.empty:
            flags.append(f"  [FLAG] Cost/income out of range [10%, 120%] for {len(bad_ci)} rows")

    # -----------------------------------------------------------------------
    # Write output
    # -----------------------------------------------------------------------
    out_path = os.path.join(DATA, "02_bank_ratios.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="ratios", index=False)
        if flags:
            pd.DataFrame({"validation_flags": flags}).to_excel(
                writer, sheet_name="validation_flags", index=False
            )

    print(f"  Written: {out_path}")
    print(f"  Rows: {len(out)} | Columns: {len(out.columns)}")
    if flags:
        print("\nValidation flags:")
        for f in flags:
            print(f)
    else:
        print("  All validation checks passed.")


if __name__ == "__main__":
    print("=== calculate_ratios.py ===\n")
    calculate()

