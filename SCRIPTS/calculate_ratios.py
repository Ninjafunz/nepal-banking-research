"""
calculate_ratios.py — Derives financial & efficiency ratios from raw bank financials.
"""

import os
import sys

import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

from config.config_loader import load_config
from SCRIPTS.validators import validate_ratios

CFG = load_config()
DATA = os.path.join(BASE, "DATA")


def safe_div(a, b):
    """Safe division returning NaN for zero, negative or missing denominators."""
    a_ser = pd.Series(a) if not isinstance(a, pd.Series) else a
    b_ser = pd.Series(b) if not isinstance(b, pd.Series) else b
    with np.errstate(divide="ignore", invalid="ignore"):
        res = np.where((b_ser == 0) | pd.isnull(b_ser), np.nan, a_ser / b_ser)
    return pd.Series(res, index=a_ser.index)


def avg_balance(df, col, key="bank_code"):
    """Computes opening/closing average balance: (t + t-1) / 2."""
    df = df.sort_values([key, "fy"]).copy()
    if CFG.get("average_balance_method", True):
        df[f"avg_{col}"] = df.groupby(key)[col].transform(lambda x: (x + x.shift(1)) / 2)
    else:
        df[f"avg_{col}"] = df[col]
    return df


def yoy_growth(df, col, key="bank_code"):
    df = df.sort_values([key, "fy"]).copy()
    df[f"{col}_growth"] = df.groupby(key)[col].pct_change() * 100
    return df


def calculate():
    bs_path = os.path.join(DATA, "01_bank_financials.xlsx")
    op_path = os.path.join(DATA, "04_operating_metrics.xlsx")

    if not os.path.exists(bs_path):
        print(f"  [ERROR] {bs_path} not found.")
        return

    bs = pd.read_excel(bs_path, sheet_name="balance_sheet")
    inc = pd.read_excel(bs_path, sheet_name="income_statement")
    op = pd.read_excel(op_path, sheet_name="operating_metrics") if os.path.exists(op_path) else pd.DataFrame()

    df = pd.merge(bs, inc, on=["bank_code", "bank_name", "fy"], how="outer", suffixes=("_bs", "_is"))

    # Average balances
    for col in ["total_assets", "shareholders_equity", "gross_loans", "net_loans", "total_deposits"]:
        if col in df.columns:
            df = avg_balance(df, col)

    # Profitability
    if "profit_after_tax" in df.columns:
        df["roa"] = safe_div(df["profit_after_tax"], df.get("avg_total_assets", df.get("total_assets"))) * 100
        df["roe"] = safe_div(df["profit_after_tax"], df.get("avg_shareholders_equity", df.get("shareholders_equity"))) * 100

    if "net_interest_income" in df.columns:
        avg_earning = df.get("avg_gross_loans", df.get("gross_loans"))
        df["nim"] = safe_div(df["net_interest_income"], avg_earning) * 100

    if all(c in df.columns for c in ["profit_after_tax", "operating_income"]):
        df["profit_margin"] = safe_div(df["profit_after_tax"], df["operating_income"]) * 100

    # Efficiency
    if all(c in df.columns for c in ["operating_expenses", "operating_income"]):
        df["cost_income"] = safe_div(df["operating_expenses"], df["operating_income"]) * 100

    # Growth YoY
    for col in ["total_assets", "gross_loans", "total_deposits", "shareholders_equity", "profit_after_tax"]:
        if col in df.columns:
            df = yoy_growth(df, col)

    df = df.rename(columns={
        "total_assets_growth": "asset_growth",
        "gross_loans_growth": "loan_growth",
        "total_deposits_growth": "deposit_growth",
        "shareholders_equity_growth": "equity_growth",
        "profit_after_tax_growth": "profit_growth"
    })

    if all(c in df.columns for c in ["gross_loans", "total_deposits"]):
        df["loan_deposit_ratio"] = safe_div(df["gross_loans"], df["total_deposits"]) * 100

    # Operational metrics
    if not op.empty and all(c in op.columns for c in ["employees", "branches"]):
        op_sub = op[["bank_code", "fy", "employees", "branches", "gross_npl_pct", "net_npl_pct", "provision_coverage_pct", "car_pct", "cet1_pct"]].copy()
        df = pd.merge(df, op_sub, on=["bank_code", "fy"], how="left")
        df["assets_per_employee"] = safe_div(df["total_assets"], df["employees"])
        df["loans_per_employee"] = safe_div(df["gross_loans"], df["employees"])
        df["deposits_per_employee"] = safe_div(df["total_deposits"], df["employees"])
        df["profit_per_employee"] = safe_div(df["profit_after_tax"], df["employees"])
        df["assets_per_branch"] = safe_div(df["total_assets"], df["branches"])

    ratio_cols = [
        "bank_code", "bank_name", "fy", "roa", "roe", "nim", "profit_margin",
        "asset_growth", "loan_growth", "deposit_growth", "equity_growth", "profit_growth",
        "cost_income", "assets_per_employee", "loans_per_employee", "deposits_per_employee",
        "profit_per_employee", "assets_per_branch", "gross_npl_pct", "net_npl_pct",
        "provision_coverage_pct", "car_pct", "cet1_pct", "loan_deposit_ratio"
    ]
    out = df[[c for c in ratio_cols if c in df.columns]].sort_values(["bank_code", "fy"]).reset_index(drop=True)

    # Validation
    _, flags = validate_ratios(out)

    out_path = os.path.join(DATA, "02_bank_ratios.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out.to_excel(writer, sheet_name="ratios", index=False)
        if flags:
            pd.DataFrame({"validation_flags": flags}).to_excel(writer, sheet_name="validation_flags", index=False)

    print(f"  [Ratios Calculated] {out_path} ({len(out)} rows)")
    return out


if __name__ == "__main__":
    calculate()
