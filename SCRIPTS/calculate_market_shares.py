"""
calculate_market_shares.py — Computes market shares, HHI, and CR-N concentration metrics.
"""

import os
import sys

import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

from config.config_loader import load_config

CFG = load_config()
DATA = os.path.join(BASE, "DATA")
ANALYSIS = os.path.join(BASE, "ANALYSIS")


def safe_div(a, b):
    with np.errstate(divide="ignore", invalid="ignore"):
        return np.where((b == 0) | pd.isnull(b), np.nan, a / b)


def compute_hhi(shares_series: pd.Series, convention: str = "percent_squared") -> float:
    """
    Computes Herfindahl-Hirschman Index.
    If convention == "percent_squared": scale 0 - 10,000 (shares sum to 100).
    If convention == "decimal": scale 0 - 1 (shares sum to 1).
    """
    clean_s = shares_series.dropna()
    if clean_s.empty:
        return 0.0
    if convention == "decimal":
        decimal_shares = clean_s / 100.0 if clean_s.sum() > 2.0 else clean_s
        return round(float((decimal_shares ** 2).sum()), 4)
    else:
        # percent squared
        pct_shares = clean_s * 100.0 if clean_s.sum() < 2.0 else clean_s
        return round(float((pct_shares ** 2).sum()), 1)


def calc_shares(bs, inc):
    industry = bs.groupby("fy")[["total_assets", "gross_loans", "total_deposits"]].sum().rename(
        columns={"total_assets": "ind_assets", "gross_loans": "ind_loans", "total_deposits": "ind_deposits"}
    )
    inc_totals = inc.groupby("fy")[["profit_after_tax"]].sum().rename(columns={"profit_after_tax": "ind_profit"})
    industry = industry.join(inc_totals)

    df = bs.merge(inc[["bank_code", "fy", "profit_after_tax"]], on=["bank_code", "fy"], how="left")
    df = df.merge(industry.reset_index(), on="fy", how="left")

    df["asset_share_pct"] = safe_div(df["total_assets"], df["ind_assets"]) * 100
    df["loan_share_pct"] = safe_div(df["gross_loans"], df["ind_loans"]) * 100
    df["deposit_share_pct"] = safe_div(df["total_deposits"], df["ind_deposits"]) * 100
    df["profit_share_pct"] = safe_div(df["profit_after_tax"], df["ind_profit"]) * 100

    df["asset_rank"] = df.groupby("fy")["total_assets"].rank(ascending=False, method="min").astype("Int64")
    df["loan_rank"] = df.groupby("fy")["gross_loans"].rank(ascending=False, method="min").astype("Int64")
    df["deposit_rank"] = df.groupby("fy")["total_deposits"].rank(ascending=False, method="min").astype("Int64")
    df["profit_rank"] = df.groupby("fy")["profit_after_tax"].rank(ascending=False, method="min").astype("Int64")

    out_cols = [
        "bank_code", "bank_name", "fy",
        "asset_share_pct", "loan_share_pct", "deposit_share_pct", "profit_share_pct",
        "asset_rank", "loan_rank", "deposit_rank", "profit_rank"
    ]
    return df[[c for c in out_cols if c in df.columns]].sort_values(["fy", "asset_rank"])


def calc_concentration(shares):
    hhi_conv = CFG.get("hhi_convention", "percent_squared")
    unconc_max = CFG.get("hhi_thresholds", {}).get("unconcentrated_max", 1000)
    mod_max = CFG.get("hhi_thresholds", {}).get("moderately_concentrated_max", 1800)

    rows = []
    for fy, grp in shares.groupby("fy"):
        row = {"fy": fy}
        for metric, col in [("assets", "asset_share_pct"), ("loans", "loan_share_pct"), ("deposits", "deposit_share_pct")]:
            s = grp[col].dropna().sort_values(ascending=False)
            if s.empty:
                continue
            row[f"hhi_{metric}"] = compute_hhi(s, hhi_conv)
            row[f"cr4_{metric}"] = round(float(s.head(4).sum()), 2)
            row[f"cr5_{metric}"] = round(float(s.head(5).sum()), 2)
            row[f"cr10_{metric}"] = round(float(s.head(10).sum()), 2)

        row["n_banks"] = len(grp)
        hhi_a = row.get("hhi_assets", 0)
        if hhi_a < unconc_max:
            row["interpretation"] = "Unconcentrated"
        elif hhi_a < mod_max:
            row["interpretation"] = "Moderately concentrated"
        else:
            row["interpretation"] = "Highly concentrated"
        rows.append(row)

    return pd.DataFrame(rows).sort_values("fy")


def run():
    bs = pd.read_excel(os.path.join(DATA, "01_bank_financials.xlsx"), sheet_name="balance_sheet")
    inc = pd.read_excel(os.path.join(DATA, "01_bank_financials.xlsx"), sheet_name="income_statement")

    shares = calc_shares(bs, inc)
    conc = calc_concentration(shares)

    out_path = os.path.join(DATA, "03_market_shares.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        shares.to_excel(writer, sheet_name="market_shares", index=False)
        conc.to_excel(writer, sheet_name="concentration", index=False)

    os.makedirs(ANALYSIS, exist_ok=True)
    ind_path = os.path.join(ANALYSIS, "industry_structure.xlsx")
    with pd.ExcelWriter(ind_path, engine="openpyxl") as writer:
        conc.to_excel(writer, sheet_name="industry_structure", index=False)
        shares.to_excel(writer, sheet_name="bank_shares_detail", index=False)

    print(f"  [Market Shares & Concentration Calculated] {ind_path}")
    return shares, conc


if __name__ == "__main__":
    run()
