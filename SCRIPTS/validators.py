"""
validators.py — Automated Data Sanity, Accounting & Regulatory Rule Validator.
"""


import numpy as np
import pandas as pd

from config.config_loader import load_config

CFG = load_config()
RULES = CFG.get("validation_rules", {})


def validate_balance_sheet(df: pd.DataFrame) -> tuple[bool, list[str]]:
    """
    Checks:
      1. total_assets > 0 for active banks
      2. total_assets > shareholders_equity > 0
      3. total_assets ≈ total_liabilities + shareholders_equity (within tolerance)
    """
    flags = []
    tol_pct = RULES.get("balance_sheet_tolerance_pct", 1.0) / 100.0

    # Inactive post-merger entities
    merged_inactive = {
        "BOKL": [2024, 2025, 2026],
        "CIVIL": [2024, 2025, 2026],
        "CCBL": [2024, 2025, 2026]
    }

    for idx, row in df.iterrows():
        b_code = row.get("bank_code", f"row_{idx}")
        fy = row.get("fy", "Unknown")
        assets = row.get("total_assets", np.nan)
        equity = row.get("shareholders_equity", np.nan)
        liab = row.get("total_liabilities", np.nan)

        # Skip known inactive post-merger years
        if b_code in merged_inactive and fy in merged_inactive[b_code]:
            continue

        if pd.isna(assets) or assets <= 0:
            flags.append(f"[BS ERROR] {b_code} FY{fy}: total_assets ({assets}) must be positive for active bank.")
            continue

        if pd.notna(equity):
            if equity <= 0:
                flags.append(f"[BS WARNING] {b_code} FY{fy}: shareholders_equity ({equity}) is non-positive.")
            elif equity >= assets:
                flags.append(f"[BS ERROR] {b_code} FY{fy}: equity ({equity}) >= total_assets ({assets}).")

        if pd.notna(liab) and pd.notna(equity) and pd.notna(assets):
            expected = liab + equity
            diff_pct = abs(assets - expected) / assets
            if diff_pct > tol_pct:
                flags.append(f"[BS MISMATCH] {b_code} FY{fy}: Assets ({assets}) != Liab+Equity ({expected}) [diff: {diff_pct*100:.2f}%]")

    is_valid = len([f for f in flags if "ERROR" in f]) == 0
    return is_valid, flags


def validate_ratios(df: pd.DataFrame) -> tuple[bool, list[str]]:
    flags = []
    roa_min = RULES.get("roa_min", -10.0)
    roa_max = RULES.get("roa_max", 15.0)
    roe_min = RULES.get("roe_min", -50.0)
    roe_max = RULES.get("roe_max", 60.0)
    nim_min = RULES.get("nim_min", 0.0)
    nim_max = RULES.get("nim_max", 12.0)
    car_min = RULES.get("car_min_regulatory", 8.5)

    if "roa" in df.columns:
        bad_roa = df[(df["roa"].notna()) & ((df["roa"] < roa_min) | (df["roa"] > roa_max))]
        for _, row in bad_roa.iterrows():
            flags.append(f"[RATIO OUTLIER] {row.get('bank_code')} FY{row.get('fy')}: ROA ({row['roa']:.2f}%) outside [{roa_min}%, {roa_max}%]")

    if "roe" in df.columns:
        bad_roe = df[(df["roe"].notna()) & ((df["roe"] < roe_min) | (df["roe"] > roe_max))]
        for _, row in bad_roe.iterrows():
            flags.append(f"[RATIO OUTLIER] {row.get('bank_code')} FY{row.get('fy')}: ROE ({row['roe']:.2f}%) outside [{roe_min}%, {roe_max}%]")

    if "nim" in df.columns:
        bad_nim = df[(df["nim"].notna()) & ((df["nim"] < nim_min) | (df["nim"] > nim_max))]
        for _, row in bad_nim.iterrows():
            flags.append(f"[RATIO OUTLIER] {row.get('bank_code')} FY{row.get('fy')}: NIM ({row['nim']:.2f}%) outside [{nim_min}%, {nim_max}%]")

    if "car_pct" in df.columns:
        breach_car = df[(df["car_pct"].notna()) & (df["car_pct"] < car_min)]
        for _, row in breach_car.iterrows():
            flags.append(f"[REGULATORY BREACH] {row.get('bank_code')} FY{row.get('fy')}: CAR ({row['car_pct']:.2f}%) below NRB Basel III minimum ({car_min}%)")

    is_valid = len(flags) == 0
    return is_valid, flags


def validate_panel_completeness(df: pd.DataFrame, expected_banks: list[str], expected_years: list[int]) -> tuple[bool, list[str]]:
    flags = []
    dupes = df[df.duplicated(subset=["bank_code", "fy"], keep=False)]
    if not dupes.empty:
        flags.append(f"[PANEL DUPLICATE] Found {len(dupes)} duplicate (bank_code, fy) records.")
    return len(flags) == 0, flags


def run_all_validations(panel_df: pd.DataFrame) -> dict:
    bs_valid, bs_flags = validate_balance_sheet(panel_df)
    r_valid, r_flags = validate_ratios(panel_df)
    all_flags = bs_flags + r_flags
    return {
        "is_valid": bs_valid and r_valid,
        "total_flags": len(all_flags),
        "flags": all_flags
    }
