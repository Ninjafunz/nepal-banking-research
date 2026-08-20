"""
build_panel.py — Master Panel Builder & Multi-Format Research Exporter (.xlsx, .csv, .parquet, .dta).
"""

import os
import sys

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

import SCRIPTS.calculate_market_shares as cms
import SCRIPTS.calculate_ratios as cr
from config.config_loader import load_config
from SCRIPTS.banks import BANKS, FISCAL_YEARS
from SCRIPTS.validators import run_all_validations

CFG = load_config()
DATA = os.path.join(BASE, "DATA")
MASTER = os.path.join(BASE, "MASTER")


def file_has_data(path, sheet=None):
    if not os.path.exists(path):
        return False
    try:
        df = pd.read_excel(path, sheet_name=sheet or 0, nrows=1)
        return not df.empty
    except Exception:
        return False


def load_if_ready(path, sheet=None, dtype_map=None, label=""):
    dtype_map = dtype_map or {"bank_code": str, "fy": int}
    if not file_has_data(path, sheet):
        return None
    return pd.read_excel(path, sheet_name=sheet or 0, dtype=dtype_map)


def export_multiformat(panel_df: pd.DataFrame, coverage_df: pd.DataFrame):
    os.makedirs(MASTER, exist_ok=True)
    formats = CFG.get("export_formats", ["xlsx", "csv", "parquet", "dta"])

    # 1. Excel
    if "xlsx" in formats:
        xlsx_path = os.path.join(MASTER, "master_bank_panel.xlsx")
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            panel_df.to_excel(writer, sheet_name="panel", index=False)
            coverage_df.to_excel(writer, sheet_name="data_coverage", index=False)
        print(f"  [Exported Excel]   {xlsx_path}")

    # 2. CSV
    if "csv" in formats:
        csv_path = os.path.join(MASTER, "master_bank_panel.csv")
        panel_df.to_csv(csv_path, index=False, encoding="utf-8")
        print(f"  [Exported CSV]     {csv_path}")

    # 3. Parquet (columnar format)
    if "parquet" in formats:
        try:
            parquet_path = os.path.join(MASTER, "master_bank_panel.parquet")
            # Convert any object columns to string for parquet serialization
            p_df = panel_df.copy()
            for col in p_df.select_dtypes(include=["object", "string", "str"]).columns:
                p_df[col] = p_df[col].astype(str)
            p_df.to_parquet(parquet_path, index=False, engine="pyarrow")
            print(f"  [Exported Parquet] {parquet_path}")
        except Exception as e:
            print(f"  [WARN] Parquet export failed: {e}")

    # 4. Stata (.dta)
    if "dta" in formats:
        try:
            dta_path = os.path.join(MASTER, "master_bank_panel.dta")
            stata_df = panel_df.copy()
            # Stata column names max 32 chars and alphanumeric/underscore only
            stata_df.columns = [str(c)[:30].replace("-", "_").replace(" ", "_") for c in stata_df.columns]
            for col in stata_df.select_dtypes(include=["object", "string", "str"]).columns:
                stata_df[col] = stata_df[col].astype(str).str.slice(0, 244)
            stata_df.to_stata(dta_path, write_index=False, version=117)
            print(f"  [Exported Stata]   {dta_path}")
        except Exception as e:
            print(f"  [WARN] Stata export failed: {e}")


def build():
    print("=" * 65)
    print("Nepal Banking Research — Master Panel Builder & Exporter")
    print("=" * 65)

    fin_path = os.path.join(DATA, "01_bank_financials.xlsx")
    if file_has_data(fin_path):
        print("\n[1/4] Running Derived Calculus (Ratios & Market Shares)...")
        cr.calculate()
        cms.run()

    print("\n[2/4] Loading Raw & Derived Datasets...")
    bs = load_if_ready(fin_path, "balance_sheet")
    inc = load_if_ready(fin_path, "income_statement")
    rat = load_if_ready(os.path.join(DATA, "02_bank_ratios.xlsx"), "ratios")
    ms = load_if_ready(os.path.join(DATA, "03_market_shares.xlsx"), "market_shares")
    op = load_if_ready(os.path.join(DATA, "04_operating_metrics.xlsx"), "operating_metrics")
    lc = load_if_ready(os.path.join(DATA, "05_loan_composition.xlsx"), "loan_composition")
    dc = load_if_ready(os.path.join(DATA, "06_deposit_composition.xlsx"), "deposit_composition")
    macro = load_if_ready(os.path.join(DATA, "07_macro_indicators.xlsx"), "macro_indicators", dtype_map={"fy": int})
    mkt = load_if_ready(os.path.join(DATA, "08_market_data.xlsx"), "market_data")
    dig = load_if_ready(os.path.join(DATA, "09_strategic_coding.xlsx"), "digital_scorecard")
    strat = load_if_ready(os.path.join(DATA, "09_strategic_coding.xlsx"), "strategic_priorities")

    # Base skeleton
    bank_df = pd.DataFrame([
        {"bank_code": b["code"], "bank_name": b["name"],
         "listed": b["listed"], "state_owned": b["state_owned"]}
        for b in BANKS
    ])
    fy_df = pd.DataFrame({"fy": FISCAL_YEARS})
    panel = bank_df.assign(key=1).merge(fy_df.assign(key=1), on="key").drop("key", axis=1)

    def merge_on(left, right, drop_cols=None):
        if right is None:
            return left
        drop_cols = drop_cols or ["bank_name"]
        right = right.drop(columns=[c for c in drop_cols if c in right.columns and c not in ["bank_code", "fy"]], errors="ignore")
        right = right.drop(columns=[c for c in right.columns if c in ["source", "notes", "source_bs", "source_is"]], errors="ignore")
        return pd.merge(left, right, on=["bank_code", "fy"], how="left")

    if bs is not None:
        bs_cols = ["bank_code", "fy", "total_assets", "cash_bank_balances", "investments", "gross_loans", "net_loans", "total_deposits", "borrowings", "total_liabilities", "shareholders_equity", "paid_up_capital", "reserves"]
        panel = merge_on(panel, bs[[c for c in bs_cols if c in bs.columns]])
    if inc is not None:
        inc_cols = ["bank_code", "fy", "interest_income", "interest_expense", "net_interest_income", "non_interest_income", "operating_income", "operating_expenses", "personnel_expenses", "provision_loan_losses", "profit_before_tax", "profit_after_tax"]
        panel = merge_on(panel, inc[[c for c in inc_cols if c in inc.columns]])
    if rat is not None:
        panel = merge_on(panel, rat)
    if ms is not None:
        ms_cols = ["bank_code", "fy", "asset_share_pct", "loan_share_pct", "deposit_share_pct", "profit_share_pct", "asset_rank", "loan_rank", "deposit_rank", "profit_rank"]
        panel = merge_on(panel, ms[[c for c in ms_cols if c in ms.columns]])
    if op is not None:
        # NPL/CAR columns come from ratios (which already includes them); only merge operational metrics
        op_cols = ["bank_code", "fy", "branches", "employees", "atms", "extension_counters", "branchless_banking_centers", "mobile_banking_users", "internet_banking_users", "debit_cards", "credit_cards", "qr_users", "digital_transactions_count"]
        panel = merge_on(panel, op[[c for c in op_cols if c in op.columns]])
    if lc is not None:
        panel = merge_on(panel, lc.drop(columns=["total_loans_check"], errors="ignore"))
    if dc is not None:
        panel = merge_on(panel, dc.drop(columns=["total_deposits_check"], errors="ignore"))
    if macro is not None:
        panel = pd.merge(panel, macro.drop(columns=["source", "notes"], errors="ignore"), on="fy", how="left")
    if mkt is not None:
        mkt_cols = ["bank_code", "fy", "ticker", "share_price_eoy", "market_cap", "pe_ratio", "pb_ratio", "eps", "bvps", "dividend_per_share", "dividend_yield_pct", "annual_return_pct", "price_volatility"]
        panel = merge_on(panel, mkt[[c for c in mkt_cols if c in mkt.columns]])
    if dig is not None:
        dig_cols = ["bank_code", "fy", "digital_account_opening", "mobile_banking", "digital_lending", "qr_ecosystem", "api_open_banking", "ai_initiatives", "digital_customer_acquisition", "core_banking_upgrade", "fintech_partnership", "digital_index"]
        panel = merge_on(panel, dig[[c for c in dig_cols if c in dig.columns]])
    if strat is not None:
        strat_cols = ["bank_code", "fy", "priority_retail", "priority_sme", "priority_corporate", "priority_digital", "priority_branch_expansion", "priority_cost_reduction", "priority_wealth_mgmt", "priority_remittance", "priority_sustainability", "strategic_score"]
        panel = merge_on(panel, strat[[c for c in strat_cols if c in strat.columns]])

    panel = panel.sort_values(["bank_code", "fy"]).reset_index(drop=True)

    print("\n[3/4] Running Automated Sanity & Validation Module...")
    v_report = run_all_validations(panel)
    if v_report["flags"]:
        print(f"  [Validation Note] Found {v_report['total_flags']} advisory flags:")
        for f in v_report["flags"][:5]:
            print(f"    {f}")
        if len(v_report["flags"]) > 5:
            print(f"    ... and {len(v_report['flags']) - 5} more.")
    else:
        print("  [Validation Passed] All balance sheet, ratio, and regulatory checks valid.")

    print("\n[4/4] Multi-Format Exporting...")
    coverage = pd.DataFrame({
        "column": panel.columns,
        "n_filled": [panel[c].notna().sum() for c in panel.columns],
        "pct_filled": [round(panel[c].notna().mean() * 100, 1) for c in panel.columns]
    })
    export_multiformat(panel, coverage)

    print("=" * 65)
    print(f"Master Panel Built: {len(panel)} rows × {len(panel.columns)} columns")
    print("=" * 65)
    return panel


if __name__ == "__main__":
    build()

