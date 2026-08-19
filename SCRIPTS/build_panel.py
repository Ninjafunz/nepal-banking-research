import glob as _glob, os as _os
for _p in _glob.glob(_os.path.expanduser(r"~\AppData\Local\Packages\PythonSoftwareFoundation.Python.3*\LocalCache\local-packages\Python3*\site-packages")) + _glob.glob(_os.path.expanduser(r"~\AppData\Roaming\Python\Python3*\site-packages")):
    import sys as _sys
    if _p not in _sys.path: _sys.path.insert(0, _p)
"""
build_panel.py
==============
Master pipeline. Merges all available datasets into:
    MASTER/master_bank_panel.xlsx

Stage-aware: gracefully skips files not yet populated.
Automatically calls calculate_ratios.py and calculate_market_shares.py
before merging if the raw financials file exists.

Usage:
    cd SCRIPTS
    python build_panel.py
"""

import sys, os, subprocess
sys.path.insert(0, os.path.dirname(__file__))

import pandas as pd
import numpy as np
from banks import BANKS, BANK_CODES, FISCAL_YEARS, MACRO_YEARS

BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA     = os.path.join(BASE, "DATA")
MASTER   = os.path.join(BASE, "MASTER")
SCRIPTS  = os.path.join(BASE, "SCRIPTS")


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------

def file_has_data(path, sheet=None):
    """Return True if the file exists and has at least one data row."""
    if not os.path.exists(path):
        return False
    try:
        df = pd.read_excel(path, sheet_name=sheet or 0, nrows=1)
        return not df.empty
    except Exception:
        return False


def load_if_ready(path, sheet=None, dtype_map=None, label=""):
    """Load an Excel sheet, return None with a message if not ready."""
    dtype_map = dtype_map or {"bank_code": str, "fy": int}
    if not file_has_data(path, sheet):
        print(f"  [SKIP] {label or os.path.basename(path)} — file empty or missing.")
        return None
    df = pd.read_excel(path, sheet_name=sheet or 0, dtype=dtype_map)
    print(f"  [LOAD] {label or os.path.basename(path)} — {len(df)} rows")
    return df


def run_script(script_name):
    script_path = os.path.join(SCRIPTS, script_name)
    result = subprocess.run(
        [sys.executable, script_path],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"  [ERROR] {script_name} failed:\n{result.stderr}")
    else:
        print(f"  [OK] {script_name} completed.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def build():
    os.makedirs(MASTER, exist_ok=True)
    print("=" * 60)
    print("Nepal Banking Research — Master Panel Builder")
    print("=" * 60)

    # -----------------------------------------------------------------------
    # Step 1: Regenerate derived files
    # -----------------------------------------------------------------------
    fin_path = os.path.join(DATA, "01_bank_financials.xlsx")
    if file_has_data(fin_path):
        print("\n[Step 1] Regenerating derived datasets from 01_bank_financials.xlsx...")
        run_script("calculate_ratios.py")
        run_script("calculate_market_shares.py")
    else:
        print("\n[Step 1] 01_bank_financials.xlsx has no data yet — skipping derived calculations.")

    # -----------------------------------------------------------------------
    # Step 2: Load all available datasets
    # -----------------------------------------------------------------------
    print("\n[Step 2] Loading datasets...")

    # Stage 1 — Core
    bs  = load_if_ready(fin_path, "balance_sheet",   label="01 balance_sheet")
    inc = load_if_ready(fin_path, "income_statement", label="01 income_statement")
    rat = load_if_ready(os.path.join(DATA, "02_bank_ratios.xlsx"),   "ratios",        label="02 ratios")
    ms  = load_if_ready(os.path.join(DATA, "03_market_shares.xlsx"), "market_shares", label="03 market_shares")

    # Stage 2 — Strategic depth
    op  = load_if_ready(os.path.join(DATA, "04_operating_metrics.xlsx"), "operating_metrics", label="04 operating")
    lc  = load_if_ready(os.path.join(DATA, "05_loan_composition.xlsx"),  "loan_composition",  label="05 loan_comp")
    dc  = load_if_ready(os.path.join(DATA, "06_deposit_composition.xlsx"),"deposit_composition",label="06 deposit_comp")

    # Stage 3 — Macro
    macro = load_if_ready(os.path.join(DATA, "07_macro_indicators.xlsx"), "macro_indicators",
                          dtype_map={"fy": int}, label="07 macro")

    # Stage 4 — Market
    mkt = load_if_ready(os.path.join(DATA, "08_market_data.xlsx"), "market_data", label="08 market_data")

    # Stage 5 — Qualitative
    dig  = load_if_ready(os.path.join(DATA, "09_strategic_coding.xlsx"), "digital_scorecard",    label="09 digital")
    strat= load_if_ready(os.path.join(DATA, "09_strategic_coding.xlsx"), "strategic_priorities", label="09 strategic")

    # -----------------------------------------------------------------------
    # Step 3: Build base skeleton (all bank x year combos)
    # -----------------------------------------------------------------------
    print("\n[Step 3] Building base skeleton...")
    bank_df = pd.DataFrame([
        {"bank_code": b["code"], "bank_name": b["name"],
         "listed": b["listed"], "state_owned": b["state_owned"]}
        for b in BANKS
    ])
    fy_df = pd.DataFrame({"fy": FISCAL_YEARS})
    panel = bank_df.assign(key=1).merge(fy_df.assign(key=1), on="key").drop("key", axis=1)
    print(f"  Base skeleton: {len(panel)} rows ({len(BANKS)} banks x {len(FISCAL_YEARS)} years)")

    # -----------------------------------------------------------------------
    # Step 4: Merge all datasets
    # -----------------------------------------------------------------------
    print("\n[Step 4] Merging datasets...")

    def merge_on(left, right, label, drop_cols=None):
        if right is None:
            return left
        drop_cols = drop_cols or ["bank_name"]
        right = right.drop(columns=[c for c in drop_cols if c in right.columns and c != "bank_code" and c != "fy"], errors="ignore")
        # Drop source/notes columns to keep panel clean
        right = right.drop(columns=[c for c in right.columns if c in ["source","notes","source_bs","source_is"]], errors="ignore")
        merged = pd.merge(left, right, on=["bank_code", "fy"], how="left")
        n_new = len(merged.columns) - len(left.columns)
        print(f"  + {label}: added {n_new} columns")
        return merged

    # Balance sheet (key columns only to avoid bloat)
    if bs is not None:
        bs_cols = ["bank_code", "fy", "total_assets", "cash_bank_balances", "investments",
                   "gross_loans", "net_loans", "total_deposits", "borrowings",
                   "total_liabilities", "shareholders_equity", "paid_up_capital", "reserves"]
        panel = merge_on(panel, bs[[c for c in bs_cols if c in bs.columns]], "balance_sheet")

    # Income statement
    if inc is not None:
        inc_cols = ["bank_code", "fy", "interest_income", "interest_expense", "net_interest_income",
                    "non_interest_income", "operating_income", "operating_expenses",
                    "personnel_expenses", "provision_loan_losses", "profit_before_tax", "profit_after_tax"]
        panel = merge_on(panel, inc[[c for c in inc_cols if c in inc.columns]], "income_statement")

    # Ratios
    if rat is not None:
        panel = merge_on(panel, rat, "ratios")

    # Market shares
    if ms is not None:
        ms_cols = ["bank_code", "fy", "asset_share_pct", "loan_share_pct", "deposit_share_pct",
                   "profit_share_pct", "asset_rank", "loan_rank", "deposit_rank", "profit_rank"]
        panel = merge_on(panel, ms[[c for c in ms_cols if c in ms.columns]], "market_shares")

    # Operating metrics
    if op is not None:
        op_cols = ["bank_code", "fy", "branches", "employees", "atms", "extension_counters",
                   "branchless_banking_centers", "mobile_banking_users", "internet_banking_users",
                   "debit_cards", "credit_cards", "qr_users", "digital_transactions_count",
                   "gross_npl_pct", "net_npl_pct", "provision_coverage_pct", "car_pct", "cet1_pct"]
        panel = merge_on(panel, op[[c for c in op_cols if c in op.columns]], "operating_metrics")

    # Loan composition
    if lc is not None:
        lc_drop = ["total_loans_check"]
        panel = merge_on(panel, lc.drop(columns=[c for c in lc_drop if c in lc.columns], errors="ignore"),
                         "loan_composition")

    # Deposit composition
    if dc is not None:
        dc_drop = ["total_deposits_check"]
        panel = merge_on(panel, dc.drop(columns=[c for c in dc_drop if c in dc.columns], errors="ignore"),
                         "deposit_composition")

    # Macro (join on fy only)
    if macro is not None:
        panel = pd.merge(panel, macro.drop(columns=["source","notes"], errors="ignore"),
                         on="fy", how="left")
        print(f"  + macro_indicators: added {len(macro.columns)-1} columns")

    # Market data
    if mkt is not None:
        mkt_cols = ["bank_code", "fy", "ticker", "share_price_eoy", "market_cap",
                    "pe_ratio", "pb_ratio", "eps", "bvps", "dividend_per_share",
                    "dividend_yield_pct", "annual_return_pct", "price_volatility"]
        panel = merge_on(panel, mkt[[c for c in mkt_cols if c in mkt.columns]], "market_data")

    # Digital scorecard
    if dig is not None:
        dig_cols = ["bank_code", "fy", "digital_account_opening", "mobile_banking", "digital_lending",
                    "qr_ecosystem", "api_open_banking", "ai_initiatives", "digital_customer_acquisition",
                    "core_banking_upgrade", "fintech_partnership", "digital_index"]
        panel = merge_on(panel, dig[[c for c in dig_cols if c in dig.columns]], "digital_scorecard")

    # Strategic priorities
    if strat is not None:
        strat_cols = ["bank_code", "fy", "priority_retail", "priority_sme", "priority_corporate",
                      "priority_digital", "priority_branch_expansion", "priority_cost_reduction",
                      "priority_wealth_mgmt", "priority_remittance", "priority_sustainability",
                      "strategic_score"]
        panel = merge_on(panel, strat[[c for c in strat_cols if c in strat.columns]], "strategic_priorities")

    panel = panel.sort_values(["bank_code", "fy"]).reset_index(drop=True)

    # -----------------------------------------------------------------------
    # Step 5: Data coverage audit
    # -----------------------------------------------------------------------
    print("\n[Step 5] Computing data coverage...")
    coverage = pd.DataFrame({
        "column":    panel.columns,
        "n_filled":  [panel[c].notna().sum() for c in panel.columns],
        "pct_filled":[round(panel[c].notna().mean() * 100, 1) for c in panel.columns],
        "stage":     ["Key" if c in ["bank_code","bank_name","fy","listed","state_owned"] else
                      "Stage1" if c in ["total_assets","gross_loans","total_deposits","shareholders_equity",
                                        "profit_after_tax","roa","roe","nim","asset_share_pct","hhi_assets"] else
                      "Stage2" if c in ["branches","employees","gross_npl_pct","casa_ratio"] else
                      "Stage3" if c.startswith("gdp") or c.startswith("inflation") or c.startswith("system_") else
                      "Stage4" if c in ["pb_ratio","pe_ratio","market_cap"] else
                      "Stage5" if c in ["digital_index","strategic_score"] else "Other"
                      for c in panel.columns]
    })

    # -----------------------------------------------------------------------
    # Step 6: Write master panel
    # -----------------------------------------------------------------------
    print("\n[Step 6] Writing master panel...")
    out_path = os.path.join(MASTER, "master_bank_panel.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        panel.to_excel(writer, sheet_name="panel", index=False)
        coverage.to_excel(writer, sheet_name="data_coverage", index=False)

    print(f"\n{'='*60}")
    print(f"  master_bank_panel.xlsx written.")
    print(f"  Rows:    {len(panel)}")
    print(f"  Columns: {len(panel.columns)}")
    print(f"  Path:    {out_path}")

    # Summary by stage
    for stage in ["Stage1","Stage2","Stage3","Stage4","Stage5"]:
        stage_cov = coverage[coverage["stage"] == stage]["pct_filled"].mean()
        if not pd.isna(stage_cov):
            print(f"  {stage} avg coverage: {stage_cov:.0f}%")
    print("="*60)


if __name__ == "__main__":
    build()

