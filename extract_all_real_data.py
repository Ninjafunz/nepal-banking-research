#!/usr/bin/env python3
"""
extract_all_real_data.py
========================
Reads NRB supervision report CSVs and populates all 10 DATA/ Excel files
with authentic data. Missing data is left as NaN (honest, not synthetic).

Pipeline schema (what build expects):
  01_bank_financials.xlsx  → balance_sheet, income_statement
  02_bank_ratios.xlsx      → ratios  (calculated by pipeline)
  03_market_shares.xlsx    → market_shares, concentration  (calculated by pipeline)
  04_operating_metrics.xlsx→ operating_metrics
  05_loan_composition.xlsx → loan_composition
  06_deposit_composition.xlsx → deposit_composition
  07_macro_indicators.xlsx → macro_indicators
  08_market_data.xlsx      → market_data
  09_strategic_coding.xlsx → digital_scorecard, strategic_priorities
  10_bank_events.xlsx      → events
"""

import math
import os
import re
from collections import defaultdict

import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "DATA")

# ═══════════════════════════════════════════════════════════════════════════
# 1. BANK REGISTRY – canonical codes & names
# ═══════════════════════════════════════════════════════════════════════════
BANK_REGISTRY = {
    # code: (canonical_name, is_state_owned, is_listed)
    "NBL":    ("Nepal Bank Limited", True, True),
    "RBB":    ("Rastriya Banijya Bank", True, False),
    "ADBL":   ("Agriculture Development Bank", True, True),
    "NABIL":  ("Nabil Bank", False, True),
    "NIMB":   ("Nepal Investment Mega Bank", False, True),
    "SCB":    ("Standard Chartered Bank Nepal", False, True),
    "HBL":    ("Himalayan Bank", False, True),
    "SBI":    ("Nepal SBI Bank", False, True),
    "EBL":    ("Everest Bank", False, True),
    "BOKL":   ("Bank of Kathmandu", False, True),
    "NICA":   ("NIC Asia Bank", False, True),
    "MBL":    ("Machhapuchchhre Bank", False, True),
    "KBL":    ("Kumari Bank", False, True),
    "LLBS":   ("Laxmi Sunrise Bank", False, True),
    "CIVIL":  ("Civil Bank", False, True),
    "CCBL":   ("Century Commercial Bank", False, True),
    "SANIMA": ("Sanima Bank", False, True),
    "SBL":    ("Siddhartha Bank", False, True),
    "GIBL":   ("Global IME Bank", False, True),
    "PCBL":   ("Prime Commercial Bank", False, True),
    "PRVU":   ("Prabhu Bank", False, True),
    "CZBIL":  ("Citizens Bank International", False, True),
    "NMB":    ("NMB Bank", False, True),
}

# Post-merger: these banks cease independent operations from given FY
MERGED_INACTIVE = {
    "BOKL":  [2024, 2025, 2026],   # merged into GIBL Jan 2023 → FY2024
    "CIVIL": [2024, 2025, 2026],   # acquired by HBL Feb 2023 → FY2024
    "CCBL":  [2024, 2025, 2026],   # merged into PRVU Jan 2023 → FY2024
}

# Pre-merger names → surviving entity codes
PRE_MERGER_MAP = {
    # 2021 report names
    "Nepal Bangladesh Bank": "NABIL",    # acquired by NABIL FY2022
    "Mega Bank Nepal": "NIMB",           # merged into NIMB FY2023
    "Nepal Investment Bank": "NIMB",     # merged into NIMB FY2023
    "Sunrise Bank": "LLBS",             # merged into LLBS FY2024
    "Laxmi Bank": "LLBS",               # merged into LLBS FY2024
    "Nepal Credit and Commerce": "KBL",  # merged into KBL FY2023
    "Janata Bank": "GIBL",              # merged into GIBL FY2020
    "Bank of Kathmandu": "BOKL",        # merged into GIBL FY2024
}

# Map from report-specific names → bank_code
# Each report uses slightly different naming conventions
def map_report_name_to_code(name_str):
    """Map a bank name from any NRB report to the canonical bank_code."""
    n = name_str.strip().lower()
    # Direct matches first
    direct_map = {
        "nepal bank": "NBL",
        "nepal bank limited": "NBL",
        "rastriya banijya bank": "RBB",
        "rastriya banijya bank ltd": "RBB",
        "rastriya banijya bank limited": "RBB",
        "agricultural development bank": "ADBL",
        "agriculture development bank": "ADBL",
        "agriculture development bank ltd": "ADBL",
        "nabil bank": "NABIL",
        "nabil bank limited": "NABIL",
        "nepal investment mega bank": "NIMB",
        "nepal investment mega bank limited": "NIMB",
        "nepal investment bank": "NIMB",
        "nepal investment bank limited": "NIMB",
        "standard chartered bank nepal": "SCB",
        "standard chartered bank nepal limited": "SCB",
        "standard chartered bank": "SCB",
        "standard chartered bank limited": "SCB",
        "himalayan bank": "HBL",
        "himalayan bank limited": "HBL",
        "himlayan bank limited": "HBL",  # typo in 2025 report
        "nepal sbi bank": "SBI",
        "nepal sbi bank limited": "SBI",
        "everest bank": "EBL",
        "everest bank limited": "EBL",
        "bank of kathmandu": "BOKL",
        "bank of kathmandu limited": "BOKL",
        "nic asia bank": "NICA",
        "nic asia bank limited": "NICA",
        "machhapuchchhre bank": "MBL",
        "machhapuchchhre bank limited": "MBL",
        "machhapuchhre bank limited": "MBL",
        "kumari bank": "KBL",
        "kumari bank limited": "KBL",
        "laxmi sunrise bank": "LLBS",
        "laxmi sunrise bank limited": "LLBS",
        "laxmi bank": "LLBS",
        "laxmi bank limited": "LLBS",
        "civil bank": "CIVIL",
        "civil bank limited": "CIVIL",
        "citizens bank international": "CZBIL",
        "citizens bank international limited": "CZBIL",
        "prime commercial bank": "PCBL",
        "prime commercial bank limited": "PCBL",
        "sanima bank": "SANIMA",
        "sanima bank limited": "SANIMA",
        "siddhartha bank": "SBL",
        "siddhartha bank limited": "SBL",
        "siddharth bank": "SBL",
        "siddharth bank limited": "SBL",
        "global ime bank": "GIBL",
        "global ime bank limited": "GIBL",
        "nmb bank": "NMB",
        "nmb bank limited": "NMB",
        "prabhu bank": "PRVU",
        "prabhu bank limited": "PRVU",
        "century commercial bank": "CCBL",
        "century commercial bank limited": "CCBL",
        "nepal bangladesh bank": "NABIL",
        "nepal bangladesh bank limited": "NABIL",
        "mega bank nepal": "NIMB",
        "mega bank nepal limited": "NIMB",
        "sunrise bank": "LLBS",
        "sunrise bank limited": "LLBS",
        "nepal credit and commerce": "KBL",
        "nepal credit and commerce bank": "KBL",
        "nepal credit and commerce bank ltd": "KBL",
        "janata bank": "GIBL",
        "janata bank limited": "GIBL",
    }
    # Try exact match
    for pattern, code in direct_map.items():
        if n == pattern or n.startswith(pattern):
            return code
    # Try fuzzy: if "nepal bank" in name
    if "nepal bank" in n and "banijya" not in n and "bangladesh" not in n:
        return "NBL"
    if "banijya" in n:
        return "RBB"
    if "agricult" in n:
        return "ADBL"
    if "nabil" in n:
        return "NABIL"
    if "investment mega" in n or "investment bank" in n:
        return "NIMB"
    if "standard chartered" in n:
        return "SCB"
    if "himalayan" in n or "himlayan" in n:
        return "HBL"
    if "sbi" in n:
        return "SBI"
    if "everest" in n:
        return "EBL"
    if "bank of kathmandu" in n:
        return "BOKL"
    if "nic asia" in n or "nic bank" in n:
        return "NICA"
    if "machhapuch" in n:
        return "MBL"
    if "kumari" in n:
        return "KBL"
    if "laxmi" in n or "sunrise" in n:
        return "LLBS"
    if "civil" in n:
        return "CIVIL"
    if "citizens" in n:
        return "CZBIL"
    if "prime" in n:
        return "PCBL"
    if "sanima" in n:
        return "SANIMA"
    if "siddhar" in n:
        return "SBL"
    if "global" in n or "ime" in n:
        return "GIBL"
    if "nmb" in n:
        return "NMB"
    if "prabhu" in n:
        return "PRVU"
    if "century" in n:
        return "CCBL"
    if "bangladesh" in n:
        return "NABIL"
    if "mega bank" in n:
        return "NIMB"
    if "nepal credit" in n or "ncc" in n:
        return "KBL"
    if "janata" in n:
        return "GIBL"
    return None

# ═══════════════════════════════════════════════════════════════════════════
# 2. REPORT METADATA – table ranges, fiscal years, bank ordering
# ═══════════════════════════════════════════════════════════════════════════

# Each report: (prefix, fy_columns, bank_table_start_numbers)
# fy_columns: the fiscal years in column order (left to right)
# bank_table_start: first table number for that bank's data

REPORTS_2025 = {
    "prefix": "BANK-SUPERVISION-REPORT-2025",
    "fy_columns": [2023, 2024, 2025],  # FY 2022-23, 2023-24, 2024-25
    "banks": [
        (37, 40, "NBL"),     # tables 37-40
        (41, 44, "RBB"),     # tables 41-44
        (45, 46, "ADBL"),    # tables 45-46 (may be partial)
        (47, 48, "NABIL"),   # tables 47-48 (BS+P&L combined in 48)
        (49, 51, "NIMB"),    # tables 49-51
        (52, 55, "SCB"),     # tables 52-55
        (56, 58, "HBL"),     # tables 56-58
        (59, 60, "SBI"),     # tables 59-60
        (61, 63, "EBL"),     # tables 61-63
        (64, 66, "NICA"),    # tables 64-66
        (67, 68, "MBL"),     # tables 67-68
        (69, 70, "KBL"),     # tables 69-70
        (71, 73, "LLBS"),    # tables 71-73
        (74, 76, "SBL"),     # tables 74-76
        (77, 79, "GIBL"),    # tables 77-79
        (80, 82, "CZBIL"),   # tables 80-82
        (83, 84, "PCBL"),    # tables 83-84
        (85, 86, "NMB"),     # tables 85-86
        (87, 89, "PRVU"),    # tables 87-89
        (90, 92, "SANIMA"),  # tables 90-92
    ],
}

REPORTS_2024 = {
    "prefix": "Annual-Bank-Supervision-Report-2024-3",
    "fy_columns": [2022, 2023, 2024],  # FY 2021-22, 2022-23, 2023-24
    "banks": [
        (42, 43, "NBL"),
        (44, 45, "RBB"),
        (46, 47, "ADBL"),
        (48, 49, "NABIL"),
        (50, 51, "NIMB"),
        (52, 53, "SCB"),
        (54, 55, "HBL"),
        (56, 57, "SBI"),
        (58, 59, "EBL"),
        (60, 61, "NICA"),
        (62, 63, "MBL"),
        (64, 65, "KBL"),
        (66, 67, "LLBS"),
        (68, 69, "SBL"),
        (70, 71, "GIBL"),
        (72, 73, "CZBIL"),
        (74, 75, "PCBL"),
        (76, 77, "NMB"),
        (78, 79, "PRVU"),
        (80, 81, "SANIMA"),
    ],
}

REPORTS_2021 = {
    "prefix": "Bank-Supervision-report-2020-21-Final",
    "fy_columns": [2019, 2020, 2021],  # FY 2018-19, 2019-20, 2020-21
    "banks": [
        (39, 41, "NBL"),
        (42, 43, "ADBL"),
        (44, 45, "NABIL"),
        (46, 48, "NIMB"),    # "Nepal Investment Bank" → NIMB
        (49, 50, "SCB"),
        (51, 52, "HBL"),
        (53, 54, "SBI"),
        (55, 57, "NABIL"),   # "Nepal Bangladesh Bank" → NABIL (second entry, skip)
        (58, 59, "EBL"),
        (60, 61, "KBL"),
        (62, 63, "LLBS"),    # "Laxmi Bank"
        (64, 65, "CZBIL"),
        (66, 67, "PCBL"),
        (68, 69, "LLBS"),    # "Sunrise Bank" → LLBS (second entry, skip)
        (70, 72, "CCBL"),
        (73, 74, "SANIMA"),
        (75, 76, "MBL"),
        (77, 79, "NICA"),
        (80, 81, "GIBL"),
        (82, 83, "NMB"),
        (84, 85, "PRVU"),
        (86, 88, "SBL"),
        (89, 90, "BOKL"),
        (91, 92, "CIVIL"),
        (93, 94, "KBL"),     # "Nepal Credit and Commerce" → KBL (second entry, skip)
        (95, 96, "RBB"),
        (97, 98, "NIMB"),    # "Mega Bank Nepal" → NIMB (second entry, skip)
    ],
}

REPORTS_2019 = {
    "prefix": "BSD-Annual-Report-2019",
    "fy_columns": [2017, 2018, 2019],  # FY 2016-17, 2017-18, 2018-19
    "banks": [
        (37, 38, "NBL"),
        (39, 41, "RBB"),
        (42, 43, "NABIL"),
        (44, 45, "NIMB"),    # "Nepal Investment Bank" → NIMB
        (46, 47, "SCB"),
        (48, 50, "HBL"),
        (51, 52, "SBI"),
        (53, 55, "NABIL"),   # "Nepal Bangladesh Bank" → NABIL (skip)
        (56, 57, "EBL"),
        (58, 60, "BOKL"),
        (61, 62, "KBL"),     # "Nepal Credit and Commerce" → KBL
        (63, 65, "NICA"),
        (66, 67, "MBL"),
        (68, 70, "KBL"),
        (71, 72, "LLBS"),    # "Laxmi Bank"
        (73, 75, "SBL"),
        (76, 77, "ADBL"),
        (78, 79, "GIBL"),
        (80, 82, "NMB"),
        (83, 84, "CZBIL"),
        (85, 87, "PCBL"),
        (88, 89, "LLBS"),    # "Sunrise Bank"
        (90, 92, "CIVIL"),
        (93, 94, "PRVU"),
        (95, 97, "GIBL"),    # "Janata Bank" → GIBL (skip)
        (98, 99, "NIMB"),    # "Mega Bank Nepal" → NIMB (skip)
    ],
}

# Reports where a bank appears multiple times (pre-merger entities).
# We only want the FIRST occurrence per bank_code per report.
# The second occurrence is a pre-merger entity that maps to the same code.
REPORT_LIST = [REPORTS_2025, REPORTS_2024, REPORTS_2021, REPORTS_2019]


# ═══════════════════════════════════════════════════════════════════════════
# 3. CSV PARSING UTILITIES
# ═══════════════════════════════════════════════════════════════════════════

def parse_number(val):
    """Parse a number from NRB report format: handles commas, parens for negatives, hyphens."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return np.nan
    s = str(val).strip()
    if s in ["", "-", "--", "nan", "None"]:
        return np.nan
    # Handle parenthetical negatives: (1,234.56) → -1234.56
    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1]
        negate = True
    else:
        negate = False
    # Remove commas
    s = s.replace(",", "")
    try:
        v = float(s)
        return -v if negate else v
    except ValueError:
        return np.nan


def read_table_csv(filepath):
    """Read a CSV table, skipping the column-index header row."""
    try:
        df = pd.read_csv(filepath, header=None)
        # First row is usually column indices (0,1,2,3...) – skip it
        if len(df) > 0 and str(df.iloc[0, 0]).strip() == "0":
            df = df.iloc[1:].reset_index(drop=True)
        return df
    except Exception:
        return None


def extract_line_items_from_tables(table_dfs):
    """
    Given a list of DataFrames (tables for one bank), extract all line items
    as a dict of {item_name: [col1_val, col2_val, col3_val]}.
    """
    items = {}
    for df in table_dfs:
        if df is None or df.empty:
            continue
        for _, row in df.iterrows():
            item_name = str(row.iloc[0]).strip()
            if not item_name or item_name.lower() in ["s.n.", "s.n", "particulars", ""]:
                continue
            # Skip header-like rows (but NOT "Total Assets", "Total Liabilities" etc.)
            item_low = item_name.lower().strip()
            if any(item_low.startswith(s) or item_low == s for s in [
                "amt. in rs", "amount in rs", "statement of financial",
                "statement of profit",
            ]) or item_low in ["assets", "liabilities", "assets,,", "assets,,,",
                              "liabilities,", "liabilities,,", "liabilities,,,"]:
                continue
            vals = []
            for c in range(1, len(row)):
                vals.append(parse_number(row.iloc[c]))
            items[item_name] = vals
    return items


def get_val(items, *keywords, col_idx=0):
    """Find a line item by keywords and return value at given column index."""
    for name, vals in items.items():
        name_lower = name.lower().replace("  ", " ")
        for kw in keywords:
            if kw.lower() in name_lower:
                if col_idx < len(vals):
                    return vals[col_idx]
    return np.nan


def detect_fiscal_years_from_tables(table_dfs):
    """Try to detect fiscal year labels from the first row of tables.
    Nepal convention: '2022-23' means FY ending 2023."""
    for df in table_dfs:
        if df is None or df.empty:
            continue
        for _, row in df.iterrows():
            cell0 = str(row.iloc[0]).strip().lower()
            if "statement" in cell0 or "particulars" in cell0:
                fys = []
                for c in range(1, len(row)):
                    val = str(row.iloc[c]).strip()
                    # Match "2022-23" or "FY 2022-23" → use ending year
                    m = re.search(r"20(\d{2})[\-–](\d{2})", val)
                    if m:
                        yr = 2000 + int(m.group(2))  # Use ending year
                        fys.append(yr)
                    else:
                        # Fallback: single year like "2023"
                        m2 = re.search(r"20(\d{2})", val)
                        if m2:
                            yr = 2000 + int(m2.group(1))
                            fys.append(yr)
                        else:
                            fys.append(None)
                return fys
    return None


# ═══════════════════════════════════════════════════════════════════════════
# 4. EXTRACT DATA FROM ALL REPORTS
# ═══════════════════════════════════════════════════════════════════════════

def extract_from_report(report_meta):
    """
    Extract balance sheet + income statement data from one report.
    Returns dict: {bank_code: {fy: {line_item: value}}}
    """
    prefix = report_meta["prefix"]
    fy_cols = report_meta["fy_columns"]
    bank_ranges = report_meta["banks"]

    # Load all relevant CSVs
    all_tables = {}
    for (start, end, _) in bank_ranges:
        for t in range(start, end + 1):
            filepath = os.path.join(BASE, f"{prefix}_table_{t}.csv")
            if os.path.exists(filepath) and t not in all_tables:
                all_tables[t] = read_table_csv(filepath)

    # Process each bank
    bank_data = defaultdict(lambda: defaultdict(dict))
    seen_banks = set()  # track which bank_codes we've already extracted

    for (start, end, target_code) in bank_ranges:
        # Skip duplicate entries (pre-merger entities mapped to same code)
        if target_code in seen_banks:
            continue
        seen_banks.add(target_code)

        table_dfs = [all_tables.get(t) for t in range(start, end + 1)]

        # Detect actual FY columns from table headers
        detected_fys = detect_fiscal_years_from_tables(table_dfs)
        if detected_fys and len(detected_fys) == len(fy_cols):
            actual_fys = [f if f else fy_cols[i] for i, f in enumerate(detected_fys)]
        else:
            actual_fys = fy_cols

        # Extract all line items
        items = extract_line_items_from_tables(table_dfs)

        if not items:
            continue

        # Map line items to schema columns for each fiscal year
        for fy_idx, fy in enumerate(actual_fys):
            if fy is None or fy < 2017 or fy > 2026:
                continue

            # Balance Sheet items
            total_assets = get_val(items, "total assets", col_idx=fy_idx)
            deposits_cust = get_val(items, "deposits from customers", col_idx=fy_idx)
            equity = get_val(items, "equity", col_idx=fy_idx)
            share_capital = get_val(items, "share capital", col_idx=fy_idx)
            reserves_val = get_val(items, "reserves", col_idx=fy_idx)
            retained = get_val(items, "retained earnings", col_idx=fy_idx)
            investments = get_val(items, "investment securities", col_idx=fy_idx)
            loans_cust = get_val(items, "loans and advances to customers", col_idx=fy_idx)
            loans_bfis = get_val(items, "loans and advances to b/fis", col_idx=fy_idx,
                                  ) if not pd.isna(get_val(items, "loans and advances to b/fis", col_idx=fy_idx)) else get_val(items, "loan and advances to b/fis", col_idx=fy_idx)
            cash_equiv = get_val(items, "cash and cash equivalents", col_idx=fy_idx)
            due_nrb = get_val(items, "due from nepal rastra bank", col_idx=fy_idx)
            placements = get_val(items, "placement with bank and financial", col_idx=fy_idx)
            due_bfis = get_val(items, "due from bank and financial", col_idx=fy_idx)
            get_val(items, "property and equipment", col_idx=fy_idx)
            get_val(items, "property & equipment", col_idx=fy_idx)
            get_val(items, "other assets", col_idx=fy_idx)
            get_val(items, "goodwill", col_idx=fy_idx)
            liab_total = get_val(items, "liabilities", col_idx=fy_idx)
            get_val(items, "due to bank and financial", col_idx=fy_idx)
            borrowings_val = get_val(items, "borrowing", col_idx=fy_idx)
            get_val(items, "debt securities issued", col_idx=fy_idx)
            get_val(items, "subordinated liabilities", col_idx=fy_idx)
            get_val(items, "other liabilities", col_idx=fy_idx)
            get_val(items, "deferred tax liabilities", col_idx=fy_idx)
            provisions_val = get_val(items, "provisions", col_idx=fy_idx)
            get_val(items, "current tax liabilities", col_idx=fy_idx)
            get_val(items, "derivative financial instruments", col_idx=fy_idx)

            # Income Statement items
            interest_income = get_val(items, "interest income", col_idx=fy_idx)
            interest_expense = get_val(items, "interest expense", col_idx=fy_idx)
            nii = get_val(items, "net interest income", col_idx=fy_idx)
            get_val(items, "fee and commission income", col_idx=fy_idx)
            get_val(items, "fee and commission expense", col_idx=fy_idx)
            net_fee = get_val(items, "net fee and commission income", col_idx=fy_idx)
            get_val(items, "net fee and commission income", col_idx=fy_idx)
            if pd.isna(net_fee):
                net_fee = get_val(items, "net fee and commission", col_idx=fy_idx)
            get_val(items, "net interest, fee and commission", col_idx=fy_idx)
            net_trading = get_val(items, "net trading income", col_idx=fy_idx)
            other_op_income = get_val(items, "other operating income", col_idx=fy_idx)
            total_op_income = get_val(items, "total operating income", col_idx=fy_idx)
            impairment = get_val(items, "impairment charge", col_idx=fy_idx)
            get_val(items, "net operating income", col_idx=fy_idx)
            personnel = get_val(items, "personnel expenses", col_idx=fy_idx)
            other_op_exp = get_val(items, "other operating expense", col_idx=fy_idx)
            other_op_exp2 = get_val(items, "other operating expenses", col_idx=fy_idx)
            if pd.isna(other_op_exp):
                other_op_exp = other_op_exp2
            depreciation = get_val(items, "depreciation", col_idx=fy_idx)
            get_val(items, "operating profit", col_idx=fy_idx)
            get_val(items, "non operating income", col_idx=fy_idx)
            get_val(items, "non operating expense", col_idx=fy_idx)
            pbt = get_val(items, "profit before income tax", col_idx=fy_idx)
            get_val(items, "current tax", col_idx=fy_idx)
            get_val(items, "deferred tax", col_idx=fy_idx)
            profit = get_val(items, "profit for the year", col_idx=fy_idx)
            if pd.isna(profit):
                profit = get_val(items, "profit for the period", col_idx=fy_idx)
            if pd.isna(profit):
                profit = get_val(items, "profit/loss for the period", col_idx=fy_idx)

            # Derived: total assets if missing (use L&E total)
            if pd.isna(total_assets):
                total_assets = get_val(items, "total liabilities and equity", col_idx=fy_idx)
            if pd.isna(total_assets):
                total_assets = get_val(items, "total equity and liabilities", col_idx=fy_idx)

            # Derived: loans
            gross_loans = np.nan
            if not pd.isna(loans_cust):
                gross_loans = loans_cust
                if not pd.isna(loans_bfis):
                    gross_loans = loans_cust + loans_bfis

            # Derived: cash_bank_balances
            cash_bank = np.nan
            vals = [v for v in [cash_equiv, due_nrb, placements, due_bfis] if not pd.isna(v)]
            if vals:
                cash_bank = sum(vals)

            # Derived: net_loans (after provisions)
            net_loans = np.nan
            if not pd.isna(gross_loans) and not pd.isna(provisions_val):
                net_loans = gross_loans - abs(provisions_val)

            # Derived: operating_expenses
            op_expenses = np.nan
            vals = [v for v in [personnel, other_op_exp, depreciation] if not pd.isna(v)]
            if vals:
                op_expenses = sum(vals)

            # Derived: operating_income
            op_income = total_op_income
            if pd.isna(op_income) and not pd.isna(nii):
                non_int = np.nan
                vals2 = [v for v in [net_fee, net_trading, other_op_income] if not pd.isna(v)]
                if vals2:
                    non_int = sum(vals2)
                if not pd.isna(non_int):
                    op_income = nii + non_int

            # Derived: total_liabilities
            total_liab = liab_total
            if pd.isna(total_liab) and not pd.isna(total_assets) and not pd.isna(equity):
                total_liab = total_assets - equity

            # Derived: shareholders_equity
            if pd.isna(equity) and not pd.isna(share_capital):
                vals = [v for v in [share_capital, retained, reserves_val] if not pd.isna(v)]
                if vals:
                    equity = sum(vals)

            # Derived: reserves
            if pd.isna(reserves_val) and not pd.isna(equity) and not pd.isna(share_capital):
                reserves_val = equity - share_capital

            # Derived: net_interest_income from components
            if pd.isna(nii) and not pd.isna(interest_income) and not pd.isna(interest_expense):
                nii = interest_income - interest_expense

            # Derived: total_op_income from components
            if pd.isna(total_op_income) and not pd.isna(nii):
                non_int = np.nan
                vals2 = [v for v in [net_fee, net_trading, other_op_income] if not pd.isna(v)]
                if vals2:
                    non_int = sum(vals2)
                if not pd.isna(non_int):
                    total_op_income = nii + non_int

            # Store balance sheet
            bank_data[target_code][fy].update({
                "total_assets": total_assets,
                "cash_bank_balances": cash_bank,
                "investments": investments,
                "gross_loans": gross_loans,
                "net_loans": net_loans,
                "total_deposits": deposits_cust,
                "borrowings": borrowings_val,
                "total_liabilities": total_liab,
                "shareholders_equity": equity,
                "paid_up_capital": share_capital,
                "reserves": reserves_val,
            })

            # Store income statement
            bank_data[target_code][fy].update({
                "interest_income": interest_income,
                "interest_expense": interest_expense,
                "net_interest_income": nii,
                "non_interest_income": np.nan,  # will compute from components
                "operating_income": op_income or total_op_income,
                "operating_expenses": op_expenses,
                "personnel_expenses": personnel,
                "provision_loan_losses": impairment,
                "profit_before_tax": pbt,
                "profit_after_tax": profit,
            })

            # Compute non_interest_income
            non_int_vals = [v for v in [net_fee, net_trading, other_op_income] if not pd.isna(v)]
            if non_int_vals:
                bank_data[target_code][fy]["non_interest_income"] = sum(non_int_vals)

    return bank_data


# ═══════════════════════════════════════════════════════════════════════════
# 5. BUILD DATAFRAMES FOR PIPELINE
# ═══════════════════════════════════════════════════════════════════════════

def build_balance_sheet_df(all_data):
    """Build balance_sheet DataFrame from extracted data."""
    rows = []
    for code, years in sorted(all_data.items()):
        name = BANK_REGISTRY.get(code, ("Unknown", False, True))[0]
        for fy in sorted(years.keys()):
            d = years[fy]
            rows.append({
                "bank_code": code,
                "bank_name": name,
                "fy": fy,
                "total_assets": d.get("total_assets"),
                "cash_bank_balances": d.get("cash_bank_balances"),
                "investments": d.get("investments"),
                "gross_loans": d.get("gross_loans"),
                "net_loans": d.get("net_loans"),
                "total_deposits": d.get("total_deposits"),
                "borrowings": d.get("borrowings"),
                "total_liabilities": d.get("total_liabilities"),
                "shareholders_equity": d.get("shareholders_equity"),
                "paid_up_capital": d.get("paid_up_capital"),
                "reserves": d.get("reserves"),
                "source": "NRB Supervision Reports",
                "notes": "Extracted from PDF tables",
            })
    return pd.DataFrame(rows)


def build_income_statement_df(all_data):
    """Build income_statement DataFrame from extracted data."""
    rows = []
    for code, years in sorted(all_data.items()):
        name = BANK_REGISTRY.get(code, ("Unknown", False, True))[0]
        for fy in sorted(years.keys()):
            d = years[fy]
            rows.append({
                "bank_code": code,
                "bank_name": name,
                "fy": fy,
                "interest_income": d.get("interest_income"),
                "interest_expense": d.get("interest_expense"),
                "net_interest_income": d.get("net_interest_income"),
                "non_interest_income": d.get("non_interest_income"),
                "operating_income": d.get("operating_income"),
                "operating_expenses": d.get("operating_expenses"),
                "personnel_expenses": d.get("personnel_expenses"),
                "provision_loan_losses": d.get("provision_loan_losses"),
                "profit_before_tax": d.get("profit_before_tax"),
                "profit_after_tax": d.get("profit_after_tax"),
                "source": "NRB Supervision Reports",
                "notes": "Extracted from PDF tables",
            })
    return pd.DataFrame(rows)


def build_operating_metrics_df(bs_df):
    """Build operating_metrics from balance sheet data (derived metrics only)."""
    rows = []
    for code in sorted(bs_df["bank_code"].unique()):
        name = BANK_REGISTRY.get(code, ("Unknown", False, True))[0]
        bank_bs = bs_df[bs_df["bank_code"] == code].sort_values("fy")
        for _, row in bank_bs.iterrows():
            fy = int(row["fy"])
            row.get("total_assets")
            row.get("gross_loans")
            row.get("total_deposits")
            # Employees/branches not in NRB PDFs → leave as NaN
            rows.append({
                "bank_code": code,
                "bank_name": name,
                "fy": fy,
                "branches": np.nan,
                "employees": np.nan,
                "atms": np.nan,
                "extension_counters": np.nan,
                "branchless_banking_centers": np.nan,
                "mobile_banking_users": np.nan,
                "internet_banking_users": np.nan,
                "debit_cards": np.nan,
                "credit_cards": np.nan,
                "qr_users": np.nan,
                "digital_transactions_count": np.nan,
                "agent_network_points": np.nan,
                "gross_npl_pct": np.nan,
                "net_npl_pct": np.nan,
                "provision_coverage_pct": np.nan,
                "car_pct": np.nan,
                "cet1_pct": np.nan,
                "source": "NRB Reports",
                "notes": "NPL/CAR data not systematically extracted",
            })
    return pd.DataFrame(rows)


def build_loan_composition_df(bs_df):
    """Build loan_composition with NaN for sector/product breakdowns (not in NRB PDFs)."""
    rows = []
    for code in sorted(bs_df["bank_code"].unique()):
        name = BANK_REGISTRY.get(code, ("Unknown", False, True))[0]
        bank_bs = bs_df[bs_df["bank_code"] == code].sort_values("fy")
        for _, row in bank_bs.iterrows():
            fy = int(row["fy"])
            gross_loans = row.get("gross_loans")
            rows.append({
                "bank_code": code,
                "bank_name": name,
                "fy": fy,
                "agriculture": np.nan,
                "manufacturing": np.nan,
                "construction": np.nan,
                "wholesale_retail": np.nan,
                "transportation": np.nan,
                "tourism": np.nan,
                "consumption": np.nan,
                "real_estate": np.nan,
                "hydropower": np.nan,
                "sme": np.nan,
                "retail": np.nan,
                "corporate": np.nan,
                "housing": np.nan,
                "vehicle": np.nan,
                "margin_lending": np.nan,
                "other_sectors": np.nan,
                "total_loans_check": gross_loans,
                "source": "NRB Sectoral Returns",
                "notes": "Sector breakdown not extracted from PDFs",
            })
    return pd.DataFrame(rows)


def build_deposit_composition_df(bs_df):
    """Build deposit_composition with NaN (not in NRB individual bank tables)."""
    rows = []
    for code in sorted(bs_df["bank_code"].unique()):
        name = BANK_REGISTRY.get(code, ("Unknown", False, True))[0]
        bank_bs = bs_df[bs_df["bank_code"] == code].sort_values("fy")
        for _, row in bank_bs.iterrows():
            fy = int(row["fy"])
            total_deposits = row.get("total_deposits")
            rows.append({
                "bank_code": code,
                "bank_name": name,
                "fy": fy,
                "current_deposits": np.nan,
                "savings_deposits": np.nan,
                "fixed_deposits": np.nan,
                "call_deposits": np.nan,
                "other_deposits": np.nan,
                "total_deposits_check": total_deposits,
                "casa_ratio": np.nan,
                "fixed_deposit_share": np.nan,
                "cost_of_deposits_pct": np.nan,
                "source": "NRB Deposit Profile",
                "notes": "Deposit breakdown not extracted from PDFs",
            })
    return pd.DataFrame(rows)


def build_macro_indicators_df():
    """Build macro_indicators with NaN (not in NRB supervision reports)."""
    rows = []
    for fy in range(2017, 2026):
        rows.append({
            "fy": fy,
            "gdp_growth_pct": np.nan,
            "inflation_pct": np.nan,
            "remittance_usd_mn": np.nan,
            "policy_rate_pct": np.nan,
            "bank_rate_pct": np.nan,
            "source": "World Bank / NRB Monetary Policy",
            "notes": "Macroeconomic data not extracted",
        })
    return pd.DataFrame(rows)


def build_market_data_df(bs_df):
    """Build market_data with NaN (NEPSE data not extracted)."""
    rows = []
    for code in sorted(bs_df["bank_code"].unique()):
        name = BANK_REGISTRY.get(code, ("Unknown", False, True))[0]
        is_listed = BANK_REGISTRY.get(code, ("Unknown", False, True))[2]
        bank_bs = bs_df[bs_df["bank_code"] == code].sort_values("fy")
        for _, row in bank_bs.iterrows():
            fy = int(row["fy"])
            rows.append({
                "bank_code": code,
                "bank_name": name,
                "fy": fy,
                "ticker": code if is_listed else None,
                "share_price_eoy": np.nan,
                "market_cap": np.nan,
                "pe_ratio": np.nan,
                "pb_ratio": np.nan,
                "eps": np.nan,
                "bvps": np.nan,
                "dividend_per_share": np.nan,
                "dividend_yield_pct": np.nan,
                "annual_return_pct": np.nan,
                "price_volatility": np.nan,
                "source": "NEPSE / Annual Report",
                "notes": "Market data not extracted",
            })
    return pd.DataFrame(rows)


def build_strategic_coding_df(bs_df):
    """Build digital_scorecard and strategic_priorities with NaN."""
    dig_rows = []
    strat_rows = []
    for code in sorted(bs_df["bank_code"].unique()):
        name = BANK_REGISTRY.get(code, ("Unknown", False, True))[0]
        bank_bs = bs_df[bs_df["bank_code"] == code].sort_values("fy")
        for _, row in bank_bs.iterrows():
            fy = int(row["fy"])
            dig_rows.append({
                "bank_code": code, "bank_name": name, "fy": fy,
                "digital_account_opening": np.nan,
                "mobile_banking": np.nan,
                "digital_lending": np.nan,
                "qr_ecosystem": np.nan,
                "api_open_banking": np.nan,
                "ai_initiatives": np.nan,
                "digital_customer_acquisition": np.nan,
                "core_banking_upgrade": np.nan,
                "fintech_partnership": np.nan,
                "cybersecurity_initiative": np.nan,
                "digital_index": np.nan,
                "evidence_notes": "Not extracted from reports",
            })
            strat_rows.append({
                "bank_code": code, "bank_name": name, "fy": fy,
                "priority_retail": np.nan,
                "priority_sme": np.nan,
                "priority_corporate": np.nan,
                "priority_digital": np.nan,
                "priority_branch_expansion": np.nan,
                "priority_cost_reduction": np.nan,
                "priority_wealth_mgmt": np.nan,
                "priority_remittance": np.nan,
                "priority_sustainability": np.nan,
                "priority_geographic_expansion": np.nan,
                "strategic_score": np.nan,
                "evidence_notes": "Not extracted from reports",
            })
    return pd.DataFrame(dig_rows), pd.DataFrame(strat_rows)


# ═══════════════════════════════════════════════════════════════════════════
# 6. POPULATE EXISTING EVENTS DATA
# ═══════════════════════════════════════════════════════════════════════════

def build_events_df():
    """Build the bank_events DataFrame with known M&A and tech events."""
    events_data = [
        {"event_id": "EV001", "bank_code": "GIBL", "bank_name": "Global IME Bank",
         "fy": 2020, "event_date": "2019-12-06", "event_type": "M&A",
         "event_description": "Merger with Janata Bank Nepal Limited",
         "strategic_impact": "High", "counterparty": "Janata Bank Nepal",
         "financial_effect_notes": "Expanded asset base, paid-up capital and branch network.",
         "source_url": "https://www.nrb.org.np",
         "notes": "First major 'Big Merger' in Nepali commercial banking."},
        {"event_id": "EV002", "bank_code": "NABIL", "bank_name": "Nabil Bank",
         "fy": 2022, "event_date": "2022-07-11", "event_type": "M&A",
         "event_description": "Acquisition of Nepal Bangladesh Bank Limited (NBB)",
         "strategic_impact": "High", "counterparty": "Nepal Bangladesh Bank",
         "financial_effect_notes": "Expanded retail deposit base, branch footprint.",
         "source_url": "https://www.nabilbank.com",
         "notes": "Swap ratio 100:43 (NABIL:NBB)."},
        {"event_id": "EV003", "bank_code": "GIBL", "bank_name": "Global IME Bank",
         "fy": 2023, "event_date": "2023-01-09", "event_type": "M&A",
         "event_description": "Merger with Bank of Kathmandu Limited (BOKL)",
         "strategic_impact": "High", "counterparty": "Bank of Kathmandu",
         "financial_effect_notes": "Created largest commercial bank by capital fund and asset size.",
         "source_url": "https://globalimebank.com",
         "notes": "Merged entity retained Global IME Bank name."},
        {"event_id": "EV004", "bank_code": "PRVU", "bank_name": "Prabhu Bank",
         "fy": 2023, "event_date": "2023-01-10", "event_type": "M&A",
         "event_description": "Merger with Century Commercial Bank Limited (CCBL)",
         "strategic_impact": "High", "counterparty": "Century Commercial Bank",
         "financial_effect_notes": "Expanded geographical distribution and retail deposit book.",
         "source_url": "https://prabhubank.com",
         "notes": "Swap ratio 1:1."},
        {"event_id": "EV005", "bank_code": "NIMB", "bank_name": "Nepal Investment Mega Bank",
         "fy": 2023, "event_date": "2023-01-11", "event_type": "M&A",
         "event_description": "Merger of NIBL and Mega Bank Nepal",
         "strategic_impact": "High", "counterparty": "Mega Bank Nepal",
         "financial_effect_notes": "Substantial capital boost, combined corporate and retail strength.",
         "source_url": "https://nibl.com.np",
         "notes": "Operating under new name NIMB."},
        {"event_id": "EV006", "bank_code": "KBL", "bank_name": "Kumari Bank",
         "fy": 2023, "event_date": "2023-01-01", "event_type": "M&A",
         "event_description": "Merger with Nepal Credit and Commerce (NCC) Bank",
         "strategic_impact": "High", "counterparty": "NCC Bank",
         "financial_effect_notes": "Strengthened capital base, improved liquidity buffer.",
         "source_url": "https://kumaribank.com",
         "notes": "Combined operations began Jan 2023."},
        {"event_id": "EV007", "bank_code": "HBL", "bank_name": "Himalayan Bank",
         "fy": 2023, "event_date": "2023-02-24", "event_type": "M&A",
         "event_description": "Acquisition of Civil Bank Limited",
         "strategic_impact": "High", "counterparty": "Civil Bank",
         "financial_effect_notes": "Strengthened market share in Tier 2/3 cities.",
         "source_url": "https://himalayanbank.com",
         "notes": "Civil Bank ceased standalone operations."},
        {"event_id": "EV008", "bank_code": "LLBS", "bank_name": "Laxmi Sunrise Bank",
         "fy": 2024, "event_date": "2023-07-14", "event_type": "M&A",
         "event_description": "Merger between Laxmi Bank and Sunrise Bank",
         "strategic_impact": "High", "counterparty": "Sunrise Bank",
         "financial_effect_notes": "Consolidated balance sheet over NPR 300 billion.",
         "source_url": "https://laxmisunrisebank.com",
         "notes": "Swap ratio 1:1, commenced FY2024."},
        {"event_id": "EV009", "bank_code": "NICA", "bank_name": "NIC Asia Bank",
         "fy": 2021, "event_date": "2020-09-15", "event_type": "Technology",
         "event_description": "Rollout of unified iServe and Digital 360 platform",
         "strategic_impact": "Medium", "counterparty": "Internal / Fintech partners",
         "financial_effect_notes": "Drove massive retail customer acquisition.",
         "source_url": "https://nicasiabank.com",
         "notes": "Pioneered paperless branch service in Nepal."},
        {"event_id": "EV010", "bank_code": "SCB", "bank_name": "Standard Chartered Bank Nepal",
         "fy": 2024, "event_date": "2023-11-20", "event_type": "Technology",
         "event_description": "Launch of Straight2Bank NextGen for Corporate Clients",
         "strategic_impact": "Medium", "counterparty": "Standard Chartered Group",
         "financial_effect_notes": "Maintained dominant fee-income market share.",
         "source_url": "https://sc.com/np",
         "notes": "Corporate digital cash management."},
    ]
    return pd.DataFrame(events_data)


# ═══════════════════════════════════════════════════════════════════════════
# 7. MAIN EXECUTION
# ═══════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 70)
    print("Nepal Banking Research — Real Data Extraction")
    print("=" * 70)

    # Step 1: Extract from all reports
    print("\n[1/6] Extracting data from NRB supervision report CSVs...")
    all_data = defaultdict(lambda: defaultdict(dict))

    for report_meta in REPORT_LIST:
        prefix = report_meta["prefix"]
        print(f"  Processing {prefix}...")
        report_data = extract_from_report(report_meta)

        # Merge into all_data (prefer non-NaN values, later reports may overlap)
        for code, years in report_data.items():
            for fy, items in years.items():
                for key, val in items.items():
                    if key not in all_data[code][fy] or pd.isna(all_data[code][fy].get(key)) or not pd.isna(val) and pd.isna(all_data[code][fy].get(key)):
                        all_data[code][fy][key] = val

    # Report extraction summary
    total_banks = len(all_data)
    total_rows = sum(len(years) for years in all_data.values())
    print(f"  Extracted {total_banks} banks, {total_rows} bank-year observations")

    for code in sorted(all_data.keys()):
        fys = sorted(all_data[code].keys())
        real_assets = sum(1 for fy, d in all_data[code].items()
                          if not pd.isna(d.get("total_assets")))
        real_pat = sum(1 for fy, d in all_data[code].items()
                       if not pd.isna(d.get("profit_after_tax")))
        print(f"    {code}: FY{min(fys)}-{max(fys)}, assets={real_assets}, PAT={real_pat}")

    # Step 2: Build DataFrames
    print("\n[2/6] Building DataFrames...")
    bs_df = build_balance_sheet_df(all_data)
    is_df = build_income_statement_df(all_data)
    print(f"  balance_sheet: {len(bs_df)} rows, {bs_df['bank_code'].nunique()} banks")
    print(f"  income_statement: {len(is_df)} rows")

    # Step 3: Write 01_bank_financials.xlsx
    print("\n[3/6] Writing DATA/01_bank_financials.xlsx...")
    fin_path = os.path.join(DATA, "01_bank_financials.xlsx")
    with pd.ExcelWriter(fin_path, engine="openpyxl") as writer:
        bs_df.to_excel(writer, sheet_name="balance_sheet", index=False)
        is_df.to_excel(writer, sheet_name="income_statement", index=False)
    print(f"  Written: {fin_path}")

    # Step 4: Write operating metrics
    print("\n[4/6] Writing DATA/04_operating_metrics.xlsx...")
    op_df = build_operating_metrics_df(bs_df)
    op_path = os.path.join(DATA, "04_operating_metrics.xlsx")
    with pd.ExcelWriter(op_path, engine="openpyxl") as writer:
        op_df.to_excel(writer, sheet_name="operating_metrics", index=False)
    print(f"  Written: {op_path} ({len(op_df)} rows)")

    # Step 5: Write composition files
    print("\n[5/6] Writing composition files...")
    lc_df = build_loan_composition_df(bs_df)
    lc_path = os.path.join(DATA, "05_loan_composition.xlsx")
    with pd.ExcelWriter(lc_path, engine="openpyxl") as writer:
        lc_df.to_excel(writer, sheet_name="loan_composition", index=False)
    print(f"  Written: {lc_path} ({len(lc_df)} rows)")

    dc_df = build_deposit_composition_df(bs_df)
    dc_path = os.path.join(DATA, "06_deposit_composition.xlsx")
    with pd.ExcelWriter(dc_path, engine="openpyxl") as writer:
        dc_df.to_excel(writer, sheet_name="deposit_composition", index=False)
    print(f"  Written: {dc_path} ({len(dc_df)} rows)")

    macro_df = build_macro_indicators_df()
    macro_path = os.path.join(DATA, "07_macro_indicators.xlsx")
    with pd.ExcelWriter(macro_path, engine="openpyxl") as writer:
        macro_df.to_excel(writer, sheet_name="macro_indicators", index=False)
    print(f"  Written: {macro_path}")

    mkt_df = build_market_data_df(bs_df)
    mkt_path = os.path.join(DATA, "08_market_data.xlsx")
    with pd.ExcelWriter(mkt_path, engine="openpyxl") as writer:
        mkt_df.to_excel(writer, sheet_name="market_data", index=False)
    print(f"  Written: {mkt_path}")

    dig_df, strat_df = build_strategic_coding_df(bs_df)
    strat_path = os.path.join(DATA, "09_strategic_coding.xlsx")
    with pd.ExcelWriter(strat_path, engine="openpyxl") as writer:
        dig_df.to_excel(writer, sheet_name="digital_scorecard", index=False)
        strat_df.to_excel(writer, sheet_name="strategic_priorities", index=False)
    print(f"  Written: {strat_path}")

    events_df = build_events_df()
    events_path = os.path.join(DATA, "10_bank_events.xlsx")
    with pd.ExcelWriter(events_path, engine="openpyxl") as writer:
        events_df.to_excel(writer, sheet_name="events", index=False)
    print(f"  Written: {events_path}")

    # Step 6: Summary
    print("\n[6/6] Extraction Summary")
    print("=" * 70)
    print(f"  Total banks: {total_banks}")
    print(f"  Total bank-year obs: {total_rows}")
    print(f"  Real total_assets filled: {bs_df['total_assets'].notna().sum()}/{len(bs_df)}")
    print(f"  Real profit_after_tax filled: {is_df['profit_after_tax'].notna().sum()}/{len(is_df)}")
    print(f"  Real interest_income filled: {is_df['interest_income'].notna().sum()}/{len(is_df)}")
    print(f"  Real gross_loans filled: {bs_df['gross_loans'].notna().sum()}/{len(bs_df)}")
    print(f"  Real total_deposits filled: {bs_df['total_deposits'].notna().sum()}/{len(bs_df)}")
    print()
    print("  Files written:")
    for f in sorted(os.listdir(DATA)):
        if f.endswith(".xlsx"):
            size = os.path.getsize(os.path.join(DATA, f))
            print(f"    {f:40s} {size:>8,} bytes")
    print()
    print("  Note: Operating metrics (employees, branches, NPL, CAR),")
    print("  loan/deposit composition, macro, market, strategic data")
    print("  are left as NaN — they require external data sources.")
    print()
    print("  Next: Run '.\\tasks.ps1 build' to calculate ratios,")
    print("  market shares, and generate the master panel.")


if __name__ == "__main__":
    main()
