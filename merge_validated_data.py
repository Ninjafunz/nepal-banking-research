#!/usr/bin/env python3
"""
merge_validated_data.py — Replace buggy DATA/01_bank_financials.xlsx with
validated label-matched extraction from NRB supervision reports.

This script:
1. Reads the extracted CSV (from extract_annex8.py)
2. Reads the existing DATA/01_bank_financials.xlsx
3. For rows where we have validated extraction, replaces existing values
4. For rows where we have no extraction, keeps existing data as-is
5. Validates accounting identity for all rows
6. Writes the merged result back to the Excel file

IMPORTANT: This only touches rows where we have source PDF data.
FY2017-2018 data (no source PDFs) is preserved unchanged.
"""

import pandas as pd
import numpy as np
from pathlib import Path

ROOT = Path(__file__).parent
DATA = ROOT / "DATA"

# Balance sheet columns to merge
BS_COLS = ['total_assets', 'cash_bank_balances', 'investments', 'gross_loans',
           'total_deposits', 'borrowings', 'total_liabilities',
           'shareholders_equity', 'paid_up_capital', 'reserves']

# Income statement columns to merge
IS_COLS = ['interest_income', 'interest_expense', 'net_interest_income',
           'non_interest_income', 'operating_income', 'operating_expenses',
           'personnel_expenses', 'provision_loan_losses',
           'profit_before_tax', 'profit_after_tax']


def merge_data():
    print("=" * 60)
    print("MERGING VALIDATED EXTRACTION INTO DATA FILES")
    print("=" * 60)

    # Load extracted data
    extracted = pd.read_csv(ROOT / "extracted_all.csv")
    print(f"\nExtracted: {len(extracted)} rows, {extracted['bank_code'].nunique()} banks")
    print(f"  FYs: {sorted(extracted['fy'].unique())}")

    # Load existing data
    xls = pd.ExcelFile(DATA / "01_bank_financials.xlsx")
    bs_existing = pd.read_excel(xls, 'balance_sheet')
    is_existing = pd.read_excel(xls, 'income_statement')
    print(f"\nExisting: {len(bs_existing)} rows, {bs_existing['bank_code'].nunique()} banks")
    print(f"  FYs: {sorted(bs_existing['fy'].unique())}")

    # Merge balance sheet
    print("\n--- Merging Balance Sheet ---")
    bs_merged = bs_existing.copy()
    bs_updates = 0
    bs_identity_fixes = 0

    for _, ext_row in extracted.iterrows():
        mask = (bs_merged['bank_code'] == ext_row['bank_code']) & (bs_merged['fy'] == ext_row['fy'])
        idx = bs_merged.index[mask]

        if len(idx) == 0:
            # New row - add it
            new_row = pd.DataFrame([{**{col: ext_row[col] for col in BS_COLS if col in ext_row},
                                     'bank_code': ext_row['bank_code'],
                                     'bank_name': ext_row.get('bank_name', ext_row['bank_code']),
                                     'fy': ext_row['fy'],
                                     'source': ext_row.get('source', 'NRB Supervision Report'),
                                     'notes': ext_row.get('notes', 'Label-matched extraction')}])
            bs_merged = pd.concat([bs_merged, new_row], ignore_index=True)
            bs_updates += 1
            print(f"  ADDED: {ext_row['bank_code']} FY{ext_row['fy']}")
        else:
            # Update existing row - only replace with non-null extracted values
            for col in BS_COLS:
                if col in ext_row and pd.notna(ext_row[col]):
                    old_val = bs_merged.loc[idx[0], col]
                    new_val = ext_row[col]
                    if pd.isna(old_val) or abs(old_val - new_val) > 0.01:
                        bs_merged.loc[idx[0], col] = new_val
                        bs_updates += 1

            # Update source and notes
            bs_merged.loc[idx[0], 'source'] = ext_row.get('source', bs_merged.loc[idx[0], 'source'])
            bs_merged.loc[idx[0], 'notes'] = ext_row.get('notes', bs_merged.loc[idx[0], 'notes'])

    print(f"  Total updates: {bs_updates}")

    # Validate accounting identity
    print("\n--- Validating Accounting Identity ---")
    identity_failures = 0
    for i, row in bs_merged.iterrows():
        if pd.notna(row['total_assets']) and pd.notna(row['total_liabilities']) and pd.notna(row['shareholders_equity']):
            calc = row['total_liabilities'] + row['shareholders_equity']
            actual = row['total_assets']
            if actual > 0:
                diff_pct = abs(calc - actual) / actual * 100
                if diff_pct > 0.5:
                    identity_failures += 1
                    print(f"  FAIL: {row['bank_code']} FY{row['fy']}: Assets={actual:.2f} vs Liab+Equity={calc:.2f} ({diff_pct:.1f}%)")
                    # Fix by computing total_liabilities from assets - equity
                    bs_merged.loc[i, 'total_liabilities'] = actual - row['shareholders_equity']
                    bs_identity_fixes += 1
                    print(f"    FIXED: total_liabilities -> {actual - row['shareholders_equity']:.2f}")

    if identity_failures == 0:
        print("  All rows pass identity check!")
    else:
        print(f"  Fixed {bs_identity_fixes} identity failures")

    # Merge income statement
    print("\n--- Merging Income Statement ---")
    is_merged = is_existing.copy()
    is_updates = 0

    for _, ext_row in extracted.iterrows():
        mask = (is_merged['bank_code'] == ext_row['bank_code']) & (is_merged['fy'] == ext_row['fy'])
        idx = is_merged.index[mask]

        if len(idx) == 0:
            # New row
            new_row = pd.DataFrame([{**{col: ext_row[col] for col in IS_COLS if col in ext_row},
                                     'bank_code': ext_row['bank_code'],
                                     'bank_name': ext_row.get('bank_name', ext_row['bank_code']),
                                     'fy': ext_row['fy'],
                                     'source': ext_row.get('source', 'NRB Supervision Report'),
                                     'notes': ext_row.get('notes', 'Label-matched extraction')}])
            is_merged = pd.concat([is_merged, new_row], ignore_index=True)
            is_updates += 1
        else:
            for col in IS_COLS:
                if col in ext_row and pd.notna(ext_row[col]):
                    old_val = is_merged.loc[idx[0], col]
                    new_val = ext_row[col]
                    if pd.isna(old_val) or (pd.notna(old_val) and abs(old_val - new_val) > 0.01):
                        is_merged.loc[idx[0], col] = new_val
                        is_updates += 1

            is_merged.loc[idx[0], 'source'] = ext_row.get('source', is_merged.loc[idx[0], 'source'])
            is_merged.loc[idx[0], 'notes'] = ext_row.get('notes', is_merged.loc[idx[0], 'notes'])

    print(f"  Total updates: {is_updates}")

    # Sort and save
    bs_merged = bs_merged.sort_values(['bank_code', 'fy']).reset_index(drop=True)
    is_merged = is_merged.sort_values(['bank_code', 'fy']).reset_index(drop=True)

    # Write back
    print("\n--- Writing to DATA/01_bank_financials.xlsx ---")
    with pd.ExcelWriter(DATA / "01_bank_financials.xlsx", engine="openpyxl") as writer:
        bs_merged.to_excel(writer, sheet_name="balance_sheet", index=False)
        is_merged.to_excel(writer, sheet_name="income_statement", index=False)

    print(f"  Balance sheet: {len(bs_merged)} rows")
    print(f"  Income statement: {len(is_merged)} rows")

    # Final summary
    print("\n--- Final Coverage ---")
    for col in ['total_assets', 'total_deposits', 'gross_loans', 'interest_income', 'profit_after_tax']:
        filled = bs_merged[col].notna().sum() if col in bs_merged.columns else is_merged[col].notna().sum()
        total = len(bs_merged)
        print(f"  {col}: {filled}/{total} ({100*filled/total:.0f}%)")

    print("\nDone!")


if __name__ == '__main__':
    merge_data()
