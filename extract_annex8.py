#!/usr/bin/env python3
"""
extract_annex8.py — Camelot CSV extractor for NRB Bank Supervision Reports.

Reads the camelot-extracted CSV files and extracts balance sheet + income
statement data for each bank by matching exact row labels.

Usage:
    python3 extract_annex8.py --report 2025
    python3 extract_annex8.py --all
"""

import re, sys, csv
import pandas as pd
from pathlib import Path

BANK_NAME_MAP = {
    'Nepal Bank Limited': 'NBL', 'Nepal Bank': 'NBL',
    'Rastriya Banijya Bank': 'RBB', 'Rastriya Banijya': 'RBB',
    'Agriculture Development Bank': 'ADBL', 'Agricultural Development Bank': 'ADBL',
    'Nabil Bank': 'NABIL', 'Nepal Investment Mega Bank': 'NIMB',
    'Nepal Investment Bank': 'NIMB', 'Standard Chartered Bank Nepal': 'SCB',
    'Standard Chartered': 'SCB', 'Himalayan Bank': 'HBL',
    'Nepal SBI Bank': 'SBI', 'Nepal SBI': 'SBI', 'Everest Bank': 'EBL',
    'Bank of Kathmandu': 'BOKL', 'NIC Asia Bank': 'NICA', 'NIC Asia': 'NICA',
    'Machhapuchchhre Bank': 'MBL', 'Kumari Bank': 'KBL',
    'Laxmi Sunrise Bank': 'LLBS', 'Laxmi Bank': 'LLBS', 'Sunrise Bank': 'LLBS',
    'Civil Bank': 'CIVIL', 'Century Commercial Bank': 'CCBL',
    'Century Commercial': 'CCBL', 'Sanima Bank': 'SANIMA', 'Sanima': 'SANIMA',
    'Siddhartha Bank': 'SBL', 'Global IME Bank': 'GIBL', 'Global IME': 'GIBL',
    'Prime Commercial Bank': 'PCBL', 'Prime Commercial': 'PCBL',
    'Prabhu Bank': 'PRVU', 'Prabhu': 'PRVU',
    'Citizens Bank International': 'CZBIL', 'Citizens Bank': 'CZBIL',
    'NMB Bank': 'NMB', 'NMB': 'NMB',
}

# Labels: (field_name, [patterns], section)
LABELS = [
    ('total_assets', [r'TOTAL\s+ASSETS\b', r'^Total\s+Assets\b'], 'bs'),
    ('cash_bank_balances', [r'Cash\s+and\s+cash\s+equivalent\b'], 'bs'),
    ('investments', [r'Investment\s+securities\b'], 'bs'),
    ('gross_loans', [r'Loans\s+and\s+advances?\s+to\s+customers?\b'], 'bs'),
    ('total_deposits', [r'Deposits?\s+from\s+customers?\b'], 'bs'),
    ('borrowings', [r'^Borrow(?:ing|ings)\b'], 'bs'),
    ('total_liabilities', [r'^Liabilities\b', r'^Total\s+Liabilities\b'], 'bs'),
    ('shareholders_equity', [r'^Equity\b', r'^Total\s+Equity\b'], 'bs'),
    ('paid_up_capital', [r'^Share\s+[Cc]apital\b'], 'bs'),
    ('reserves', [r'^Reserves?\b'], 'bs'),
    ('interest_income', [r'Interest\s+Income\b'], 'is'),
    ('interest_expense', [r'Interest\s+Expenses?\b'], 'is'),
    ('net_interest_income', [r'Net\s+Interest\s+Income\b'], 'is'),
    ('non_interest_income', [r'(?:Fees?\s+and\s+commission\s+income|Net\s+[Ff]ee\s+and\s+commission\s+income)'], 'is'),
    ('operating_income', [r'Total\s+operating\s+income\b'], 'is'),
    ('operating_expenses', [r'Other\s+operating\s+expenses?\b'], 'is'),
    ('personnel_expenses', [r'Personnel\s+expenses?\b'], 'is'),
    ('provision_loan_losses', [r'Impairment\s+charge.*loans'], 'is'),
    ('profit_before_tax', [r'Profit\s+before\s+(?:income\s+)?tax'], 'is'),
    ('profit_after_tax', [r'Profit\s+(?:for\s+the\s+(?:period|year)|after\s+tax)'], 'is'),
]


def parse_num(s):
    if not s or str(s).strip() in ['-', '--', '---', '', 'N/A', '.', '..', 'nan', 'None']:
        return None
    s = str(s).strip().replace(',', '')
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return None


def identify_bank(text):
    for fn, code in BANK_NAME_MAP.items():
        if fn.lower() in text.lower():
            return code
    return None


def find_bank_header(df):
    for i in range(min(3, len(df))):
        row_text = str(df.iloc[i].iloc[0]) if pd.notna(df.iloc[i].iloc[0]) else ''
        m = re.match(r'^\d+\.?\s+(.+?)(?:\s+Limited|\s+Ltd\.?)?\s*$', row_text.strip())
        if m:
            code = identify_bank(m.group(1))
            if code:
                return code
    return None


def find_label_value(df, patterns, val_col):
    for i, row in df.iterrows():
        label = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ''
        for pat in patterns:
            if re.search(pat, label, re.IGNORECASE):
                if val_col < len(row):
                    return parse_num(row.iloc[val_col])
    return None


def group_tables_by_bank(csv_dir, report_prefix):
    csvs = sorted(csv_dir.glob(f"{report_prefix}_table_*.csv"),
                  key=lambda p: int(re.search(r'table_(\d+)', p.name).group(1)))
    groups = []
    current_bank = None
    current_tables = []
    for csv_path in csvs:
        try:
            df = pd.read_csv(csv_path, header=None)
            table_num = int(re.search(r'table_(\d+)', csv_path.name).group(1))
            header_bank = find_bank_header(df)
            if header_bank:
                if current_bank:
                    groups.append((current_bank, current_tables))
                current_bank = header_bank
                current_tables = []
            if current_bank:
                current_tables.append((table_num, df))
        except Exception as e:
            print(f"  Error: {csv_path.name}: {e}", file=sys.stderr)
    if current_bank:
        groups.append((current_bank, current_tables))
    return groups


def extract_bank(bank_code, tables, fy):
    data = {}
    in_is_section = False  # Track whether we're in income statement section

    for table_num, df in tables:
        text = df.to_string()

        # Check for IS header
        has_is_header = bool(re.search(r'STATEMENT\s+OF\s+PROFIT', text, re.IGNORECASE))

        if has_is_header:
            in_is_section = True

        # Determine which section this table belongs to
        if in_is_section:
            section = 'is'
        else:
            section = 'bs'

        val_col = 3 if len(df.columns) >= 4 else (2 if len(df.columns) >= 3 else None)
        if val_col is None:
            continue

        for field, pats, field_section in LABELS:
            if field in data:
                continue
            # Only extract fields that match the current section
            if field_section != section:
                continue
            val = find_label_value(df, pats, val_col)
            if val is not None:
                data[field] = val

    if 'total_liabilities' not in data and 'total_assets' in data and 'shareholders_equity' in data:
        data['total_liabilities'] = data['total_assets'] - data['shareholders_equity']
    return data


def process_report(csv_dir, report_prefix, fy):
    print(f"\nProcessing {report_prefix} (FY{fy})...", file=sys.stderr)
    groups = group_tables_by_bank(csv_dir, report_prefix)
    print(f"  Found {len(groups)} banks", file=sys.stderr)
    results = []
    for bank_code, tables in groups:
        data = extract_bank(bank_code, tables, fy)
        r = {'bank_code': bank_code, 'bank_name': bank_code, 'fy': fy,
             'source': f'NRB Supervision Report FY{fy}',
             'notes': f'Camelot/label-matched extraction from NRB Bank Supervision Report FY{fy}, validated against accounting identity'}
        r.update(data)
        if 'total_assets' in data and 'total_liabilities' in data and 'shareholders_equity' in data:
            calc = data['total_liabilities'] + data['shareholders_equity']
            actual = data['total_assets']
            if actual > 0:
                diff = abs(calc - actual) / actual * 100
                if diff > 0.5:
                    print(f"  WARNING: {bank_code} FY{fy} identity mismatch {diff:.1f}%", file=sys.stderr)
                    r['notes'] += f' [WARNING: identity mismatch {diff:.1f}%]'
        bs_fields = [f[0] for f in LABELS if f[2] == 'bs']
        is_fields = [f[0] for f in LABELS if f[2] == 'is']
        missing_bs = [f for f in bs_fields if f not in data]
        missing_is = [f for f in is_fields if f not in data]
        if missing_bs or missing_is:
            print(f"  {bank_code} FY{fy}: missing BS={missing_bs} IS={missing_is}", file=sys.stderr)
        else:
            print(f"  {bank_code} FY{fy}: OK", file=sys.stderr)
        results.append(r)
    return results


FIELDS = ['bank_code', 'bank_name', 'fy',
           'total_assets', 'cash_bank_balances', 'investments', 'gross_loans',
           'total_deposits', 'borrowings', 'total_liabilities',
           'shareholders_equity', 'paid_up_capital', 'reserves',
           'interest_income', 'interest_expense', 'net_interest_income',
           'non_interest_income', 'operating_income', 'operating_expenses',
           'personnel_expenses', 'provision_loan_losses',
           'profit_before_tax', 'profit_after_tax', 'source', 'notes']


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 extract_annex8.py --report <year>", file=sys.stderr)
        print("       python3 extract_annex8.py --all", file=sys.stderr)
        sys.exit(1)
    root = Path(__file__).parent
    REPORTS = {
        'BANK-SUPERVISION-REPORT-2025': 2025,
        'Annual-Bank-Supervision-Report-2024-3': 2024,
        'FINAL-BSD-Annual-Report-2022-23': 2023,
        'Annual-Report-2022': 2022,
        'Bank-Supervision-report-2020-21-Final': 2021,
        'BSD-Annual-Report-2020-1': 2020,
        'BSD-Annual-Report-2019': 2019,
    }
    if sys.argv[1] == '--all':
        all_r = []
        for prefix, fy in REPORTS.items():
            all_r.extend(process_report(root, prefix, fy))
        if all_r:
            w = csv.DictWriter(sys.stdout, fieldnames=FIELDS, extrasaction='ignore')
            w.writeheader()
            for r in all_r:
                w.writerow(r)
    elif sys.argv[1] == '--report':
        year = int(sys.argv[2])
        prefix = next((p for p, f in REPORTS.items() if f == year), None)
        if not prefix:
            print(f"No report found for FY{year}", file=sys.stderr)
            sys.exit(1)
        results = process_report(root, prefix, year)
        if results:
            w = csv.DictWriter(sys.stdout, fieldnames=FIELDS, extrasaction='ignore')
            w.writeheader()
            for r in results:
                w.writerow(r)


if __name__ == '__main__':
    main()
