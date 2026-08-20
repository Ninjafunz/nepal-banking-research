"""
extract_nabil_annual.py — Extract verified data from NABIL annual reports.

Reads the pdftotext output of each annual report and extracts:
- Employees and branches (for operating_metrics)
- NPL ratios (gross_npl_pct, net_npl_pct)
- CAR (capital adequacy ratio)

Then merges into DATA files, preserving existing data where we have it.
"""

import re
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
RAW = ROOT / "RAW" / "Nabil"
DATA = ROOT / "DATA"

# Map FY to annual report file
# Nepal FY ends mid-July, so FY2020 = July 2019-July 2020
ANNUAL_REPORTS = {
    2020: "Annual Report 2019-20 (English).pdf",
    2021: "Annual Report 2020-21 (English) - For Website.pdf",
    2022: "Annual Report 2021-22 (English).pdf",
    2023: "Annual Report 2022-23 (English).pdf",
    2024: "Annual Report 2023-24 (English).pdf",
    2025: "Annual Report 2024-25 (English).pdf",
}


def extract_text(pdf_path):
    """Extract text from PDF using pdftotext."""
    import subprocess
    txt_path = str(pdf_path).replace('.pdf', '.txt')
    subprocess.run(['pdftotext', '-layout', str(pdf_path), txt_path], capture_output=True)
    with open(txt_path, 'r', encoding='utf-8', errors='replace') as f:
        return f.readlines()


def parse_num(s):
    if not s or s.strip() in ['-', '--', '---', '', 'N/A']:
        return None
    s = s.strip().replace(',', '')
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return None


def extract_employees_branches(lines):
    """Extract employee count and branch count from annual report text."""
    employees = None
    branches = None
    atms = None
    
    for i, line in enumerate(lines):
        # Employee count: "No of employees  2,404  2,235"
        m = re.search(r'No\s+of\s+employees\s+(\d[\d,]*)', line, re.IGNORECASE)
        if m:
            employees = parse_num(m.group(1))
        
        # Branch count: "Branch network (Count)" or "branches"
        m = re.search(r'Branch\s+network\s+\(?Count\)?\s+(\d+)', line, re.IGNORECASE)
        if m:
            branches = parse_num(m.group(1))
        
        # ATM count: "ATM (Count)  268  265"
        m = re.search(r'ATM\s+\(?Count\)?\s+(\d+)', line, re.IGNORECASE)
        if m:
            atms = parse_num(m.group(1))
    
    return employees, branches, atms


def extract_npl(lines):
    """Extract NPL ratios from annual report text."""
    gross_npl = None
    net_npl = None
    
    for i, line in enumerate(lines):
        # Gross NPA: "Gross NPAs  6.82%  6.86%"
        m = re.search(r'Gross\s+NPAs?\s+([\d.]+)%', line, re.IGNORECASE)
        if m:
            gross_npl = parse_num(m.group(1))
        
        # Net NPA: "Net NPAs/net worth" or "Net NPA ratio"
        m = re.search(r'Net\s+NPAs?/?net\s+worth\s+([\d.]+)%', line, re.IGNORECASE)
        if m:
            net_npl = parse_num(m.group(1))
    
    return gross_npl, net_npl


def extract_car(lines):
    """Extract Capital Adequacy Ratio from annual report text."""
    car = None
    
    for i, line in enumerate(lines):
        # CAR: "Capital adequacy ratio  11.40%  10.90%"
        m = re.search(r'Capital\s+adequacy\s+ratio\s+([\d.]+)%', line, re.IGNORECASE)
        if m:
            car = parse_num(m.group(1))
    
    return car


def process_report(fy, filename):
    """Process a single annual report."""
    pdf_path = RAW / filename
    if not pdf_path.exists():
        print(f"  File not found: {filename}")
        return {}
    
    print(f"  Processing {filename}...")
    lines = extract_text(pdf_path)
    
    employees, branches, atms = extract_employees_branches(lines)
    gross_npl, net_npl = extract_npl(lines)
    car = extract_car(lines)
    
    result = {
        'employees': employees,
        'branches': branches,
        'atms': atms,
        'gross_npl_pct': gross_npl,
        'net_npl_pct': net_npl,
        'car_pct': car,
    }
    
    print(f"    Employees: {employees}, Branches: {branches}, ATMs: {atms}")
    print(f"    Gross NPL: {gross_npl}%, Net NPL: {net_npl}%, CAR: {car}%")
    
    return result


def main():
    print("=" * 60)
    print("EXTRACTING NABIL DATA FROM ANNUAL REPORTS")
    print("=" * 60)
    
    # Extract data from each annual report
    extracted = {}
    for fy, filename in ANNUAL_REPORTS.items():
        extracted[fy] = process_report(fy, filename)
    
    # Load existing data
    om = pd.read_excel(DATA / "04_operating_metrics.xlsx", sheet_name="operating_metrics")
    
    print("\n" + "=" * 60)
    print("MERGING INTO DATA FILES")
    print("=" * 60)
    
    # Update operating_metrics for NABIL
    updated = 0
    for fy, data in extracted.items():
        mask = (om['bank_code'] == 'NABIL') & (om['fy'] == fy)
        idx = om.index[mask]
        
        if len(idx) == 0:
            print(f"  FY{fy}: No row found in operating_metrics")
            continue
        
        for field in ['employees', 'branches', 'atms', 'gross_npl_pct', 'net_npl_pct', 'car_pct']:
            if data.get(field) is not None:
                old_val = om.loc[idx[0], field]
                if pd.isna(old_val):
                    om.loc[idx[0], field] = data[field]
                    updated += 1
                    print(f"  FY{fy}: {field} = {data[field]}")
    
    # Save
    om.to_excel(DATA / "04_operating_metrics.xlsx", sheet_name="operating_metrics", index=False)
    print(f"\nUpdated {updated} cells in 04_operating_metrics.xlsx")
    
    # Print summary
    print("\n" + "=" * 60)
    print("FINAL NABIL DATA")
    print("=" * 60)
    nabil_om = om[om['bank_code'] == 'NABIL']
    for _, row in nabil_om.iterrows():
        print(f"  FY{row['fy']}: employees={row['employees']}, branches={row['branches']}, "
              f"NPL={row['gross_npl_pct']}, CAR={row['car_pct']}")


if __name__ == '__main__':
    main()
