"""
generate_descriptives.py — Generates publication-grade descriptive statistics tables (LaTeX & CSV).
"""

import os
import sys

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

ANALYSIS = os.path.join(BASE, "ANALYSIS")
MASTER = os.path.join(BASE, "MASTER")


def generate_summary_tables():
    os.makedirs(ANALYSIS, exist_ok=True)
    panel_path = os.path.join(MASTER, "master_bank_panel.csv")
    if not os.path.exists(panel_path):
        panel_path = os.path.join(MASTER, "master_bank_panel.xlsx")
        df = pd.read_excel(panel_path, sheet_name="panel")
    else:
        df = pd.read_csv(panel_path)

    # Core continuous variables for descriptive analysis
    vars_to_summarize = {
        "total_assets": "Total Assets (NPR M)",
        "gross_loans": "Gross Loans (NPR M)",
        "total_deposits": "Total Deposits (NPR M)",
        "shareholders_equity": "Shareholders' Equity (NPR M)",
        "profit_after_tax": "Net Profit PAT (NPR M)",
        "roa": "Return on Assets ROA (%)",
        "roe": "Return on Equity ROE (%)",
        "nim": "Net Interest Margin NIM (%)",
        "cost_income": "Cost-to-Income Ratio (%)",
        "gross_npl_pct": "Gross NPL Ratio (%)",
        "car_pct": "Capital Adequacy Ratio CAR (%)",
        "casa_ratio": "CASA Deposit Ratio (%)",
        "digital_index": "Digital Maturity Index (0-10)"
    }

    # Filter variables present
    avail_vars = [v for v in vars_to_summarize if v in df.columns]

    # Overall Summary Statistics
    summary_rows = []
    for v in avail_vars:
        s = df[v].dropna()
        if s.empty:
            continue
        summary_rows.append({
            "Variable": vars_to_summarize[v],
            "Obs": len(s),
            "Mean": round(s.mean(), 2),
            "Std Dev": round(s.std(), 2),
            "Min": round(s.min(), 2),
            "P25": round(s.quantile(0.25), 2),
            "Median": round(s.median(), 2),
            "P75": round(s.quantile(0.75), 2),
            "Max": round(s.max(), 2)
        })

    summary_df = pd.DataFrame(summary_rows)

    # Export to CSV
    csv_out = os.path.join(ANALYSIS, "summary_statistics.csv")
    summary_df.to_csv(csv_out, index=False)
    print(f"  [Summary Stats CSV]   {csv_out}")

    # Export to LaTeX Table
    tex_out = os.path.join(ANALYSIS, "summary_statistics.tex")
    with open(tex_out, "w", encoding="utf-8") as f:
        f.write("% =========================================================================\n")
        f.write("% Summary Statistics — Nepal Commercial Banking Panel (FY2020 - FY2025)\n")
        f.write("% =========================================================================\n")
        f.write("\\begin{table}[htbp]\n\\centering\n")
        f.write("\\caption{Descriptive Statistics of Key Financial and Operating Variables}\n")
        f.write("\\label{tab:summary_stats}\n")
        f.write("\\begin{tabular}{lrrrrrrrr}\n\\hline\\hline\n")
        f.write("Variable & N & Mean & SD & Min & P25 & Median & P75 & Max \\\\\n\\hline\n")
        f.writelines(f"{r['Variable']} & {r['Obs']} & {r['Mean']} & {r['Std Dev']} & {r['Min']} & {r['P25']} & {r['Median']} & {r['P75']} & {r['Max']} \\\\\n" for _, r in summary_df.iterrows())
        f.write("\\hline\\hline\n\\end{tabular}\n")
        f.write("\\begin{tablenotes}\n\\small\n\\item \\textit{Notes:} Sample covers 23 licensed Class A commercial banks in Nepal across FY2020--FY2025. Monetary values in NPR millions. Ratios in percent.\n\\end{tablenotes}\n")
        f.write("\\end{table}\n")
    print(f"  [Summary Stats LaTeX] {tex_out}")

    # By-Year Summary (Means)
    by_year = df.groupby("fy")[avail_vars].mean().round(2).reset_index()
    by_year.rename(columns=vars_to_summarize, inplace=True)
    by_year_out = os.path.join(ANALYSIS, "annual_mean_trajectory.csv")
    by_year.to_csv(by_year_out, index=False)
    print(f"  [Annual Trajectory]   {by_year_out}")


if __name__ == "__main__":
    generate_summary_tables()
