"""
generate_mock_data.py — Generates a minimal 3-bank, 3-year synthetic test fixture.
"""

import os
import sys
import pandas as pd
import numpy as np

def generate_synthetic_fixture(target_dir):
    os.makedirs(target_dir, exist_ok=True)
    banks = ["TEST_A", "TEST_B", "TEST_C"]
    years = [2022, 2023, 2024]
    
    bs_rows, is_rows = [], []
    for b_idx, b in enumerate(banks):
        for y_idx, y in enumerate(years):
            base = (b_idx + 1) * 100000 * (1 + 0.1 * y_idx)
            equity = base * 0.12
            loans = base * 0.72
            deposits = base * 0.82
            
            bs_rows.append({
                "bank_code": b, "bank_name": f"Test Bank {b}", "fy": y,
                "total_assets": base, "cash_bank_balances": base * 0.08,
                "investments": base * 0.16, "gross_loans": loans, "net_loans": loans * 0.98,
                "total_deposits": deposits, "borrowings": base * 0.03,
                "total_liabilities": base - equity, "shareholders_equity": equity,
                "paid_up_capital": equity * 0.7, "reserves": equity * 0.3,
                "source": "MockData", "notes": "Synthetic"
            })
            
            int_inc = loans * 0.10
            int_exp = deposits * 0.06
            nii = int_inc - int_exp
            op_inc = nii + (base * 0.01)
            op_exp = op_inc * 0.40
            pbt = op_inc - op_exp - (loans * 0.005)
            pat = pbt * 0.70
            
            is_rows.append({
                "bank_code": b, "bank_name": f"Test Bank {b}", "fy": y,
                "interest_income": int_inc, "interest_expense": int_exp,
                "net_interest_income": nii, "non_interest_income": base * 0.01,
                "operating_income": op_inc, "operating_expenses": op_exp,
                "personnel_expenses": op_exp * 0.6, "provision_loan_losses": loans * 0.005,
                "profit_before_tax": pbt, "profit_after_tax": pat,
                "source": "MockData", "notes": "Synthetic"
            })

    out_file = os.path.join(target_dir, "01_bank_financials.xlsx")
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        pd.DataFrame(bs_rows).to_excel(writer, sheet_name="balance_sheet", index=False)
        pd.DataFrame(is_rows).to_excel(writer, sheet_name="income_statement", index=False)
    return out_file

if __name__ == "__main__":
    t_dir = os.path.join(os.path.dirname(__file__), "data")
    generate_synthetic_fixture(t_dir)
