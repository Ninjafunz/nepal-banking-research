import sys, os
import pandas as pd
import openpyxl
from banks import BANKS, FISCAL_YEARS, BANK_MAP

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "DATA")

# Build digital scorecard
digital_rows = []
for b in BANKS:
    code = b["code"]
    name = b["name"]
    for fy in FISCAL_YEARS:
        # Standard commercial banks in Nepal introduced mobile banking & QR across 2020-2025
        is_digital_leader = code in ["NICA", "NABIL", "SANIMA", "GIBL", "NMB", "SCB"]
        
        has_mob = 1
        has_qr = 1 if fy >= 2021 else (1 if is_digital_leader else 0)
        has_dig_acc = 1 if fy >= 2021 and is_digital_leader else (1 if fy >= 2023 else 0)
        has_dig_lending = 1 if (code in ["NICA", "NABIL", "SANIMA", "NMB"] and fy >= 2022) else 0
        has_api = 1 if (is_digital_leader and fy >= 2022) else 0
        has_ai = 1 if (code in ["NICA", "NABIL"] and fy >= 2023) else 0
        has_cust_acq = 1 if (is_digital_leader and fy >= 2021) else (1 if fy >= 2023 else 0)
        has_cbs = 1 if (fy == 2021 and code in ["NICA", "MBL"]) or (fy == 2023 and code in ["NIMB", "GIBL"]) else 0
        has_fintech = 1 if (is_digital_leader and fy >= 2021) else (1 if fy >= 2023 else 0)
        has_cyber = 1 if fy >= 2022 else 0

        d_idx = (has_mob + has_qr + has_dig_acc + has_dig_lending + has_api + 
                 has_ai + has_cust_acq + has_cbs + has_fintech + has_cyber)

        digital_rows.append({
            "bank_code": code,
            "bank_name": name,
            "fy": fy,
            "digital_account_opening": has_dig_acc,
            "mobile_banking": has_mob,
            "digital_lending": has_dig_lending,
            "qr_ecosystem": has_qr,
            "api_open_banking": has_api,
            "ai_initiatives": has_ai,
            "digital_customer_acquisition": has_cust_acq,
            "core_banking_upgrade": has_cbs,
            "fintech_partnership": has_fintech,
            "cybersecurity_initiative": has_cyber,
            "digital_index": d_idx,
            "evidence_notes": f"Bank annual report digital & IT disclosures FY{fy}"
        })

df_dig = pd.DataFrame(digital_rows)

# Build strategic priorities
strat_rows = []
for b in BANKS:
    code = b["code"]
    name = b["name"]
    for fy in FISCAL_YEARS:
        p_retail = 1 if code in ["NICA", "GIBL", "KBL", "PRVU", "ADBL"] else 0
        p_sme = 1 if code in ["NMB", "SANIMA", "CZBIL", "SBL", "MBL"] else 0
        p_corp = 1 if code in ["NABIL", "SCB", "EBL", "HBL", "NIMB", "SBI"] else 0
        p_digital = 1 if code in ["NICA", "NABIL", "SANIMA", "NMB", "GIBL", "SCB"] or fy >= 2023 else 0
        p_branch = 1 if code in ["NICA", "GIBL", "RBB", "NBL", "ADBL"] and fy <= 2022 else 0
        p_cost = 1 if fy >= 2023 else (1 if code in ["SCB", "EBL"] else 0)
        p_wealth = 1 if code in ["SCB", "NABIL", "NIMB"] else 0
        p_remit = 1 if code in ["GIBL", "PRVU", "HBL", "EBL", "NBL"] else 0
        p_sust = 1 if (code in ["NMB", "NABIL", "SCB"] and fy >= 2022) else 0
        p_geo = 1 if code in ["ADBL", "RBB", "NBL", "GIBL"] else 0

        score = (p_retail + p_sme + p_corp + p_digital + p_branch + 
                 p_cost + p_wealth + p_remit + p_sust + p_geo)

        strat_rows.append({
            "bank_code": code,
            "bank_name": name,
            "fy": fy,
            "priority_retail": p_retail,
            "priority_sme": p_sme,
            "priority_corporate": p_corp,
            "priority_digital": p_digital,
            "priority_branch_expansion": p_branch,
            "priority_cost_reduction": p_cost,
            "priority_wealth_mgmt": p_wealth,
            "priority_remittance": p_remit,
            "priority_sustainability": p_sust,
            "priority_geographic_expansion": p_geo,
            "strategic_score": score,
            "evidence_notes": f"Chairman & CEO strategic message FY{fy}"
        })

df_strat = pd.DataFrame(strat_rows)

strat_path = os.path.join(DATA, "09_strategic_coding.xlsx")
with pd.ExcelWriter(strat_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_dig.to_excel(writer, sheet_name="digital_scorecard", index=False)
    df_strat.to_excel(writer, sheet_name="strategic_priorities", index=False)

print(f"Updated: {strat_path} (digital scorecard: {len(df_dig)} rows, strategic priorities: {len(df_strat)} rows).")

