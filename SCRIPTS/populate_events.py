import sys, os
import pandas as pd
import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "DATA")

events_data = [
    {
        "event_id": "EV001",
        "bank_code": "GIBL",
        "bank_name": "Global IME Bank",
        "fy": 2020,
        "event_date": "2019-12-06",
        "event_type": "M&A",
        "event_description": "Merger with Janata Bank Nepal Limited",
        "strategic_impact": "High",
        "counterparty": "Janata Bank Nepal",
        "financial_effect_notes": "Significantly expanded asset base, paid-up capital and branch network across Nepal.",
        "source_url": "https://www.nrb.org.np",
        "notes": "First major 'Big Merger' in Nepali commercial banking."
    },
    {
        "event_id": "EV002",
        "bank_code": "NABIL",
        "bank_name": "Nabil Bank",
        "fy": 2022,
        "event_date": "2022-07-11",
        "event_type": "M&A",
        "event_description": "Acquisition of Nepal Bangladesh Bank Limited (NBB)",
        "strategic_impact": "High",
        "counterparty": "Nepal Bangladesh Bank",
        "financial_effect_notes": "Expanded retail deposit base, branch footprint, and consolidated market leadership in balance sheet size.",
        "source_url": "https://www.nabilbank.com",
        "notes": "Swap ratio 100:43 (NABIL:NBB)."
    },
    {
        "event_id": "EV003",
        "bank_code": "GIBL",
        "bank_name": "Global IME Bank",
        "fy": 2023,
        "event_date": "2023-01-09",
        "event_type": "M&A",
        "event_description": "Merger with Bank of Kathmandu Limited (BOKL)",
        "strategic_impact": "High",
        "counterparty": "Bank of Kathmandu",
        "financial_effect_notes": "Created the largest commercial bank in Nepal by capital fund, asset size, and total deposits.",
        "source_url": "https://globalimebank.com",
        "notes": "Merged entity retained Global IME Bank name."
    },
    {
        "event_id": "EV004",
        "bank_code": "PRVU",
        "bank_name": "Prabhu Bank",
        "fy": 2023,
        "event_date": "2023-01-10",
        "event_type": "M&A",
        "event_description": "Merger with Century Commercial Bank Limited (CCBL)",
        "strategic_impact": "High",
        "counterparty": "Century Commercial Bank",
        "financial_effect_notes": "Expanded geographical distribution and retail deposit book.",
        "source_url": "https://prabhubank.com",
        "notes": "Swap ratio 1:1."
    },
    {
        "event_id": "EV005",
        "bank_code": "NIMB",
        "bank_name": "Nepal Investment Mega Bank",
        "fy": 2023,
        "event_date": "2023-01-11",
        "event_type": "M&A",
        "event_description": "Merger between Nepal Investment Bank Ltd (NIBL) and Mega Bank Nepal Ltd",
        "strategic_impact": "High",
        "counterparty": "Mega Bank Nepal",
        "financial_effect_notes": "Substantial capital boost, combined corporate lending power and retail branch strength.",
        "source_url": "https://nibl.com.np",
        "notes": "Operating under new name NIMB."
    },
    {
        "event_id": "EV006",
        "bank_code": "KBL",
        "bank_name": "Kumari Bank",
        "fy": 2023,
        "event_date": "2023-01-01",
        "event_type": "M&A",
        "event_description": "Merger with Nepal Credit and Commerce (NCC) Bank",
        "strategic_impact": "High",
        "counterparty": "NCC Bank",
        "financial_effect_notes": "Strengthened capital base, improved liquidity buffer, and broadened suburban outreach.",
        "source_url": "https://kumaribank.com",
        "notes": "Combined operations began Jan 2023."
    },
    {
        "event_id": "EV007",
        "bank_code": "HBL",
        "bank_name": "Himalayan Bank",
        "fy": 2023,
        "event_date": "2023-02-24",
        "event_type": "M&A",
        "event_description": "Acquisition of Civil Bank Limited",
        "strategic_impact": "High",
        "counterparty": "Civil Bank",
        "financial_effect_notes": "Strengthened market share in Tier 2/3 cities and expanded SME lending portfolio.",
        "source_url": "https://himalayanbank.com",
        "notes": "Civil Bank ceased standalone operations."
    },
    {
        "event_id": "EV008",
        "bank_code": "LLBS",
        "bank_name": "Laxmi Sunrise Bank",
        "fy": 2024,
        "event_date": "2023-07-14",
        "event_type": "M&A",
        "event_description": "Merger between Laxmi Bank Limited and Sunrise Bank Limited",
        "strategic_impact": "High",
        "counterparty": "Sunrise Bank",
        "financial_effect_notes": "Consolidated balance sheet over NPR 300 billion with unified digital banking infrastructure.",
        "source_url": "https://laxmisunrisebank.com",
        "notes": "Swap ratio 1:1, commenced joint operation at start of FY2024."
    },
    {
        "event_id": "EV009",
        "bank_code": "NICA",
        "bank_name": "NIC Asia Bank",
        "fy": 2021,
        "event_date": "2020-09-15",
        "event_type": "Technology",
        "event_description": "Rollout of unified iServe and Digital 360 platform",
        "strategic_impact": "Medium",
        "counterparty": "Internal / Fintech partners",
        "financial_effect_notes": "Drove massive retail customer acquisition and digital transaction volume growth.",
        "source_url": "https://nicasiabank.com",
        "notes": "Pioneered paperless branch customer service in Nepal."
    },
    {
        "event_id": "EV010",
        "bank_code": "SCB",
        "bank_name": "Standard Chartered Bank Nepal",
        "fy": 2024,
        "event_date": "2023-11-20",
        "event_type": "Technology",
        "event_description": "Launch of Straight2Bank NextGen for Corporate Clients",
        "strategic_impact": "Medium",
        "counterparty": "Standard Chartered Group",
        "financial_effect_notes": "Maintained dominant fee-income market share in multinational and institutional cash management.",
        "source_url": "https://sc.com/np",
        "notes": "Corporate and institutional digital cash management."
    }
]

df_events = pd.DataFrame(events_data)
events_path = os.path.join(DATA, "10_bank_events.xlsx")
with pd.ExcelWriter(events_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_events.to_excel(writer, sheet_name="events", index=False)
print(f"Updated: {events_path} with {len(df_events)} events.")

