# Nepal Banking Research — Data Architecture

A reproducible multi-dataset research infrastructure for analysing Nepal's
commercial banking sector (FY2020–FY2025) across 23 Class A commercial banks.

## Quick Start

```bash
cd SCRIPTS

# 1. Install dependencies
pip install -r requirements.txt

# 2. Create all Excel templates (run once)
python setup_templates.py

# 3. Fill in raw data — start with DATA/01_bank_financials.xlsx
#    Sources: NRB Banking & Financial Statistics, bank annual reports

# 4. Regenerate ratios + market shares
python calculate_ratios.py
python calculate_market_shares.py

# 5. Build master panel (runs steps 4 automatically)
python build_panel.py
```

## Project Structure

```
nepal-banking-research/
├── DATA/
│   ├── raw/                          <- Source files (NRB, annual reports)
│   ├── 01_bank_financials.xlsx       <- FILL IN (Stage 1)
│   ├── 02_bank_ratios.xlsx           <- AUTO-GENERATED
│   ├── 03_market_shares.xlsx         <- AUTO-GENERATED
│   ├── 04_operating_metrics.xlsx     <- FILL IN (Stage 2)
│   ├── 05_loan_composition.xlsx      <- FILL IN (Stage 2)
│   ├── 06_deposit_composition.xlsx   <- FILL IN (Stage 2)
│   ├── 07_macro_indicators.xlsx      <- FILL IN (Stage 3)
│   ├── 08_market_data.xlsx           <- FILL IN (Stage 4)
│   ├── 09_strategic_coding.xlsx      <- FILL IN (Stage 5)
│   └── 10_bank_events.xlsx           <- FILL IN (Stage 5)
├── MASTER/
│   └── master_bank_panel.xlsx        <- AUTO-GENERATED (final output)
├── ANALYSIS/
│   └── industry_structure.xlsx       <- AUTO-GENERATED (HHI, CR4/CR5/CR10)
└── SCRIPTS/
    ├── banks.py                      <- Bank registry (23 Class A banks)
    ├── setup_templates.py            <- Creates all Excel templates
    ├── calculate_ratios.py           <- Generates 02_bank_ratios.xlsx
    ├── calculate_market_shares.py    <- Generates 03_market_shares.xlsx
    ├── build_panel.py                <- Master pipeline
    └── requirements.txt
```

## Banks Covered (23 NRB Class A Commercial Banks)

| Code   | Bank Name                        | State-owned |
|--------|----------------------------------|-------------|
| NBL    | Nepal Bank Limited               | Yes         |
| RBB    | Rastriya Banijya Bank            | Yes         |
| ADBL   | Agriculture Development Bank     | Yes         |
| NABIL  | Nabil Bank                       | No          |
| NIMB   | Nepal Investment Mega Bank       | No          |
| SCB    | Standard Chartered Bank Nepal    | No          |
| HBL    | Himalayan Bank                   | No          |
| SBI    | Nepal SBI Bank                   | No          |
| EBL    | Everest Bank                     | No          |
| BOKL   | Bank of Kathmandu Lumbini        | No          |
| NICA   | NIC Asia Bank                    | No          |
| MBL    | Machhapuchchhre Bank             | No          |
| KBL    | Kumari Bank                      | No          |
| LLBS   | Laxmi Sunrise Bank               | No          |
| CIVIL  | Civil Bank                       | No          |
| CCBL   | Century Commercial Bank          | No          |
| SANIMA | Sanima Bank                      | No          |
| SBL    | Siddhartha Bank                  | No          |
| GIBL   | Global IME Bank                  | No          |
| PCBL   | Prime Commercial Bank            | No          |
| PRVU   | Prabhu Bank                      | No          |
| CZBIL  | Citizens Bank International      | No          |
| NMB    | NMB Bank                         | No          |

## Data Conventions

| Convention     | Value               |
|----------------|---------------------|
| Monetary unit  | NPR millions        |
| Fiscal year    | Gregorian year-end  |
| ROA/ROE/NIM    | Average-balance basis |
| HHI convention | % squared (max 10,000) |
| NPL            | % of gross loans    |

## Data Sources

| Dataset | Primary source | Secondary source |
|---------|----------------|-----------------|
| 01 Bank financials | NRB Banking & Financial Statistics | Annual reports |
| 04 Operating metrics | NRB, Annual reports | — |
| 05-06 Composition | Annual reports | NRB sector data |
| 07 Macro | NRB, CBS, MOF | World Bank |
| 08 Market data | NEPSE, Merolagani | Annual reports |
| 09 Strategic | Annual reports | — |
| 10 Events | Annual reports, press | NRB notices |

## Data Collection Stages

| Stage | Datasets | What it unlocks |
|-------|----------|-----------------|
| 1 — MVD | 01 | ROA, ROE, NIM, growth, HHI, rankings |
| 2 — Depth | + 04, 05, 06 | Business model, efficiency, funding |
| 3 — Macro | + 07 | Regression controls, scenario analysis |
| 4 — Market | + 08 | Investor signals, P/B validation |
| 5 — Qualitative | + 09, 10 | Digital index, event studies |
