# 🏦 Nepal Commercial Banking Research Platform (FY2020 – FY2025)

[![CI Pipeline](https://github.com/Ninjafunz/nepal-banking-research/actions/workflows/ci.yml/badge.svg)](https://github.com/Ninjafunz/nepal-banking-research/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Code Style: Ruff](https://img.shields.io/badge/code%20style-ruff-000000.svg)](https://github.com/astral-sh/ruff)

A benchmark-quality empirical research dataset, automated calculation pipeline, and interactive analytical dashboard for the commercial banking sector in Nepal across **FY2020 to FY2025** (mid-July 2019 to mid-July 2025).

---

## 📌 Key Highlights

- **Complete Sample:** All 23 Class A commercial banks licensed by Nepal Rastra Bank (NRB) across FY2020–FY2025 (accounting for the 2022–2023 big merger wave down to 20 surviving banks).
- **136 Research Indicators:** Balance sheet items, income statement metrics, profitability/efficiency ratios, loan & deposit mix, macro controls, market valuations, and qualitative digital maturity scores.
- **Multi-Format Research Export:** Master panel auto-exported as **Excel (`.xlsx`)**, **CSV (`.csv`)**, **Apache Parquet (`.parquet`)**, and **Stata 13+ (`.dta`)**.
- **Publication-Ready Tables:** Auto-generates LaTeX (`.tex`) and CSV summary statistics formatted for academic journals (*NRB Economic Review*, *Journal of Banking & Finance*).
- **Automated Validation:** Rigorous accounting identities ($Assets = Liabilities + Equity$), Basel III regulatory checks (CAR $\ge 8.5\%$), and economic boundary validation.
- **Interactive Web App:** Full Streamlit + Plotly visual dashboard for exploratory data analysis.

---

## 📂 Repository Architecture

```
nepal-banking-research/
├── config.yaml                        # Central research configuration (years, units, HHI conventions)
├── pyproject.toml                     # Modern standard Python packaging & dependencies
├── bank_registry.csv                  # Complete 23-bank master registry & merger chronology
├── data_dictionary.md                 # Full documentation of all 136 panel variables
├── Makefile                           # Task runner for Linux / macOS
├── tasks.ps1                          # Task runner for Windows PowerShell
├── app.py                             # Interactive Streamlit Web Dashboard
│
├── config/
│   └── config_loader.py               # Typed configuration loader
│
├── DATA/
│   ├── 01_bank_financials.xlsx       # Core balance sheet & income statement
│   ├── 02_bank_ratios.xlsx           # Financial, growth & efficiency ratios
│   ├── 03_market_shares.xlsx         # Bank market shares & rankings
│   ├── 04_operating_metrics.xlsx     # Branch networks, staff, NPL%, CAR%
│   ├── 05_loan_composition.xlsx      # Sectoral & segmented loan books
│   ├── 06_deposit_composition.xlsx   # Current, Savings, Fixed deposits & CASA%
│   ├── 07_macro_indicators.xlsx      # World Bank GDP, CPI inflation, remittance & policy rates
│   ├── 08_market_data.xlsx           # NEPSE share prices, P/E, P/B, EPS, Market Cap
│   ├── 09_strategic_coding.xlsx      # Digital capability scorecard & strategic priorities
│   └── 10_bank_events.xlsx           # Documented M&A, tech launches & restructuring events
│
├── MASTER/
│   ├── master_bank_panel.xlsx        # Master analytical panel (Excel)
│   ├── master_bank_panel.csv         # Master analytical panel (CSV)
│   ├── master_bank_panel.parquet     # Columnar format for Python & R
│   └── master_bank_panel.dta         # Stata 13+ format for econometric analysis
│
├── ANALYSIS/
│   ├── industry_structure.xlsx       # HHI & CR4/CR5/CR10 concentration trends
│   ├── summary_statistics.csv        # Full descriptive statistics
│   ├── summary_statistics.tex        # LaTeX table for academic papers
│   └── annual_mean_trajectory.csv    # Annual metric trajectories
│
├── SCRIPTS/
│   ├── banks.py                      # Bank entity mapping & fiscal year bounds
│   ├── calculate_ratios.py           # Financial ratio derivation engine
│   ├── calculate_market_shares.py    # Market share & HHI concentration calculator
│   ├── build_panel.py                # Master multi-format builder & validator
│   ├── validators.py                 # Accounting identity & boundary validator
│   ├── generate_descriptives.py      # LaTeX & CSV descriptive table generator
│   ├── fetch_macro.py                # Live World Bank API macro extractor
│   └── setup_templates.py            # Blank template builder
│
├── tests/
│   ├── test_math.py                  # Unit tests for division, growth & average balance
│   ├── test_hhi.py                   # Analytical tests for HHI & concentration
│   ├── test_validators.py            # Unit tests for balance sheet & ratio checks
│   ├── test_pipeline.py              # End-to-end panel build test
│   └── generate_mock_data.py         # Synthetic toy data generator
│
└── notebooks/
    └── 01_exploratory_analysis.ipynb # Visual exploratory analysis notebook
```

---

## ⚡ Quickstart

### 1. Installation

```bash
# Clone the repository
git clone https://github.com/Ninjafunz/nepal-banking-research.git
cd nepal-banking-research

# Install in editable mode with development dependencies
pip install -e ".[dev]"
```

### 2. Run the Full Analytical Pipeline

```bash
# Linux / macOS
make build

# Windows PowerShell
.\tasks.ps1 build
```

This sequentially:
1. Calculates all financial ratios and productivity metrics (`02_bank_ratios.xlsx`)
2. Computes market shares, rankings, and HHI concentration (`03_market_shares.xlsx`)
3. Validates accounting identities and regulatory bounds (`validators.py`)
4. Assembles and exports the master panel in **Excel, CSV, Parquet, and Stata** (`MASTER/`)
5. Generates summary statistics and LaTeX tables (`ANALYSIS/summary_statistics.tex`)

### 3. Run Test Suite

```bash
pytest tests/ -v
```

### 4. Launch the Web Dashboard

```bash
streamlit run app.py
```

---

## 🏛️ Bank Coverage & Consolidation History

The dataset spans all 23 commercial banks in Nepal across the study period, tracking the consolidation down to 20 banks:

| Bank Code | Bank Name | Ownership | Listed | Consolidation / Merger Note |
|---|---|---|---|---|
| **NBL** | Nepal Bank Limited | State-Owned | Yes | Pioneer bank (est. 1937) |
| **RBB** | Rastriya Banijya Bank | State-Owned | No | 100% Government-owned; merged NIDC in 2018 |
| **ADBL** | Agriculture Development Bank | State-Owned | Yes | Specialized agricultural & rural credit |
| **NABIL** | Nabil Bank | Private | Yes | Acquired Nepal Bangladesh Bank (NBB) in FY2022 |
| **NIMB** | Nepal Investment Mega Bank | Private | Yes | Formed from merger of NIBL & Mega Bank in FY2023 |
| **SCB** | Standard Chartered Bank Nepal | Multinational | Yes | Subsidiary of Standard Chartered PLC |
| **HBL** | Himalayan Bank | Private | Yes | Acquired Civil Bank (CIVIL) in FY2023 |
| **SBI** | Nepal SBI Bank | Joint Venture | Yes | Joint venture with State Bank of India |
| **EBL** | Everest Bank | Joint Venture | Yes | Joint venture with Punjab National Bank India |
| **GIBL** | Global IME Bank | Private | Yes | Merged with Janata Bank (FY2020) & Bank of Kathmandu (FY2023) |
| **PRVU** | Prabhu Bank | Private | Yes | Merged with Century Commercial Bank (CCBL) in FY2023 |
| **KBL** | Kumari Bank | Private | Yes | Merged with NCC Bank in FY2023 |
| **LLBS** | Laxmi Sunrise Bank | Private | Yes | Formed from merger of Laxmi Bank & Sunrise Bank in FY2024 |
| **NICA** | NIC Asia Bank | Private | Yes | Standalone retail leader |
| **SANIMA** | Sanima Bank | Private | Yes | Standalone |
| **SBL** | Siddhartha Bank | Private | Yes | Standalone |
| **PCBL** | Prime Commercial Bank | Private | Yes | Standalone |
| **CZBIL** | Citizens Bank International | Private | Yes | Standalone |
| **NMB** | NMB Bank | Private | Yes | Standalone |
| **MBL** | Machhapuchchhre Bank | Private | Yes | Standalone |
| *BOKL* | Bank of Kathmandu | Private | Yes | *Merged into Global IME Bank in FY2023* |
| *CIVIL* | Civil Bank | Private | Yes | *Acquired by Himalayan Bank in FY2023* |
| *CCBL* | Century Commercial Bank | Private | Yes | *Merged into Prabhu Bank in FY2023* |

---

## 📄 License & Citation

Distributed under the **MIT License**.

If you use this dataset or code in academic research, please cite:
```bibtex
@misc{khadka2025nepalbanking,
  author = {Khadka, Manjil},
  title = {Nepal Commercial Banking Research Platform (FY2020--FY2025): Master Panel, Analytical Pipeline, and Market Concentration Analysis},
  year = {2025},
  publisher = {GitHub},
  journal = {GitHub repository},
  howpublished = {\url{https://github.com/Ninjafunz/nepal-banking-research}}
}
```
