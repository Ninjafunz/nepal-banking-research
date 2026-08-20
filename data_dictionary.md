# Nepal Banking Research Platform — Data Dictionary

This data dictionary provides technical definitions, measurement units, mathematical formulas, and data source mappings for all variables in the **Master Bank-Year Panel** (`master_bank_panel.xlsx`, `.csv`, `.parquet`, `.dta`).

---

## 1. Identifiers & Entity Attributes

| Variable | Type | Description | Measurement Unit | Source / Mapping |
|---|---|---|---|---|
| `bank_code` | String | NRB / NEPSE Standard bank ticker symbol | — | Central Bank licensed code |
| `bank_name` | String | Full legal corporate name of the bank | — | NRB Bank Registry |
| `fy` | Integer | Ending Gregorian year of Nepal Fiscal Year (e.g. 2024 = 2080/81 BS) | Year | NRB Calendar |
| `listed` | Boolean | True if bank shares trade on Nepal Stock Exchange (NEPSE) | Boolean | NEPSE |
| `state_owned` | Boolean | True if majority government/public ownership (NBL, RBB, ADBL) | Boolean | NRB Ownership Profile |

---

## 2. Core Balance Sheet (NPR Millions)

| Variable | Type | Description | Formula / Accounting Definition |
|---|---|---|---|
| `total_assets` | Float | Total consolidated balance sheet assets | Cash + Investments + Net Loans + Other Assets |
| `cash_bank_balances` | Float | Cash in hand + balances with NRB & domestic/foreign banks | Line item in NRB Statement 1 |
| `investments` | Float | Government treasury bills, development bonds & securities | Mark-to-market / amortized cost securities |
| `gross_loans` | Float | Total loans and advances before provisions | Total loan portfolio at principal |
| `net_loans` | Float | Total loans and advances after loan loss provisions | `gross_loans - provisions` |
| `total_deposits` | Float | Total customer deposits across all liability types | Current + Savings + Fixed + Call + Other Deposits |
| `borrowings` | Float | Interbank, repo, and institutional borrowings | NRB Statement 1 Liabilities |
| `total_liabilities` | Float | Total liabilities excluding shareholder equity | `total_assets - shareholders_equity` |
| `shareholders_equity`| Float | Total Tier 1 capital + reserves & surplus | Paid-up Capital + General & Free Reserves |
| `paid_up_capital` | Float | Issued and fully paid-up equity capital | Equity share capital at par (NPR 100) |
| `reserves` | Float | Retained earnings, statutory general reserves, capital reserves | `shareholders_equity - paid_up_capital` |

---

## 3. Core Income Statement (NPR Millions)

| Variable | Type | Description | Formula |
|---|---|---|---|
| `interest_income` | Float | Total gross interest earned on loans and investments | Reported Interest Income |
| `interest_expense` | Float | Total interest paid on deposits and borrowings | Reported Interest Expense |
| `net_interest_income` | Float | Core lending revenue margin (NII) | `interest_income - interest_expense` |
| `non_interest_income` | Float | Fees, commissions, FX trading gains, dividends | Non-interest operating revenue |
| `operating_income` | Float | Total net operating income | `net_interest_income + non_interest_income` |
| `operating_expenses` | Float | Total administrative, staff, and general overhead costs | Personnel + General Admin + Depreciation |
| `personnel_expenses` | Float | Salaries, bonuses, retirement benefits, staff training | Employee benefit expenses |
| `provision_loan_losses`| Float | Loan loss impairment charge against profit | Provisions for pass, watch, sub, doubtful & loss |
| `profit_before_tax` | Float | Profit before statutory corporate taxes (PBT) | `operating_income - operating_expenses - provision_loan_losses` |
| `profit_after_tax` | Float | Net profit available to shareholders (PAT) | `profit_before_tax - income_tax_expense` (approx 30% tax) |

---

## 4. Financial & Operating Ratios

| Variable | Unit | Description | Analytical Formula |
|---|---|---|---|
| `roa` | % | Return on Assets (average balance basis) | `(profit_after_tax / avg_total_assets) * 100` |
| `roe` | % | Return on Equity (average balance basis) | `(profit_after_tax / avg_shareholders_equity) * 100` |
| `nim` | % | Net Interest Margin | `(net_interest_income / avg_earning_assets) * 100` |
| `profit_margin` | % | Net profit margin on total operating income | `(profit_after_tax / operating_income) * 100` |
| `cost_income` | % | Cost-to-Income Efficiency Ratio | `(operating_expenses / operating_income) * 100` |
| `asset_growth` | % YoY | Year-over-year percentage growth in total assets | `((total_assets_t / total_assets_t-1) - 1) * 100` |
| `loan_growth` | % YoY | Year-over-year percentage growth in gross loans | `((gross_loans_t / gross_loans_t-1) - 1) * 100` |
| `deposit_growth` | % YoY | Year-over-year percentage growth in total deposits | `((total_deposits_t / total_deposits_t-1) - 1) * 100` |
| `gross_npl_pct` | % | Non-Performing Loan ratio to gross loans | `(non_performing_loans / gross_loans) * 100` |
| `net_npl_pct` | % | Net NPL after specific provisions | `(net_npl / net_loans) * 100` |
| `provision_coverage_pct`| % | Total loan loss provisions to gross NPL | `(total_provisions / gross_npl) * 100` |
| `car_pct` | % | Capital Adequacy Ratio (Basel III total capital) | `(total_capital_fund / risk_weighted_assets) * 100` |
| `cet1_pct` | % | Common Equity Tier 1 Ratio | `(tier_1_capital / risk_weighted_assets) * 100` |
| `loan_deposit_ratio` | % | Credit-to-Deposit (CD) ratio | `(gross_loans / total_deposits) * 100` |
| `casa_ratio` | % | Low-cost deposit ratio | `((current_deposits + savings_deposits) / total_deposits) * 100` |
| `assets_per_employee` | NPR M | Asset productivity per FTE staff | `total_assets / employees` |
| `profit_per_employee` | NPR M | Net profit generated per FTE staff | `profit_after_tax / employees` |

---

## 5. Market Position & Concentration

| Variable | Unit | Description | Formula |
|---|---|---|---|
| `asset_share_pct` | % | Bank's share of total commercial banking assets | `(bank_assets / system_assets) * 100` |
| `loan_share_pct` | % | Bank's share of total commercial banking credit | `(bank_loans / system_loans) * 100` |
| `deposit_share_pct`| % | Bank's share of total commercial banking deposits | `(bank_deposits / system_deposits) * 100` |
| `hhi_assets` | Index | Herfindahl-Hirschman Index for Asset market | `sum(asset_share_pct_i ^ 2)` (Scale: 0 - 10,000) |
| `cr4_assets` | % | Combined asset market share of top 4 banks | `sum(asset_share_pct of top 4 banks)` |
| `cr10_assets` | % | Combined asset market share of top 10 banks | `sum(asset_share_pct of top 10 banks)` |

---

## 6. Macroeconomic & Monetary Indicators

| Variable | Unit | Description | Source |
|---|---|---|---|
| `gdp_growth_pct` | % | Annual Real Gross Domestic Product Growth | World Bank (NY.GDP.MKTP.KD.ZG) |
| `inflation_pct` | % | Annual Consumer Price Index (CPI) Inflation | World Bank (FP.CPI.TOTL.ZG) |
| `remittance_usd_mn` | USD M | Personal Remittances received in USD Millions | World Bank (BX.TRF.PWKR.CD.DT) |
| `policy_rate_pct` | % | NRB Policy Repo / Standing Liquidity Facility Rate | NRB Monetary Policy |
| `bank_rate_pct` | % | NRB Bank Rate for lender of last resort | NRB Monetary Policy |
