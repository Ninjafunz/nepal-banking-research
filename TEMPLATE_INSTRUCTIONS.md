# TEMPLATE INSTRUCTIONS - How to Fill in Missing Data
# ====================================================
#
# This project has 3 scripts for filling the remaining NaN gaps in the
# master panel. Here is how to use them, step by step.
#
#
# STEP 1: Generate the blank template
# ------------------------------------
#
# Open a terminal in the project folder and run:
#
#     python generate_template.py
#
# This creates TEMPLATE_DATA.xlsx with 5 sheets:
#
#   Sheet 1: 1_employees_branches  (23 banks x 9 years = 207 rows)
#     - employees: total staff count at year-end
#     - branches: number of branches
#
#   Sheet 2: 2_share_prices  (23 banks x 6 years = 138 rows)
#     - share_price_eoy: NEPSE closing price on the last trading day
#       of the fiscal year (typically mid-July)
#
#   Sheet 3: 3_per_bank_npl  (23 banks x 9 years = 207 rows)
#     - gross_npl_pct: gross NPL as % of total loans
#     - net_npl_pct: net NPL as % of total loans
#     - provision_coverage_pct: provisions / gross NPL * 100
#
#   Sheet 4: 4_macro_indicators  (9 rows, one per FY)
#     - remittance_usd_mn: total remittance inflows in USD millions
#     (also pre-fills existing GDP/inflation/policy rate values)
#
#   Sheet 5: LEGEND - explains each column, its source, and units
#
#
# STEP 2: Fill in the template
# ----------------------------
#
# Open TEMPLATE_DATA.xlsx in Excel.
#
# IMPORTANT RULES:
#   - Do NOT change bank_code, bank_name, or fy values
#   - Only fill the BLANK cells (left as empty strings)
#   - Enter numbers only (no % sign, no commas in large numbers)
#   - Leave cells blank if you cant find the data - the script
#     only overwrites empty cells, so your existing real data is safe
#
# Where to find each data point:
#
#   employees        - Each bank Annual Report (Financial Highlights)
#                      or NRB Bankers Almanac (published yearly)
#   branches         - Same sources as employees
#   share_price_eoy  - NEPSE daily trading data:
#                        nepalstock.com (Downloads section)
#                        SmartCharts: nepalstock.com.np/smartcharts
#                      Use closing price on LAST trading day before
#                      mid-July of each year
#   gross_npl_pct    - Individual bank Annual Report, Note 15 or
#                      Risk Management section
#                      Formula: (Gross NPL / Total Loans) x 100
#   net_npl_pct      - Same source: (Gross NPL - Provisions) /
#                      Total Loans x 100
#   provision_cov    - Same source: Provisions / Gross NPL x 100
#   remittance_usd   - NRB Balance of Payments statistics:
#                        nrb.org.np (Financial Statistics)
#                        World Bank: data.worldbank.org
#                        indicator BX.TRF.PWKR.CD.DT (Nepal)
#
# TIP: Start with the easiest data. You dont have to fill everything at
# once - you can run the fill script multiple times as you collect data.
#
# FY NOTE: Nepal fiscal year runs mid-July to mid-July.
#   FY2020 = July 2019 to July 2020
#   FY2025 = July 2024 to July 2025
# The fy column uses the ENDING year.
#
#
# STEP 3: Run the fill script
# ---------------------------
#
# After filling in whatever data you have:
#
#     python fill_from_template.py
#
# This script:
#   1. Shows BEFORE coverage rates (what is currently filled)
#   2. Merges your filled data into the pipeline DATA files
#   3. Recomputes derived fields (EPS, BVPS, P/B, P/E)
#   4. Re-runs the full pipeline (ratios -> market shares -> panel -> descriptives)
#   5. Shows AFTER coverage rates with deltas
#
# Your data is SAFE: the script only updates cells that were previously
# blank. It never overwrites existing real data.
#
#
# STEP 4: Verify
# --------------
#
# After the script finishes:
#   - Check MASTER/master_bank_panel.xlsx
#   - Run: python SCRIPTS/generate_descriptives.py
#   - The script will print the final coverage stats
#
#
# QUICK REFERENCE: Bank codes
# ---------------------------
#
# NBL     Nepal Bank Limited (state-owned)
# RBB     Rastriya Banijya Bank (state-owned)
# ADBL    Agriculture Development Bank (state-owned)
# NABIL   Nabil Bank
# NIMB    Nepal Investment Mega Bank
# SCB     Standard Chartered Bank Nepal
# HBL     Himalayan Bank
# SBI     Nepal SBI Bank
# EBL     Everest Bank
# BOKL    Bank of Kathmandu (merged into GIBL in FY2023)
# NICA    NIC Asia Bank
# MBL     Machhapuchchhre Bank
# KBL     Kumari Bank
# LLBS    Laxmi Sunrise Bank (formed July 2023)
# CIVIL   Civil Bank (acquired by HBL in FY2023)
# CCBL    Century Commercial Bank (merged into PRVU in FY2023)
# SANIMA  Sanima Bank
# SBL     Siddhartha Bank
# GIBL    Global IME Bank
# PCBL    Prime Commercial Bank
# PRVU    Prabhu Bank
# CZBIL   Citizens Bank International
# NMB     NMB Bank
#
# Note: BOKL, CIVIL, CCBL are no longer active - you only need data
# for the years they existed independently.
