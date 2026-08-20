import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import os

# Set page configuration
st.set_page_config(
    page_title="Nepal Banking Research Platform",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border-radius: 8px;
        padding: 15px;
        border-left: 4px solid #2563EB;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)

# Data loader functions
@st.cache_data
def load_all_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "DATA")
    master_path = os.path.join(base_dir, "MASTER", "master_bank_panel.xlsx")
    ind_path = os.path.join(base_dir, "ANALYSIS", "industry_structure.xlsx")

    panel = pd.read_excel(master_path, sheet_name="panel")
    industry = pd.read_excel(ind_path, sheet_name="industry_structure")
    events = pd.read_excel(os.path.join(data_dir, "10_bank_events.xlsx"), sheet_name="events")
    
    return panel, industry, events

try:
    panel_df, ind_df, events_df = load_all_data()
except Exception as e:
    st.error(f"Error loading datasets: {e}")
    st.stop()

# Sidebar Navigation
st.sidebar.image("https://img.icons8.com/color/96/bank-building.png", width=64)
st.sidebar.title("Nepal Banking Research")
st.sidebar.markdown("**FY2020 – FY2025 Analytics**")

nav = st.sidebar.radio(
    "Navigation Menu",
    [
        "🏛️ Industry Overview",
        "📊 Bank Peer Comparison",
        "📈 Market Concentration (HHI)",
        "💼 Business Model (Loans & CASA)",
        "🌐 Macroeconomic Environment",
        "📱 Digital & Strategic Scorecard",
        "📋 Master Data Explorer"
    ]
)

# Year filter in sidebar
selected_fy = st.sidebar.slider("Select Fiscal Year", min_value=2020, max_value=2025, value=2024, step=1)
fy_filtered = panel_df[panel_df["fy"] == selected_fy]

# ─────────────────────────────────────────────────────────────────────────────
# 1. Industry Overview
# ─────────────────────────────────────────────────────────────────────────────
if nav == "🏛️ Industry Overview":
    st.markdown('<div class="main-header">Nepal Commercial Banking System</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="sub-header">Executive Performance Dashboard — <b>FY{selected_fy}</b> (All Monetary Values in NPR Millions)</div>', unsafe_allow_html=True)

    # Key Metrics
    tot_assets = fy_filtered["total_assets"].sum()
    tot_deposits = fy_filtered["total_deposits"].sum()
    tot_loans = fy_filtered["gross_loans"].sum()
    avg_roe = fy_filtered["roe"].mean()
    avg_roa = fy_filtered["roa"].mean()
    avg_npl = fy_filtered["gross_npl_pct"].mean()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total System Assets", f"NPR {tot_assets/1000:,.1f}B", f"{(tot_assets/panel_df[panel_df['fy']==selected_fy-1]['total_assets'].sum() - 1)*100:.1f}% YoY" if selected_fy > 2020 else None)
    c2.metric("Total Deposits", f"NPR {tot_deposits/1000:,.1f}B")
    c3.metric("Gross Loans", f"NPR {tot_loans/1000:,.1f}B")
    c4.metric("Average ROE", f"{avg_roe:.2f}%")
    c5.metric("Average NPL", f"{avg_npl:.2f}%")

    st.markdown("---")

    col_l, col_r = st.columns([3, 2])
    
    with col_l:
        st.subheader(f"Top 10 Commercial Banks by Asset Size (FY{selected_fy})")
        top10_assets = fy_filtered.sort_values("total_assets", ascending=False).head(10)
        fig_assets = px.bar(
            top10_assets,
            x="total_assets",
            y="bank_name",
            orientation="h",
            color="total_assets",
            color_continuous_scale="Blues",
            labels={"total_assets": "Total Assets (NPR Millions)", "bank_name": "Bank Name"},
            text_auto=",.0f"
        )
        fig_assets.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False, height=420)
        st.plotly_chart(fig_assets, use_container_width=True)

    with col_r:
        st.subheader(f"Profitability vs Scale (ROA vs Assets)")
        fig_scatter = px.scatter(
            fy_filtered,
            x="total_assets",
            y="roa",
            size="gross_loans",
            color="bank_code",
            hover_name="bank_name",
            labels={"total_assets": "Assets (NPR M)", "roa": "ROA (%)"},
            title=f"Asset Scale vs. Return on Assets (FY{selected_fy})"
        )
        fig_scatter.update_layout(showlegend=False, height=420)
        st.plotly_chart(fig_scatter, use_container_width=True)

    # 5-Year Aggregate Growth Chart
    st.subheader("5-Year Aggregate Growth Trend (FY2020 – FY2025)")
    trend_df = panel_df.groupby("fy")[["total_assets", "gross_loans", "total_deposits", "shareholders_equity"]].sum().reset_index()
    fig_trend = px.line(
        trend_df,
        x="fy",
        y=["total_assets", "gross_loans", "total_deposits"],
        markers=True,
        labels={"value": "NPR Millions", "variable": "Balance Sheet Metric", "fy": "Fiscal Year"},
        title="Aggregate Banking Sector Growth"
    )
    st.plotly_chart(fig_trend, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# 2. Bank Peer Comparison
# ─────────────────────────────────────────────────────────────────────────────
elif nav == "📊 Bank Peer Comparison":
    st.markdown('<div class="main-header">Bank-by-Bank Peer Benchmarking</div>', unsafe_allow_html=True)
    
    bank_list = sorted(panel_df["bank_name"].unique())
    selected_bank = st.selectbox("Select Target Bank for Deep Dive", bank_list, index=0)
    
    bank_data = panel_df[panel_df["bank_name"] == selected_bank].sort_values("fy")
    b_code = bank_data["bank_code"].iloc[0]
    
    st.markdown(f"### Historical Trajectory — **{selected_bank} ({b_code})**")
    
    c1, c2, c3, c4 = st.columns(4)
    latest = bank_data[bank_data["fy"] == selected_fy]
    if not latest.empty:
        c1.metric("Assets", f"NPR {latest['total_assets'].values[0]:,.0f}M")
        c2.metric("Deposits", f"NPR {latest['total_deposits'].values[0]:,.0f}M")
        c3.metric("Net Profit", f"NPR {latest['profit_after_tax'].values[0]:,.0f}M")
        c4.metric("ROE", f"{latest['roe'].values[0]:.2f}%" if pd.notna(latest['roe'].values[0]) else "N/A")

    col1, col2 = st.columns(2)
    with col1:
        fig_b1 = px.line(
            bank_data, x="fy", y=["total_assets", "total_deposits", "gross_loans"],
            markers=True, title=f"{b_code}: Balance Sheet Growth (NPR Millions)"
        )
        st.plotly_chart(fig_b1, use_container_width=True)

    with col2:
        fig_b2 = px.line(
            bank_data, x="fy", y=["roa", "roe", "nim"],
            markers=True, title=f"{b_code}: Profitability Ratios (%)"
        )
        st.plotly_chart(fig_b2, use_container_width=True)

    st.subheader("Cross-Bank Efficiency & Productivity Radar (FY" + str(selected_fy) + ")")
    comp_metrics = ["cost_income", "gross_npl_pct", "car_pct", "casa_ratio"]
    fig_comp = px.box(fy_filtered, y=["cost_income", "gross_npl_pct", "car_pct", "casa_ratio"], points="all", title="Industry Spread of Key Ratios")
    st.plotly_chart(fig_comp, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# 3. Market Concentration (HHI)
# ─────────────────────────────────────────────────────────────────────────────
elif nav == "📈 Market Concentration (HHI)":
    st.markdown('<div class="main-header">Industry Structure & Concentration Analysis</div>', unsafe_allow_html=True)
    st.markdown("Measurement of market concentration using Herfindahl-Hirschman Index (HHI) and Concentration Ratios (CR4, CR5, CR10).")

    col1, col2 = st.columns([3, 2])
    with col1:
        fig_hhi = px.line(
            ind_df,
            x="fy",
            y=["hhi_assets", "hhi_loans", "hhi_deposits"],
            markers=True,
            title="Herfindahl-Hirschman Index (HHI) Trend (FY2020 – FY2025)",
            labels={"value": "HHI Points (0 - 10,000)", "variable": "Market Segment"}
        )
        fig_hhi.add_hline(y=1000, line_dash="dash", line_color="green", annotation_text="Unconcentrated (<1,000)")
        fig_hhi.add_hline(y=1800, line_dash="dash", line_color="orange", annotation_text="Moderately Concentrated (1,000 - 1,800)")
        st.plotly_chart(fig_hhi, use_container_width=True)

    with col2:
        fig_cr = px.bar(
            ind_df,
            x="fy",
            y=["cr4_assets", "cr10_assets"],
            barmode="group",
            title="Top 4 vs Top 10 Market Share (CR4 & CR10)",
            labels={"value": "Market Share (%)", "variable": "Ratio"}
        )
        st.plotly_chart(fig_cr, use_container_width=True)

    st.subheader("Major M&A and Consolidation Events (2019 – 2024)")
    st.dataframe(events_df[["fy", "event_date", "bank_name", "event_type", "event_description", "counterparty", "strategic_impact"]], use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# 4. Business Model (Loans & Deposits)
# ─────────────────────────────────────────────────────────────────────────────
elif nav == "💼 Business Model (Loans & CASA)":
    st.markdown('<div class="main-header">Lending Portfolio & Deposit Structure</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader(f"Sectoral Lending Composition (System Total FY{selected_fy})")
        sector_cols = ["agriculture", "manufacturing", "construction", "wholesale_retail", "transportation", "tourism", "consumption", "real_estate", "hydropower"]
        sector_sums = fy_filtered[sector_cols].sum()
        fig_pie_loans = px.pie(
            values=sector_sums.values,
            names=[s.replace("_", " ").title() for s in sector_sums.index],
            hole=0.4,
            title="Lending by Economic Sector"
        )
        st.plotly_chart(fig_pie_loans, use_container_width=True)

    with col2:
        st.subheader(f"Deposit Mix: Fixed vs. Low-Cost CASA (FY{selected_fy})")
        dep_cols = ["current_deposits", "savings_deposits", "fixed_deposits", "call_deposits"]
        dep_sums = fy_filtered[dep_cols].sum()
        fig_pie_dep = px.pie(
            values=dep_sums.values,
            names=["Current (0%)", "Savings (Low-cost)", "Fixed (High-cost)", "Call Deposits"],
            hole=0.4,
            title="Deposit Structure (NPR Millions)"
        )
        st.plotly_chart(fig_pie_dep, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# 5. Macroeconomic Environment
# ─────────────────────────────────────────────────────────────────────────────
elif nav == "🌐 Macroeconomic Environment":
    st.markdown('<div class="main-header">Macroeconomic & Monetary Context</div>', unsafe_allow_html=True)
    st.markdown("Grounding banking sector performance against broader macroeconomic variables sourced directly from the World Bank and NRB.")

    macro_df = panel_df[["fy", "gdp_growth_pct", "inflation_pct", "remittance_usd_mn", "policy_rate_pct"]].drop_duplicates().sort_values("fy")

    col1, col2 = st.columns(2)
    with col1:
        fig_macro1 = px.bar(
            macro_df,
            x="fy",
            y="gdp_growth_pct",
            color="gdp_growth_pct",
            color_continuous_scale="Viridis",
            title="Nepal Real GDP Growth (%)",
            labels={"gdp_growth_pct": "GDP Growth (%)", "fy": "Fiscal Year"}
        )
        st.plotly_chart(fig_macro1, use_container_width=True)

    with col2:
        fig_macro2 = px.line(
            macro_df,
            x="fy",
            y=["inflation_pct", "policy_rate_pct"],
            markers=True,
            title="Inflation CPI vs NRB Policy Rate (%)"
        )
        st.plotly_chart(fig_macro2, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# 6. Digital & Strategic Scorecard
# ─────────────────────────────────────────────────────────────────────────────
elif nav == "📱 Digital & Strategic Scorecard":
    st.markdown('<div class="main-header">Strategic & Digital Maturity Scorecards</div>', unsafe_allow_html=True)
    st.markdown(f"Evaluated digital capabilities (0–10 scale) and strategic orientation for **FY{selected_fy}**.")

    fig_dig = px.bar(
        fy_filtered.sort_values("digital_index", ascending=False),
        x="digital_index",
        y="bank_name",
        orientation="h",
        color="digital_index",
        color_continuous_scale="Teal",
        title=f"Digital Banking Maturity Index (FY{selected_fy})"
    )
    fig_dig.update_layout(yaxis={"categoryorder": "total ascending"}, height=550)
    st.plotly_chart(fig_dig, use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# 7. Master Data Explorer
# ─────────────────────────────────────────────────────────────────────────────
elif nav == "📋 Master Data Explorer":
    st.markdown('<div class="main-header">Master Bank-Year Panel Explorer</div>', unsafe_allow_html=True)
    st.markdown(f"Accessing **{len(panel_df)} bank-year observations** across **{len(panel_df.columns)} indicators**.")

    selected_cols = st.multiselect(
        "Select Columns to Display",
        options=list(panel_df.columns),
        default=["bank_code", "bank_name", "fy", "total_assets", "gross_loans", "total_deposits", "profit_after_tax", "roa", "roe", "gross_npl_pct", "car_pct", "digital_index"]
    )

    st.dataframe(panel_df[selected_cols], use_container_width=True)

    csv = panel_df[selected_cols].to_csv(index=False).encode('utf-8')
    st.download_button(
        "📥 Download Filtered Data as CSV",
        data=csv,
        file_name="nepal_bank_panel_export.csv",
        mime="text/csv",
    )
