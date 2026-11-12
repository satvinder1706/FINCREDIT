
import io
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="FINCREDIT | Financial Analysis",
    page_icon="📊",
    layout="wide",
)

st.title("FINCREDIT")
st.caption("Financial Analysis & Working Capital Assessment")
st.info(
    "Prototype for preliminary analysis only. Confirm all figures, classifications, "
    "RBI directions and your bank's current credit policy before using the output in a lending decision."
)

def safe_divide(a, b):
    return a / b if b not in (0, 0.0, None) else np.nan

def money(value, unit):
    if pd.isna(value):
        return "Not available"
    return f"{value:,.2f} {unit}"

def make_excel_report(company, unit, bs_assets, bs_liabilities, history, ratios,
                      oc_summary, wc_summary, observations):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(
            [{"Particulars": k, "Amount": v} for k, v in bs_assets.items()]
            + [{"Particulars": k, "Amount": v} for k, v in bs_liabilities.items()]
        ).to_excel(writer, sheet_name="Balance Sheet", index=False)
        history.to_excel(writer, sheet_name="Financial Trends", index=False)
        ratios.to_excel(writer, sheet_name="Financial Ratios", index=False)
        pd.DataFrame([oc_summary]).to_excel(writer, sheet_name="Operating Cycle", index=False)
        pd.DataFrame([wc_summary]).to_excel(writer, sheet_name="Working Capital", index=False)
        pd.DataFrame({"Preliminary Observations": observations}).to_excel(
            writer, sheet_name="Observations", index=False
        )
        pd.DataFrame([
            ["Company", company],
            ["Reporting unit", unit],
            ["Generated on", date.today().isoformat()],
            ["Disclaimer", "Indicative analysis; subject to verification and applicable bank policy."],
        ], columns=["Field", "Value"]).to_excel(writer, sheet_name="Cover", index=False)
    output.seek(0)
    return output.getvalue()

with st.sidebar:
    st.header("Borrower profile")
    company_name = st.text_input("Company / borrower name", "ABC Industries Pvt Ltd")
    business_activity = st.text_input("Business activity", "Manufacturing")
    constitution = st.selectbox(
        "Constitution",
        ["Proprietorship", "Partnership", "LLP", "Private Limited Company",
         "Public Limited Company", "Other"],
    )
    financial_year = st.text_input("Latest financial year", "2025-26")
    unit = st.selectbox("Amount unit", ["₹ lakh", "₹ thousand", "₹ crore"])
    st.divider()
    st.caption("Enter all monetary amounts in the same selected unit.")

tab_profile, tab_history, tab_cycle, tab_wc, tab_analysis = st.tabs([
    "Company & Balance Sheet",
    "Financial History",
    "Operating Cycle",
    "Working Capital & CC",
    "Analysis & Reports",
])

with tab_profile:
    st.subheader("Latest balance sheet")
    st.write("Enter the latest available balance-sheet figures.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Assets")
        cash_bank = st.number_input("Cash and bank balances", min_value=0.0, value=15.0, step=1.0)
        debtors_bs = st.number_input("Sundry debtors", min_value=0.0, value=45.0, step=1.0)
        inventory_bs = st.number_input("Inventory / stock", min_value=0.0, value=60.0, step=1.0)
        current_advances = st.number_input("Current loans and advances", min_value=0.0, value=10.0, step=1.0)
        fixed_assets = st.number_input("Fixed assets (net)", min_value=0.0, value=80.0, step=1.0)
        other_assets = st.number_input("Other assets", min_value=0.0, value=5.0, step=1.0)
    with c2:
        st.markdown("#### Liabilities and equity")
        share_capital = st.number_input("Share capital / proprietor's capital", min_value=0.0, value=50.0, step=1.0)
        reserves = st.number_input("Reserves and surplus", min_value=0.0, value=40.0, step=1.0)
        term_loans = st.number_input("Term loans", min_value=0.0, value=50.0, step=1.0)
        cc_borrowings = st.number_input("Cash-credit / working-capital borrowings", min_value=0.0, value=25.0, step=1.0)
        sundry_creditors = st.number_input("Sundry creditors", min_value=0.0, value=35.0, step=1.0)
        other_current_liabilities = st.number_input("Other current liabilities", min_value=0.0, value=10.0, step=1.0)
        other_liabilities = st.number_input("Other liabilities", min_value=0.0, value=5.0, step=1.0)

    st.markdown("#### TNW adjustments (enter amounts as per applicable policy)")
    t1, t2, t3, t4 = st.columns(4)
    with t1:
        intangible_assets = st.number_input("Intangible assets", min_value=0.0, value=0.0, step=1.0)
    with t2:
        accumulated_losses = st.number_input("Accumulated losses", min_value=0.0, value=0.0, step=1.0)
    with t3:
        misc_expenditure = st.number_input("Unwritten-off misc. expenditure", min_value=0.0, value=0.0, step=1.0)
    with t4:
        quasi_equity = st.number_input("Eligible quasi-equity*", min_value=0.0, value=0.0, step=1.0)
    st.caption("*Include quasi-equity only where recognised under the applicable bank policy.")

with tab_history:
    st.subheader("Historical and projected financial information")
    st.write("Edit the sample figures directly. Keep three historical years followed by projected years, where available.")
    default_history = pd.DataFrame({
        "Financial Year": ["2023-24", "2024-25", "2025-26", "2026-27 (P)", "2027-28 (P)"],
        "Sales": [300.0, 375.0, 450.0, 540.0, 620.0],
        "Capital": [60.0, 75.0, 92.0, 110.0, 130.0],
        "Net Profit": [15.0, 22.0, 28.0, 35.0, 43.0],
        "Borrowings": [90.0, 120.0, 160.0, 190.0, 220.0],
        "Net Working Capital": [40.0, 50.0, 60.0, 75.0, 90.0],
        "Gross Profit": [75.0, 94.0, 112.0, 135.0, 155.0],
        "EBIT": [35.0, 44.0, 55.0, 67.0, 78.0],
        "Interest Expense": [10.0, 12.0, 15.0, 17.0, 19.0],
        "Depreciation": [8.0, 9.0, 10.0, 12.0, 13.0],
        "COGS": [225.0, 281.0, 338.0, 405.0, 465.0],
    })
    history = st.data_editor(
        default_history,
        num_rows="dynamic",
        use_container_width=True,
        key="financial_history_editor",
        column_config={
            "Financial Year": st.column_config.TextColumn("Financial Year", required=True),
            **{col: st.column_config.NumberColumn(col, min_value=0.0, format="%.2f")
               for col in default_history.columns if col != "Financial Year"}
        },
    )

with tab_cycle:
    st.subheader("Operating-cycle information")
    st.caption("Use opening and closing balances for the same period as the activity figures.")
    days_in_year = st.selectbox("Days in year", [365, 360], index=0)
    a, b, c = st.columns(3)
    with a:
        opening_inventory = st.number_input("Opening inventory", min_value=0.0, value=50.0)
        closing_inventory = st.number_input("Closing inventory", min_value=0.0, value=60.0)
        cogs_cycle = st.number_input("Cost of goods sold for period", min_value=0.0, value=300.0)
    with b:
        opening_debtors = st.number_input("Opening debtors", min_value=0.0, value=35.0)
        closing_debtors = st.number_input("Closing debtors", min_value=0.0, value=45.0)
        credit_sales = st.number_input("Credit sales for period", min_value=0.0, value=400.0)
    with c:
        opening_creditors = st.number_input("Opening creditors", min_value=0.0, value=25.0)
        closing_creditors = st.number_input("Closing creditors", min_value=0.0, value=35.0)
        credit_purchases = st.number_input("Credit purchases for period", min_value=0.0, value=280.0)
    annual_operating_expenditure = st.number_input(
        "Annual operating expenditure (funding estimate)", min_value=0.0, value=350.0
    )

with tab_wc:
    st.subheader("Projected working-capital assessment")
    st.caption("Enter projected current assets and other current liabilities for the assessment period.")
    d1, d2 = st.columns(2)
    with d1:
        st.markdown("#### Projected current assets")
        raw_material = st.number_input("Raw material", min_value=0.0, value=40.0)
        work_in_progress = st.number_input("Work in progress", min_value=0.0, value=20.0)
        finished_goods = st.number_input("Finished goods", min_value=0.0, value=50.0)
        cma_debtors = st.number_input("Projected sundry debtors", min_value=0.0, value=60.0)
        cma_cash = st.number_input("Projected cash and bank", min_value=0.0, value=10.0)
        cma_other_assets = st.number_input("Other current assets", min_value=0.0, value=20.0)
    with d2:
        st.markdown("#### Other current liabilities")
        cma_creditors = st.number_input("Projected sundry creditors", min_value=0.0, value=40.0)
        outstanding_expenses = st.number_input("Outstanding expenses", min_value=0.0, value=10.0)
        cma_other_cl = st.number_input("Other current liabilities (excluding bank borrowing)", min_value=0.0, value=10.0)
        st.markdown("#### Cash-credit request")
        projected_turnover = st.number_input("Projected annual turnover", min_value=0.0, value=500.0)
        existing_cc = st.number_input("Existing CC limit", min_value=0.0, value=80.0)
        proposed_cc = st.number_input("Proposed CC limit", min_value=0.0, value=110.0)
        assessment_method = st.selectbox("Assessment method", [
            "Turnover Method", "MPBF Method I", "MPBF Method II"
        ])

with tab_profile:
    st.caption("All figures are in the selected unit. Review the balance-sheet tally in the Analysis tab.")

# ---------------- CALCULATIONS ----------------
assets = {
    "Cash and Bank Balance": cash_bank,
    "Sundry Debtors": debtors_bs,
    "Inventory": inventory_bs,
    "Current Loans and Advances": current_advances,
    "Fixed Assets": fixed_assets,
    "Other Assets": other_assets,
}
liabilities = {
    "Share Capital": share_capital,
    "Reserves and Surplus": reserves,
    "Term Loans": term_loans,
    "Cash Credit Borrowings": cc_borrowings,
    "Sundry Creditors": sundry_creditors,
    "Other Current Liabilities": other_current_liabilities,
    "Other Liabilities": other_liabilities,
}
total_assets = sum(assets.values())
total_liabilities = sum(liabilities.values())
current_assets = cash_bank + debtors_bs + inventory_bs + current_advances
current_liabilities = cc_borrowings + sundry_creditors + other_current_liabilities
net_working_capital = current_assets - current_liabilities
current_ratio = safe_divide(current_assets, current_liabilities)
quick_ratio = safe_divide(current_assets - inventory_bs, current_liabilities)

adjusted_tnw = (
    share_capital + reserves - intangible_assets - accumulated_losses
    - misc_expenditure + quasi_equity
)
total_borrowings = term_loans + cc_borrowings
total_outside_liabilities = (
    term_loans + cc_borrowings + sundry_creditors
    + other_current_liabilities + other_liabilities
)
debt_equity = safe_divide(total_borrowings, adjusted_tnw) if adjusted_tnw > 0 else np.nan
tol_tnw = safe_divide(total_outside_liabilities, adjusted_tnw) if adjusted_tnw > 0 else np.nan

# Validate and prepare history.
required_numeric = [c for c in default_history.columns if c != "Financial Year"]
history_valid = history.copy()
for col in required_numeric:
    history_valid[col] = pd.to_numeric(history_valid[col], errors="coerce")
history_valid = history_valid.dropna(subset=["Financial Year"])
numeric_bad = history_valid[required_numeric].isna().any().any()
negative_bad = (history_valid[required_numeric] < 0).any().any()
history_valid["Financial Year"] = history_valid["Financial Year"].astype(str)

ratios = pd.DataFrame(columns=["Ratio", "Result", "Unit", "Formula"])
observations = []
oc_summary = {}
wc_summary = {}
selected_assessment = np.nan

if st.button("Analyse Financial Statements", type="primary", use_container_width=True):
    st.session_state["run_analysis"] = True

if st.session_state.get("run_analysis", False):
    with tab_analysis:
        st.subheader("Balance sheet check")
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Total assets", money(total_assets, unit))
        col_b.metric("Total liabilities & equity", money(total_liabilities, unit))
        difference = total_assets - total_liabilities
        col_c.metric("Difference", money(difference, unit))
        if abs(difference) <= 0.01:
            st.success("Balance sheet tallies within the 0.01 tolerance.")
        else:
            st.error("Balance sheet does not tally. Resolve the difference before relying on the analysis.")

    if numeric_bad or negative_bad:
        st.error("Financial history contains blank, non-numeric or negative figures. Correct the table and analyse again.")
        st.stop()
    if history_valid.empty or len(history_valid) < 2:
        st.error("Enter at least two financial years to calculate trends.")
        st.stop()
    if history_valid["Financial Year"].duplicated().any():
        st.error("Financial year labels must be unique.")
        st.stop()

    latest = history_valid.iloc[-1]
    latest_sales = float(latest["Sales"])
    latest_gp = float(latest["Gross Profit"])
    latest_ebit = float(latest["EBIT"])
    latest_interest = float(latest["Interest Expense"])
    latest_np = float(latest["Net Profit"])
    latest_cogs = float(latest["COGS"])
    ebitda = latest_ebit + float(latest["Depreciation"])

    avg_inventory = (opening_inventory + closing_inventory) / 2
    avg_debtors = (opening_debtors + closing_debtors) / 2
    avg_creditors = (opening_creditors + closing_creditors) / 2

    inventory_days = safe_divide(avg_inventory * days_in_year, cogs_cycle)
    debtor_days = safe_divide(avg_debtors * days_in_year, credit_sales)
    creditor_days = safe_divide(avg_creditors * days_in_year, credit_purchases)
    gross_cycle = inventory_days + debtor_days if pd.notna(inventory_days) and pd.notna(debtor_days) else np.nan
    net_cycle = gross_cycle - creditor_days if pd.notna(gross_cycle) and pd.notna(creditor_days) else np.nan
    cycle_funding = (
        annual_operating_expenditure / days_in_year * net_cycle
        if pd.notna(net_cycle) else np.nan
    )

    total_cma_ca = sum([
        raw_material, work_in_progress, finished_goods, cma_debtors, cma_cash, cma_other_assets
    ])
    total_other_cl = cma_creditors + outstanding_expenses + cma_other_cl
    wc_gap = total_cma_ca - total_other_cl
    turnover_wcf = 0.20 * projected_turnover
    method1_raw = 0.75 * wc_gap
    method1 = max(method1_raw, 0)
    method2_raw = 0.75 * total_cma_ca - total_other_cl
    method2 = max(method2_raw, 0)
    assessment_map = {
        "Turnover Method": turnover_wcf,
        "MPBF Method I": method1,
        "MPBF Method II": method2,
    }
    selected_assessment = assessment_map[assessment_method]

    inventory_turnover = safe_divide(latest_cogs, avg_inventory)
    debtor_turnover = safe_divide(credit_sales, avg_debtors)
    creditor_turnover = safe_divide(credit_purchases, avg_creditors)
    interest_coverage = safe_divide(latest_ebit, latest_interest)
    gross_margin = safe_divide(latest_gp * 100, latest_sales)
    ebitda_margin = safe_divide(ebitda * 100, latest_sales)
    operating_margin = safe_divide(latest_ebit * 100, latest_sales)
    net_margin = safe_divide(latest_np * 100, latest_sales)

    ratios = pd.DataFrame([
        ["Current Ratio", current_ratio, "times", "Current Assets / Current Liabilities"],
        ["Quick Ratio", quick_ratio, "times", "(Current Assets − Inventory) / Current Liabilities"],
        ["Debt–Equity", debt_equity, "times", "Total Borrowings / Adjusted TNW"],
        ["TOL / TNW", tol_tnw, "times", "Total Outside Liabilities / Adjusted TNW"],
        ["Interest Coverage", interest_coverage, "times", "EBIT / Interest Expense"],
        ["Gross Profit Margin", gross_margin, "%", "Gross Profit / Sales × 100"],
        ["EBITDA Margin", ebitda_margin, "%", "(EBIT + Depreciation) / Sales × 100"],
        ["Operating Profit Margin", operating_margin, "%", "EBIT / Sales × 100"],
        ["Net Profit Margin", net_margin, "%", "Net Profit / Sales × 100"],
        ["Inventory Turnover", inventory_turnover, "times", "COGS / Average Inventory"],
        ["Debtor Turnover", debtor_turnover, "times", "Credit Sales / Average Debtors"],
        ["Creditor Turnover", creditor_turnover, "times", "Credit Purchases / Average Creditors"],
    ], columns=["Ratio", "Result", "Unit", "Formula"])

    # Trend table with year-on-year absolute and percentage changes.
    trend = history_valid.copy()
    for metric in ["Sales", "Capital", "Net Profit", "Borrowings", "Net Working Capital"]:
        trend[f"{metric} Change"] = trend[metric].diff()
        trend[f"{metric} Growth (%)"] = trend[metric].pct_change() * 100
    for metric in ["Gross Profit", "EBIT", "Net Profit"]:
        trend[f"{metric} Margin (%)"] = np.where(
            trend["Sales"] != 0, trend[metric] / trend["Sales"] * 100, np.nan
        )

    observations = []
    if abs(difference) > 0.01:
        observations.append("Balance sheet does not tally; resolve the difference.")
    if adjusted_tnw <= 0:
        observations.append("Adjusted TNW is zero or negative; solvency requires detailed review.")
    if pd.notna(current_ratio) and current_ratio < 1:
        observations.append(f"Current ratio is {current_ratio:.2f}:1; current liabilities exceed current assets.")
    if pd.notna(interest_coverage) and interest_coverage < 1.5:
        observations.append("Interest coverage is below 1.5 times; review debt-servicing headroom.")
    if len(trend) >= 2:
        latest_row, previous_row = trend.iloc[-1], trend.iloc[-2]
        if pd.notna(latest_row["Borrowings Growth (%)"]) and pd.notna(latest_row["Sales Growth (%)"]):
            if latest_row["Borrowings Growth (%)"] > latest_row["Sales Growth (%)"]:
                observations.append("Borrowings grew faster than sales in the latest entered period; examine end-use and working-capital cycle.")
        if pd.notna(latest_row["Net Profit Margin (%)"]) and pd.notna(previous_row["Net Profit Margin (%)"]):
            if latest_row["Net Profit Margin (%)"] < previous_row["Net Profit Margin (%)"]:
                observations.append("Net profit margin declined in the latest entered period.")
    if pd.notna(net_cycle) and net_cycle < 0:
        observations.append("Net operating cycle is negative; verify balances and credit terms.")
    if proposed_cc > selected_assessment:
        observations.append(
            f"Proposed CC limit exceeds the selected indicative assessment by {money(proposed_cc-selected_assessment, unit)}."
        )

    oc_summary = {
        "Inventory days": inventory_days,
        "Debtor days": debtor_days,
        "Creditor days": creditor_days,
        "Gross operating cycle": gross_cycle,
        "Net operating cycle": net_cycle,
        "Indicative cycle funding estimate": cycle_funding,
    }
    wc_summary = {
        "Projected current assets": total_cma_ca,
        "Other current liabilities": total_other_cl,
        "Working capital gap": wc_gap,
        "Turnover Method (illustrative)": turnover_wcf,
        "MPBF Method I (illustrative)": method1,
        "MPBF Method II (illustrative)": method2,
        "Selected method": assessment_method,
        "Selected indicative assessment": selected_assessment,
        "Existing CC limit": existing_cc,
        "Proposed CC limit": proposed_cc,
    }

    with tab_analysis:
        st.header("Analysis results")
        if company_name.strip() == "":
            st.warning("Enter a borrower name in the sidebar.")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Current ratio", f"{current_ratio:.2f}:1" if pd.notna(current_ratio) else "N/A")
        k2.metric("Quick ratio", f"{quick_ratio:.2f}:1" if pd.notna(quick_ratio) else "N/A")
        k3.metric("Adjusted TNW", money(adjusted_tnw, unit))
        k4.metric("Selected WC assessment", money(selected_assessment, unit))

        st.subheader("Financial ratios")
        ratio_display = ratios.copy()
        ratio_display["Result"] = ratio_display["Result"].map(lambda x: "N/A" if pd.isna(x) else f"{x:,.2f}")
        st.dataframe(ratio_display, use_container_width=True, hide_index=True)

        st.subheader("Financial trend analysis")
        st.dataframe(trend, use_container_width=True, hide_index=True)
        st.line_chart(trend.set_index("Financial Year")[["Sales", "Capital", "Net Profit", "Borrowings"]])

        st.subheader("Operating cycle")
        st.dataframe(pd.DataFrame([oc_summary]), use_container_width=True, hide_index=True)
        if pd.notna(net_cycle) and net_cycle < 0:
            st.warning("Negative net operating cycle: verify the input assumptions.")
        st.caption("Operating-cycle funding estimate is a simplified indicator, not a sanctioned limit or drawing power.")

        st.subheader("Working-capital assessment")
        st.dataframe(pd.DataFrame([wc_summary]), use_container_width=True, hide_index=True)
        if wc_gap < 0:
            st.warning("Other current liabilities exceed projected current assets; working-capital gap is negative.")
        if method1_raw < 0 or method2_raw < 0:
            st.warning("A raw MPBF calculation is negative; displayed finance is floored at zero.")
        st.caption(
            "Turnover Method and MPBF formulas shown are illustrative conventional calculations. "
            "Use only the method and parameters applicable under your bank's current policy and borrower category."
        )

        st.subheader("Preliminary analyst observations")
        if observations:
            for item in observations:
                st.warning(item)
        else:
            st.success("No rule-based flags were triggered. This does not mean the proposal is risk-free.")
        st.caption("These observations are prompts for analyst review, not a credit rating or sanction recommendation.")

        report = make_excel_report(
            company_name, unit, assets, liabilities, trend, ratios,
            oc_summary, wc_summary, observations
        )
        st.download_button(
            "Download analysis workbook (Excel)",
            data=report,
            file_name=f"{company_name.strip().replace(' ', '_') or 'borrower'}_FINCREDIT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
else:
    st.warning("Enter or review the sample data in the tabs, then click **Analyse Financial Statements**.")
    st.caption("The sample borrower and amounts are illustrative. Replace them with verified borrower data.")
