
import io
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from xml.sax.saxutils import escape

st.set_page_config(
    page_title="FINCREDIT | Financial Analysis",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
html, body, [class*='css'], .stApp { font-family: 'Times New Roman', Times, serif; }
:root { --fc-ink:#172b4d; --fc-line:#e2e8f0; }
.block-container {padding-top:1.5rem; padding-bottom:3rem; max-width:1450px;}
[data-testid="stAppViewContainer"] {background:linear-gradient(180deg,#f5f8fd 0%,#ffffff 38%);}
[data-testid="stSidebar"] {background:#f1f5fb; border-right:1px solid #e2e8f0;}
h1,h2,h3 {color:var(--fc-ink); letter-spacing:-.025em;}
[data-testid="stMetric"] {background:#fff; border:1px solid var(--fc-line); padding:14px 16px; border-radius:14px; box-shadow:0 3px 12px rgba(25,45,80,.045);}
[data-testid="stMetricLabel"] {color:#64748b; font-weight:600; font-size:.88rem;}
[data-testid="stMetricValue"] {color:#172b4d; font-weight:700; font-size:1.25rem !important; line-height:1.25;}
.stTabs [data-baseweb="tab-list"] {gap:7px; border-bottom:1px solid #e5eaf2;}
.stTabs [data-baseweb="tab"] {background:#f0f4fa; border-radius:9px 9px 0 0; padding:10px 14px; font-weight:650;}
.stTabs [aria-selected="true"] {background:#e3edff !important; color:#174a91 !important;}
.stButton>button,.stDownloadButton>button {border-radius:10px; font-weight:700; padding:.6rem 1rem;}
[data-testid="stDataFrame"] {border:1px solid var(--fc-line); border-radius:12px; overflow:hidden;}
.fc-hero {background:linear-gradient(115deg,#102b50,#2457a6); color:white; padding:22px 25px; border-radius:18px; margin:8px 0 20px; box-shadow:0 8px 24px rgba(20,55,100,.14);}
.fc-hero p {color:#dbeafe; margin:4px 0 0;}
.fc-kicker {font-size:11px; text-transform:uppercase; letter-spacing:.14em; color:#bcd4f6; font-weight:800;}
.fc-note {background:#f1f6ff; border:1px solid #d9e7fb; border-left:4px solid #3974c6; border-radius:10px; padding:12px 15px; color:#334155; margin:10px 0 16px;}
@media(max-width:700px){.block-container{padding-left:1rem;padding-right:1rem}.fc-hero{padding:18px}}
</style>
""", unsafe_allow_html=True)

st.markdown("""<div class="fc-hero"><div class="fc-kicker">Credit appraisal workspace</div><h1 style="color:white;margin:5px 0 0 0">FINCREDIT</h1><p>Financial intelligence • Working-capital assessment • Drawing power</p></div>""", unsafe_allow_html=True)
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
                      oc_summary, wc_summary, dp_summary, observations):
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
        pd.DataFrame([dp_summary]).to_excel(writer, sheet_name="Drawing Power", index=False)
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


def make_pdf_report(company, activity, constitution, financial_year, unit,
                    bs_assets, bs_liabilities, history, ratios,
                    oc_summary, wc_summary, dp_summary, observations):
    """Create a PDF report using the Times font family."""
    output = io.BytesIO()

    # Built-in PDF fonts are always available in ReportLab.
    pdf_regular = "Times-Roman"
    pdf_bold = "Times-Bold"

    doc = SimpleDocTemplate(
        output, pagesize=A4, rightMargin=17*mm, leftMargin=17*mm,
        topMargin=17*mm, bottomMargin=17*mm,
        title=f"FINCREDIT Financial Analysis - {company}",
        author="FINCREDIT",
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="FC_Title", parent=styles["Title"], fontName=pdf_bold,
        fontSize=21, leading=26, textColor=colors.HexColor("#17365D"),
        alignment=TA_LEFT, spaceAfter=4*mm
    ))
    styles.add(ParagraphStyle(
        name="FC_Subtitle", parent=styles["Normal"], fontName=pdf_regular,
        fontSize=9, leading=13, textColor=colors.HexColor("#64748B"),
        spaceAfter=5*mm
    ))
    styles.add(ParagraphStyle(
        name="FC_H2", parent=styles["Heading2"], fontName=pdf_bold,
        fontSize=12, leading=16, textColor=colors.HexColor("#1D4E89"),
        spaceBefore=5*mm, spaceAfter=2*mm
    ))
    styles.add(ParagraphStyle(
        name="FC_Body", parent=styles["BodyText"], fontName=pdf_regular,
        fontSize=8, leading=11, textColor=colors.HexColor("#263445")
    ))
    styles.add(ParagraphStyle(
        name="FC_Small", parent=styles["BodyText"], fontName=pdf_regular,
        fontSize=7, leading=9, textColor=colors.HexColor("#64748B")
    ))

    def para(value, style="FC_Body"):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            value = "Not available"
        return Paragraph(escape(str(value)).replace("\n", "<br/>"), styles[style])

    def table_from_mapping(mapping, first="Particulars", second="Amount"):
        rows = [[para(first, "FC_Body"), para(second, "FC_Body")]]
        for key, value in mapping.items():
            if isinstance(value, (float, np.floating)):
                value = "N/A" if pd.isna(value) else f"{value:,.2f} {unit}"
            rows.append([para(key), para(value)])
        return Table(rows, colWidths=[92*mm, 78*mm], repeatRows=1, hAlign="LEFT")

    def styled_table(data, widths=None):
        tbl = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), pdf_bold),
            ("FONTNAME", (0, 1), (-1, -1), pdf_regular),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("LEADING", (0, 0), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D8E1EC")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#F3F6FA")]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        return tbl

    story = [
        Paragraph("FINCREDIT", styles["FC_Title"]),
        Paragraph("Financial Analysis &amp; Working Capital Assessment", styles["FC_Subtitle"]),
        para(f"Borrower: {company}  |  Activity: {activity}  |  Constitution: {constitution}"),
        para(f"Financial year: {financial_year}  |  Reporting unit: {unit}  |  Generated: {date.today().isoformat()}"),
        Spacer(1, 3*mm),
        para("Prototype report for preliminary analysis only. Verify source documents, classifications, "
             "eligibility, sanction terms and current bank policy before relying on these figures.", "FC_Small"),
        Paragraph("Balance Sheet", styles["FC_H2"]),
    ]
    bs_rows = [[para("Particulars"), para("Amount")]]
    for label, amount in list(bs_assets.items()) + list(bs_liabilities.items()):
        bs_rows.append([para(label), para(f"{amount:,.2f} {unit}")])
    story.append(styled_table(bs_rows, [92*mm, 78*mm]))

    story.append(Paragraph("Financial History &amp; Trends", styles["FC_H2"]))
    hist = history.copy()
    hist = hist.fillna("")
    hist_data = [[para(c) for c in hist.columns]]
    for row in hist.astype(str).values.tolist():
        hist_data.append([para(v) for v in row])
    col_width = 170*mm / max(len(hist.columns), 1)
    story.append(styled_table(hist_data, [col_width]*len(hist.columns)))

    story.append(Paragraph("Financial Ratios", styles["FC_H2"]))
    ratio_data = [[para(c) for c in ratios.columns]]
    for row in ratios.fillna("").astype(str).values.tolist():
        ratio_data.append([para(v) for v in row])
    story.append(styled_table(ratio_data, [38*mm, 27*mm, 20*mm, 85*mm]))

    story.append(Paragraph("Operating Cycle", styles["FC_H2"]))
    story.append(table_from_mapping(oc_summary))
    story.append(Paragraph("Working-Capital Assessment", styles["FC_H2"]))
    story.append(table_from_mapping(wc_summary))
    story.append(Paragraph("Drawing Power", styles["FC_H2"]))
    story.append(table_from_mapping(dp_summary))

    story.append(Paragraph("Preliminary Analyst Observations", styles["FC_H2"]))
    if observations:
        for item in observations:
            story.append(Paragraph("• " + escape(str(item)), styles["FC_Body"]))
            story.append(Spacer(1, 1.5*mm))
    else:
        story.append(para("No rule-based flags were triggered. This does not mean the proposal is risk-free."))

    story.append(Spacer(1, 5*mm))
    story.append(para(
        "Important: FINCREDIT is an analytical aid, not a credit rating, legal opinion or sanction recommendation. "
        "The 25% MSME drawing-power margin is a configurable default in this application, not a universal "
        "RBI-prescribed margin. Apply the margin, debtor eligibility, ageing criteria and creditor deductions "
        "specified by the bank's current policy and the borrower's sanction terms.",
        "FC_Small"
    ))

    def add_page_number(canvas, document):
        canvas.saveState()
        canvas.setFont(pdf_regular, 7)
        canvas.setFillColor(colors.HexColor("#64748B"))
        canvas.drawString(17*mm, 10*mm, "FINCREDIT | Confidential borrower analysis")
        canvas.drawRightString(A4[0]-17*mm, 10*mm, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
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

tab_profile, tab_history, tab_cycle, tab_wc, tab_dp, tab_analysis = st.tabs([
    "🏢 Profile & Balance Sheet", "📈 Financial History", "🔄 Operating Cycle",
    "💼 Working Capital", "🧮 Drawing Power", "📊 Analysis & Reports",
])

# Guided workflow tracker: progress is based on sections the user explicitly marks reviewed.
workflow_steps = [
    ("profile_reviewed", "Profile & Balance Sheet"),
    ("history_reviewed", "Financial History"),
    ("cycle_reviewed", "Operating Cycle"),
    ("wc_reviewed", "Working Capital"),
    ("dp_reviewed", "Drawing Power"),
]
completed_steps = sum(bool(st.session_state.get(key, False)) for key, _ in workflow_steps)
progress_fraction = completed_steps / len(workflow_steps)
st.markdown("### Your appraisal progress")
progress_col, status_col = st.columns([4, 1])
with progress_col:
    st.progress(progress_fraction, text=f"{completed_steps} of {len(workflow_steps)} sections reviewed")
with status_col:
    st.metric("Completion", f"{progress_fraction:.0%}")
step_cols = st.columns(len(workflow_steps))
for col, (key, label) in zip(step_cols, workflow_steps):
    with col:
        done = bool(st.session_state.get(key, False))
        st.markdown(f"{'✅' if done else '⬜'} **{label}**")
st.caption("Mark a section as reviewed after checking its entries. Progress tracks review status, not the correctness or completeness of the financial information.")

with tab_profile:
    st.checkbox("I have reviewed the borrower profile and balance-sheet entries", key="profile_reviewed")
    st.subheader("Latest balance sheet")
    st.write("Enter the latest available balance-sheet figures.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Assets")
        cash_bank = st.number_input("Cash & Bank", min_value=0.0, value=15.0, step=1.0)
        debtors_bs = st.number_input("Sundry Debtors", min_value=0.0, value=45.0, step=1.0)
        stocks_bs = st.number_input("Stock", min_value=0.0, value=60.0, step=1.0)
        current_advances = st.number_input("Other Current Assets", min_value=0.0, value=10.0, step=1.0)
        fixed_assets = st.number_input("Fixed Assets (after depreciation)", min_value=0.0, value=80.0, step=1.0)
        other_assets = st.number_input("Other Assets (non-current)", min_value=0.0, value=5.0, step=1.0)
    with c2:
        st.markdown("#### Liabilities")
        share_capital = st.number_input("Capital", min_value=0.0, value=50.0, step=1.0)
        reserves = st.number_input("Reserves / Surplus (if shown)", min_value=0.0, value=40.0, step=1.0)
        term_loans = st.number_input("Term Loans", min_value=0.0, value=50.0, step=1.0)
        cc_borrowings = st.number_input("Bank CC / Working Capital Loan", min_value=0.0, value=25.0, step=1.0)
        sundry_creditors = st.number_input("Sundry Creditors", min_value=0.0, value=35.0, step=1.0)
        other_current_liabilities = st.number_input("Expenses Payable & Other Current Liabilities", min_value=0.0, value=10.0, step=1.0)
        other_liabilities = st.number_input("Other Liabilities (non-current)", min_value=0.0, value=5.0, step=1.0)
    st.caption("Enter only the items shown in the borrower's balance sheet. Leave items not shown at zero.")

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
    st.info("Net Working Capital is calculated automatically as Total Current Assets minus Total Current Liabilities. Current assets include Stock, Sundry Debtors, Cash & Bank and Other Current Assets. Current liabilities include Bank CC / Working Capital Loan, Sundry Creditors and Expenses Payable / Other Current Liabilities.")

with tab_history:
    st.checkbox("I have reviewed the actual and projected financial figures", key="history_reviewed")
    st.subheader("Historical and projected financial information")
    st.write("Select how many actual / previous years and projected years you want to enter. The table adjusts automatically.")
    choose1, choose2 = st.columns(2)
    with choose1:
        actual_year_count = st.selectbox("Number of actual / previous years", [1, 2, 3, 4, 5], index=2, key="actual_year_count")
    with choose2:
        projected_year_count = st.selectbox("Number of projected years", [0, 1, 2, 3, 4, 5], index=2, key="projected_year_count")
    actual_labels = ["2023-24", "2024-25", "2025-26", "2026-27", "2027-28"][:actual_year_count]
    projected_labels = ["2026-27 (P)", "2027-28 (P)", "2028-29 (P)", "2029-30 (P)", "2030-31 (P)"][:projected_year_count]
    labels = actual_labels + projected_labels
    sample = {
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
    }
    default_history = pd.DataFrame({"Financial Year": labels, **{k: v[:len(labels)] for k, v in sample.items()}})
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
    st.checkbox("I have reviewed the operating-cycle inputs", key="cycle_reviewed")
    st.subheader("Operating-cycle information")
    st.caption("Choose a financial year. Sales and COGS are fetched from financial history; enter total purchases. Debtor and creditor periods use total sales and purchases as simplified proxies when credit-only figures are unavailable.")
    days_in_year = st.selectbox("Days in year", [365, 360], index=0)
    cy = history.copy()
    cy["Financial Year"] = cy["Financial Year"].astype(str)
    years = [v for v in cy["Financial Year"].tolist() if "(P)" not in v.upper()] or cy["Financial Year"].tolist()
    cycle_year = st.selectbox("Financial year", years, index=len(years)-1)
    row = cy.loc[cy["Financial Year"] == cycle_year].iloc[-1]
    cogs_value = pd.to_numeric(row["COGS"], errors="coerce")
    sales_value = pd.to_numeric(row["Sales"], errors="coerce")
    cogs_cycle = 0.0 if pd.isna(cogs_value) else float(cogs_value)
    sales_cycle = 0.0 if pd.isna(sales_value) else float(sales_value)
    a,b,c = st.columns(3)
    with a:
        opening_inventory = st.number_input("Opening Stock", min_value=0.0, value=50.0)
        closing_inventory = st.number_input("Closing Stock", min_value=0.0, value=60.0)
        st.metric("COGS from financial history", f"{cogs_cycle:,.2f} {unit}")
    with b:
        opening_debtors = st.number_input("Opening Debtors", min_value=0.0, value=35.0)
        closing_debtors = st.number_input("Closing Debtors", min_value=0.0, value=45.0)
        st.metric("Sales from financial history", f"{sales_cycle:,.2f} {unit}")
    with c:
        opening_creditors = st.number_input("Opening Creditors", min_value=0.0, value=25.0)
        closing_creditors = st.number_input("Closing Creditors", min_value=0.0, value=35.0)
        total_purchases = st.number_input("Purchases", min_value=0.0, value=300.0)
    operating_expenses_ex_cogs = st.number_input("Operating expenses excluding COGS, depreciation and interest", min_value=0.0, value=50.0)
    annual_operating_expenditure = cogs_cycle + operating_expenses_ex_cogs
    st.metric("Calculated Annual Operating Expenditure", f"{annual_operating_expenditure:,.2f} {unit}")


with tab_wc:
    st.checkbox("I have reviewed the working-capital assessment inputs", key="wc_reviewed")
    st.subheader("Projected working-capital assessment")
    st.caption("Enter projected current assets and other current liabilities for the assessment period.")
    d1, d2 = st.columns(2)
    with d1:
        st.markdown("#### Projected current assets")
        cma_stock_basis = st.radio("Stock information available", ["Consolidated stock figure", "Detailed stock breakup"], horizontal=True, key="cma_stock_basis")
        if cma_stock_basis == "Consolidated stock figure":
            cma_stock_total = st.number_input("Projected stocks / inventory as per balance sheet", min_value=0.0, value=110.0, key="cma_consolidated_stock")
        else:
            st.caption("Enter only the categories disclosed by the borrower. Leave undisclosed or inapplicable categories at zero.")
            cma_rm = st.number_input("Raw materials (optional)", min_value=0.0, value=0.0, key="cma_rm")
            cma_wip = st.number_input("Work in progress (optional)", min_value=0.0, value=0.0, key="cma_wip")
            cma_fg = st.number_input("Finished goods (optional)", min_value=0.0, value=0.0, key="cma_fg")
            cma_trade = st.number_input("Stock-in-trade (optional)", min_value=0.0, value=0.0, key="cma_trade")
            cma_other_stock = st.number_input("Other stocks (optional)", min_value=0.0, value=0.0, key="cma_other_stock")
            cma_stock_total = sum([cma_rm, cma_wip, cma_fg, cma_trade, cma_other_stock])
        st.metric("Total projected stocks", f"{cma_stock_total:,.2f} {unit}")
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

with tab_dp:
    st.checkbox("I have reviewed the drawing-power inputs and eligibility", key="dp_reviewed")
    st.subheader("Drawing Power calculator")
    st.markdown("""<div class="fc-note">Calculate DP from eligible stocks and eligible book debts, less the creditor amount required to be deducted under the bank's policy. The MSME margin defaults to 25% and remains editable.</div>""", unsafe_allow_html=True)
    category = st.selectbox("Borrower category", ["MSME", "Other"], key="dp_category")
    default_margin = 25.0 if category == "MSME" else 30.0
    margin_pct = st.number_input("Margin (%)", min_value=0.0, max_value=100.0, value=default_margin, step=1.0, key="dp_margin")
    st.markdown("#### 1. Eligible stocks")
    dp_stock_basis = st.radio("Stock information available", ["Consolidated stock figure", "Detailed stock breakup"], horizontal=True, key="dp_stock_basis")
    if dp_stock_basis == "Consolidated stock figure":
        gross_stock = st.number_input("Gross stocks as per latest stock statement / balance sheet", min_value=0.0, value=110.0, key="dp_consolidated_stock")
    else:
        st.caption("Enter only the stock categories disclosed. Leave unavailable or inapplicable categories at zero.")
        q1,q2,q3 = st.columns(3)
        with q1: dp_rm=st.number_input("Raw materials (optional)",min_value=0.0,value=0.0,key="dp_rm")
        with q2: dp_wip=st.number_input("Work in progress (optional)",min_value=0.0,value=0.0,key="dp_wip")
        with q3: dp_fg=st.number_input("Finished goods (optional)",min_value=0.0,value=0.0,key="dp_fg")
        q4,q5 = st.columns(2)
        with q4: dp_trade=st.number_input("Stock-in-trade (optional)",min_value=0.0,value=0.0,key="dp_trade")
        with q5: dp_other=st.number_input("Other eligible stock (optional)",min_value=0.0,value=0.0,key="dp_other")
        gross_stock=sum([dp_rm,dp_wip,dp_fg,dp_trade,dp_other])
    ineligible_stock=st.number_input("Less: ineligible / obsolete stock",min_value=0.0,value=0.0,key="dp_bad_stock")
    eligible_stock=max(gross_stock-ineligible_stock,0.0)
    st.metric("Eligible Stocks",f"{eligible_stock:,.2f} {unit}")
    st.markdown("#### 2. Eligible book debts and creditor adjustment")
    r1,r2,r3=st.columns(3)
    with r1: gross_debts=st.number_input("Book debts / receivables",min_value=0.0,value=60.0,key="dp_debts")
    with r2: bad_debts=st.number_input("Less: ineligible / overdue book debts",min_value=0.0,value=0.0,key="dp_bad_debts")
    with r3: dp_creditors=st.number_input("Sundry creditors to be deducted",min_value=0.0,value=40.0,key="dp_creditors",help="Enter the amount required to be deducted under the sanction terms and bank policy.")
    eligible_debts=max(gross_debts-bad_debts,0.0)
    st.metric("Eligible Book Debts",f"{eligible_debts:,.2f} {unit}")
    st.markdown("#### 3. Facility position")
    f1,f2=st.columns(2)
    with f1: dp_sanctioned=st.number_input("Sanctioned CC limit",min_value=0.0,value=80.0,key="dp_sanctioned")
    with f2: dp_outstanding=st.number_input("Present CC outstanding",min_value=0.0,value=25.0,key="dp_outstanding")
    net_dp_base=eligible_stock+eligible_debts-dp_creditors
    raw_dp=max(net_dp_base,0.0)*(1-margin_pct/100.0)
    usable_dp=min(raw_dp,dp_sanctioned)
    available_dp=max(usable_dp-dp_outstanding,0.0)
    excess_dp=max(dp_outstanding-usable_dp,0.0)
    dp_summary={"Borrower category":category,"Margin (%)":margin_pct,"Gross stocks":gross_stock,"Ineligible stocks":ineligible_stock,"Eligible stocks":eligible_stock,"Gross book debts":gross_debts,"Ineligible book debts":bad_debts,"Eligible book debts":eligible_debts,"Sundry creditors deducted":dp_creditors,"Net eligible assets before margin":net_dp_base,"Calculated DP":raw_dp,"Sanctioned CC limit":dp_sanctioned,"Usable DP":usable_dp,"CC outstanding":dp_outstanding,"Available DP":available_dp,"Outstanding above usable DP":excess_dp}
    m1,m2,m3,m4=st.columns(4)
    m1.metric("Calculated DP",f"{raw_dp:,.2f} {unit}")
    m2.metric("Usable DP",f"{usable_dp:,.2f} {unit}")
    m3.metric("Outstanding",f"{dp_outstanding:,.2f} {unit}")
    m4.metric("Available DP",f"{available_dp:,.2f} {unit}")
    st.dataframe(pd.DataFrame([dp_summary]),use_container_width=True,hide_index=True)
    if ineligible_stock>gross_stock or bad_debts>gross_debts: st.warning("An ineligible amount exceeds its gross balance. Please review the entries.")
    if net_dp_base<0: st.warning("Net eligible assets are negative; DP has been floored at zero before applying the margin.")
    if excess_dp>0: st.error(f"Outstanding exceeds usable DP by {excess_dp:,.2f} {unit}.")
    st.caption("The 25% MSME margin is a configurable default for this application, not a universal RBI-prescribed margin. Apply sanction-specific margins, ageing criteria and creditor deductions.")

with tab_profile:
    st.caption("All figures are in the selected unit. Review the balance-sheet tally in the Analysis tab.")

# ---------------- CALCULATIONS ----------------
assets = {
    "Cash and Bank Balance": cash_bank,
    "Sundry Debtors": debtors_bs,
    "Stocks": stocks_bs,
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
current_assets = cash_bank + debtors_bs + stocks_bs + current_advances
current_liabilities = cc_borrowings + sundry_creditors + other_current_liabilities
net_working_capital = current_assets - current_liabilities
current_ratio = safe_divide(current_assets, current_liabilities)
quick_ratio = safe_divide(current_assets - stocks_bs, current_liabilities)

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

if st.button("Run FINCREDIT Analysis  →", type="primary", use_container_width=True):
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
    if history_valid.empty:
        st.error("Enter at least one financial year.")
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

    stock_days = safe_divide(avg_inventory * days_in_year, cogs_cycle)
    debtor_days = safe_divide(avg_debtors * days_in_year, sales_cycle)
    creditor_days = safe_divide(avg_creditors * days_in_year, total_purchases)
    gross_cycle = stock_days + debtor_days if pd.notna(stock_days) and pd.notna(debtor_days) else np.nan
    net_cycle = gross_cycle - creditor_days if pd.notna(gross_cycle) and pd.notna(creditor_days) else np.nan
    cycle_funding = (
        annual_operating_expenditure / days_in_year * net_cycle
        if pd.notna(net_cycle) else np.nan
    )

    total_cma_ca = sum([
        cma_stock_total, cma_debtors, cma_cash, cma_other_assets
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

    # Display automatically calculated working-capital figures in the WC tab.
    with tab_wc:
        st.markdown("---")
        st.markdown("#### Automatically calculated working capital")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total projected current assets", money(total_cma_ca, unit))
        m2.metric("Other current liabilities (excluding bank borrowing)", money(total_other_cl, unit))
        m3.metric("Working Capital Gap", money(wc_gap, unit))
        st.caption("Calculated as projected current assets minus other current liabilities, excluding bank borrowing. This is the projected working-capital gap used in the MPBF calculations.")

    stock_turnover = safe_divide(latest_cogs, avg_inventory)
    debtor_turnover = safe_divide(sales_cycle, avg_debtors)
    creditor_turnover = safe_divide(total_purchases, avg_creditors)
    interest_coverage = safe_divide(latest_ebit, latest_interest)
    gross_margin = safe_divide(latest_gp * 100, latest_sales)
    ebitda_margin = safe_divide(ebitda * 100, latest_sales)
    operating_margin = safe_divide(latest_ebit * 100, latest_sales)
    net_margin = safe_divide(latest_np * 100, latest_sales)

    ratios = pd.DataFrame([
        ["Current Ratio", current_ratio, "times", "Current Assets / Current Liabilities"],
        ["Quick Ratio", quick_ratio, "times", "(Current Assets − Stocks) / Current Liabilities"],
        ["Debt–Equity", debt_equity, "times", "Total Borrowings / Adjusted TNW"],
        ["TOL / TNW", tol_tnw, "times", "Total Outside Liabilities / Adjusted TNW"],
        ["Interest Coverage", interest_coverage, "times", "EBIT / Interest Expense"],
        ["Gross Profit Margin", gross_margin, "%", "Gross Profit / Sales × 100"],
        ["EBITDA Margin", ebitda_margin, "%", "(EBIT + Depreciation) / Sales × 100"],
        ["Operating Profit Margin", operating_margin, "%", "EBIT / Sales × 100"],
        ["Net Profit Margin", net_margin, "%", "Net Profit / Sales × 100"],
        ["Stock Turnover Ratio", stock_turnover, "times", "COGS / Average Stocks"],
        ["Debtor Turnover", debtor_turnover, "times", "Sales / Average Debtors (simplified)"],
        ["Creditor Turnover", creditor_turnover, "times", "Purchases / Average Creditors (simplified)"],
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
        "Stock holding period (days)": stock_days,
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
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Current ratio", f"{current_ratio:.2f}:1" if pd.notna(current_ratio) else "N/A")
        k2.metric("Quick ratio", f"{quick_ratio:.2f}:1" if pd.notna(quick_ratio) else "N/A")
        k3.metric("Net Working Capital (NWC)", money(net_working_capital, unit))
        k4.metric("Adjusted TNW", money(adjusted_tnw, unit))
        k5.metric("Selected WC assessment", money(selected_assessment, unit))

        st.subheader("Financial ratios")
        ratio_display = ratios.copy()
        ratio_display["Result"] = ratio_display["Result"].map(lambda x: "N/A" if pd.isna(x) else f"{x:,.2f}")
        st.dataframe(ratio_display, use_container_width=True, hide_index=True)

        st.subheader("Financial trend analysis")
        st.dataframe(trend, use_container_width=True, hide_index=True)
        if len(trend) >= 2:
            st.line_chart(trend.set_index("Financial Year")[["Sales", "Capital", "Net Profit", "Borrowings"]])
        else:
            st.info("Only one financial year is available. Financial ratios and assessments are calculated, but year-on-year trends require at least two years.")

        st.subheader("Operating cycle")
        st.dataframe(pd.DataFrame([oc_summary]), use_container_width=True, hide_index=True)
        if pd.notna(net_cycle) and net_cycle < 0:
            st.warning("Negative net operating cycle: verify the input assumptions.")
        st.caption("Operating-cycle funding estimate is a simplified indicator, not a sanctioned limit or drawing power.")

        st.subheader("Working-capital assessment")
        st.dataframe(pd.DataFrame([wc_summary]), use_container_width=True, hide_index=True)
        st.subheader("Drawing Power")
        st.dataframe(pd.DataFrame([dp_summary]), use_container_width=True, hide_index=True)
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
            oc_summary, wc_summary, dp_summary, observations
        )
        st.download_button(
            "Download analysis workbook (Excel)",
            data=report,
            file_name=f"{company_name.strip().replace(' ', '_') or 'borrower'}_FINCREDIT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        try:
            pdf_report = make_pdf_report(
                company_name, business_activity, constitution, financial_year, unit,
                assets, liabilities, trend, ratios, oc_summary, wc_summary,
                dp_summary, observations
            )
            st.download_button(
                "Download credit analysis report (PDF)",
                data=pdf_report,
                file_name=f"{company_name.strip().replace(' ', '_') or 'borrower'}_FINCREDIT_Report.pdf",
                mime="application/pdf",
                type="primary",
            )
        except Exception as exc:
            st.error(f"PDF report could not be generated: {exc}")

else:
    st.warning("Enter or review the sample data in the tabs, then click **Analyse Financial Statements**.")
    st.caption("The sample borrower and amounts are illustrative. Replace them with verified borrower data.")
