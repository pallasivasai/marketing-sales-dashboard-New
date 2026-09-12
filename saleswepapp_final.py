import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# =========================================================
# MARKETING SALES + EMPLOYEE PERFORMANCE DASHBOARD
# =========================================================

st.set_page_config(
    page_title="Marketing Sales & Employee Performance Dashboard",
    page_icon="📊",
    layout="wide"
)


# ================= FONT =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR, "fonts")

try:
    pdfmetrics.registerFont(
        TTFont("DejaVu", os.path.join(FONT_DIR, "DejaVuSans.ttf"))
    )
    pdfmetrics.registerFont(
        TTFont("DejaVu-Bold", os.path.join(FONT_DIR, "DejaVuSans-Bold.ttf"))
    )
    PDF_FONT = "DejaVu"
    PDF_BOLD = "DejaVu-Bold"
except Exception:
    PDF_FONT = "Helvetica"
    PDF_BOLD = "Helvetica-Bold"


# ================= CONFIG =================
SALES_SHEET = "MAIN_COPY"
TARGET_SHEET = "MARKETING TARGET"
MAKE_TARGET_SHEET = "MAKE TARGET"
NEW_CUSTOMER_SHEET = "Merge1"

USERS = {
    "admin": {
        "password": "admin@123",
        "marketing": "ALL",
        "role": "Administrator / HR"
    },
    "ashok": {
        "password": "ashok@123",
        "marketing": "Ashok Marketing",
        "role": "Marketing Employee"
    },
    "suresh": {
        "password": "suresh@123",
        "marketing": "Suresh - Marketing",
        "role": "Marketing Employee"
    },
    "ho": {
        "password": "ho@123",
        "marketing": "H O - Marketing",
        "role": "Marketing Employee"
    },
}

MONTH_MAP = {
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
    "JAN": 1,
    "FEB": 2,
    "MAR": 3
}


# ================= CUSTOM CSS =================
st.markdown(
    """
    <style>
    .main-title {
        font-size: 34px;
        font-weight: 700;
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 17px;
        color: #666;
        margin-bottom: 20px;
    }
    .status-reached {
        font-weight: 700;
    }
    .status-not-reached {
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ================= EXCEL LOADER =================
@st.cache_data(show_spinner="Loading Excel data...")
def load_excel_cached(file_bytes):

    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = xl.sheet_names

    if SALES_SHEET not in sheets:
        raise ValueError(
            f"Worksheet '{SALES_SHEET}' not found. "
            f"Available sheets: {', '.join(sheets)}"
        )

    sales_df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=SALES_SHEET
    )

    if TARGET_SHEET in sheets:
        target_raw = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=TARGET_SHEET
        )
    else:
        target_raw = pd.DataFrame()

    if MAKE_TARGET_SHEET in sheets:
        make_target_df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=MAKE_TARGET_SHEET
        )
    else:
        make_target_df = pd.DataFrame()

    if NEW_CUSTOMER_SHEET in sheets:
        new_customer_df = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=NEW_CUSTOMER_SHEET
        )
    else:
        new_customer_df = pd.DataFrame()

    return (
        sales_df,
        target_raw,
        make_target_df,
        new_customer_df,
        sheets
    )


# ================= TARGET STATUS HELPER =================

def add_target_status_columns(df, target_col="Target", sales_col="Value"):
    """Add achievement, gap, and explicit target-reach status dynamically."""
    out = df.copy()
    out[target_col] = pd.to_numeric(out[target_col], errors="coerce").fillna(0)
    out[sales_col] = pd.to_numeric(out[sales_col], errors="coerce").fillna(0)

    out["Achievement %"] = 0.0
    mask = out[target_col] > 0
    out.loc[mask, "Achievement %"] = (
        out.loc[mask, sales_col] / out.loc[mask, target_col] * 100
    )
    out["Achievement %"] = out["Achievement %"].round(1)

    out["Target Gap"] = (
        out[target_col] - out[sales_col]
    ).clip(lower=0).round(0)

    out["Target Reach"] = out["Achievement %"].apply(
        lambda x: "YES" if x >= 100 else "NO"
    )

    out["Target Status"] = out["Achievement %"].apply(
        lambda x: "✅ TARGET REACHED" if x >= 100 else "❌ NOT REACHED"
    )
    return out

# ================= PDF =================
def generate_employee_pdf(employee_name, employee_df):

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4
    )

    styles = getSampleStyleSheet()
    styles["Title"].fontName = PDF_BOLD
    styles["Normal"].fontName = PDF_FONT

    elements = []

    elements.append(
        Paragraph(
            "EMPLOYEE PERFORMANCE REPORT",
            styles["Title"]
        )
    )

    elements.append(Spacer(1, 15))

    total_target = employee_df["Target"].sum()
    total_sales = employee_df["Sales"].sum()
    gap = max(total_target - total_sales, 0)

    achievement = (
        total_sales / total_target * 100
        if total_target else 0
    )

    reached_months = int(
        (employee_df["Achievement %"] >= 100).sum()
    )

    total_months = len(employee_df)

    summary_data = [
        ["Employee", employee_name],
        ["Total Target", f"₹ {total_target:,.0f}"],
        ["Total Sales", f"₹ {total_sales:,.0f}"],
        ["Target Gap", f"₹ {gap:,.0f}"],
        ["Achievement", f"{achievement:.1f}%"],
        ["Months Target Reached", f"{reached_months}/{total_months}"],
    ]

    summary = Table(
        summary_data,
        colWidths=[180, 280]
    )

    summary.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("BACKGROUND", (0, 0), (0, -1), colors.lightgrey),
            ("FONTNAME", (0, 0), (0, -1), PDF_BOLD),
            ("FONTNAME", (1, 0), (1, -1), PDF_FONT),
        ])
    )

    elements.append(summary)
    elements.append(Spacer(1, 20))

    detail_data = [
        ["Month", "Target", "Sales", "Achievement", "Target Reach", "Status"]
    ]

    for _, row in employee_df.iterrows():

        status = (
            "TARGET REACHED"
            if row["Achievement %"] >= 100
            else "NOT REACHED"
        )

        detail_data.append([
            row["Month"],
            f"₹ {row['Target']:,.0f}",
            f"₹ {row['Sales']:,.0f}",
            f"{row['Achievement %']:.1f}%",
            "YES" if row["Achievement %"] >= 100 else "NO",
            status
        ])

    detail = Table(
        detail_data,
        repeatRows=1
    )

    detail.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("FONTNAME", (0, 0), (-1, 0), PDF_BOLD),
        ])
    )

    elements.append(detail)

    doc.build(elements)

    buffer.seek(0)

    return buffer


# ================= LOGIN =================
def login():

    st.markdown(
        '<div class="main-title">🔐 Marketing & Employee Performance Login</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sub-title">Secure role-based dashboard access</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        username = st.text_input(
            "Username",
            placeholder="Enter username"
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter password"
        )

        if st.button(
            "🔓 Login",
            use_container_width=True
        ):

            if (
                username in USERS
                and USERS[username]["password"] == password
            ):

                st.session_state["user"] = username
                st.session_state["marketing"] = USERS[username]["marketing"]
                st.session_state["role"] = USERS[username]["role"]

                st.rerun()

            else:

                st.error("Invalid username or password.")


# ================= EMPLOYEE PERFORMANCE =================
def employee_performance_section(
    sales_df,
    target_df,
    is_admin
):

    st.subheader("👥 Employee Performance Analytics")

    if target_df.empty:

        st.warning(
            "MARKETING TARGET sheet is required for employee target performance."
        )

        return

    employees = sorted(
        target_df["MARK"]
        .dropna()
        .unique()
    )

    if not employees:

        st.warning("No employees found in MARKETING TARGET.")
        return

    selected_employee = st.selectbox(
        "Select Employee",
        employees,
        key="employee_performance_select"
    )

    employee_sales = sales_df[
        sales_df["MARK"] == selected_employee
    ]

    monthly_sales = (
        employee_sales
        .groupby(
            ["Safe_YearMonth", "Month_Text"],
            as_index=False
        )["Value"]
        .sum()
    )

    employee_target = target_df[
        target_df["MARK"] == selected_employee
    ][
        ["Safe_YearMonth", "Target"]
    ]

    employee_report = pd.merge(
        employee_target,
        monthly_sales,
        on="Safe_YearMonth",
        how="left"
    )

    employee_report["Value"] = (
        employee_report["Value"]
        .fillna(0)
    )

    employee_report["Target"] = (
        employee_report["Target"]
        .fillna(0)
    )

    employee_report["Achievement %"] = 0.0

    mask = employee_report["Target"] > 0

    employee_report.loc[mask, "Achievement %"] = (
        employee_report.loc[mask, "Value"]
        / employee_report.loc[mask, "Target"]
        * 100
    )

    employee_report["Achievement %"] = (
        employee_report["Achievement %"]
        .round(1)
    )

    employee_report["Status"] = employee_report[
        "Achievement %"
    ].apply(
        lambda x: "✅ TARGET REACHED"
        if x >= 100
        else "❌ NOT REACHED"
    )

    employee_report["Target Reach"] = employee_report[
        "Achievement %"
    ].apply(lambda x: "YES" if x >= 100 else "NO")

    employee_report["Gap"] = (
        employee_report["Target"]
        - employee_report["Value"]
    ).clip(lower=0)

    employee_report = employee_report.sort_values(
        "Safe_YearMonth"
    )

    # ---- Summary ----
    total_target = employee_report["Target"].sum()
    total_sales = employee_report["Value"].sum()
    target_gap = max(total_target - total_sales, 0)

    overall_achievement = (
        total_sales / total_target * 100
        if total_target else 0
    )

    reached_months = int(
        (employee_report["Achievement %"] >= 100).sum()
    )

    missed_months = int(
        (employee_report["Achievement %"] < 100).sum()
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric(
        "🎯 Total Target",
        f"₹ {total_target:,.0f}"
    )

    c2.metric(
        "💰 Achieved Sales",
        f"₹ {total_sales:,.0f}"
    )

    c3.metric(
        "📊 Achievement",
        f"{overall_achievement:.1f}%"
    )

    c4.metric(
        "🎯 Months Reached",
        reached_months
    )

    c5.metric(
        "❌ Months Missed",
        missed_months
    )

    st.markdown("---")

    # ---- Overall status ----
    if overall_achievement >= 100:

        st.success(
            f"🏆 {selected_employee} has reached the overall target."
        )

    else:

        st.warning(
            f"⚠️ {selected_employee} has not reached the overall target. "
            f"Current gap: ₹ {target_gap:,.0f}"
        )

    # ---- Monthly table ----
    st.markdown("### 📋 Monthly Employee Performance")

    display_df = employee_report[
        [
            "Month_Text",
            "Target",
            "Value",
            "Achievement %",
            "Gap",
            "Target Reach",
            "Status"
        ]
    ].copy()

    display_df.columns = [
        "Month",
        "Target",
        "Sales",
        "Achievement %",
        "Target Gap",
        "Target Reach",
        "Status"
    ]

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )

    # ---- Charts ----
    st.markdown("### 📈 Employee Target vs Actual")

    chart_df = employee_report[
        ["Month_Text", "Target", "Value"]
    ].copy()

    chart_df = chart_df.set_index(
        "Month_Text"
    )

    chart_df.columns = [
        "Target",
        "Sales"
    ]

    st.bar_chart(chart_df)

    st.markdown("### 📊 Achievement % Trend")

    achievement_chart = employee_report[
        ["Month_Text", "Achievement %"]
    ].copy()

    achievement_chart = achievement_chart.set_index(
        "Month_Text"
    )

    st.line_chart(achievement_chart)

    # ---- HR interpretation ----
    st.markdown("### 🧑‍💼 HR Performance Summary")

    if overall_achievement >= 100:

        st.success(
            f"**{selected_employee}** is currently at "
            f"**{overall_achievement:.1f}%** of the assigned target."
        )

    elif overall_achievement >= 90:

        st.info(
            f"**{selected_employee}** is at "
            f"**{overall_achievement:.1f}%**. "
            f"The employee is close to the overall target."
        )

    elif overall_achievement >= 75:

        st.warning(
            f"**{selected_employee}** is at "
            f"**{overall_achievement:.1f}%**. "
            f"Performance is below target and should be reviewed."
        )

    else:

        st.error(
            f"**{selected_employee}** is at "
            f"**{overall_achievement:.1f}%**. "
            f"A detailed performance review may be appropriate."
        )

    # ---- PDF ----
    pdf_df = employee_report[
        ["Month_Text", "Target", "Value", "Achievement %", "Target Reach"]
    ].copy()

    pdf_df.columns = [
        "Month",
        "Target",
        "Sales",
        "Achievement %"
    ]

    pdf = generate_employee_pdf(
        selected_employee,
        pdf_df
    )

    st.download_button(
        "📄 Download Employee Performance PDF",
        pdf,
        file_name=(
            selected_employee
            .replace(" ", "_")
            .replace("-", "_")
            + "_Performance_Report.pdf"
        ),
        mime="application/pdf"
    )


# ================= EMPLOYEE RANKING =================
def employee_ranking_section(sales_df, target_df):

    st.subheader("🏆 Employee Performance Ranking")

    if target_df.empty:
        st.info(
            "Add MARKETING TARGET sheet to enable employee ranking."
        )
        return

    sales_summary = (
        sales_df
        .groupby("MARK", as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "Sales"})
    )

    target_summary = (
        target_df
        .groupby("MARK", as_index=False)["Target"]
        .sum()
    )

    ranking = pd.merge(
        target_summary,
        sales_summary,
        on="MARK",
        how="left"
    )

    ranking["Sales"] = (
        ranking["Sales"]
        .fillna(0)
    )

    ranking["Achievement %"] = 0.0

    mask = ranking["Target"] > 0

    ranking.loc[mask, "Achievement %"] = (
        ranking.loc[mask, "Sales"]
        / ranking.loc[mask, "Target"]
        * 100
    )

    ranking["Achievement %"] = (
        ranking["Achievement %"]
        .round(1)
    )

    ranking["Target Reach"] = ranking[
        "Achievement %"
    ].apply(
        lambda x: "YES" if x >= 100 else "NO"
    )

    ranking["Status"] = ranking[
        "Achievement %"
    ].apply(
        lambda x: "✅ Reached"
        if x >= 100
        else "❌ Not Reached"
    )

    ranking = ranking.sort_values(
        "Achievement %",
        ascending=False
    ).reset_index(drop=True)

    ranking.insert(
        0,
        "Rank",
        ranking.index + 1
    )

    st.dataframe(
        ranking.rename(
            columns={
                "MARK": "Employee",
                "Target": "Total Target",
                "Sales": "Total Sales"
            }
        ),
        use_container_width=True,
        hide_index=True
    )


# ================= DASHBOARD =================
def dashboard():

    marketing = st.session_state["marketing"]
    role = st.session_state["role"]

    is_admin = marketing == "ALL"

    # ---- Header ----
    st.markdown(
        '<div class="main-title">📊 Marketing Sales & Employee Performance Dashboard</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        f'<div class="sub-title">Logged in as: <b>{marketing}</b> | Role: <b>{role}</b></div>',
        unsafe_allow_html=True
    )

    # ---- Sidebar ----
    st.sidebar.title("Dashboard Menu")

    if st.sidebar.button(
        "🚪 Logout",
        use_container_width=True
    ):

        st.session_state.clear()
        st.rerun()

    # ---- Upload ----
    if is_admin:

        st.sidebar.markdown("---")

        uploaded_file = st.sidebar.file_uploader(
            "📤 Upload Excel Data",
            type=["xlsx"]
        )

        if uploaded_file is not None:

            st.session_state["file_bytes"] = (
                uploaded_file.getvalue()
            )

    if "file_bytes" not in st.session_state:

        st.info(
            "👈 Admin: Upload the Marketing_Sales_Dashboard_Data.xlsx file from the sidebar."
        )

        st.stop()

    # ---- Load ----
    try:

        (
            sales_df,
            target_raw,
            make_target_df,
            new_customer_df,
            available_sheets
        ) = load_excel_cached(
            st.session_state["file_bytes"]
        )

    except Exception as e:

        st.error("❌ Excel loading failed.")

        st.code(str(e))

        st.stop()

    # ================= VALIDATION =================
    required_sales_columns = [
        "Date",
        "MARK",
        "make",
        "HELPER",
        "Value",
        "CUSTOMER NAME"
    ]

    missing = [
        c for c in required_sales_columns
        if c not in sales_df.columns
    ]

    if missing:

        st.error(
            "❌ Required columns are missing from MAIN_COPY."
        )

        st.write(
            "Missing columns:",
            missing
        )

        st.write(
            "Available columns:",
            list(sales_df.columns)
        )

        st.stop()

    # ================= SALES CLEANING =================
    sales_df["Date"] = pd.to_datetime(
        sales_df["Date"],
        errors="coerce"
    )

    sales_df = sales_df.dropna(
        subset=["Date"]
    )

    sales_df["Safe_Year"] = (
        sales_df["Date"].dt.year
    )

    sales_df["Safe_Month"] = (
        sales_df["Date"].dt.month
    )

    sales_df["Safe_YearMonth"] = (
        sales_df["Safe_Year"] * 100
        + sales_df["Safe_Month"]
    )

    sales_df["Month_Text"] = (
        sales_df["Date"]
        .dt.strftime("%b")
        .str.upper()
    )

    sales_df["MARK"] = (
        sales_df["MARK"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    sales_df["make"] = (
        sales_df["make"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    sales_df["HELPER"] = (
        sales_df["HELPER"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    sales_df = sales_df[
        sales_df["HELPER"].isin(
            ["NOFILL", "GREEN"]
        )
    ]

    sales_df["Value"] = pd.to_numeric(
        sales_df["Value"],
        errors="coerce"
    ).fillna(0)

    sales_df["CUSTOMER NAME"] = (
        sales_df["CUSTOMER NAME"]
        .astype(str)
        .str.upper()
        .str.strip()
    )

    # ================= TARGET PROCESSING =================
    if (
        not target_raw.empty
        and "Marketing Person" in target_raw.columns
    ):

        target_df = target_raw.melt(
            id_vars=["Marketing Person"],
            var_name="Month",
            value_name="Target"
        )

        target_df = target_df.rename(
            columns={
                "Marketing Person": "MARK"
            }
        )

        target_df["Target"] = (
            target_df["Target"]
            .astype(str)
            .str.replace("₹", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )

        target_df["Target"] = pd.to_numeric(
            target_df["Target"],
            errors="coerce"
        ).fillna(0)

        target_df["Month_No"] = (
            target_df["Month"]
            .astype(str)
            .str.upper()
            .map(MONTH_MAP)
        )

        target_df = target_df.dropna(
            subset=["Month_No"]
        )

        target_df["Month_No"] = (
            target_df["Month_No"]
            .astype(int)
        )

        target_df["Year"] = target_df[
            "Month_No"
        ].apply(
            lambda x: 2025 if x >= 4 else 2026
        )

        target_df["Safe_YearMonth"] = (
            target_df["Year"] * 100
            + target_df["Month_No"]
        )

        target_df["MARK"] = (
            target_df["MARK"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

    else:

        target_df = pd.DataFrame(
            columns=[
                "MARK",
                "Month",
                "Target",
                "Month_No",
                "Year",
                "Safe_YearMonth"
            ]
        )

    # ================= MARKETING FILTER =================
    selected_marketing = "ALL"

    if is_admin:

        if not target_df.empty:

            options = sorted(
                target_df["MARK"]
                .dropna()
                .unique()
            )

        else:

            options = sorted(
                sales_df["MARK"]
                .dropna()
                .unique()
            )

        selected_marketing = st.selectbox(
            "Select Marketing Person",
            ["ALL"] + options
        )

        if selected_marketing != "ALL":

            sales_df = sales_df[
                sales_df["MARK"]
                == selected_marketing
            ]

            if not target_df.empty:

                target_df = target_df[
                    target_df["MARK"]
                    == selected_marketing
                ]

    else:

        sales_df = sales_df[
            sales_df["MARK"]
            == marketing.upper()
        ]

        if not target_df.empty:

            target_df = target_df[
                target_df["MARK"]
                == marketing.upper()
            ]

    # ================= TOP KPI =================
    st.subheader("📌 Dashboard Overview")

    total_sales = sales_df["Value"].sum()
    total_orders = len(sales_df)
    total_customers = sales_df[
        "CUSTOMER NAME"
    ].nunique()

    total_cost = 0

    if "Cost" in sales_df.columns:

        total_cost = pd.to_numeric(
            sales_df["Cost"],
            errors="coerce"
        ).fillna(0).sum()

    profit = total_sales - total_cost

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "💰 Total Sales",
        f"₹ {total_sales:,.0f}"
    )

    c2.metric(
        "📦 Total Orders",
        f"{total_orders:,}"
    )

    c3.metric(
        "👥 Customers",
        f"{total_customers:,}"
    )

    c4.metric(
        "💵 Profit",
        f"₹ {profit:,.0f}"
    )

    # ================= MONTHLY REPORT =================
    st.subheader("📊 Sales Performance Report")

    monthly_sales = (
        sales_df
        .groupby(
            ["MARK", "Safe_YearMonth", "Month_Text"],
            as_index=False
        )["Value"]
        .sum()
    )

    if not target_df.empty:

        monthly_report = pd.merge(
            monthly_sales,
            target_df[
                ["MARK", "Safe_YearMonth", "Target"]
            ],
            on=[
                "MARK",
                "Safe_YearMonth"
            ],
            how="left"
        )

    else:

        monthly_report = monthly_sales.copy()

        monthly_report["Target"] = 0

    monthly_report["Target"] = (
        monthly_report["Target"]
        .fillna(0)
    )

    monthly_report["Achievement %"] = 0.0

    mask = monthly_report["Target"] > 0

    monthly_report.loc[mask, "Achievement %"] = (
        monthly_report.loc[mask, "Value"]
        / monthly_report.loc[mask, "Target"]
        * 100
    )

    monthly_report["Achievement %"] = (
        monthly_report["Achievement %"]
        .round(1)
    )

    monthly_report["Target Gap"] = (
        monthly_report["Target"] - monthly_report["Value"]
    ).clip(lower=0).round(0)

    monthly_report["Target Reach"] = monthly_report[
        "Achievement %"
    ].apply(lambda x: "YES" if x >= 100 else "NO")

    # ================= TARGET STATUS =================
    total_target = monthly_report["Target"].sum()

    achievement = (
        total_sales / total_target * 100
        if total_target else 0
    )

    target_gap = max(
        total_target - total_sales,
        0
    )

    st.markdown("### 🎯 Overall Target Status")

    k1, k2, k3, k4 = st.columns(4)

    k1.metric(
        "Total Target",
        f"₹ {total_target:,.0f}"
    )

    k2.metric(
        "Achieved",
        f"₹ {total_sales:,.0f}"
    )

    k3.metric(
        "Achievement %",
        f"{achievement:.1f}%"
    )

    k4.metric(
        "Target Gap",
        f"₹ {target_gap:,.0f}"
    )

    if total_target > 0:

        if achievement >= 100:

            st.success(
                "🏆 Overall target reached!"
            )

        else:

            st.warning(
                f"⚠️ Overall target not reached. "
                f"Gap: ₹ {target_gap:,.0f}"
            )

    # ================= MONTHLY STATUS TABLE =================
    st.markdown("### 📋 Month-wise Target Achievement")

    status_df = monthly_report.copy()

    status_df["Status"] = status_df[
        "Achievement %"
    ].apply(
        lambda x: (
            "✅ TARGET REACHED"
            if x >= 100
            else "❌ NOT REACHED"
        )
    )

    status_display = status_df[
        [
            "MARK",
            "Month_Text",
            "Target",
            "Value",
            "Achievement %",
            "Target Gap",
            "Target Reach",
            "Status"
        ]
    ].rename(
        columns={
            "MARK": "Employee",
            "Month_Text": "Month",
            "Value": "Sales",
            "Target Gap": "Target Gap",
            "Target Reach": "Target Reach"
        }
    )

    st.dataframe(
        status_display,
        use_container_width=True,
        hide_index=True
    )

    # ================= CHART =================
    st.subheader("📈 Month-wise Target vs Sales")

    if not monthly_report.empty:

        chart_df = (
            monthly_report
            .groupby(
                "Month_Text",
                as_index=False
            )[["Target", "Value"]]
            .sum()
        )

        chart_df = chart_df.set_index(
            "Month_Text"
        )

        chart_df.columns = [
            "Target",
            "Sales"
        ]

        st.bar_chart(chart_df)

    # ================= EMPLOYEE RANKING =================
    if is_admin:

        employee_ranking_section(
            sales_df=sales_df if selected_marketing == "ALL"
            else pd.concat([]) if False else (
                # Re-load all filtered data isn't necessary for single selection.
                sales_df
            ),
            target_df=target_df
        )

        # For admin + ALL, ranking above works on all employees.
        # For a selected employee, it intentionally shows that employee only.

    # ================= EMPLOYEE PERFORMANCE =================
    employee_performance_section(
        sales_df=sales_df,
        target_df=target_df,
        is_admin=is_admin
    )

    # ================= NEW CUSTOMER =================
    st.subheader("🆕 New Customer Report")

    if (
        not new_customer_df.empty
        and "CUSTOMER NAME" in new_customer_df.columns
    ):

        new_customer_df["CUSTOMER NAME"] = (
            new_customer_df["CUSTOMER NAME"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

        new_customer_sales = sales_df[
            sales_df["CUSTOMER NAME"].isin(
                new_customer_df["CUSTOMER NAME"]
            )
        ]

        new_customer_count = (
            new_customer_sales[
                "CUSTOMER NAME"
            ].nunique()
        )

        new_customer_sales_value = (
            new_customer_sales["Value"].sum()
        )

    else:

        new_customer_count = 0
        new_customer_sales_value = 0

    n1, n2 = st.columns(2)

    n1.metric(
        "New Customers",
        new_customer_count
    )

    n2.metric(
        "New Customer Sales",
        f"₹ {new_customer_sales_value:,.0f}"
    )

    # ================= BRAND WISE =================
    st.subheader("🏷️ Brand-wise Sales")

    if (
        not make_target_df.empty
        and "Make" in make_target_df.columns
        and "Target" in make_target_df.columns
    ):

        make_target_df["Make"] = (
            make_target_df["Make"]
            .astype(str)
            .str.upper()
            .str.strip()
        )

        make_target_df["Target"] = pd.to_numeric(
            make_target_df["Target"],
            errors="coerce"
        ).fillna(0)

        brand_rows = []

        months_count = max(
            sales_df["Safe_YearMonth"].nunique(),
            1
        )

        for _, row in make_target_df.iterrows():

            make = row["Make"]

            make_target = (
                row["Target"] * months_count
            )

            make_sales = sales_df[
                sales_df["make"]
                .str.contains(
                    make,
                    na=False,
                    regex=False
                )
            ]["Value"].sum()

            make_pct = (
                make_sales / make_target * 100
                if make_target else 0
            )

            brand_rows.append({
                "Brand": make,
                "Sales": make_sales,
                "Target": make_target,
                "Achievement %": round(
                    make_pct,
                    1
                )
            })

        if brand_rows:

            st.dataframe(
                pd.DataFrame(brand_rows),
                use_container_width=True,
                hide_index=True
            )

    else:

        st.info(
            "MAKE TARGET sheet is optional and not available."
        )

    # ================= PDF =================
    if not monthly_report.empty:

        pdf_df = monthly_report[
            [
                "Month_Text",
                "Target",
                "Value",
                "Achievement %"
            ]
        ].copy()

        pdf_df.columns = [
            "Month",
            "Target",
            "Sales",
            "Achievement %"
        ]

        pdf = generate_employee_pdf(
            selected_marketing
            if selected_marketing != "ALL"
            else "All Marketing Employees",
            pdf_df
        )

        st.download_button(
            "📄 Download Overall Performance PDF",
            pdf,
            file_name=(
                selected_marketing
                if selected_marketing != "ALL"
                else "All_Marketing"
            )
            + "_Performance_Report.pdf",
            mime="application/pdf"
        )

    # ================= FILE INFO =================
    with st.expander("🔎 Excel File Information"):

        st.write(
            "Available sheets:",
            available_sheets
        )


# ================= MAIN =================
if "user" not in st.session_state:

    login()

else:

    dashboard()
