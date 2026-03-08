import json
from io import BytesIO
from typing import Optional

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="NMS Grand View Financial Dashboard", page_icon="🏨", layout="wide")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

EXPECTED_MONTHLY_COLS = ["Month", "Total Revenue", "Total Expenses", "Net Profit"]
EXPECTED_EXPENSE_COLS = ["Date", "Category", "Description", "Amount", "Payment Mode", "Remarks"]
EXPECTED_BOOKING_COLS = ["Date", "Booking Name", "Booking Engine", "Amount"]


def fmt_currency(value: float) -> str:
    return f"₹{value:,.0f}"


def clean_money(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.strip()
        .replace({"": None, "-": None, "nan": None})
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0)
    )


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_resource(show_spinner=False)
def get_gspread_client() -> gspread.Client:
    if "gcp_service_account" not in st.secrets:
        raise KeyError("Missing [gcp_service_account] in Streamlit secrets.")
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    return gspread.authorize(creds)


@st.cache_data(ttl=300, show_spinner=False)
def load_private_worksheet(sheet_key: str, worksheet_name: str) -> pd.DataFrame:
    client = get_gspread_client()
    ws = client.open_by_key(sheet_key).worksheet(worksheet_name)
    rows = ws.get_all_records()
    return normalize_columns(pd.DataFrame(rows))


@st.cache_data(ttl=300, show_spinner=False)
def load_public_csv(csv_url: str) -> pd.DataFrame:
    return normalize_columns(pd.read_csv(csv_url))


def load_sheet(config_key: str, default_tab_name: str) -> pd.DataFrame:
    data_source = st.secrets.get("data_source", "private_google_sheet")

    if data_source == "public_csv":
        csv_url = st.secrets.get(config_key)
        if not csv_url:
            raise KeyError(f"Missing secret: {config_key}")
        return load_public_csv(csv_url)

    sheet_key = st.secrets.get("google_sheet_key")
    worksheet_name = st.secrets.get(config_key, default_tab_name)
    if not sheet_key:
        raise KeyError("Missing secret: google_sheet_key")
    return load_private_worksheet(sheet_key, worksheet_name)


@st.cache_data(show_spinner=False)
def prep_monthly(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    rename_map = {}
    for col in df.columns:
        low = col.lower().strip()
        if low in {"month", "months"}:
            rename_map[col] = "Month"
        elif low in {"total revenue", "revenue", "income"}:
            rename_map[col] = "Total Revenue"
        elif low in {"total expenses", "expenses", "expense"}:
            rename_map[col] = "Total Expenses"
        elif low in {"net profit", "profit", "net"}:
            rename_map[col] = "Net Profit"
        elif low in {"% increase", "percent increase", "increase"}:
            rename_map[col] = "% Increase"
    df = df.rename(columns=rename_map)

    missing = [c for c in EXPECTED_MONTHLY_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            "Monthly sheet is missing required columns: " + ", ".join(missing)
        )

    for col in ["Total Revenue", "Total Expenses", "Net Profit"]:
        df[col] = clean_money(df[col])

    if "% Increase" not in df.columns:
        df["% Increase"] = (df["Net Profit"].pct_change() * 100).fillna(0)
    else:
        df["% Increase"] = pd.to_numeric(df["% Increase"], errors="coerce").fillna(0)

    month_order = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]
    df["Month"] = pd.Categorical(df["Month"], categories=month_order, ordered=True)
    return df.sort_values("Month").reset_index(drop=True)


@st.cache_data(show_spinner=False)
def prep_expenses(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    rename_map = {}
    for col in df.columns:
        low = col.lower().strip()
        if low in {"date", "expense date"}:
            rename_map[col] = "Date"
        elif low == "category":
            rename_map[col] = "Category"
        elif low == "description":
            rename_map[col] = "Description"
        elif low == "amount":
            rename_map[col] = "Amount"
        elif low in {"payment mode", "payment method"}:
            rename_map[col] = "Payment Mode"
        elif low in {"remarks", "remark", "notes"}:
            rename_map[col] = "Remarks"
    df = df.rename(columns=rename_map)

    required = ["Date", "Category", "Description", "Amount"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "Expenses sheet is missing required columns: " + ", ".join(missing)
        )

    df["Amount"] = clean_money(df["Amount"])
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    df["Month"] = df["Date"].dt.month_name()
    return df.sort_values("Date", ascending=False).reset_index(drop=True)


@st.cache_data(show_spinner=False)
def prep_bookings(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    rename_map = {}
    for col in df.columns:
        low = col.lower().strip()
        if low in {"date", "booking date"}:
            rename_map[col] = "Date"
        elif low in {"booking name", "name", "guest name"}:
            rename_map[col] = "Booking Name"
        elif low in {"booking engine", "engine", "channel", "source"}:
            rename_map[col] = "Booking Engine"
        elif low == "amount":
            rename_map[col] = "Amount"
    df = df.rename(columns=rename_map)

    missing = [c for c in EXPECTED_BOOKING_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            "Bookings sheet is missing required columns: " + ", ".join(missing)
        )

    df["Amount"] = clean_money(df["Amount"])
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    df["Month"] = df["Date"].dt.month_name()
    df["Month_num"] = df["Date"].dt.month
    df["Day"] = df["Date"].dt.day
    return df.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)


def dataframe_download_bytes(df: pd.DataFrame, name: str) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=name[:31])
    return out.getvalue()


def render_overview(monthly_df: pd.DataFrame, expenses_df: pd.DataFrame) -> None:
    latest = monthly_df.iloc[-1]
    previous = monthly_df.iloc[-2] if len(monthly_df) > 1 else None

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Revenue", fmt_currency(latest["Total Revenue"]))
    c2.metric("Expenses", fmt_currency(latest["Total Expenses"]))
    c3.metric("Net Profit", fmt_currency(latest["Net Profit"]))
    margin = (latest["Net Profit"] / latest["Total Revenue"] * 100) if latest["Total Revenue"] else 0
    c4.metric("Profit Margin", f"{margin:.1f}%")

    if previous is not None:
        st.caption(
            f"Month-over-month profit change: **{latest['% Increase']:.2f}%** compared with {previous['Month']}."
        )

    left, right = st.columns((1.3, 1))
    with left:
        fig_rev_exp = px.bar(
            monthly_df,
            x="Month",
            y=["Total Revenue", "Total Expenses"],
            barmode="group",
            title="Revenue vs Expenses",
            text_auto=True,
        )
        fig_rev_exp.update_layout(legend_title_text="")
        st.plotly_chart(fig_rev_exp, use_container_width=True)

    with right:
        fig_profit = px.bar(
            monthly_df,
            x="Month",
            y="Net Profit",
            title="Net Profit by Month",
            text_auto=True,
        )
        st.plotly_chart(fig_profit, use_container_width=True)

    cat_totals = (
        expenses_df.groupby("Category", dropna=False, as_index=False)["Amount"]
        .sum()
        .sort_values("Amount", ascending=False)
    )

    left, right = st.columns((1, 1))
    with left:
        pie = px.pie(
            cat_totals,
            names="Category",
            values="Amount",
            title="Expense Mix by Category",
            hole=0.35,
        )
        st.plotly_chart(pie, use_container_width=True)

    with right:
        bar = px.bar(
            cat_totals,
            x="Category",
            y="Amount",
            title="Total Expenditure by Category",
            text_auto=True,
        )
        bar.update_xaxes(tickangle=-25)
        st.plotly_chart(bar, use_container_width=True)


def render_expense_analysis(expenses_df: pd.DataFrame) -> None:
    st.subheader("Expense Analysis")

    categories = sorted([c for c in expenses_df["Category"].dropna().unique().tolist()])
    months = sorted([m for m in expenses_df["Month"].dropna().unique().tolist()])

    c1, c2 = st.columns(2)
    selected_categories = c1.multiselect("Category", categories, default=categories)
    selected_months = c2.multiselect("Month", months, default=months)

    filtered = expenses_df.copy()
    if selected_categories:
        filtered = filtered[filtered["Category"].isin(selected_categories)]
    if selected_months:
        filtered = filtered[filtered["Month"].isin(selected_months)]

    spend_total = filtered["Amount"].sum()
    txn_count = len(filtered)
    avg_txn = filtered["Amount"].mean() if txn_count else 0
    k1, k2, k3 = st.columns(3)
    k1.metric("Filtered Spend", fmt_currency(spend_total))
    k2.metric("Transactions", f"{txn_count:,}")
    k3.metric("Avg Transaction", fmt_currency(avg_txn))

    daily = (
        filtered.dropna(subset=["Date"])
        .groupby("Date", as_index=False)["Amount"]
        .sum()
        .sort_values("Date")
    )
    if not daily.empty:
        line = px.line(daily, x="Date", y="Amount", title="Daily Expense Trend", markers=True)
        st.plotly_chart(line, use_container_width=True)

    cat_table = (
        filtered.groupby("Category", dropna=False, as_index=False)["Amount"]
        .sum()
        .sort_values("Amount", ascending=False)
        .rename(columns={"Amount": "Total Spent"})
    )
    st.dataframe(cat_table, use_container_width=True, hide_index=True)

    st.download_button(
        "Download filtered expenses (.xlsx)",
        data=dataframe_download_bytes(filtered, "FilteredExpenses"),
        file_name="filtered_expenses.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.dataframe(filtered, use_container_width=True, hide_index=True)


def render_booking_engine(bookings_df: pd.DataFrame) -> None:
    st.subheader("Bookings by Engine")

    months_available = (
        bookings_df.dropna(subset=["Month"])
        .drop_duplicates(subset=["Month", "Month_num"])
        .sort_values("Month_num")["Month"]
        .tolist()
    )

    if not months_available:
        st.info("No booking data available.")
        return

    selected_month = st.selectbox("Select Month", ["All Months"] + months_available, key="engine_month")

    if selected_month == "All Months":
        filtered = bookings_df
        chart_title = "Bookings by Engine – All Months"
    else:
        filtered = bookings_df[bookings_df["Month"] == selected_month]
        chart_title = f"Bookings by Engine – {selected_month}"

    if filtered.empty:
        st.info("No bookings for the selected period.")
        return

    engine_totals = (
        filtered.groupby("Booking Engine", dropna=False, as_index=False)["Amount"]
        .sum()
        .sort_values("Amount", ascending=False)
    )

    left, right = st.columns((1, 1))
    with left:
        pie = px.pie(
            engine_totals,
            names="Booking Engine",
            values="Amount",
            title=chart_title,
            hole=0.35,
        )
        pie.update_traces(textinfo="percent+label")
        st.plotly_chart(pie, use_container_width=True)

    with right:
        bar = px.bar(
            engine_totals,
            x="Booking Engine",
            y="Amount",
            title="Revenue by Booking Engine",
            text_auto=True,
        )
        bar.update_xaxes(tickangle=-25)
        st.plotly_chart(bar, use_container_width=True)

    st.dataframe(
        engine_totals.rename(columns={"Amount": "Total Revenue (₹)"}),
        use_container_width=True,
        hide_index=True,
    )


def render_daily_revenue(bookings_df: pd.DataFrame) -> None:
    st.subheader("Daily Booking Revenue")

    months_available = (
        bookings_df.dropna(subset=["Month"])
        .drop_duplicates(subset=["Month", "Month_num"])
        .sort_values("Month_num")["Month"]
        .tolist()
    )

    if not months_available:
        st.info("No booking data available.")
        return

    selected_month = st.selectbox("Select Month", months_available, key="daily_rev_month")

    filtered = bookings_df[bookings_df["Month"] == selected_month]

    if filtered.empty:
        st.info(f"No bookings found for {selected_month}.")
        return

    daily = (
        filtered.groupby("Day", as_index=False)["Amount"]
        .sum()
        .sort_values("Day")
    )

    total = daily["Amount"].sum()
    peak_row = daily.loc[daily["Amount"].idxmax()]
    k1, k2 = st.columns(2)
    k1.metric("Total Revenue", fmt_currency(total))
    k2.metric("Peak Day", f"Day {int(peak_row['Day'])} – {fmt_currency(peak_row['Amount'])}")

    fig = px.bar(
        daily,
        x="Day",
        y="Amount",
        title=f"Revenue per Day – {selected_month}",
        text_auto=True,
        labels={"Day": "Day of Month", "Amount": "Total Revenue (₹)"},
    )
    fig.update_xaxes(dtick=1)
    st.plotly_chart(fig, use_container_width=True)


def render_raw_data(monthly_df: pd.DataFrame, expenses_df: pd.DataFrame, bookings_df: pd.DataFrame) -> None:
    st.subheader("Monthly Summary")
    st.dataframe(monthly_df, use_container_width=True, hide_index=True)
    st.download_button(
        "Download monthly summary (.xlsx)",
        data=dataframe_download_bytes(monthly_df, "MonthlySummary"),
        file_name="monthly_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Expenses")
    st.dataframe(expenses_df, use_container_width=True, hide_index=True)

    st.subheader("Bookings")
    st.dataframe(bookings_df, use_container_width=True, hide_index=True)





def main() -> None:
    st.title("🏨 NMS Grand View Financial Dashboard")
    st.caption("Live financial dashboard connected to Google Sheets.")

    with st.sidebar:
        st.header("Dashboard Settings")
        refresh = st.button("Refresh data")
        if refresh:
            st.cache_data.clear()
            st.cache_resource.clear()
            st.success("Cache cleared. Data will reload.")


    try:
        monthly_raw = load_sheet("monthly_sheet_name", "MonthlySummary")
        expenses_raw = load_sheet("expenses_sheet_name", "Expenses")
        bookings_raw = load_sheet("bookings_sheet_name", "BookingSummary")
        monthly_df = prep_monthly(monthly_raw)
        expenses_df = prep_expenses(expenses_raw)
        bookings_df = prep_bookings(bookings_raw)
    except Exception as exc:
        st.error(f"Could not load dashboard data: {exc}")
        st.stop()

    overview_tab, expenses_tab, engine_tab, daily_rev_tab, raw_tab = st.tabs(
        ["Overview", "Expense Analysis", "Bookings by Engine", "Daily Revenue", "Raw Data"]
    )

    with overview_tab:
        render_overview(monthly_df, expenses_df)

    with expenses_tab:
        render_expense_analysis(expenses_df)

    with engine_tab:
        render_booking_engine(bookings_df)

    with daily_rev_tab:
        render_daily_revenue(bookings_df)

    with raw_tab:
        render_raw_data(monthly_df, expenses_df, bookings_df)


if __name__ == "__main__":
    main()
