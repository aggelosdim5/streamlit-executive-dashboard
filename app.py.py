import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ----------------------------
# CONFIG
# ----------------------------
DATA_FILE = "Data Set_11.xlsx"

st.set_page_config(
    page_title="Executive Dashboard (Data Set_11)",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------
# DATA LOADING + PREPROCESSING
# ----------------------------
@st.cache_data
def load_and_prepare(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath)
    df.columns = [c.strip() for c in df.columns]

    if "InvoiceDate" in df.columns:
        df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"], errors="coerce")

    for c in ["Quantity", "UnitPrice", "CustomerID"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df["Order ID"] = df.get("InvoiceNo")
    df["Product Name"] = df.get("Description")
    df["Customer ID"] = df.get("CustomerID")
    df["Region"] = df.get("Country")
    df["Segment"] = df.get("Type")
    df["Category"] = df.get("Type")
    df["Sub-Category"] = df.get("StockCode")
    df["Order Date"] = df.get("InvoiceDate")

    df["Sales"] = df["Quantity"].fillna(0) * df["UnitPrice"].fillna(0)

    margin_map = {"Retail": 0.22, "B2B": 0.14}
    df["Profit"] = df["Sales"] * df["Segment"].map(margin_map).fillna(0.18)

    df["Discount"] = 0.0
    df["Year"] = df["Order Date"].dt.year
    df["Month"] = df["Order Date"].dt.month
    df["Quarter"] = df["Order Date"].dt.quarter
    df["Year-Month"] = df["Order Date"].dt.to_period("M").astype(str)

    df["Profit Margin"] = df.apply(
        lambda r: (r["Profit"] / r["Sales"]) * 100 if r["Sales"] != 0 else 0,
        axis=1
    )

    df = df.dropna(subset=["Order Date"])
    df["Customer ID"] = df["Customer ID"].fillna(-1).astype(int)

    return df


def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    out = df.copy()

    if filters["year"] != "All":
        out = out[out["Year"] == filters["year"]]

    out = out[
        out["Category"].isin(filters["category"]) &
        out["Region"].isin(filters["region"]) &
        out["Segment"].isin(filters["segment"])
    ]

    start, end = filters["date_range"]
    out = out[
        (out["Order Date"] >= pd.to_datetime(start)) &
        (out["Order Date"] <= pd.to_datetime(end))
    ]

    return out


def calculate_kpis(df: pd.DataFrame) -> dict:
    sales = df["Sales"].sum()
    profit = df["Profit"].sum()

    return {
        "total_sales": sales,
        "total_profit": profit,
        "total_orders": df["Order ID"].nunique(),
        "total_customers": df["Customer ID"].nunique(),
        "avg_order_value": df["Sales"].mean(),
        "profit_margin": (profit / sales * 100) if sales != 0 else 0
    }


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buffer.getvalue()


# ----------------------------
# LOAD DATA
# ----------------------------
df = load_and_prepare(DATA_FILE)

# ----------------------------
# SIDEBAR
# ----------------------------
st.sidebar.title("Navigation")
page = st.sidebar.radio(
    "Page",
    ["Overview", "Details", "Insights", "What-If"]
)

st.sidebar.markdown("---")
st.sidebar.header("Filters")

years = ["All"] + sorted(df["Year"].dropna().unique().tolist(), reverse=True)
filters = {
    "year": st.sidebar.selectbox("Year", years),
    "category": st.sidebar.multiselect(
        "Category",
        sorted(df["Category"].unique()),
        default=sorted(df["Category"].unique())
    ),
    "region": st.sidebar.multiselect(
        "Region",
        sorted(df["Region"].unique()),
        default=sorted(df["Region"].unique())
    ),
    "segment": st.sidebar.multiselect(
        "Segment",
        sorted(df["Segment"].unique()),
        default=sorted(df["Segment"].unique())
    ),
    "date_range": st.sidebar.date_input(
        "Date Range",
        (df["Order Date"].min().date(), df["Order Date"].max().date())
    )
}

filtered_df = apply_filters(df, filters)

# ----------------------------
# HEADER + KPIs
# ----------------------------
st.title("Executive Dashboard")
st.caption(f"Rows displayed: {len(filtered_df):,}")
st.markdown("---")

kpis = calculate_kpis(filtered_df)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Sales", f"${kpis['total_sales']:,.0f}")
c2.metric("Total Profit", f"${kpis['total_profit']:,.0f}")
c3.metric("Orders", f"{kpis['total_orders']:,}")
c4.metric("Customers", f"{kpis['total_customers']:,}")
c5.metric("Profit Margin", f"{kpis['profit_margin']:.1f}%")

st.markdown("---")

# ----------------------------
# OVERVIEW
# ----------------------------
if page == "Overview":
    col1, col2 = st.columns(2)

    with col1:
        monthly = filtered_df.groupby("Year-Month")["Sales"].sum().reset_index()
        fig = px.line(monthly, x="Year-Month", y="Sales", markers=True)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        mix = filtered_df.groupby("Segment")["Sales"].sum().reset_index()
        fig = px.pie(mix, values="Sales", names="Segment", hole=0.4)
        st.plotly_chart(fig, use_container_width=True)

    pivot = filtered_df.pivot_table(
        values="Sales",
        index="Segment",
        columns="Region",
        aggfunc="sum"
    )
    fig = px.imshow(pivot, text_auto=".2s")
    st.plotly_chart(fig, use_container_width=True)

# ----------------------------
# DETAILS
# ----------------------------
elif page == "Details":
    groupby = st.selectbox(
        "Group by",
        ["Product Name", "Sub-Category", "Region", "Segment", "Customer ID"]
    )

    grouped = (
        filtered_df.groupby(groupby)
        .agg(
            Sales=("Sales", "sum"),
            Profit=("Profit", "sum"),
            Quantity=("Quantity", "sum"),
            Orders=("Order ID", "count")
        )
        .sort_values("Sales", ascending=False)
    )

    grouped["Profit Margin %"] = grouped["Profit"] / grouped["Sales"] * 100
    grouped = grouped.fillna(0)

    st.dataframe(
        grouped.style.format({
            "Sales": "${:,.0f}",
            "Profit": "${:,.0f}",
            "Quantity": "{:,.0f}",
            "Orders": "{:,.0f}",
            "Profit Margin %": "{:.1f}%"
        }),
        use_container_width=True
    )

    st.download_button(
        "Download Excel",
        to_excel_bytes(grouped.reset_index()),
        "analysis.xlsx"
    )

# ----------------------------
# INSIGHTS
# ----------------------------
elif page == "Insights":
    fig = px.scatter(
        filtered_df,
        x="UnitPrice",
        y="Quantity",
        size="Sales",
        color="Segment",
        hover_data=["Product Name", "Region"]
    )
    st.plotly_chart(fig, use_container_width=True)

    corr = filtered_df[
        ["Sales", "Profit", "Quantity", "UnitPrice", "Profit Margin"]
    ].corr(numeric_only=True)

    fig = px.imshow(corr, text_auto=".2f")
    st.plotly_chart(fig, use_container_width=True)

# ----------------------------
# WHAT-IF
# ----------------------------
elif page == "What-If":
    current_sales = filtered_df["Sales"].sum()
    current_profit = filtered_df["Profit"].sum()

    price_change = st.slider("Price change (%)", -30, 30, 0)
    elasticity = -1.5

    factor = (1 + price_change / 100) * (1 + elasticity * price_change / 100)
    est_sales = current_sales * max(factor, 0)
    est_profit = current_profit * max(factor, 0)

    c1, c2, c3 = st.columns(3)
    c1.metric("Current Sales", f"${current_sales:,.0f}")
    c2.metric("Estimated Sales", f"${est_sales:,.0f}")
    c3.metric("Estimated Profit", f"${est_profit:,.0f}")

# ----------------------------
# RAW DATA
# ----------------------------
with st.expander("View filtered data"):
    st.dataframe(filtered_df, use_container_width=True)

st.markdown("---")
st.caption("Built with Streamlit and Plotly")
