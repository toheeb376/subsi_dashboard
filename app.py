import streamlit as st
import pandas as pd
import plotly.express as px

# --------------------------------------------------
# Page Configuration
# --------------------------------------------------
st.set_page_config(
    page_title="Subsi E-commerce Fulfilment Dashboard",
    layout="wide"
)

# --------------------------------------------------
# Background Theme
# --------------------------------------------------
st.markdown(
    """
    <style>
    .stApp {
        background-color: rgb(67, 61, 122);
    }
    h1, h2, h3, h4, h5, h6, p, label {
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --------------------------------------------------
# Header with Logo
# --------------------------------------------------
col1, col2 = st.columns([1, 6])
with col1:
    st.image("subsi.jpg", width=90)
with col2:
    st.title("Subsi E-commerce Order Fulfilment Dashboard")
    st.caption("Operational performance and delivery insights")

# --------------------------------------------------
# Load & Clean Data
# --------------------------------------------------
@st.cache_data
def load_data():
    df = pd.read_excel("Subsi_Ecommerce_Weekly_Order_Fulfilment.xlsx")

    # Standardize column names
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )

    # Convert dates
    date_columns = [
        "order_date",
        "expected_delivery_date",
        "actual_delivery_date"
    ]

    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col])

    return df

df = load_data()

# --------------------------------------------------
# Sidebar Filters
# --------------------------------------------------
st.sidebar.header("Filters")

# Delivery Status Filter
status_options = df["delivery_status"].unique()
selected_status = st.sidebar.multiselect(
    "Delivery Status",
    options=status_options,
    default=status_options
)

df = df[df["delivery_status"].isin(selected_status)]

# Date Filter
start_date, end_date = st.sidebar.date_input(
    "Order Date Range",
    [df["order_date"].min(), df["order_date"].max()]
)

df = df[
    (df["order_date"] >= pd.to_datetime(start_date)) &
    (df["order_date"] <= pd.to_datetime(end_date))
]

# Location Filter
location_options = df["customer_location"].unique()
selected_locations = st.sidebar.multiselect(
    "Customer Location",
    options=location_options,
    default=location_options
)

df = df[df["customer_location"].isin(selected_locations)]

# --------------------------------------------------
# KPI Calculations
# --------------------------------------------------
total_orders = len(df)
delivered_orders = df["delivery_status"].str.lower().eq("delivered").sum()
pending_orders = df["delivery_status"].str.lower().eq("in transit").sum()
cancelled_orders = df["delivery_status"].str.lower().eq("cancelled").sum()
delivery_rate = (delivered_orders / total_orders * 100) if total_orders > 0 else 0

# --------------------------------------------------
# KPI Display
# --------------------------------------------------
k1, k2, k3, k4, k5 = st.columns(5)

k1.metric("Total Orders", total_orders)
k2.metric("Delivered Orders", delivered_orders)
k3.metric("In Transit Orders", pending_orders)
k4.metric("Cancelled Orders", cancelled_orders)
k5.metric("Delivery Success Rate", f"{delivery_rate:.1f}%")

st.divider()

# --------------------------------------------------
# Visualizations
# --------------------------------------------------
colA, colB = st.columns(2)

# Delivery Status Distribution
fig_status = px.bar(
    df,
    x="delivery_status",
    title="Delivery Status Distribution",
    color="delivery_status"
)
colA.plotly_chart(fig_status, use_container_width=True)

# Orders Over Time
orders_by_date = df.groupby("order_date").size().reset_index(name="orders")
fig_trend = px.line(
    orders_by_date,
    x="order_date",
    y="orders",
    title="Orders Over Time"
)
colB.plotly_chart(fig_trend, use_container_width=True)

# --------------------------------------------------
# Additional Insights
# --------------------------------------------------
colC, colD = st.columns(2)

# Payment Status Breakdown
fig_payment = px.pie(
    df,
    names="payment_status",
    title="Payment Status Breakdown",
    hole=0.4
)
colC.plotly_chart(fig_payment, use_container_width=True)

# Delivery Method Usage
fig_method = px.bar(
    df,
    x="delivery_method",
    title="Delivery Method Usage",
    color="delivery_method"
)
colD.plotly_chart(fig_method, use_container_width=True)

# --------------------------------------------------
# Raw Data View
# --------------------------------------------------
with st.expander("View Raw Dataset"):
    st.dataframe(df)

# --------------------------------------------------
# Business Insights
# --------------------------------------------------
st.subheader("Key Business Insights")

st.markdown(
    f"""
    • **{delivered_orders} orders** were successfully delivered  
    • **{pending_orders} orders** are still in transit and need monitoring  
    • **{cancelled_orders} orders** were cancelled and may indicate logistics issues  
    • Overall delivery success rate is **{delivery_rate:.1f}%**  
    """
)

