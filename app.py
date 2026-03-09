import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import warnings

warnings.filterwarnings("ignore")

# =============================================================================
# PAGE CONFIG — MUST be the very first Streamlit call
# =============================================================================
st.set_page_config(
    page_title="Subsi | Order Fulfillment Intelligence",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# BRAND PALETTE — chart color sequence
# =============================================================================
CHART_COLORS = [
    "#40397A", "#68599C", "#9380B0", "#C4BED7",
    "#4B4380", "#7D72A2", "#ABA4C6", "#E0DCEA",
    "#7168A8", "#B8B0D4",
]

# =============================================================================
# CSS INJECTION
# =============================================================================
def inject_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=Nunito+Sans:wght@400;600;700&display=swap');

        .stApp {
            background-color: rgb(240,238,244);
            font-family: 'Nunito Sans', sans-serif;
        }
        .main .block-container {
            padding-top: 1.4rem;
            padding-bottom: 2rem;
            max-width: 100%;
        }

        /* Sidebar */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, rgb(64,59,122) 0%, rgb(104,93,147) 100%);
            border-right: 2px solid rgb(171,164,198);
        }
        section[data-testid="stSidebar"] * {
            color: rgb(255,255,255) !important;
            font-family: 'Nunito Sans', sans-serif !important;
        }
        section[data-testid="stSidebar"] label {
            font-weight: 700 !important;
            font-size: 0.8rem !important;
            letter-spacing: 0.05em !important;
            text-transform: uppercase !important;
        }
        section[data-testid="stSidebar"] hr {
            border-color: rgba(255,255,255,0.25) !important;
        }

        /* KPI cards */
        [data-testid="stMetric"] {
            background: linear-gradient(135deg, rgb(196,190,215) 0%, rgb(224,220,234) 100%);
            border: 1px solid rgb(171,164,198);
            border-radius: 12px;
            padding: 14px 18px !important;
            box-shadow: 0 2px 8px rgba(64,59,122,0.12);
        }
        [data-testid="stMetric"] label {
            color: rgb(64,59,122) !important;
            font-weight: 700 !important;
            font-size: 0.76rem !important;
            text-transform: uppercase !important;
            letter-spacing: 0.06em !important;
        }
        [data-testid="stMetricValue"] {
            color: rgb(64,59,122) !important;
            font-family: 'Nunito', sans-serif !important;
            font-weight: 900 !important;
            font-size: 1.5rem !important;
        }

        h1,h2,h3,h4 {
            font-family: 'Nunito', sans-serif !important;
            color: rgb(64,59,122) !important;
        }
        h1 { font-weight: 900 !important; }
        h2 { font-weight: 800 !important; }
        h3 { font-weight: 700 !important; }

        hr { border-color: rgb(171,164,198) !important; margin: 0.8rem 0 !important; }

        [data-testid="stExpander"] {
            background-color: rgb(255,255,255);
            border: 1px solid rgb(171,164,198) !important;
            border-radius: 10px !important;
        }
        [data-testid="stExpander"] summary {
            color: rgb(64,59,122) !important;
            font-weight: 700 !important;
        }

        .section-card {
            background: rgb(255,255,255);
            border: 1px solid rgb(171,164,198);
            border-radius: 14px;
            padding: 18px;
            margin-bottom: 14px;
            box-shadow: 0 2px 6px rgba(64,59,122,0.07);
        }
        .section-title {
            font-family: 'Nunito', sans-serif;
            font-size: 1rem;
            font-weight: 800;
            color: rgb(64,59,122);
            border-left: 4px solid rgb(104,93,147);
            padding-left: 9px;
            margin-bottom: 4px;
        }
        .dash-header {
            background: linear-gradient(135deg, rgb(64,59,122) 0%, rgb(104,93,147) 60%, rgb(147,138,180) 100%);
            border-radius: 16px;
            padding: 22px 30px;
            margin-bottom: 20px;
            box-shadow: 0 4px 20px rgba(64,59,122,0.28);
        }
        .dash-header h1 { color: white !important; font-size: 1.8rem !important; margin: 0 !important; }
        .dash-header p  { color: rgb(224,220,234) !important; margin: 4px 0 0 0 !important; font-size: 0.92rem !important; }

        .kpi-label {
            font-family: 'Nunito', sans-serif;
            font-size: 0.7rem;
            font-weight: 800;
            color: rgb(125,114,162);
            text-transform: uppercase;
            letter-spacing: 0.1em;
            margin-bottom: 8px;
        }

        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: rgb(240,238,244); }
        ::-webkit-scrollbar-thumb { background: rgb(147,138,180); border-radius: 3px; }

        #MainMenu, footer, header { visibility: hidden; }
        .stDeployButton { display: none; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# DATA LOADING
# FIX: use `df[col].dtype == object` instead of select_dtypes("str")
# which breaks on pandas 2.x / 3.x
# =============================================================================
@st.cache_data(show_spinner=False)
def load_data():
    filepath = "subsi_ecommerce_dataset.xlsx"

    if not os.path.exists(filepath):
        return None, (
            "File not found: subsi_ecommerce_dataset.xlsx — "
            "place it in the same folder as app.py and restart."
        )

    try:
        df = pd.read_excel(filepath, engine="openpyxl")
    except Exception as exc:
        return None, f"Could not read Excel file: {exc}"

    # --- Strip whitespace from text columns (pandas-version-safe) ---
    # Checking dtype == object works on ALL pandas versions (1.x, 2.x, 3.x)
    # and avoids the removed "str" dtype alias in newer pandas.
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()

    # --- Parse dates ---
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")

    # --- Derived time columns ---
    df["Month_dt"] = df["Order_Date"].dt.to_period("M").dt.to_timestamp()
    df["Month"]    = df["Order_Date"].dt.to_period("M").astype(str)
    df["Year"]     = df["Order_Date"].dt.year

    # --- Numeric safety ---
    for col in ["Quantity", "Unit_Price_NGN", "Total_Value_NGN"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # --- Revenue contribution % ---
    total_rev = df["Total_Value_NGN"].sum()
    df["Rev_Pct"] = (df["Total_Value_NGN"] / total_rev * 100) if total_rev > 0 else 0.0

    # --- Ordinal date for 3-D chart ---
    df["Date_Ordinal"] = df["Order_Date"].apply(
        lambda x: x.toordinal() if pd.notnull(x) else None
    )

    return df, None


# =============================================================================
# PLOTLY THEME HELPER
# =============================================================================
def theme(fig, title="", height=380):
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(family="Nunito", size=14, color="rgb(64,59,122)"),
            x=0.01, xanchor="left",
        ),
        paper_bgcolor="rgb(255,255,255)",
        plot_bgcolor="rgb(240,238,244)",
        font=dict(family="Nunito Sans", color="rgb(64,59,122)", size=12),
        height=height,
        margin=dict(l=10, r=10, t=44, b=10),
        legend=dict(
            font=dict(size=11, color="rgb(64,59,122)"),
            bgcolor="rgba(255,255,255,0.85)",
            bordercolor="rgb(171,164,198)",
            borderwidth=1,
        ),
        xaxis=dict(
            gridcolor="rgb(224,220,234)",
            linecolor="rgb(171,164,198)",
            tickfont=dict(color="rgb(104,93,147)", size=11),
            title_font=dict(color="rgb(64,59,122)", size=12),
        ),
        yaxis=dict(
            gridcolor="rgb(224,220,234)",
            linecolor="rgb(171,164,198)",
            tickfont=dict(color="rgb(104,93,147)", size=11),
            title_font=dict(color="rgb(64,59,122)", size=12),
        ),
        hoverlabel=dict(
            bgcolor="rgb(64,59,122)",
            font_color="white",
            font_family="Nunito Sans",
            bordercolor="rgb(147,138,180)",
        ),
    )
    return fig


# =============================================================================
# SIDEBAR
# =============================================================================
def render_sidebar(df):
    with st.sidebar:
        logo = "subsi_ecommerce_dataset.png"
        if os.path.exists(logo):
            st.image(logo, use_container_width=True)
        else:
            st.markdown("## 🛒 subsi")

        st.markdown("---")
        st.markdown(
            "<div style='color:rgba(255,255,255,0.7);font-size:0.73rem;"
            "text-transform:uppercase;letter-spacing:0.1em;font-weight:700;"
            "margin-bottom:10px'>️ Dashboard Filters</div>",
            unsafe_allow_html=True,
        )

        def opts(col):
            return sorted(df[col].dropna().astype(str).unique().tolist())

        statuses   = opts("Order_Status")
        payments   = opts("Payment_Method")
        categories = opts("Product_Category")
        states     = opts("Delivery_State")

        sel_status   = st.multiselect("Order Status",     statuses,   default=statuses,   key="fs")
        sel_payment  = st.multiselect("Payment Method",   payments,   default=payments,   key="fp")
        sel_category = st.multiselect("Product Category", categories, default=categories, key="fc")
        sel_state    = st.multiselect("Delivery State",   states,     default=states,     key="fd")

        st.markdown("**Order Date Range**")
        min_d = df["Order_Date"].min().date()
        max_d = df["Order_Date"].max().date()
        date_start = st.date_input("From", value=min_d, min_value=min_d, max_value=max_d, key="ds")
        date_end   = st.date_input("To",   value=max_d, min_value=min_d, max_value=max_d, key="de")

        st.markdown("---")
        st.markdown(
            "<div style='color:rgba(255,255,255,0.5);font-size:0.68rem;"
            "text-align:center;line-height:1.7'>Subsi Intelligence Dashboard<br>"
            "© 2025 Subsi E-Commerce</div>",
            unsafe_allow_html=True,
        )

    return sel_status, sel_payment, sel_category, sel_state, date_start, date_end


# =============================================================================
# FILTER
# =============================================================================
def apply_filters(df, sel_status, sel_payment, sel_category, sel_state, date_start, date_end):
    fdf = df.copy()
    if sel_status:   fdf = fdf[fdf["Order_Status"].isin(sel_status)]
    if sel_payment:  fdf = fdf[fdf["Payment_Method"].isin(sel_payment)]
    if sel_category: fdf = fdf[fdf["Product_Category"].isin(sel_category)]
    if sel_state:    fdf = fdf[fdf["Delivery_State"].isin(sel_state)]
    fdf = fdf[fdf["Order_Date"].notna()]
    fdf = fdf[
        (fdf["Order_Date"].dt.date >= date_start) &
        (fdf["Order_Date"].dt.date <= date_end)
    ]
    return fdf.reset_index(drop=True)


# =============================================================================
# KPI CARDS
# =============================================================================
def render_kpis(fdf):
    n        = len(fdf)
    rev      = fdf["Total_Value_NGN"].sum()
    aov      = rev / n if n else 0
    deliv    = (fdf["Order_Status"] == "Delivered").sum()
    cancel   = (fdf["Order_Status"] == "Cancelled").sum()
    ful_pct  = deliv / n * 100 if n else 0
    can_pct  = cancel / n * 100 if n else 0
    qty      = int(fdf["Quantity"].sum())
    custs    = fdf["Customer_ID"].nunique()
    top_cat  = (
        fdf.groupby("Product_Category")["Total_Value_NGN"].sum().idxmax()
        if n > 0 else "—"
    )

    st.markdown("<div class='kpi-label'>Key Performance Indicators</div>", unsafe_allow_html=True)

    r1 = st.columns(4)
    r1[0].metric(" Total Orders",     f"{n:,}")
    r1[1].metric(" Total Revenue",    f"₦{rev:,.0f}")
    r1[2].metric(" Avg Order Value", f"₦{aov:,.0f}")
    r1[3].metric(" Top Category",     top_cat)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    r2 = st.columns(4)
    r2[0].metric(" Fulfilment Rate",   f"{ful_pct:.1f}%")
    r2[1].metric(" Cancellation Rate", f"{can_pct:.1f}%")
    r2[2].metric(" Total Qty Sold",    f"{qty:,}")
    r2[3].metric(" Unique Customers",  f"{custs:,}")


# =============================================================================
# CHARTS
# =============================================================================

def chart_status(fdf):
    d = fdf["Order_Status"].value_counts().reset_index()
    d.columns = ["Status", "Count"]
    fig = px.bar(d, x="Status", y="Count", color="Status",
                 color_discrete_sequence=CHART_COLORS, text="Count")
    fig.update_traces(textposition="outside", marker_line_width=0)
    fig = theme(fig, "Order Status Distribution")
    fig.update_layout(showlegend=False)
    return fig


def chart_rev_category(fdf):
    d = (fdf.groupby("Product_Category")["Total_Value_NGN"]
         .sum().reset_index().sort_values("Total_Value_NGN"))
    d.columns = ["Category", "Revenue"]
    fig = px.bar(d, y="Category", x="Revenue", orientation="h",
                 color="Revenue",
                 color_continuous_scale=["rgb(196,190,215)", "rgb(64,59,122)"],
                 text=d["Revenue"].apply(lambda x: f"₦{x/1e6:.1f}M"))
    fig.update_traces(textposition="outside", marker_line_width=0)
    fig.update_coloraxes(showscale=False)
    return theme(fig, "Revenue by Product Category")


def chart_monthly(fdf):
    d = (fdf.groupby("Month_dt")["Total_Value_NGN"]
         .sum().reset_index().sort_values("Month_dt"))
    d.columns = ["Month", "Revenue"]
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=d["Month"], y=d["Revenue"],
        mode="lines+markers",
        line=dict(color="rgb(64,59,122)", width=3),
        marker=dict(size=7, color="rgb(104,93,147)", line=dict(width=2, color="white")),
        fill="tozeroy", fillcolor="rgba(147,138,180,0.18)",
        hovertemplate="<b>%{x|%b %Y}</b><br>₦%{y:,.0f}<extra></extra>",
    ))
    fig = theme(fig, "Monthly Revenue Trend")
    fig.update_yaxes(tickprefix="₦", tickformat=",.0f")
    return fig


def chart_payment(fdf):
    d = fdf["Payment_Method"].value_counts().reset_index()
    d.columns = ["Method", "Count"]
    fig = px.pie(d, names="Method", values="Count", hole=0.55,
                 color_discrete_sequence=CHART_COLORS)
    fig.update_traces(textposition="outside", textfont_size=12, pull=[0.03]*len(d))
    fig = theme(fig, "Payment Method Breakdown", height=360)
    fig.update_layout(
        legend=dict(orientation="v", x=1.0, y=0.5),
        annotations=[dict(text="Payment<br>Split", x=0.5, y=0.5,
                          font_size=13, font_color="rgb(64,59,122)",
                          font_family="Nunito", showarrow=False)],
    )
    return fig


def chart_top10(fdf):
    d = (fdf.groupby("Product_Name")["Total_Value_NGN"]
         .sum().nlargest(10).reset_index()
         .sort_values("Total_Value_NGN"))
    d.columns = ["Product", "Revenue"]
    fig = px.bar(d, y="Product", x="Revenue", orientation="h",
                 color="Revenue",
                 color_continuous_scale=["rgb(196,190,215)", "rgb(64,59,122)"],
                 text=d["Revenue"].apply(lambda x: f"₦{x/1e6:.1f}M"))
    fig.update_traces(textposition="outside", marker_line_width=0)
    fig.update_coloraxes(showscale=False)
    return theme(fig, "Top 10 Products by Revenue", height=400)


def chart_by_state(fdf):
    d = (fdf.groupby("Delivery_State")
         .agg(Orders=("Order_ID","count"), Revenue=("Total_Value_NGN","sum"))
         .reset_index().sort_values("Revenue", ascending=False))
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Revenue (₦)", x=d["Delivery_State"], y=d["Revenue"],
        marker_color="rgb(64,59,122)", yaxis="y",
        hovertemplate="<b>%{x}</b><br>₦%{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        name="Orders", x=d["Delivery_State"], y=d["Orders"],
        mode="lines+markers",
        line=dict(color="rgb(147,138,180)", width=2, dash="dot"),
        marker=dict(size=8, color="rgb(104,93,147)"),
        yaxis="y2",
        hovertemplate="<b>%{x}</b><br>Orders: %{y}<extra></extra>",
    ))
    fig = theme(fig, "Revenue & Orders by Delivery State")
    fig.update_layout(
        yaxis=dict(title="Revenue (₦)", tickprefix="₦", tickformat=",.0f"),
        yaxis2=dict(title="Orders", overlaying="y", side="right", showgrid=False),
    )
    return fig


def chart_treemap(fdf):
    d = (fdf.groupby(["Product_Category","Product_Subcategory"])["Total_Value_NGN"]
         .sum().reset_index())
    d.columns = ["Category","Subcategory","Revenue"]
    fig = px.treemap(d, path=["Category","Subcategory"], values="Revenue",
                     color="Revenue",
                     color_continuous_scale=["rgb(224,220,234)","rgb(104,93,147)","rgb(64,59,122)"])
    fig.update_traces(
        textfont=dict(family="Nunito", size=12, color="white"),
        hovertemplate="<b>%{label}</b><br>₦%{value:,.0f}<extra></extra>",
    )
    fig.update_coloraxes(showscale=False)
    fig = theme(fig, "Revenue by Product Subcategory (Treemap)", height=420)
    fig.update_layout(margin=dict(l=10, r=10, t=44, b=10))
    return fig


def chart_qty_category(fdf):
    d = (fdf.groupby(["Product_Category","Order_Status"])["Quantity"]
         .sum().reset_index())
    d.columns = ["Category","Status","Quantity"]
    fig = px.bar(d, x="Category", y="Quantity", color="Status",
                 barmode="group", color_discrete_sequence=CHART_COLORS)
    fig.update_traces(marker_line_width=0)
    fig = theme(fig, "Quantity Sold by Category & Status")
    fig.update_xaxes(tickangle=-20)
    return fig


def chart_area(fdf):
    d = (fdf.groupby(["Order_Date","Order_Status"])["Order_ID"]
         .count().reset_index().sort_values("Order_Date"))
    d.columns = ["Date","Status","Orders"]
    fig = px.area(d, x="Date", y="Orders", color="Status",
                  color_discrete_sequence=CHART_COLORS, line_group="Status")
    return theme(fig, "Daily Order Volume by Status", height=350)


def chart_3d(fdf):
    d = fdf.dropna(subset=["Date_Ordinal","Quantity","Total_Value_NGN"]).copy()
    if d.empty:
        return None

    status_colors = {
        "Delivered":  "#40397A",
        "Shipped":    "#68599C",
        "Processing": "#9380B0",
        "Cancelled":  "#C0392B",
    }
    fig = go.Figure()
    for status, grp in d.groupby("Order_Status"):
        fig.add_trace(go.Scatter3d(
            x=grp["Date_Ordinal"],
            y=grp["Quantity"],
            z=grp["Total_Value_NGN"],
            mode="markers",
            name=str(status),
            marker=dict(
                size=5,
                color=status_colors.get(str(status), "#BDC3C7"),
                opacity=0.82,
                line=dict(width=0.5, color="white"),
            ),
            customdata=grp[["Order_ID","Product_Name","Delivery_State","Payment_Method"]].values,
            hovertemplate=(
                "<b>%{customdata[1]}</b><br>"
                "Order: %{customdata[0]}<br>"
                "State: %{customdata[2]}<br>"
                "Payment: %{customdata[3]}<br>"
                "Qty: %{y}<br>"
                "Revenue: ₦%{z:,.0f}<extra></extra>"
            ),
        ))

    fig.update_layout(
        title=dict(
            text="3D Intelligence Scatter — Date × Quantity × Revenue",
            font=dict(family="Nunito", size=14, color="rgb(64,59,122)"),
            x=0.01,
        ),
        paper_bgcolor="rgb(255,255,255)",
        height=520,
        margin=dict(l=0, r=0, t=50, b=0),
        scene=dict(
            bgcolor="rgb(240,238,244)",
            xaxis=dict(title="Order Date (Ordinal)",
                       backgroundcolor="rgb(224,220,234)",
                       gridcolor="rgb(196,190,215)",
                       color="rgb(64,59,122)", showbackground=True),
            yaxis=dict(title="Quantity",
                       backgroundcolor="rgb(224,220,234)",
                       gridcolor="rgb(196,190,215)",
                       color="rgb(64,59,122)", showbackground=True),
            zaxis=dict(title="Revenue (₦)",
                       backgroundcolor="rgb(224,220,234)",
                       gridcolor="rgb(196,190,215)",
                       color="rgb(64,59,122)", showbackground=True),
        ),
        legend=dict(
            font=dict(size=11, color="rgb(64,59,122)"),
            bgcolor="rgba(255,255,255,0.85)",
            bordercolor="rgb(171,164,198)", borderwidth=1,
        ),
        hoverlabel=dict(
            bgcolor="rgb(64,59,122)", font_color="white",
            font_family="Nunito Sans",
        ),
    )
    return fig


# =============================================================================
# EXECUTIVE INSIGHTS
# =============================================================================
def render_insights(fdf):
    n         = len(fdf)
    rev       = fdf["Total_Value_NGN"].sum()
    del_pct   = (fdf["Order_Status"] == "Delivered").sum() / n * 100 if n else 0
    can_pct   = (fdf["Order_Status"] == "Cancelled").sum() / n * 100 if n else 0
    top_state = fdf.groupby("Delivery_State")["Total_Value_NGN"].sum().idxmax() if n else "—"
    top_pay   = fdf["Payment_Method"].value_counts().idxmax() if n else "—"
    top_cat   = fdf.groupby("Product_Category")["Total_Value_NGN"].sum().idxmax() if n else "—"

    with st.expander(" Executive Insight Summary — Click to expand", expanded=False):
        st.markdown(
            f"""
            <div style='font-family:Nunito Sans,sans-serif;color:rgb(64,59,122);line-height:1.85'>
            <h4 style='color:rgb(64,59,122);font-family:Nunito,sans-serif;margin-top:0'>
                 Operational Intelligence Overview
            </h4>
            <p>Filtered view: <strong>{n:,} orders</strong> &nbsp;·&nbsp;
            Total Revenue: <strong>&#x20A6;{rev:,.0f}</strong></p>
            <hr style='border-color:rgb(171,164,198)'>
            <p> <strong>Fulfilment:</strong> {del_pct:.1f}% delivered.
            {'Strong performance.' if del_pct >= 70 else 'Improvement opportunity.'}
            Cancellation rate {can_pct:.1f}%
            {'— within acceptable range.' if can_pct <= 15 else '— requires operational review.'}</p>
            <p> <strong>Category leader:</strong> <em>{top_cat}</em> —
            concentrate inventory investment and marketing here.</p>
            <p> <strong>Top state:</strong> <em>{top_state}</em> —
            prioritise logistics and warehousing resources here.</p>
            <p> <strong>Dominant payment channel:</strong> <em>{top_pay}</em> —
            expand payment alternatives to reduce conversion friction.</p>
            <p> <strong>Seasonality:</strong> Monitor the Monthly Revenue Trend chart
            to time stock procurement around demand peaks.</p>
            <p> <strong>Action:</strong> Operations teams should review the state
            fulfilment chart daily. Finance teams can use the Treemap and grouped
            bar charts to assess category and geographic revenue concentration risk.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )


# =============================================================================
# CHART SECTION HELPER
# =============================================================================
def section_chart(title, fig):
    st.markdown(f"<div class='section-title'>{title}</div>", unsafe_allow_html=True)
    if fig is not None:
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    else:
        st.info("Not enough data for this chart with current filters.")


# =============================================================================
# MAIN
# =============================================================================
def main():
    inject_css()

    # Load data
    with st.spinner("Loading Subsi dataset…"):
        df, err = load_data()

    if err:
        st.error(err)
        st.info(
            "**Quick-start guide**\n\n"
            "1. `pip install streamlit pandas plotly openpyxl`\n"
            "2. Place `subsi_ecommerce_dataset.xlsx`, "
            "`subsi_ecommerce_dataset.png`, and `app.py` in the **same folder**\n"
            "3. `streamlit run app.py`"
        )
        return

    # Sidebar
    sel_status, sel_payment, sel_category, sel_state, date_start, date_end = (
        render_sidebar(df)
    )

    # Filter
    fdf = apply_filters(df, sel_status, sel_payment, sel_category,
                        sel_state, date_start, date_end)

    # Header banner
    st.markdown(
        f"""
        <div class='dash-header'>
            <h1>🛒 Subsi Order Fulfillment Intelligence</h1>
            <p>Showing <strong>{len(fdf):,}</strong> of <strong>{len(df):,}</strong> orders
            &nbsp;·&nbsp; {date_start.strftime('%d %b %Y')} &#x2192; {date_end.strftime('%d %b %Y')}
            &nbsp;·&nbsp; All filters active</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if len(fdf) == 0:
        st.warning("⚠️ No records match the current filters. Please adjust the sidebar.")
        return

    # KPI section
    render_kpis(fdf)
    st.markdown("---")

    # Row 1 — Status | Monthly Revenue
    c1, c2 = st.columns([1, 1.6])
    with c1:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Order Status Distribution", chart_status(fdf))
        st.markdown("</div>", unsafe_allow_html=True)
    with c2:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Monthly Revenue Trend", chart_monthly(fdf))
        st.markdown("</div>", unsafe_allow_html=True)

    # Row 2 — Revenue by Category | Payment Method
    c3, c4 = st.columns([1.6, 1])
    with c3:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Revenue by Product Category", chart_rev_category(fdf))
        st.markdown("</div>", unsafe_allow_html=True)
    with c4:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Payment Method Breakdown", chart_payment(fdf))
        st.markdown("</div>", unsafe_allow_html=True)

    # Row 3 — Top 10 Products | By State
    c5, c6 = st.columns(2)
    with c5:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Top 10 Products by Revenue", chart_top10(fdf))
        st.markdown("</div>", unsafe_allow_html=True)
    with c6:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Revenue & Orders by Delivery State", chart_by_state(fdf))
        st.markdown("</div>", unsafe_allow_html=True)

    # Row 4 — Treemap | Qty by Category
    c7, c8 = st.columns([1.2, 1])
    with c7:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Revenue by Product Subcategory", chart_treemap(fdf))
        st.markdown("</div>", unsafe_allow_html=True)
    with c8:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Quantity Sold by Category & Status", chart_qty_category(fdf))
        st.markdown("</div>", unsafe_allow_html=True)

    # Row 5 — Area chart (full width)
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    section_chart("Daily Order Volume by Status", chart_area(fdf))
    st.markdown("</div>", unsafe_allow_html=True)

    # Row 6 — 3-D Scatter (full width)
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    section_chart(
        "3D Intelligence Scatter — Date × Quantity × Revenue",
        chart_3d(fdf),
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    # Executive insights
    render_insights(fdf)

    # Footer
    st.markdown(
        "<div style='text-align:center;padding:18px 0 8px;"
        "color:rgb(125,114,162);font-size:0.73rem;"
        "font-family:Nunito Sans,sans-serif'>"
        "Subsi Order Fulfillment Intelligence &nbsp;·&nbsp; "
        "Streamlit + Plotly &nbsp;·&nbsp; © 2025 Subsi E-Commerce"
        "</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
