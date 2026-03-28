import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import base64
import warnings

warnings.filterwarnings("ignore")

# =============================================================================
# PAGE CONFIG — must be the very first Streamlit call
# =============================================================================
st.set_page_config(
    page_title="Subsi | Order Fulfillment Intelligence",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# BRAND PALETTE
# Primary   : RGB(64,  59, 122) → #403B7A  deep purple
# White     : RGB(255,255,255) → #FFFFFF
# Dark BG   : #1C1A36           deep navy
# Mid BG    : #2C2860           card surface
# Light BG  : #3A3580           hover / accent
# Muted     : #C8C5E8           labels / gridlines
# Border    : rgba(100,95,170,.30)
# =============================================================================
CHART_COLORS = [
    "#403B7A", "#6B64A8", "#9590C8", "#C8C5E8",
    "#504A99", "#7B75BB", "#ABA4D4", "#DDDAEA",
    "#3D3870", "#B0ACDD",
]

# =============================================================================
# LOGO — base64 embed (works on Streamlit Cloud, no HTTP needed)
# =============================================================================
def _logo_b64(path: str) -> str:
    if os.path.exists(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

_LOGO_B64 = _logo_b64("subsi.png")
LOGO_SRC  = f"data:image/png;base64,{_LOGO_B64}" if _LOGO_B64 else ""


# =============================================================================
# CSS INJECTION
# =============================================================================
def inject_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=Nunito+Sans:wght@400;600;700&display=swap');

        /* ── Root variables ─────────────────────────────────────────── */
        :root {
            --primary      : #403B7A;
            --primary-light: #6B64A8;
            --primary-dark : #1C1A36;
            --mid-bg       : #2C2860;
            --accent       : #9590C8;
            --white        : #FFFFFF;
            --muted        : #C8C5E8;
            --border       : rgba(100, 95, 170, 0.30);
            --shadow       : 0 4px 24px rgba(64, 59, 122, 0.35);
            --shadow-hover  : 0 8px 36px rgba(64, 59, 122, 0.55);
            --radius       : 14px;
            --radius-sm    : 8px;
        }

        /* ── App background ─────────────────────────────────────────── */
        html, body, [class*="css"] {
            font-family: 'Nunito Sans', sans-serif !important;
        }
        .stApp {
            background: linear-gradient(160deg, #1C1A36 0%, #252060 55%, #1C1A36 100%) !important;
            background-attachment: fixed !important;
        }

        /* ── Main container ─────────────────────────────────────────── */
        .main .block-container {
            padding-top    : 2.5rem  !important;
            padding-bottom : 2rem    !important;
            padding-left   : 2rem    !important;
            padding-right  : 2rem    !important;
            max-width      : 100%    !important;
        }

        /* ── Sidebar ────────────────────────────────────────────────── */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #1C1A36 0%, #2C2860 100%) !important;
            border-right: 1px solid var(--border) !important;
        }
        section[data-testid="stSidebar"] * {
            color: var(--white) !important;
            font-family: 'Nunito Sans', sans-serif !important;
        }
        section[data-testid="stSidebar"] label {
            font-weight    : 700 !important;
            font-size      : 0.78rem !important;
            letter-spacing : 0.08em !important;
            text-transform : uppercase !important;
            color          : var(--muted) !important;
        }
        section[data-testid="stSidebar"] hr {
            border-color: var(--border) !important;
        }

        /* ── Sidebar multiselect tags ───────────────────────────────── */
        .stMultiSelect [data-baseweb="tag"] {
            background    : var(--primary) !important;
            border-radius : 20px !important;
            color         : var(--white) !important;
            font-weight   : 600 !important;
            font-size     : 0.73rem !important;
        }
        [data-testid="stMultiSelect"] > div,
        [data-testid="stSelectbox"] > div {
            background    : rgba(44,40,96,0.7) !important;
            border        : 1px solid var(--border) !important;
            border-radius : var(--radius-sm) !important;
            color         : var(--white) !important;
        }

        /* ── Date input ─────────────────────────────────────────────── */
        [data-testid="stDateInput"] input {
            background    : rgba(44,40,96,0.7) !important;
            border        : 1px solid var(--border) !important;
            color         : var(--white) !important;
            border-radius : var(--radius-sm) !important;
        }

        /* ── KPI / metric cards ─────────────────────────────────────── */
        [data-testid="stMetric"] {
            background    : linear-gradient(135deg, #2C2860 0%, #1C1A36 100%) !important;
            border        : 1px solid var(--border) !important;
            border-radius : var(--radius) !important;
            padding       : 14px 18px !important;
            box-shadow    : var(--shadow) !important;
            position      : relative;
            overflow      : hidden;
            transition    : transform .2s ease, box-shadow .2s ease;
        }
        [data-testid="stMetric"]::before {
            content    : '';
            position   : absolute;
            top: 0; left: 0; right: 0;
            height     : 3px;
            background : linear-gradient(90deg, var(--primary), var(--accent), var(--primary));
        }
        [data-testid="stMetric"]:hover {
            transform  : translateY(-2px);
            box-shadow : var(--shadow-hover) !important;
        }
        [data-testid="stMetricLabel"] {
            color          : var(--muted) !important;
            font-weight    : 700 !important;
            font-size      : 0.72rem !important;
            text-transform : uppercase !important;
            letter-spacing : 0.08em !important;
        }
        [data-testid="stMetricValue"] {
            color       : var(--white) !important;
            font-family : 'Nunito', sans-serif !important;
            font-weight : 900 !important;
            font-size   : 1.5rem !important;
        }
        [data-testid="stMetricDelta"] { font-size: 0.75rem !important; font-weight: 700 !important; }

        /* ── Headings ───────────────────────────────────────────────── */
        h1, h2, h3, h4 {
            font-family : 'Nunito', sans-serif !important;
            color       : var(--white) !important;
        }
        h1 { font-weight: 900 !important; }
        h2 { font-weight: 800 !important; }
        h3 { font-weight: 700 !important; }

        /* ── Dividers ───────────────────────────────────────────────── */
        hr { border: none !important; border-top: 1px solid var(--border) !important; margin: 1rem 0 !important; }

        /* ── Expander ───────────────────────────────────────────────── */
        [data-testid="stExpander"] {
            background    : linear-gradient(135deg, #2C2860, #1C1A36) !important;
            border        : 1px solid var(--border) !important;
            border-radius : var(--radius) !important;
            box-shadow    : var(--shadow) !important;
        }
        [data-testid="stExpander"] summary {
            color       : var(--white) !important;
            font-weight : 800 !important;
            font-family : 'Nunito', sans-serif !important;
            font-size   : 0.9rem !important;
        }
        [data-testid="stExpander"] p, [data-testid="stExpander"] li {
            color       : var(--muted) !important;
            font-size   : 0.85rem !important;
            line-height : 1.7 !important;
        }

        /* ── Section card wrapper ───────────────────────────────────── */
        .section-card {
            background    : linear-gradient(135deg, #2C2860 0%, #1C1A36 100%);
            border        : 1px solid var(--border);
            border-radius : var(--radius);
            padding       : 18px;
            margin-bottom : 14px;
            box-shadow    : var(--shadow);
            transition    : box-shadow .2s ease;
        }
        .section-card:hover { box-shadow: var(--shadow-hover); }

        /* ── Section title ──────────────────────────────────────────── */
        .section-title {
            font-family    : 'Nunito', sans-serif;
            font-size      : 0.92rem;
            font-weight    : 800;
            color          : var(--white);
            border-left    : 4px solid var(--primary-light);
            padding-left   : 10px;
            margin-bottom  : 6px;
            text-transform : uppercase;
            letter-spacing : 0.06em;
        }

        /* ── Dashboard header banner ────────────────────────────────── */
        .dash-header {
            background    : linear-gradient(135deg, #403B7A 0%, #2C2860 60%, #1C1A36 100%);
            border-radius : var(--radius);
            padding       : 22px 28px;
            margin-bottom : 20px;
            box-shadow    : var(--shadow);
            display       : flex;
            align-items   : center;
            gap           : 18px;
            border        : 1px solid var(--border);
        }
        .dash-header img {
            height : 54px;
            filter : brightness(1.1) drop-shadow(0 2px 8px rgba(0,0,0,.4));
        }
        .dash-header h1 {
            color       : var(--white) !important;
            font-size   : 1.7rem !important;
            margin      : 0 !important;
            text-shadow : 0 2px 10px rgba(0,0,0,.3);
        }
        .dash-header p {
            color     : var(--muted) !important;
            margin    : 4px 0 0 0 !important;
            font-size : 0.88rem !important;
        }

        /* ── KPI row label ──────────────────────────────────────────── */
        .kpi-label {
            font-family    : 'Nunito', sans-serif;
            font-size      : 0.72rem;
            font-weight    : 800;
            color          : var(--muted);
            text-transform : uppercase;
            letter-spacing : 0.12em;
            margin-bottom  : 10px;
        }

        /* ── DataFrame table ────────────────────────────────────────── */
        [data-testid="stDataFrame"] {
            background    : #1C1A36 !important;
            border        : 1px solid var(--border) !important;
            border-radius : var(--radius) !important;
            overflow      : hidden !important;
        }
        [data-testid="stDataFrame"] th {
            background     : var(--primary) !important;
            color          : var(--white) !important;
            font-family    : 'Nunito', sans-serif !important;
            font-weight    : 700 !important;
            font-size      : 0.76rem !important;
            text-transform : uppercase;
            letter-spacing : 0.06em;
        }
        [data-testid="stDataFrame"] td {
            color       : var(--muted) !important;
            font-size   : 0.82rem !important;
            border-color: var(--border) !important;
        }
        [data-testid="stDataFrame"] tr:hover td {
            background : rgba(64,59,122,.35) !important;
            color      : var(--white) !important;
        }

        /* ── Plotly chart containers ────────────────────────────────── */
        .stPlotlyChart > div {
            background    : transparent !important;
            border-radius : var(--radius) !important;
        }

        /* ── Scrollbar ──────────────────────────────────────────────── */
        ::-webkit-scrollbar       { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #1C1A36; }
        ::-webkit-scrollbar-thumb { background: var(--primary); border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: var(--primary-light); }

        /* ── Hide Streamlit chrome ──────────────────────────────────── */
        #MainMenu, footer, header { visibility: hidden !important; }
        .stDeployButton { display: none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# DATA LOADING
# FIX: use `df[col].dtype == object` — works on pandas 2.x / 3.x
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

    # Strip whitespace from text columns (pandas-version-safe)
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()

    # Parse dates
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")

    # Derived time columns
    df["Month_dt"] = df["Order_Date"].dt.to_period("M").dt.to_timestamp()
    df["Month"]    = df["Order_Date"].dt.to_period("M").astype(str)
    df["Year"]     = df["Order_Date"].dt.year

    # Numeric safety
    for col in ["Quantity", "Unit_Price_NGN", "Total_Value_NGN"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Revenue contribution %
    total_rev = df["Total_Value_NGN"].sum()
    df["Rev_Pct"] = (df["Total_Value_NGN"] / total_rev * 100) if total_rev > 0 else 0.0

    # Ordinal date for 3-D chart
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
            font=dict(family="Nunito", size=14, color="#FFFFFF"),
            x=0.01, xanchor="left",
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor ="rgba(44,40,96,0.45)",
        font=dict(family="Nunito Sans", color="#C8C5E8", size=12),
        height=height,
        margin=dict(l=10, r=10, t=46, b=10),
        legend=dict(
            font=dict(size=11, color="#C8C5E8"),
            bgcolor="rgba(28,26,54,0.85)",
            bordercolor="rgba(100,95,170,0.35)",
            borderwidth=1,
        ),
        xaxis=dict(
            gridcolor     ="rgba(100,95,170,0.18)",
            linecolor     ="rgba(100,95,170,0.35)",
            zerolinecolor ="rgba(100,95,170,0.25)",
            tickfont      =dict(color="#C8C5E8", size=11),
            title_font    =dict(color="#FFFFFF",  size=12),
        ),
        yaxis=dict(
            gridcolor     ="rgba(100,95,170,0.18)",
            linecolor     ="rgba(100,95,170,0.35)",
            zerolinecolor ="rgba(100,95,170,0.25)",
            tickfont      =dict(color="#C8C5E8", size=11),
            title_font    =dict(color="#FFFFFF",  size=12),
        ),
        hoverlabel=dict(
            bgcolor      ="rgba(64,59,122,0.95)",
            font_color   ="#FFFFFF",
            font_family  ="Nunito Sans",
            bordercolor  ="rgba(149,144,200,0.5)",
        ),
        colorway=CHART_COLORS,
    )
    return fig


# =============================================================================
# SIDEBAR
# =============================================================================
def render_sidebar(df):
    with st.sidebar:
        # Logo
        if LOGO_SRC:
            st.markdown(
                f"""
                <div style="text-align:center;padding:1.2rem 0 1rem;
                            border-bottom:1px solid rgba(100,95,170,.3);
                            margin-bottom:1rem">
                    <img src="{LOGO_SRC}" style="width:150px;
                         filter:brightness(1.1) drop-shadow(0 2px 8px rgba(100,95,170,.5))"/>
                    <div style="font-size:.62rem;font-weight:700;letter-spacing:2px;
                                text-transform:uppercase;color:rgba(200,197,232,.6);
                                margin-top:.4rem">E-Commerce Analytics</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown("## 🛒 subsi")

        st.markdown(
            "<div style='color:rgba(200,197,232,.65);font-size:.72rem;"
            "text-transform:uppercase;letter-spacing:.1em;font-weight:700;"
            "margin-bottom:10px'>🔍 Dashboard Filters</div>",
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

        st.markdown(
            "<div style='color:rgba(200,197,232,.65);font-size:.72rem;"
            "text-transform:uppercase;letter-spacing:.1em;font-weight:700;"
            "margin:10px 0 6px'>📅 Order Date Range</div>",
            unsafe_allow_html=True,
        )
        min_d = df["Order_Date"].min().date()
        max_d = df["Order_Date"].max().date()
        date_start = st.date_input("From", value=min_d, min_value=min_d, max_value=max_d, key="ds")
        date_end   = st.date_input("To",   value=max_d, min_value=min_d, max_value=max_d, key="de")

        st.markdown("---")
        st.markdown(
            "<div style='color:rgba(200,197,232,.45);font-size:.66rem;"
            "text-align:center;line-height:1.8'>Subsi Intelligence Dashboard<br>"
            "Built by <strong style=\"color:rgba(200,197,232,.7)\">ToheebBI</strong><br>"
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
    n       = len(fdf)
    rev     = fdf["Total_Value_NGN"].sum()
    aov     = rev / n if n else 0
    deliv   = (fdf["Order_Status"] == "Delivered").sum()
    cancel  = (fdf["Order_Status"] == "Cancelled").sum()
    ful_pct = deliv / n * 100 if n else 0
    can_pct = cancel / n * 100 if n else 0
    qty     = int(fdf["Quantity"].sum())
    custs   = fdf["Customer_ID"].nunique()
    top_cat = (
        fdf.groupby("Product_Category")["Total_Value_NGN"].sum().idxmax()
        if n > 0 else "—"
    )

    st.markdown("<div class='kpi-label'>📊 Key Performance Indicators</div>", unsafe_allow_html=True)

    r1 = st.columns(4)
    r1[0].metric("🛒 Total Orders",     f"{n:,}")
    r1[1].metric("💰 Total Revenue",    f"₦{rev:,.0f}")
    r1[2].metric("🧾 Avg Order Value",  f"₦{aov:,.0f}")
    r1[3].metric("🏆 Top Category",     top_cat)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    r2 = st.columns(4)
    r2[0].metric("✅ Fulfilment Rate",   f"{ful_pct:.1f}%")
    r2[1].metric("❌ Cancellation Rate", f"{can_pct:.1f}%")
    r2[2].metric("📦 Total Qty Sold",    f"{qty:,}")
    r2[3].metric("👥 Unique Customers",  f"{custs:,}")


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
                 color_continuous_scale=["#2C2860", "#403B7A", "#9590C8"],
                 text=d["Revenue"].apply(lambda x: f"₦{x/1e6:.1f}M"))
    fig.update_traces(textposition="outside", marker_line_width=0,
                      textfont=dict(color="#FFFFFF"))
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
        line=dict(color="#FFFFFF", width=3),
        marker=dict(size=8, color="#9590C8",
                    line=dict(width=2, color="#403B7A")),
        fill="tozeroy",
        fillcolor="rgba(64,59,122,0.30)",
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
    fig.update_traces(textposition="outside", textfont_size=12,
                      textfont_color="#FFFFFF",
                      pull=[0.03] * len(d))
    fig = theme(fig, "Payment Method Breakdown", height=360)
    fig.update_layout(
        legend=dict(orientation="v", x=1.0, y=0.5),
        annotations=[dict(
            text="Payment<br>Split", x=0.5, y=0.5,
            font_size=13, font_color="#FFFFFF",
            font_family="Nunito", showarrow=False,
        )],
    )
    return fig


def chart_top10(fdf):
    d = (fdf.groupby("Product_Name")["Total_Value_NGN"]
         .sum().nlargest(10).reset_index()
         .sort_values("Total_Value_NGN"))
    d.columns = ["Product", "Revenue"]
    fig = px.bar(d, y="Product", x="Revenue", orientation="h",
                 color="Revenue",
                 color_continuous_scale=["#2C2860", "#403B7A", "#9590C8"],
                 text=d["Revenue"].apply(lambda x: f"₦{x/1e6:.1f}M"))
    fig.update_traces(textposition="outside", marker_line_width=0,
                      textfont=dict(color="#FFFFFF"))
    fig.update_coloraxes(showscale=False)
    return theme(fig, "Top 10 Products by Revenue", height=400)


def chart_by_state(fdf):
    d = (fdf.groupby("Delivery_State")
         .agg(Orders=("Order_ID", "count"), Revenue=("Total_Value_NGN", "sum"))
         .reset_index().sort_values("Revenue", ascending=False))
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Revenue (₦)", x=d["Delivery_State"], y=d["Revenue"],
        marker_color="#403B7A", yaxis="y",
        hovertemplate="<b>%{x}</b><br>₦%{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        name="Orders", x=d["Delivery_State"], y=d["Orders"],
        mode="lines+markers",
        line=dict(color="#9590C8", width=2, dash="dot"),
        marker=dict(size=8, color="#C8C5E8",
                    line=dict(width=1.5, color="#403B7A")),
        yaxis="y2",
        hovertemplate="<b>%{x}</b><br>Orders: %{y}<extra></extra>",
    ))
    fig = theme(fig, "Revenue & Orders by Delivery State")
    fig.update_layout(
        yaxis=dict(title="Revenue (₦)", tickprefix="₦", tickformat=",.0f"),
        yaxis2=dict(title="Orders", overlaying="y", side="right",
                    showgrid=False, tickfont=dict(color="#C8C5E8")),
    )
    return fig


def chart_treemap(fdf):
    d = (fdf.groupby(["Product_Category", "Product_Subcategory"])["Total_Value_NGN"]
         .sum().reset_index())
    d.columns = ["Category", "Subcategory", "Revenue"]
    fig = px.treemap(
        d, path=["Category", "Subcategory"], values="Revenue",
        color="Revenue",
        color_continuous_scale=["#2C2860", "#403B7A", "#9590C8", "#C8C5E8"],
    )
    fig.update_traces(
        textfont=dict(family="Nunito", size=12, color="white"),
        hovertemplate="<b>%{label}</b><br>₦%{value:,.0f}<extra></extra>",
    )
    fig.update_coloraxes(showscale=False)
    fig = theme(fig, "Revenue by Product Subcategory (Treemap)", height=420)
    fig.update_layout(margin=dict(l=10, r=10, t=46, b=10))
    return fig


def chart_qty_category(fdf):
    d = (fdf.groupby(["Product_Category", "Order_Status"])["Quantity"]
         .sum().reset_index())
    d.columns = ["Category", "Status", "Quantity"]
    fig = px.bar(d, x="Category", y="Quantity", color="Status",
                 barmode="group",
                 color_discrete_sequence=CHART_COLORS)
    fig.update_traces(marker_line_width=0)
    fig = theme(fig, "Quantity Sold by Category & Status")
    fig.update_xaxes(tickangle=-20)
    return fig


def chart_area(fdf):
    d = (fdf.groupby(["Order_Date", "Order_Status"])["Order_ID"]
         .count().reset_index().sort_values("Order_Date"))
    d.columns = ["Date", "Status", "Orders"]
    fig = px.area(d, x="Date", y="Orders", color="Status",
                  color_discrete_sequence=CHART_COLORS,
                  line_group="Status")
    fig.update_traces(line=dict(width=2))
    return theme(fig, "Daily Order Volume by Status", height=350)


def chart_3d(fdf):
    d = fdf.dropna(subset=["Date_Ordinal", "Quantity", "Total_Value_NGN"]).copy()
    if d.empty:
        return None

    status_colors = {
        "Delivered":  "#403B7A",
        "Shipped":    "#6B64A8",
        "Processing": "#9590C8",
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
                color=status_colors.get(str(status), "#9590C8"),
                opacity=0.85,
                line=dict(width=0.5, color="rgba(255,255,255,.3)"),
            ),
            customdata=grp[["Order_ID", "Product_Name",
                            "Delivery_State", "Payment_Method"]].values,
            hovertemplate=(
                "<b>%{customdata[1]}</b><br>"
                "Order: %{customdata[0]}<br>"
                "State: %{customdata[2]}<br>"
                "Payment: %{customdata[3]}<br>"
                "Qty: %{y}<br>"
                "Revenue: ₦%{z:,.0f}<extra></extra>"
            ),
        ))

    # FIX: Plotly 5.x — use nested title dict for 3D axes and colorbar
    fig.update_layout(
        title=dict(
            text="3D Intelligence Scatter — Date × Quantity × Revenue",
            font=dict(family="Nunito", size=14, color="#FFFFFF"),
            x=0.01,
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        height=520,
        margin=dict(l=0, r=0, t=50, b=0),
        scene=dict(
            bgcolor="rgba(28,26,54,0.95)",
            xaxis=dict(
                title=dict(text="Order Date (Ordinal)",
                           font=dict(color="#FFFFFF", size=11)),
                backgroundcolor="rgba(44,40,96,0.8)",
                gridcolor="rgba(100,95,170,0.3)",
                color="#C8C5E8",
                showbackground=True,
            ),
            yaxis=dict(
                title=dict(text="Quantity",
                           font=dict(color="#FFFFFF", size=11)),
                backgroundcolor="rgba(44,40,96,0.8)",
                gridcolor="rgba(100,95,170,0.3)",
                color="#C8C5E8",
                showbackground=True,
            ),
            zaxis=dict(
                title=dict(text="Revenue (₦)",
                           font=dict(color="#FFFFFF", size=11)),
                backgroundcolor="rgba(44,40,96,0.8)",
                gridcolor="rgba(100,95,170,0.3)",
                color="#C8C5E8",
                showbackground=True,
            ),
        ),
        legend=dict(
            font=dict(size=11, color="#C8C5E8"),
            bgcolor="rgba(28,26,54,0.85)",
            bordercolor="rgba(100,95,170,0.35)",
            borderwidth=1,
        ),
        hoverlabel=dict(
            bgcolor="rgba(64,59,122,0.95)",
            font_color="#FFFFFF",
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

    with st.expander("📋 Executive Insight Summary — Click to expand", expanded=False):
        st.markdown(
            f"""
            <div style='font-family:Nunito Sans,sans-serif;color:#C8C5E8;line-height:1.85'>
            <h4 style='color:#FFFFFF;font-family:Nunito,sans-serif;margin-top:0'>
                🧠 Operational Intelligence Overview
            </h4>
            <p>Filtered view: <strong style="color:#FFFFFF">{n:,} orders</strong>
            &nbsp;·&nbsp; Total Revenue:
            <strong style="color:#FFFFFF">&#x20A6;{rev:,.0f}</strong></p>
            <hr style='border-color:rgba(100,95,170,.3)'>
            <p>✅ <strong style="color:#FFFFFF">Fulfilment:</strong>
            {del_pct:.1f}% delivered.
            {'Strong operational performance.' if del_pct >= 70 else 'Improvement opportunity identified.'}
            Cancellation rate {can_pct:.1f}%
            {'— within acceptable range.' if can_pct <= 15 else '— requires immediate operational review.'}</p>
            <p>🏆 <strong style="color:#FFFFFF">Category leader:</strong>
            <em>{top_cat}</em> — concentrate inventory investment and marketing budget here.</p>
            <p>📍 <strong style="color:#FFFFFF">Top state:</strong>
            <em>{top_state}</em> — prioritise logistics and warehousing resources in this region.</p>
            <p>💳 <strong style="color:#FFFFFF">Dominant payment channel:</strong>
            <em>{top_pay}</em> — expand payment alternatives to reduce conversion friction.</p>
            <p>📈 <strong style="color:#FFFFFF">Seasonality:</strong>
            Monitor the Monthly Revenue Trend chart to time stock procurement around demand peaks.</p>
            <p>⚡ <strong style="color:#FFFFFF">Action:</strong>
            Operations teams should review the state fulfilment chart daily.
            Finance teams can use the Treemap and grouped bar charts to assess
            category and geographic revenue concentration risk.</p>
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
        st.plotly_chart(fig, use_container_width=True,
                        config={"displayModeBar": False})
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
            "2. Place `subsi_ecommerce_dataset.xlsx`, `subsi.png`, "
            "and `app.py` in the **same folder**\n"
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

    # ── Dashboard header banner ──────────────────────────────────────────────
    logo_html = (
        f'<img src="{LOGO_SRC}" alt="Subsi"/>'
        if LOGO_SRC
        else '<span style="font-size:2.2rem">🛒</span>'
    )
    st.markdown(
        f"""
        <div class='dash-header'>
            {logo_html}
            <div>
                <h1>Subsi Order Fulfillment Intelligence</h1>
                <p>Showing <strong>{len(fdf):,}</strong> of
                <strong>{len(df):,}</strong> orders
                &nbsp;·&nbsp; {date_start.strftime('%d %b %Y')} →
                {date_end.strftime('%d %b %Y')}
                &nbsp;·&nbsp; All filters active
                &nbsp;·&nbsp; Powered by <strong>ToheebBI</strong></p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if len(fdf) == 0:
        st.warning("⚠️ No records match the current filters. Please adjust the sidebar.")
        return

    # ── KPI section ──────────────────────────────────────────────────────────
    render_kpis(fdf)
    st.markdown("---")

    # ── Row 1 — Status | Monthly Revenue ────────────────────────────────────
    c1, c2 = st.columns([1, 1.6])
    with c1:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Order Status Distribution", chart_status(fdf))
        st.markdown("</div>", unsafe_allow_html=True)
    with c2:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Monthly Revenue Trend", chart_monthly(fdf))
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Row 2 — Revenue by Category | Payment Method ─────────────────────────
    c3, c4 = st.columns([1.6, 1])
    with c3:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Revenue by Product Category", chart_rev_category(fdf))
        st.markdown("</div>", unsafe_allow_html=True)
    with c4:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Payment Method Breakdown", chart_payment(fdf))
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Row 3 — Top 10 Products | By State ───────────────────────────────────
    c5, c6 = st.columns(2)
    with c5:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Top 10 Products by Revenue", chart_top10(fdf))
        st.markdown("</div>", unsafe_allow_html=True)
    with c6:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Revenue & Orders by Delivery State", chart_by_state(fdf))
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Row 4 — Treemap | Qty by Category ────────────────────────────────────
    c7, c8 = st.columns([1.2, 1])
    with c7:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Revenue by Product Subcategory", chart_treemap(fdf))
        st.markdown("</div>", unsafe_allow_html=True)
    with c8:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section_chart("Quantity Sold by Category & Status", chart_qty_category(fdf))
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Row 5 — Area chart (full width) ──────────────────────────────────────
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    section_chart("Daily Order Volume by Status", chart_area(fdf))
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Row 6 — 3D Scatter (full width) ──────────────────────────────────────
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    section_chart(
        "3D Intelligence Scatter — Date × Quantity × Revenue",
        chart_3d(fdf),
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")

    # ── Executive insights ────────────────────────────────────────────────────
    render_insights(fdf)

    # ── Footer ────────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style='text-align:center;padding:18px 0 8px;
             color:rgba(200,197,232,.45);font-size:.72rem;
             font-family:Nunito Sans,sans-serif;
             border-top:1px solid rgba(100,95,170,.25);margin-top:1rem'>
            Subsi Order Fulfillment Intelligence &nbsp;·&nbsp;
            Streamlit + Plotly &nbsp;·&nbsp;
            Built by <strong style="color:rgba(200,197,232,.7)">ToheebBI</strong>
            &nbsp;·&nbsp; © 2025 Subsi E-Commerce
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
