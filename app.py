import math
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="OEM Digital Performance Dashboard",
    page_icon="🚗",
    layout="wide",
)


DATA_FILE = Path(__file__).parent / "OEM Sales Data 2024-2025.xlsx"

# Valtech-inspired palette requested: grey + blue
VALTECH_BLUE = "#009FE3"
VALTECH_GREY = "#6F6F6F"
VALTECH_LIGHT_GREY = "#F3F3F3"
BLACK = "#000000"
WHITE = "#FFFFFF"


TOYOTA_SET = [
    "Toyota", "VW", "Ford", "Peugeot", "Renault",
    "Hyundai", "Kia", "Nissan", "Skoda", "Seat", "Dacia"
]

LEXUS_SET = [
    "Lexus", "BMW", "Mercedes-Benz", "Audi",
    "Volvo", "Tesla", "Jaguar", "Land Rover"
]

CHINESE_SET = [
    "BYD", "MG", "XPENG", "NIO", "Geely",
    "Omoda", "Jaecoo", "Leapmotor", "Ora", "Aiways"
]


LOGO_DOMAIN_MAP = {
    "Toyota": "toyota.com",
    "Lexus": "lexus.com",
    "VW": "volkswagen.com",
    "Volkswagen": "volkswagen.com",
    "BMW": "bmw.com",
    "Mercedes-Benz": "mercedes-benz.com",
    "Mercedes": "mercedes-benz.com",
    "Audi": "audi.com",
    "Volvo": "volvocars.com",
    "Tesla": "tesla.com",
    "Ford": "ford.com",
    "Peugeot": "peugeot.com",
    "Renault": "renault.com",
    "Hyundai": "hyundai.com",
    "Kia": "kia.com",
    "Nissan": "nissan-global.com",
    "Skoda": "skoda-auto.com",
    "Škoda": "skoda-auto.com",
    "Seat": "seat.com",
    "Dacia": "dacia.com",
    "Jaguar": "jaguar.com",
    "Land Rover": "landrover.com",
    "BYD": "byd.com",
    "MG": "mg.co.uk",
    "XPENG": "xpeng.com",
    "Xpeng": "xpeng.com",
    "NIO": "nio.com",
    "Geely": "geely.com",
    "Omoda": "omodaauto.com",
    "Jaecoo": "jaecoo.com",
    "Leapmotor": "leapmotor.com",
    "Ora": "gwm-global.com",
    "GWM Ora": "gwm-global.com",
    "Aiways": "ai-ways.eu",
    "Citroen": "citroen.com",
    "Citroën": "citroen.com",
    "Fiat": "fiat.com",
    "Mini": "mini.com",
    "MINI": "mini.com",
    "Mazda": "mazda.com",
    "Honda": "honda.com",
    "Suzuki": "suzuki.com",
    "Opel": "opel.com",
    "Vauxhall": "vauxhall.co.uk",
    "Porsche": "porsche.com",
    "Alfa Romeo": "alfaromeo.com",
    "Cupra": "cupraofficial.com",
    "DS": "dsautomobiles.com",
    "DS Automobiles": "dsautomobiles.com",
    "Jeep": "jeep.com",
    "Mitsubishi": "mitsubishi-motors.com",
    "Smart": "smart.com",
    "Polestar": "polestar.com",
    "Subaru": "subaru.com",
    "Maserati": "maserati.com",
}


def get_brandfetch_client_id():
    try:
        return st.secrets.get("BRANDFETCH_CLIENT_ID", "")
    except Exception:
        return ""


def get_logo_url(oem):
    client_id = get_brandfetch_client_id()
    domain = LOGO_DOMAIN_MAP.get(oem)
    if not client_id or not domain:
        return None
    return f"https://cdn.brandfetch.io/{domain}?c={client_id}"


st.markdown(
    f"""
    <style>
    html, body, [class*="css"] {{
        font-family: "Helvetica Neue", Arial, sans-serif;
    }}

    .stApp {{
        background: #ffffff;
    }}

    section[data-testid="stSidebar"] {{
        background: #000000;
    }}

    section[data-testid="stSidebar"] * {{
        color: white;
    }}

    section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] *,
    section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] *,
    section[data-testid="stSidebar"] input {{
        color: black !important;
    }}

    .hero {{
        background: #000000;
        color: #ffffff;
        padding: 28px 32px;
        border-radius: 22px;
        margin-bottom: 22px;
        position: relative;
        overflow: hidden;
        border-bottom: 8px solid {VALTECH_BLUE};
    }}

    .hero::after {{
        content: "";
        position: absolute;
        width: 430px;
        height: 430px;
        border-radius: 50%;
        background: linear-gradient(135deg, {VALTECH_BLUE}, {VALTECH_GREY});
        right: -180px;
        top: -230px;
        opacity: 0.55;
    }}

    .hero-logo {{
        position: relative;
        z-index: 2;
        margin-bottom: 20px;
    }}

    .hero-logo img {{
        height: 42px;
        width: auto;
        background: #ffffff;
        padding: 8px 12px;
        border-radius: 8px;
    }}

    .hero-title {{
        position: relative;
        z-index: 2;
        font-size: 42px;
        line-height: 1.05;
        font-weight: 800;
        letter-spacing: -0.04em;
        max-width: 920px;
    }}

    .hero-subtitle {{
        position: relative;
        z-index: 2;
        margin-top: 12px;
        color: #e6e6e6;
        font-size: 17px;
        max-width: 980px;
    }}

    .caveat-box {{
        background: {VALTECH_LIGHT_GREY};
        border-left: 8px solid {VALTECH_BLUE};
        color: #000000;
        padding: 18px 20px;
        border-radius: 14px;
        margin-bottom: 18px;
        font-size: 16px;
    }}

    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
    }}

    .stTabs [data-baseweb="tab"] {{
        background: #ffffff;
        border-radius: 999px;
        border: 1px solid #d8d8d8;
        padding: 8px 18px;
    }}

    .stTabs [aria-selected="true"] {{
        background: #000000 !important;
        color: #ffffff !important;
        border-color: #000000 !important;
    }}

    .stButton button, .stDownloadButton button {{
        background: #000000;
        color: white;
        border-radius: 999px;
        border: 1px solid #000000;
    }}

    .stButton button:hover, .stDownloadButton button:hover {{
        background: {VALTECH_BLUE};
        color: #000000;
        border-color: {VALTECH_BLUE};
    }}

    div[data-testid="stMetricValue"] {{
        color: #000000;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def find_column(columns, candidates):
    lower = {str(c).strip().lower(): c for c in columns}
    for candidate in candidates:
        for key, original in lower.items():
            if candidate in key:
                return original
    raise ValueError(f"Could not find any of these columns: {candidates}")


@st.cache_data
def load_data():
    raw = pd.read_excel(DATA_FILE)
    raw.columns = [str(c).strip() for c in raw.columns]

    brand_col = find_column(raw.columns, ["brand", "oem"])
    year_col = find_column(raw.columns, ["year"])
    market_col = find_column(raw.columns, ["market", "country"])
    sales_col = find_column(raw.columns, ["sales"])
    uv_col = find_column(raw.columns, ["unique", "uv"])

    df = raw[[brand_col, market_col, year_col, sales_col, uv_col]].copy()
    df = df.rename(
        columns={
            brand_col: "OEM",
            market_col: "Market",
            year_col: "Year",
            sales_col: "Sales",
            uv_col: "UniqueVisitors",
        }
    )

    df["OEM"] = df["OEM"].astype(str).str.strip()
    df["Market"] = df["Market"].astype(str).str.strip()
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")
    df["UniqueVisitors"] = pd.to_numeric(df["UniqueVisitors"], errors="coerce")

    df = df.dropna(subset=["OEM", "Market", "Year", "Sales", "UniqueVisitors"])
    df = df[(df["UniqueVisitors"] > 0) & (df["Sales"] >= 0)]
    df["Year"] = df["Year"].astype(int)

    market_data = (
        df.groupby(["OEM", "Market", "Year"], as_index=False)
        .agg({"Sales": "sum", "UniqueVisitors": "sum"})
    )
    market_data["ConversionPct"] = market_data["Sales"] / market_data["UniqueVisitors"] * 100

    all_mm5 = (
        df.groupby(["OEM", "Year"], as_index=False)
        .agg({"Sales": "sum", "UniqueVisitors": "sum"})
    )
    all_mm5["Market"] = "MM5"
    all_mm5["ConversionPct"] = all_mm5["Sales"] / all_mm5["UniqueVisitors"] * 100

    return pd.concat(
        [
            all_mm5[["OEM", "Market", "Year", "Sales", "UniqueVisitors", "ConversionPct"]],
            market_data[["OEM", "Market", "Year", "Sales", "UniqueVisitors", "ConversionPct"]],
        ],
        ignore_index=True,
    )


def preset_selection(preset_name, available_oems):
    if preset_name == "Toyota volume competitors":
        preset = TOYOTA_SET
    elif preset_name == "Lexus premium competitors":
        preset = LEXUS_SET
    elif preset_name == "Chinese disruptors":
        preset = CHINESE_SET
    else:
        preset = available_oems
    return [oem for oem in preset if oem in available_oems]


def add_logo_images(fig, chart_df, x_max, y_max):
    for _, row in chart_df.iterrows():
        logo = get_logo_url(row["OEM"])
        if not logo:
            continue
        fig.add_layout_image(
            dict(
                source=logo,
                xref="x",
                yref="y",
                x=row["UniqueVisitors"],
                y=row["ConversionPct"],
                sizex=x_max * 0.035,
                sizey=y_max * 0.055,
                xanchor="center",
                yanchor="middle",
                sizing="contain",
                opacity=0.95,
                layer="above",
            )
        )


def build_chart(chart_df, selected_oems, market, year_view, show_logos):
    fig = go.Figure()

    if chart_df.empty:
        return fig

    max_sales = max(chart_df["Sales"].max(), 1)
    max_x = max(chart_df["UniqueVisitors"].max() * 1.15, 1)
    max_y = max(chart_df["ConversionPct"].max() * 1.20, 0.1)

    if year_view == "2024 and 2025 + shift":
        for oem in selected_oems:
            d24 = chart_df[(chart_df["OEM"] == oem) & (chart_df["Year"] == 2024)]
            d25 = chart_df[(chart_df["OEM"] == oem) & (chart_df["Year"] == 2025)]
            if len(d24) and len(d25):
                r24 = d24.iloc[0]
                r25 = d25.iloc[0]
                fig.add_trace(
                    go.Scatter(
                        x=[r24["UniqueVisitors"], r25["UniqueVisitors"]],
                        y=[r24["ConversionPct"], r25["ConversionPct"]],
                        mode="lines",
                        line=dict(width=2, dash="dot", color="rgba(111,111,111,0.6)"),
                        hoverinfo="skip",
                        showlegend=False,
                    )
                )

    years = [2024, 2025] if year_view == "2024 and 2025 + shift" else [int(year_view)]
    colors = {2024: VALTECH_BLUE, 2025: VALTECH_GREY}

    for year in years:
        year_df = chart_df[chart_df["Year"] == year].copy()
        if year_df.empty:
            continue

        marker_sizes = year_df["Sales"].apply(lambda s: max(18, math.sqrt(s / max_sales) * 74))

        fig.add_trace(
            go.Scatter(
                x=year_df["UniqueVisitors"],
                y=year_df["ConversionPct"],
                mode="markers+text" if not show_logos else "markers",
                text=year_df["OEM"] if not show_logos else None,
                textposition="top center",
                name=str(year),
                marker=dict(
                    size=marker_sizes,
                    color=colors.get(year, VALTECH_BLUE),
                    opacity=0.68,
                    line=dict(width=1.3, color="rgba(0,0,0,0.95)"),
                ),
                customdata=year_df[["OEM", "Market", "Year", "Sales", "UniqueVisitors", "ConversionPct"]],
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Market: %{customdata[1]}<br>"
                    "Year: %{customdata[2]}<br>"
                    "Sales: %{customdata[3]:,.0f}<br>"
                    "Unique visitors: %{customdata[4]:,.0f}<br>"
                    "Conversion rate: %{customdata[5]:.2f}%"
                    "<extra></extra>"
                ),
            )
        )

    if show_logos:
        add_logo_images(fig, chart_df, max_x, max_y)

    fig.update_layout(
        title=dict(
            text=f"{market} | Unique visitors vs conversion rate",
            x=0.02,
            font=dict(size=24, color=BLACK),
        ),
        xaxis_title="Unique visitors",
        yaxis_title="Conversion rate (%)",
        height=720,
        hovermode="closest",
        legend=dict(orientation="h", y=1.08, x=0),
        margin=dict(l=70, r=40, t=90, b=70),
        plot_bgcolor=WHITE,
        paper_bgcolor=WHITE,
        font=dict(family="Helvetica Neue, Arial, sans-serif", color=BLACK),
    )

    fig.update_xaxes(
        range=[0, max_x],
        tickformat=",",
        gridcolor="#e6e6e6",
        zeroline=False,
        linecolor=BLACK,
    )
    fig.update_yaxes(
        range=[0, max_y],
        ticksuffix="%",
        gridcolor="#e6e6e6",
        zeroline=False,
        linecolor=BLACK,
    )

    return fig


def render_market_tab(data, market, selected_oems, year_view, show_logos):
    chart_data = data[(data["Market"] == market) & (data["OEM"].isin(selected_oems))].copy()

    if year_view != "2024 and 2025 + shift":
        chart_data = chart_data[chart_data["Year"] == int(year_view)]

    if chart_data.empty:
        st.warning(f"No data available for {market} with the current OEM/year selection.")
        return

    latest_year = chart_data["Year"].max()
    latest = chart_data[chart_data["Year"] == latest_year]
    total_sales = latest["Sales"].sum()
    total_uv = latest["UniqueVisitors"].sum()
    weighted_conv = (total_sales / total_uv * 100) if total_uv else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Selected sales", f"{total_sales:,.0f}")
    c2.metric("Selected unique visitors", f"{total_uv:,.0f}")
    c3.metric("Weighted conversion", f"{weighted_conv:.2f}%")

    fig = build_chart(chart_data, selected_oems, market, year_view, show_logos)
    st.plotly_chart(fig, use_container_width=True)

    display_data = chart_data.sort_values(["OEM", "Year"]).copy()
    display_data["ConversionPct"] = display_data["ConversionPct"].round(2)

    with st.expander("View selected data", expanded=False):
        st.dataframe(
            display_data.rename(
                columns={
                    "ConversionPct": "Conversion rate (%)",
                    "UniqueVisitors": "Unique visitors",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        csv = display_data.to_csv(index=False).encode("utf-8")
        st.download_button(
            f"Download {market} selected data as CSV",
            data=csv,
            file_name=f"{market.lower().replace(' ', '_')}_selected_oem_data.csv",
            mime="text/csv",
            key=f"download_{market}",
        )


st.markdown(
    """
    <div class="hero">
        <div class="hero-logo">
            <img src="https://mma.prnewswire.com/media/2728124/Valtech_Logo.jpg" alt="Valtech logo">
        </div>
        <div class="hero-title">OEM Digital Performance Dashboard</div>
        <div class="hero-subtitle">
            Passenger car digital sales view across MM5. Explore how OEM traffic scale, conversion efficiency and 2024–2025 movement differ by market.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="caveat-box">
        This is not a total market-share dashboard. It uses passenger car sales and unique visitor website data; it does not include fleet, LCV or tactical registrations.
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("Data definition and caveats", expanded=False):
    st.markdown(
        """
        - **Markets:** UK, France, Italy, Spain and Germany.
        - **MM5:** aggregated view across those five markets.
        - **Sales:** passenger car sales - Marklines.
        - **Unique visitor:** Similarweb.
        - **Conversion rate:** sales divided by unique visitors.
        """
    )

data = load_data()
available_oems = sorted(data["OEM"].unique())

st.sidebar.header("Filters")

year_view = st.sidebar.selectbox(
    "Year view",
    ["2024 and 2025 + shift", "2024", "2025"],
    index=0,
)

preset = st.sidebar.selectbox(
    "Preset",
    ["Toyota volume competitors", "Lexus premium competitors", "Chinese disruptors", "All OEMs"],
    index=0,
)

default_oems = preset_selection(preset, available_oems)

selected_oems = st.sidebar.multiselect(
    "OEMs",
    available_oems,
    default=default_oems,
)

show_logos = st.sidebar.toggle(
    "Show OEM logos",
    value=True,
    help="Logos require internet access and a Brandfetch Client ID in Streamlit Secrets.",
)

if show_logos and not get_brandfetch_client_id():
    st.sidebar.warning(
        "Brandfetch Client ID not found. Add BRANDFETCH_CLIENT_ID in Streamlit Secrets or turn logos off."
    )

if not selected_oems:
    st.warning("Select at least one OEM in the sidebar.")
    st.stop()

tabs = st.tabs(["MM5", "UK", "France", "Germany", "Italy", "Spain"])

tab_markets = [
    ("MM5", tabs[0]),
    ("UK", tabs[1]),
    ("France", tabs[2]),
    ("Germany", tabs[3]),
    ("Italy", tabs[4]),
    ("Spain", tabs[5]),
]

for market, tab in tab_markets:
    with tab:
        render_market_tab(data, market, selected_oems, year_view, show_logos)

st.markdown("---")
st.caption(
    "Recommended executive flow: start with MM5, then use the market tabs to identify where the conversion gap is most material."
)
