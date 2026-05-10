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

VALTECH_BLUE = "#00bdfa"
VALTECH_CYAN = "#00fffa"
VALTECH_ROSE = "#ff1b78"
VALTECH_TURQUOISE = "#0dfdb4"
VALTECH_YELLOW = "#fff011"
BLACK = "#000000"
WHITE = "#ffffff"
DARK = "#101010"
MID_GREY = "#343434"
LIGHT_GREY = "#f4f4f4"

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


LOGO_MAP = {
    "Toyota": "https://logo.clearbit.com/toyota.com",
    "Lexus": "https://logo.clearbit.com/lexus.com",
    "VW": "https://logo.clearbit.com/volkswagen.com",
    "BMW": "https://logo.clearbit.com/bmw.com",
    "Mercedes-Benz": "https://logo.clearbit.com/mercedes-benz.com",
    "Audi": "https://logo.clearbit.com/audi.com",
    "Volvo": "https://logo.clearbit.com/volvocars.com",
    "Tesla": "https://logo.clearbit.com/tesla.com",
    "Ford": "https://logo.clearbit.com/ford.com",
    "Peugeot": "https://logo.clearbit.com/peugeot.com",
    "Renault": "https://logo.clearbit.com/renault.com",
    "Hyundai": "https://logo.clearbit.com/hyundai.com",
    "Kia": "https://logo.clearbit.com/kia.com",
    "Nissan": "https://logo.clearbit.com/nissan-global.com",
    "Skoda": "https://logo.clearbit.com/skoda-auto.com",
    "Seat": "https://logo.clearbit.com/seat.com",
    "Dacia": "https://logo.clearbit.com/dacia.com",
    "Jaguar": "https://logo.clearbit.com/jaguar.com",
    "Land Rover": "https://logo.clearbit.com/landrover.com",
    "BYD": "https://logo.clearbit.com/byd.com",
    "MG": "https://logo.clearbit.com/mgmotor.com",
    "XPENG": "https://logo.clearbit.com/xpeng.com",
    "NIO": "https://logo.clearbit.com/nio.com",
    "Geely": "https://logo.clearbit.com/geely.com",
    "Omoda": "https://logo.clearbit.com/omodaauto.com",
    "Jaecoo": "https://logo.clearbit.com/jaecoo.com",
}


st.markdown(
    f"""
    <style>
    html, body, [class*="css"] {{
        font-family: "Helvetica Neue", Arial, sans-serif;
    }}

    .stApp {{
        background: linear-gradient(180deg, #ffffff 0%, #f7f7f7 100%);
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

    .valtech-header {{
        background: #000000;
        color: white;
        padding: 28px 32px;
        border-radius: 22px;
        margin-bottom: 20px;
        position: relative;
        overflow: hidden;
    }}

    .valtech-header:after {{
        content: "";
        position: absolute;
        width: 420px;
        height: 420px;
        border-radius: 50%;
        background: linear-gradient(135deg, {VALTECH_BLUE}, {VALTECH_TURQUOISE});
        right: -180px;
        top: -230px;
        opacity: .9;
    }}

    .valtech-logo {{
        font-size: 30px;
        font-weight: 800;
        letter-spacing: -1px;
        margin-bottom: 18px;
        position: relative;
        z-index: 2;
    }}

    .valtech-logo span {{
        color: {VALTECH_BLUE};
    }}

    .valtech-title {{
        font-size: 42px;
        line-height: 1.05;
        font-weight: 800;
        max-width: 900px;
        position: relative;
        z-index: 2;
    }}

    .valtech-subtitle {{
        color: #dcdcdc;
        font-size: 17px;
        margin-top: 12px;
        max-width: 1000px;
        position: relative;
        z-index: 2;
    }}

    .metric-card {{
        background: #ffffff;
        border: 1px solid #e5e5e5;
        border-radius: 18px;
        padding: 18px;
        box-shadow: 0 2px 12px rgba(0,0,0,.04);
    }}

    .metric-label {{
        color: #555;
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: .06em;
        margin-bottom: 8px;
    }}

    .metric-value {{
        font-size: 28px;
        font-weight: 800;
        color: #000;
    }}

    .insight-box {{
        background: #000;
        color: #fff;
        padding: 18px 20px;
        border-left: 8px solid {VALTECH_BLUE};
        border-radius: 14px;
        margin-bottom: 18px;
    }}

    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
    }}

    .stTabs [data-baseweb="tab"] {{
        background: #ffffff;
        border-radius: 999px;
        border: 1px solid #dddddd;
        padding: 8px 18px;
    }}

    .stTabs [aria-selected="true"] {{
        background: #000000 !important;
        color: #ffffff !important;
        border-color: #000000 !important;
    }}

    div[data-testid="stMetricValue"] {{
        color: #000000;
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

    hr {{
        border: none;
        border-top: 1px solid #e5e5e5;
        margin: 24px 0;
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
def load_data(uploaded_file=None):
    if uploaded_file is not None:
        raw = pd.read_excel(uploaded_file)
    else:
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
            uv_col: "UniqueVisits",
        }
    )

    df["OEM"] = df["OEM"].astype(str).str.strip()
    df["Market"] = df["Market"].astype(str).str.strip()
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")
    df["UniqueVisits"] = pd.to_numeric(df["UniqueVisits"], errors="coerce")

    df = df.dropna(subset=["OEM", "Market", "Year", "Sales", "UniqueVisits"])
    df = df[(df["UniqueVisits"] > 0) & (df["Sales"] >= 0)]
    df["Year"] = df["Year"].astype(int)

    market_data = (
        df.groupby(["OEM", "Market", "Year"], as_index=False)
        .agg({"Sales": "sum", "UniqueVisits": "sum"})
    )
    market_data["ConversionRate"] = market_data["Sales"] / market_data["UniqueVisits"]
    market_data["ConversionPct"] = market_data["ConversionRate"] * 100

    all_mm5 = (
        df.groupby(["OEM", "Year"], as_index=False)
        .agg({"Sales": "sum", "UniqueVisits": "sum"})
    )
    all_mm5["Market"] = "MM5"
    all_mm5["ConversionRate"] = all_mm5["Sales"] / all_mm5["UniqueVisits"]
    all_mm5["ConversionPct"] = all_mm5["ConversionRate"] * 100

    combined = pd.concat(
        [
            all_mm5[["OEM", "Market", "Year", "Sales", "UniqueVisits", "ConversionPct"]],
            market_data[["OEM", "Market", "Year", "Sales", "UniqueVisits", "ConversionPct"]],
        ],
        ignore_index=True,
    )

    return combined


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
        logo = LOGO_MAP.get(row["OEM"])
        if not logo:
            continue

        size_x = x_max * 0.035
        size_y = y_max * 0.055

        fig.add_layout_image(
            dict(
                source=logo,
                xref="x",
                yref="y",
                x=row["UniqueVisits"],
                y=row["ConversionPct"],
                sizex=size_x,
                sizey=size_y,
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
    max_x = max(chart_df["UniqueVisits"].max() * 1.15, 1)
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
                        x=[r24["UniqueVisits"], r25["UniqueVisits"]],
                        y=[r24["ConversionPct"], r25["ConversionPct"]],
                        mode="lines",
                        line=dict(width=2, dash="dot", color="rgba(0,0,0,0.38)"),
                        hoverinfo="skip",
                        showlegend=False,
                    )
                )

    years = [2024, 2025] if year_view == "2024 and 2025 + shift" else [int(year_view)]

    colors = {
        2024: VALTECH_BLUE,
        2025: VALTECH_ROSE,
    }

    for year in years:
        year_df = chart_df[chart_df["Year"] == year].copy()
        if year_df.empty:
            continue

        marker_sizes = year_df["Sales"].apply(lambda s: max(18, math.sqrt(s / max_sales) * 74))

        fig.add_trace(
            go.Scatter(
                x=year_df["UniqueVisits"],
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
                customdata=year_df[["OEM", "Market", "Year", "Sales", "UniqueVisits", "ConversionPct"]],
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "Market: %{customdata[1]}<br>"
                    "Year: %{customdata[2]}<br>"
                    "Sales: %{customdata[3]:,.0f}<br>"
                    "Unique visits: %{customdata[4]:,.0f}<br>"
                    "Conversion rate: %{customdata[5]:.2f}%"
                    "<extra></extra>"
                ),
            )
        )

    if show_logos:
        add_logo_images(fig, chart_df, max_x, max_y)

    fig.update_layout(
        title=dict(
            text=f"{market} | Unique visits vs conversion rate",
            x=0.02,
            font=dict(size=24, color=BLACK),
        ),
        xaxis_title="Unique visits",
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

    latest = chart_data[chart_data["Year"] == chart_data["Year"].max()]
    total_sales = latest["Sales"].sum()
    total_uv = latest["UniqueVisits"].sum()
    weighted_conv = (total_sales / total_uv * 100) if total_uv else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Selected sales", f"{total_sales:,.0f}")
    c2.metric("Selected unique visits", f"{total_uv:,.0f}")
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
                    "UniqueVisits": "Unique visits",
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
    <div class="valtech-header">
        <div class="valtech-logo">valtech<span>_</span></div>
        <div class="valtech-title">OEM Digital Performance Dashboard</div>
        <div class="valtech-subtitle">
            Passenger car digital sales view across MM5. Explore how OEM traffic scale, conversion efficiency and 2024–2025 movement differ by market.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="insight-box">
        <b>Read this correctly:</b> this is a digital performance dashboard, not a total market-share dashboard.
        It uses passenger car sales and unique visits; it should not be used as proof of total brand share where fleet, LCV or tactical registrations are out of scope.
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("Data definition and caveats", expanded=False):
    st.markdown(
        """
        - **Markets:** UK, France, Italy, Spain and Germany.
        - **MM5:** aggregated view across those five markets.
        - **Sales:** passenger car sales from the uploaded workbook.
        - **Unique visits:** Similarweb traffic metric from the uploaded workbook.
        - **Conversion rate:** sales divided by unique visits.
        - **Exclusions:** do not position this as full market share unless fleet, LCV and tactical registrations are confirmed to be included.
        - **Toyota and Lexus:** treated as separate brands.
        """
    )

data = load_data(None)
available_oems = sorted(data["OEM"].unique())

st.sidebar.header("Filters")
st.sidebar.caption("Using the bundled Excel dataset from GitHub.")

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
    help="Logos require internet access. Turn off if any logos fail to load."
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
    "Valtech-styled dashboard. Recommended executive flow: start with MM5, then use the market tabs to identify where the conversion gap is most material."
)
