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

VALTECH_BLUE = "#009FE3"
VALTECH_GREY = "#6F6F6F"
VALTECH_LIGHT_GREY = "#F3F3F3"
BLACK = "#000000"
WHITE = "#FFFFFF"
GREEN_HEADER = "#A6FF4D"


TOYOTA_SET = [
    "Toyota", "VW", "Ford", "Peugeot", "Renault",
    "Hyundai", "Kia", "Nissan", "Skoda", "Seat", "Dacia"
]

LEXUS_SET = [
    "Lexus", "BMW", "Mercedes-Benz", "Audi",
    "Volvo", "Tesla", "Jaguar", "Land Rover", "Porsche", "Polestar", "Jaguar"
]

CHINESE_SET = [
    "BYD", "BYD Auto", "MG", "XPENG", "Xpeng", "NIO", "Geely",
    "Omoda", "Jaecoo", "Leapmotor", "Ora", "GWM Ora", "Aiways"
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
    "BYD Auto": "byd.com",
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

    .summary-card {{
        background: #ffffff;
        border: 1px solid #e2e2e2;
        border-radius: 18px;
        padding: 20px;
        box-shadow: 0 2px 12px rgba(0,0,0,.04);
        min-height: 170px;
    }}

    .summary-card h3 {{
        margin-top: 0;
        color: #000000;
        font-size: 20px;
    }}

    .summary-card p, .summary-card li {{
        color: #222222;
        font-size: 15px;
    }}

    .ai-box {{
        background: #000000;
        color: #ffffff;
        border-left: 8px solid {VALTECH_BLUE};
        padding: 20px;
        border-radius: 18px;
        margin: 18px 0;
    }}

    .ai-box h3 {{
        color: #ffffff;
        margin-top: 0;
    }}

    .scorecard-title {{
        font-size: 28px;
        font-weight: 800;
        margin-bottom: 8px;
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

    .dataframe tbody tr th {{
        display: none;
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


def get_market_data(data, market, year, selected_oems=None):
    df = data[(data["Market"] == market) & (data["Year"] == year)].copy()
    if selected_oems:
        df = df[df["OEM"].isin(selected_oems)]
    return df


def calculate_scorecard(data, market, year, selected_oems=None):
    df = get_market_data(data, market, year, selected_oems)
    if df.empty:
        return df

    score = df.copy()
    score["Visits to Sale"] = score["UniqueVisitors"] / score["Sales"]
    score["Conv Ranking"] = score["ConversionPct"].rank(method="first", ascending=False).astype(int)
    score = score.sort_values("Conv Ranking")

    score = score.rename(
        columns={
            "OEM": "Brand",
            "UniqueVisitors": "Website Visitors",
            "Sales": "Passenger Sales",
            "ConversionPct": "Conv Rate",
        }
    )

    return score[["Brand", "Website Visitors", "Passenger Sales", "Conv Rate", "Visits to Sale", "Conv Ranking"]]


def rank_badge(rank):
    if rank == 1:
        return "🥇"
    if rank == 2:
        return "🥈"
    if rank == 3:
        return "🥉"
    return str(rank)


def style_scorecard(scorecard):
    display = scorecard.copy()
    display["Website Visitors"] = display["Website Visitors"].map(lambda x: f"{x:,.0f}")
    display["Passenger Sales"] = display["Passenger Sales"].map(lambda x: f"{x:,.0f}")
    display["Conv Rate"] = display["Conv Rate"].map(lambda x: f"{x:.2f}%")
    display["Visits to Sale"] = display["Visits to Sale"].map(lambda x: f"{x:,.0f}")
    display["Conv Ranking"] = display["Conv Ranking"].map(rank_badge)

    return (
        display.style
        .set_table_styles([
            {
                "selector": "thead th",
                "props": [
                    ("background-color", GREEN_HEADER),
                    ("color", "black"),
                    ("font-weight", "800"),
                    ("font-size", "18px"),
                    ("text-align", "center"),
                    ("border", "1px solid black"),
                ],
            },
            {
                "selector": "tbody td",
                "props": [
                    ("font-size", "16px"),
                    ("text-align", "center"),
                    ("border", "1px solid #333333"),
                ],
            },
            {
                "selector": "tbody tr",
                "props": [
                    ("height", "46px"),
                ],
            },
        ])
    )


def generate_ai_commentary(data, market, selected_oems):
    latest = get_market_data(data, market, 2025, selected_oems)
    previous = get_market_data(data, market, 2024, selected_oems)

    if latest.empty:
        return ["No data is available for the selected view."]

    total_sales = latest["Sales"].sum()
    total_uv = latest["UniqueVisitors"].sum()
    conv_2025 = total_sales / total_uv * 100 if total_uv else 0

    prev_sales = previous["Sales"].sum() if not previous.empty else 0
    prev_uv = previous["UniqueVisitors"].sum() if not previous.empty else 0
    conv_2024 = prev_sales / prev_uv * 100 if prev_uv else 0

    latest_sorted = latest.sort_values("ConversionPct", ascending=False)
    top = latest_sorted.iloc[0]
    low = latest_sorted.iloc[-1]
    biggest_sales = latest.sort_values("Sales", ascending=False).iloc[0]
    biggest_traffic = latest.sort_values("UniqueVisitors", ascending=False).iloc[0]

    common = latest.merge(previous, on=["OEM", "Market"], suffixes=("_2025", "_2024"))
    if not common.empty:
        common["ConvChange"] = common["ConversionPct_2025"] - common["ConversionPct_2024"]
        common["SalesChangePct"] = (common["Sales_2025"] / common["Sales_2024"] - 1) * 100
        common["TrafficChangePct"] = (common["UniqueVisitors_2025"] / common["UniqueVisitors_2024"] - 1) * 100
        biggest_improver = common.sort_values("ConvChange", ascending=False).iloc[0]
        biggest_decliner = common.sort_values("ConvChange", ascending=True).iloc[0]
    else:
        biggest_improver = None
        biggest_decliner = None

    direction = "improved" if conv_2025 > conv_2024 else "weakened"
    delta = conv_2025 - conv_2024

    bullets = [
        f"In {market}, selected OEM conversion {direction} from {conv_2024:.2f}% in 2024 to {conv_2025:.2f}% in 2025 ({delta:+.2f}pp).",
        f"{top['OEM']} leads the selected set on conversion at {top['ConversionPct']:.2f}%, while {low['OEM']} is lowest at {low['ConversionPct']:.2f}%.",
        f"{biggest_sales['OEM']} delivers the largest passenger sales volume in the selected set ({biggest_sales['Sales']:,.0f}), while {biggest_traffic['OEM']} has the highest website visitor scale ({biggest_traffic['UniqueVisitors']:,.0f}).",
    ]

    if biggest_improver is not None:
        bullets.append(
            f"The sharpest conversion improvement is {biggest_improver['OEM']} ({biggest_improver['ConvChange']:+.2f}pp), while the biggest decline is {biggest_decliner['OEM']} ({biggest_decliner['ConvChange']:+.2f}pp)."
        )

    # Hard-nosed guidance
    if delta < 0:
        bullets.append("The uncomfortable read: traffic is not automatically translating into sales. Optimisation should focus on conversion quality, not just visitor growth.")
    else:
        bullets.append("The positive read: conversion is moving in the right direction, but the spread between best and worst performers shows there is still material headroom.")

    return bullets


def render_hero():
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


def render_data_definitions():
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


def render_exec_summary(data, market, selected_oems):
    st.subheader("Exec summary")

    latest = get_market_data(data, market, 2025, selected_oems)
    previous = get_market_data(data, market, 2024, selected_oems)

    if latest.empty:
        st.warning("No data available for this selection.")
        return

    total_sales = latest["Sales"].sum()
    total_uv = latest["UniqueVisitors"].sum()
    conv_2025 = total_sales / total_uv * 100 if total_uv else 0

    prev_sales = previous["Sales"].sum() if not previous.empty else 0
    prev_uv = previous["UniqueVisitors"].sum() if not previous.empty else 0
    conv_2024 = prev_sales / prev_uv * 100 if prev_uv else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("2025 passenger sales", f"{total_sales:,.0f}")
    c2.metric("2025 unique visitors", f"{total_uv:,.0f}")
    c3.metric("2025 conversion", f"{conv_2025:.2f}%")
    c4.metric("Conversion change", f"{conv_2025 - conv_2024:+.2f}pp")

    commentary = generate_ai_commentary(data, market, selected_oems)

    st.markdown(
        """
        <div class="ai-box">
            <h3>AI commentary</h3>
        """,
        unsafe_allow_html=True,
    )
    for item in commentary:
        st.markdown(f"- {item}")
    st.markdown("</div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        score = calculate_scorecard(data, market, 2025, selected_oems)
        if not score.empty:
            st.markdown("### 2025 conversion leaders")
            st.dataframe(
                score.head(5)[["Brand", "Conv Rate", "Visits to Sale", "Conv Ranking"]]
                .assign(**{"Conv Rate": lambda x: x["Conv Rate"].map(lambda v: f"{v:.2f}%")}),
                use_container_width=True,
                hide_index=True,
            )

    with col2:
        if not previous.empty:
            trend = latest.merge(previous, on=["OEM", "Market"], suffixes=("_2025", "_2024"))
            if not trend.empty:
                trend["Conversion change"] = trend["ConversionPct_2025"] - trend["ConversionPct_2024"]
                trend = trend.sort_values("Conversion change", ascending=False)
                st.markdown("### Biggest conversion movers")
                mover_display = trend[["OEM", "Conversion change"]].head(5).copy()
                mover_display["Conversion change"] = mover_display["Conversion change"].map(lambda v: f"{v:+.2f}pp")
                st.dataframe(mover_display, use_container_width=True, hide_index=True)


def render_bubble_page(data, selected_oems, year_view, show_logos):
    st.subheader("Bubble chart")

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
            chart_data = data[(data["Market"] == market) & (data["OEM"].isin(selected_oems))].copy()

            if year_view != "2024 and 2025 + shift":
                chart_data = chart_data[chart_data["Year"] == int(year_view)]

            if chart_data.empty:
                st.warning(f"No data available for {market} with the current OEM/year selection.")
                continue

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


def render_scorecard_page(data, selected_oems):
    st.subheader("Scorecard")

    col1, col2 = st.columns([1, 1])
    with col1:
        market = st.selectbox("Scorecard market", ["MM5", "UK", "France", "Germany", "Italy", "Spain"], index=0)
    with col2:
        year = st.selectbox("Scorecard year", [2025, 2024], index=0)

    scorecard = calculate_scorecard(data, market, year, selected_oems)

    if scorecard.empty:
        st.warning("No scorecard data available for this selection.")
        return

    st.markdown(f'<div class="scorecard-title">{market} {year} conversion scorecard</div>', unsafe_allow_html=True)
    st.caption("Ranking is based on conversion rate. Visits to sale = unique visitors divided by passenger sales.")

    st.dataframe(style_scorecard(scorecard), use_container_width=True, hide_index=True)

    csv = scorecard.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download scorecard as CSV",
        data=csv,
        file_name=f"{market.lower()}_{year}_conversion_scorecard.csv",
        mime="text/csv",
    )


render_hero()
render_data_definitions()

data = load_data()
available_oems = sorted(data["OEM"].unique())

st.sidebar.header("Filters")

page = st.sidebar.radio(
    "Dashboard page",
    ["Exec summary", "Bubble chart", "Scorecard"],
    index=0,
)

summary_market = st.sidebar.selectbox(
    "Exec summary market",
    ["MM5", "UK", "France", "Germany", "Italy", "Spain"],
    index=0,
)

year_view = st.sidebar.selectbox(
    "Bubble year view",
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

if page == "Exec summary":
    render_exec_summary(data, summary_market, selected_oems)
elif page == "Bubble chart":
    render_bubble_page(data, selected_oems, year_view, show_logos)
else:
    render_scorecard_page(data, selected_oems)
