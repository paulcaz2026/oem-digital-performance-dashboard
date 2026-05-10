import math
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="OEM Omnichannel Conversion Funnel",
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
OPPORTUNITY = "#12C76B"
RISK = "#FFB000"
INTELLIGENCE = "#2563EB"
PINK = "#FF5C8A"


TOYOTA_SET = [
    "Toyota", "VW", "Ford", "Peugeot", "Renault",
    "Hyundai", "Kia", "Nissan", "Skoda", "Seat", "Dacia"
]

LEXUS_SET = [
    "Lexus", "BMW", "Mercedes-Benz", "Audi",
    "Volvo", "Tesla", "Jaguar", "Land Rover", "Porsche", "Polestar"
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

    .section-kicker {{
        color: #8D96A0;
        font-size: 13px;
        font-weight: 800;
        letter-spacing: .32em;
        text-transform: uppercase;
        margin: 22px 0 18px;
        display: flex;
        align-items: center;
        gap: 14px;
    }}

    .section-kicker::after {{
        content: "";
        height: 1px;
        background: #E2E6EA;
        flex: 1;
    }}

    .insight-card {{
        background: #ffffff;
        border: 1px solid #e6e9ed;
        border-radius: 14px;
        padding: 20px 22px;
        min-height: 235px;
        box-shadow: 0 1px 8px rgba(0,0,0,.035);
        position: relative;
        overflow: hidden;
    }}

    .insight-card.opportunity {{
        border-left: 5px solid {OPPORTUNITY};
    }}

    .insight-card.risk {{
        border-left: 5px solid {RISK};
    }}

    .insight-card.intelligence {{
        border-left: 5px solid {INTELLIGENCE};
    }}

    .insight-label {{
        font-size: 12px;
        letter-spacing: .22em;
        text-transform: uppercase;
        font-weight: 800;
        margin-bottom: 12px;
    }}

    .insight-label.opportunity {{
        color: {OPPORTUNITY};
    }}

    .insight-label.risk {{
        color: {RISK};
    }}

    .insight-label.intelligence {{
        color: {INTELLIGENCE};
    }}

    .insight-dot {{
        display: inline-block;
        width: 12px;
        height: 12px;
        border-radius: 50%;
        margin-right: 8px;
        vertical-align: -1px;
        box-shadow: inset 0 0 3px rgba(0,0,0,.25);
    }}

    .dot-opportunity {{
        background: linear-gradient(180deg, #4ce36d, #00a32a);
    }}

    .dot-risk {{
        background: linear-gradient(180deg, #ffe45c, #ffb000);
    }}

    .dot-intelligence {{
        background: linear-gradient(180deg, #3aa0ff, #005bd3);
    }}

    .insight-title {{
        font-size: 18px;
        line-height: 1.25;
        font-weight: 800;
        color: #0A2342;
        margin-bottom: 10px;
    }}

    .insight-copy {{
        color: #8B95A1;
        font-size: 14px;
        line-height: 1.55;
        min-height: 78px;
    }}

    .insight-metric {{
        color: #0A2342;
        font-size: 27px;
        margin-top: 16px;
        font-weight: 500;
    }}

    .tag {{
        display: inline-block;
        background: #EEF3FF;
        color: #2563EB;
        border-radius: 5px;
        padding: 4px 10px;
        font-size: 12px;
        margin-top: 10px;
    }}

    .scorecard-title {{
        font-size: 28px;
        font-weight: 800;
        margin-bottom: 8px;
    }}

    div[data-testid="stPlotlyChart"] {{
        background: #ffffff;
        border: 1px solid #e6e9ed;
        border-radius: 14px;
        padding: 10px;
        box-shadow: 0 1px 8px rgba(0,0,0,.035);
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

    .top-table-wrap {{
        background: #ffffff;
        border: 1px solid #e6e9ed;
        border-radius: 14px;
        padding: 18px 22px;
        box-shadow: 0 1px 8px rgba(0,0,0,.035);
        margin: 16px 0 22px 0;
    }}

    .top-table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
    }}

    .top-table thead th {{
        background: #f5f6f8;
        color: #8D96A0;
        text-transform: uppercase;
        letter-spacing: .18em;
        font-size: 11px;
        padding: 13px 14px;
        text-align: left;
        border-bottom: 1px solid #e1e5ea;
    }}

    .top-table tbody td {{
        padding: 13px 14px;
        border-bottom: 1px solid #edf0f2;
        color: #0A2342;
    }}

    .top-table tbody tr:nth-child(1) {{
        background: #f7f9fc;
        font-weight: 700;
    }}

    .rank-cell {{
        color: #8D96A0;
        width: 50px;
    }}

    .brand-dot {{
        display: inline-block;
        width: 9px;
        height: 9px;
        border-radius: 50%;
        margin-right: 8px;
        background: #2563EB;
    }}

    .badge-pos {{
        background: #DDF8EC;
        color: #12C76B;
        padding: 5px 9px;
        border-radius: 6px;
        font-weight: 700;
        display: inline-block;
    }}

    .badge-neg {{
        background: #FFE5EF;
        color: #FF2F6D;
        padding: 5px 9px;
        border-radius: 6px;
        font-weight: 700;
        display: inline-block;
    }}

    .badge-neutral {{
        background: #EEF2F6;
        color: #6F6F6F;
        padding: 5px 9px;
        border-radius: 6px;
        font-weight: 700;
        display: inline-block;
    }}

    .benchmark-card {{
        background: #ffffff;
        border: 1px solid #e6e9ed;
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 1px 8px rgba(0,0,0,.035);
        min-height: 155px;
    }}

    .benchmark-title {{
        font-size: 17px;
        font-weight: 800;
        color: #0A2342;
        margin-bottom: 8px;
    }}

    .benchmark-copy {{
        color: #6F7782;
        line-height: 1.45;
        font-size: 14px;
    }}

    .benchmark-metric {{
        font-size: 26px;
        color: #0A2342;
        margin-top: 10px;
    }}

    .weakness-table-wrap {{
        background: #ffffff;
        border: 1px solid #e6e9ed;
        border-radius: 14px;
        padding: 18px 22px;
        box-shadow: 0 1px 8px rgba(0,0,0,.035);
        margin: 16px 0 22px 0;
    }}

    .weakness-table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
    }}

    .weakness-table thead th {{
        background: #f5f6f8;
        color: #8D96A0;
        text-transform: uppercase;
        letter-spacing: .14em;
        font-size: 11px;
        padding: 13px 14px;
        text-align: left;
        border-bottom: 1px solid #e1e5ea;
    }}

    .weakness-table tbody td {{
        padding: 13px 14px;
        border-bottom: 1px solid #edf0f2;
        color: #0A2342;
    }}

    .freshness {{
        display: inline-block;
        background: #F3F3F3;
        border-left: 5px solid #009FE3;
        padding: 10px 14px;
        border-radius: 10px;
        font-size: 13px;
        color: #0A2342;
        margin-bottom: 18px;
    }}


    .market-flag {{
        display: inline-block;
        background: #F3F7FF;
        color: #005F9E;
        border: 1px solid #CFE8FF;
        border-radius: 999px;
        padding: 4px 10px;
        font-size: 12px;
        font-weight: 800;
        margin-bottom: 10px;
    }}

    .footer-note {{
        margin-top: 36px;
        padding: 18px 20px;
        background: #F3F3F3;
        border-left: 6px solid #009FE3;
        border-radius: 12px;
        font-size: 13px;
        color: #0A2342;
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


def get_market_data(data, market, year, selected_oems=None):
    df = data[(data["Market"] == market) & (data["Year"] == year)].copy()
    if selected_oems:
        df = df[df["OEM"].isin(selected_oems)]
    return df


def get_yoy_table(data, market, selected_oems=None):
    d24 = get_market_data(data, market, 2024, selected_oems)
    d25 = get_market_data(data, market, 2025, selected_oems)

    merged = d25.merge(d24, on=["OEM", "Market"], suffixes=("_2025", "_2024"))

    if merged.empty:
        return merged

    merged["Sales YoY %"] = (merged["Sales_2025"] / merged["Sales_2024"] - 1) * 100
    merged["Visitors YoY %"] = (merged["UniqueVisitors_2025"] / merged["UniqueVisitors_2024"] - 1) * 100
    merged["Conv Var pp"] = merged["ConversionPct_2025"] - merged["ConversionPct_2024"]
    merged["Visits to Sale 2024"] = merged["UniqueVisitors_2024"] / merged["Sales_2024"]
    merged["Visits to Sale 2025"] = merged["UniqueVisitors_2025"] / merged["Sales_2025"]
    merged["Visits to Sale Var"] = merged["Visits to Sale 2025"] - merged["Visits to Sale 2024"]

    return merged


def calculate_scorecard(data, market, selected_oems=None):
    merged = get_yoy_table(data, market, selected_oems)

    if merged.empty:
        return merged

    score = merged.copy()
    score["Conv Ranking"] = score["ConversionPct_2025"].rank(method="first", ascending=False).astype(int)
    score = score.sort_values("Conv Ranking")

    score = score.rename(
        columns={
            "OEM": "Brand",
            "UniqueVisitors_2024": "Website Visitors 2024",
            "UniqueVisitors_2025": "Website Visitors 2025",
            "Sales_2024": "Passenger Sales 2024",
            "Sales_2025": "Passenger Sales 2025",
            "ConversionPct_2024": "Conv Rate 2024",
            "ConversionPct_2025": "Conv Rate 2025",
        }
    )

    return score[
        [
            "Brand",
            "Website Visitors 2024",
            "Website Visitors 2025",
            "Visitors YoY %",
            "Passenger Sales 2024",
            "Passenger Sales 2025",
            "Sales YoY %",
            "Conv Rate 2024",
            "Conv Rate 2025",
            "Conv Var pp",
            "Visits to Sale 2024",
            "Visits to Sale 2025",
            "Visits to Sale Var",
            "Conv Ranking",
        ]
    ]


def rank_badge(rank):
    if rank == 1:
        return "🥇"
    if rank == 2:
        return "🥈"
    if rank == 3:
        return "🥉"
    return str(rank)


def fmt_pct_delta(v):
    if pd.isna(v):
        return "n/a"
    return f"{v:+.1f}%"


def fmt_pp(v):
    if pd.isna(v):
        return "n/a"
    return f"{v:+.2f}pp"


def style_scorecard(scorecard):
    raw = scorecard.copy()
    display = scorecard.copy()

    for col in ["Website Visitors 2024", "Website Visitors 2025", "Passenger Sales 2024", "Passenger Sales 2025"]:
        display[col] = display[col].map(lambda x: f"{x:,.0f}")

    for col in ["Visitors YoY %", "Sales YoY %"]:
        display[col] = display[col].map(fmt_pct_delta)

    for col in ["Conv Rate 2024", "Conv Rate 2025"]:
        display[col] = display[col].map(lambda x: f"{x:.2f}%")

    display["Conv Var pp"] = display["Conv Var pp"].map(fmt_pp)

    for col in ["Visits to Sale 2024", "Visits to Sale 2025", "Visits to Sale Var"]:
        display[col] = display[col].map(lambda x: f"{x:,.0f}")

    display["Conv Ranking"] = display["Conv Ranking"].map(rank_badge)

    yoy_cols = ["Visitors YoY %", "Sales YoY %", "Conv Var pp"]

    def yoy_badge_style(value):
        text = str(value)
        if text.startswith("+"):
            return (
                "background-color: #DDF8EC; "
                "color: #12C76B; "
                "font-weight: 700; "
                "border-radius: 7px; "
                "text-align: center;"
            )
        if text.startswith("-"):
            return (
                "background-color: #FFE5EF; "
                "color: #FF2F6D; "
                "font-weight: 700; "
                "border-radius: 7px; "
                "text-align: center;"
            )
        return (
            "background-color: #EEF2F6; "
            "color: #6F6F6F; "
            "font-weight: 700; "
            "border-radius: 7px; "
            "text-align: center;"
        )

    styler = (
        display.style
        .set_table_styles([
            {
                "selector": "thead th",
                "props": [
                    ("background-color", GREEN_HEADER),
                    ("color", "black"),
                    ("font-weight", "800"),
                    ("font-size", "14px"),
                    ("text-align", "center"),
                    ("border", "1px solid black"),
                    ("white-space", "normal"),
                ],
            },
            {
                "selector": "tbody td",
                "props": [
                    ("font-size", "13px"),
                    ("text-align", "center"),
                    ("border", "1px solid #333333"),
                    ("vertical-align", "middle"),
                ],
            },
            {
                "selector": "tbody tr",
                "props": [
                    ("height", "42px"),
                ],
            },
        ])
        .map(yoy_badge_style, subset=yoy_cols)
    )

    return styler


def get_row(yoy, brand):
    row = yoy[yoy["OEM"].str.lower() == brand.lower()]
    if row.empty:
        return None
    return row.iloc[0]


def pick_brand(yoy, candidates):
    for brand in candidates:
        row = get_row(yoy, brand)
        if row is not None:
            return row
    return None



def make_market_insight(kind, market, title, copy, metric, tag):
    dot_class = {
        "opportunity": "dot-opportunity",
        "risk": "dot-risk",
        "intelligence": "dot-intelligence",
    }[kind]
    label = {
        "opportunity": "Opportunity",
        "risk": "Risk Alert",
        "intelligence": "Intelligence",
    }[kind]

    return f"""
    <div class="insight-card {kind}">
        <div class="market-flag">{market}</div>
        <div class="insight-label {kind}"><span class="insight-dot {dot_class}"></span>{label}</div>
        <div class="insight-title">{title}</div>
        <div class="insight-copy">{copy}</div>
        <div class="insight-metric">{metric}</div>
        <div class="tag">{tag}</div>
    </div>
    """


def get_market_action_narrative(brand, market, row, leader, gap):
    sales_yoy = row["Sales YoY %"]
    visitor_yoy = row["Visitors YoY %"]
    conv_var = row["Conv Var pp"]

    if market == "Germany":
        return (
            f"Germany is the toughest conversion gap for {brand}. {leader['OEM']} sets the benchmark at {leader['ConversionPct_2025']:.2f}% versus {brand} at {row['ConversionPct_2025']:.2f}%. "
            "Recommendation: prioritise retailer hand-off quality, availability messaging and high-intent model pages; Germany needs fewer generic visits and more purchase-ready journeys."
        )

    if market == "UK":
        return (
            f"In the UK, {brand} is losing efficiency against {leader['OEM']} with a {gap:+.2f}pp conversion gap. "
            "Recommendation: tighten the lower funnel around enquiry forms, test-drive booking, used/new stock routing and finance CTAs; this is where visitor intent is most likely leaking."
        )

    if market == "France":
        return (
            f"France shows a conversion gap of {gap:+.2f}pp for {brand} versus {leader['OEM']}. "
            "Recommendation: localise offer visibility and dealer follow-up prompts; the opportunity is to turn browsing into configured, dealer-ready demand."
        )

    if market == "Italy":
        return (
            f"In Italy, {brand} conversion is {row['ConversionPct_2025']:.2f}% with sales YoY at {sales_yoy:+.1f}% and visitor YoY at {visitor_yoy:+.1f}%. "
            "Recommendation: focus on reducing friction between model discovery and contact request; Italy looks like a journey-efficiency problem rather than a pure awareness problem."
        )

    if market == "Spain":
        return (
            f"Spain is relatively more resilient, but {brand} still trails {leader['OEM']} by {gap:+.2f}pp. "
            "Recommendation: protect momentum by scaling the highest-converting traffic sources and making offer, stock and lead response paths more prominent."
        )

    if visitor_yoy > sales_yoy and conv_var < 0:
        return (
            f"{brand} is attracting more traffic than sales growth justifies. Visitor YoY is {visitor_yoy:+.1f}% versus sales YoY of {sales_yoy:+.1f}%. "
            "Recommendation: audit traffic quality and landing-page intent; growth is only useful if it moves customers closer to contract."
        )

    return (
        f"{brand} has a {gap:+.2f}pp gap to {leader['OEM']} and conversion moved {conv_var:+.2f}pp YoY. "
        "Recommendation: prioritise the highest-intent steps in the contract funnel and remove friction from enquiry, finance and stock discovery."
    )


def generate_toyota_lexus_market_cards(data):
    cards = []
    for market in ["UK", "France", "Germany", "Italy", "Spain"]:
        yoy = get_yoy_table(data, market, None)
        if yoy.empty:
            continue

        for brand, cohort in [("Toyota", TOYOTA_SET), ("Lexus", LEXUS_SET)]:
            row = get_row(yoy, brand)
            cohort_df = yoy[yoy["OEM"].isin(cohort)].copy()
            if row is None or cohort_df.empty:
                continue

            leader = cohort_df.sort_values("ConversionPct_2025", ascending=False).iloc[0]
            gap = row["ConversionPct_2025"] - leader["ConversionPct_2025"]

            severity = abs(gap) + (0.5 if row["Sales YoY %"] < 0 else 0) + (0.25 if row["Visitors YoY %"] > row["Sales YoY %"] else 0)
            kind = "risk" if gap < -0.75 or row["Sales YoY %"] < 0 else "opportunity"

            title = f"{brand}: {market} conversion gap to {leader['OEM']}"
            copy = get_market_action_narrative(brand, market, row, leader, gap)
            metric = f"{gap:+.2f}pp gap"
            cards.append((severity, make_market_insight(kind, market, title, copy, metric, f"{brand} recommendation")))

    cards = sorted(cards, key=lambda x: x[0], reverse=True)
    return [card for _, card in cards[:8]]




def make_insight(kind, title, copy, metric, tag):
    dot_class = {
        "opportunity": "dot-opportunity",
        "risk": "dot-risk",
        "intelligence": "dot-intelligence",
    }[kind]
    label = {
        "opportunity": "Opportunity",
        "risk": "Risk Alert",
        "intelligence": "Intelligence",
    }[kind]

    return f"""
    <div class="insight-card {kind}">
        <div class="insight-label {kind}"><span class="insight-dot {dot_class}"></span>{label}</div>
        <div class="insight-title">{title}</div>
        <div class="insight-copy">{copy}</div>
        <div class="insight-metric">{metric}</div>
        <div class="tag">{tag}</div>
    </div>
    """


def generate_insight_cards(data, market, selected_oems):
    yoy = get_yoy_table(data, market, selected_oems)

    if yoy.empty:
        return []

    cards = []

    yoy["Sales YoY % safe"] = yoy["Sales YoY %"].replace([float("inf"), float("-inf")], pd.NA)
    yoy["Visitors YoY % safe"] = yoy["Visitors YoY %"].replace([float("inf"), float("-inf")], pd.NA)
    yoy["Sales Efficiency Gap"] = yoy["Sales YoY % safe"] - yoy["Visitors YoY % safe"]

    top_conv = yoy.sort_values("ConversionPct_2025", ascending=False).iloc[0]
    top_sales_growth = yoy.sort_values("Sales YoY % safe", ascending=False).iloc[0]
    top_traffic_growth = yoy.sort_values("Visitors YoY % safe", ascending=False).iloc[0]
    worst_conv_decline = yoy.sort_values("Conv Var pp", ascending=True).iloc[0]
    weakest_conv = yoy.sort_values("ConversionPct_2025", ascending=True).iloc[0]

    toyota = pick_brand(yoy, ["Toyota"])
    lexus = pick_brand(yoy, ["Lexus"])

    if toyota is not None:
        cards.append(make_insight(
            "opportunity" if toyota["Conv Var pp"] >= 0 else "risk",
            "Toyota: conversion is the priority lever",
            f"Toyota delivered {toyota['Sales_2025']:,.0f} passenger sales in 2025 with {toyota['UniqueVisitors_2025']:,.0f} unique visitors. Conversion moved {toyota['Conv Var pp']:+.2f}pp YoY, while visitor growth was {toyota['Visitors YoY %']:+.1f}%. The recommendation is to focus on journey quality and model-level intent capture, not simply more traffic.",
            f"{toyota['ConversionPct_2025']:.2f}% conv",
            "Toyota recommendation"
        ))

    if lexus is not None:
        cards.append(make_insight(
            "risk" if lexus["ConversionPct_2025"] < yoy["ConversionPct_2025"].median() else "opportunity",
            "Lexus: premium consideration is not converting hard enough",
            f"Lexus generated {lexus['UniqueVisitors_2025']:,.0f} visitors and {lexus['Sales_2025']:,.0f} sales in 2025. Conversion is {lexus['ConversionPct_2025']:.2f}%, with a YoY movement of {lexus['Conv Var pp']:+.2f}pp. The recommendation is sharper premium lead handling, finance/stock visibility and stronger lower-funnel retail prompts.",
            f"{lexus['Sales YoY %']:+.1f}% sales YoY",
            "Lexus recommendation"
        ))

    cards.append(make_insight(
        "opportunity",
        f"{top_conv['OEM']}: strongest conversion performer",
        f"{top_conv['OEM']} leads the selected cohort on 2025 conversion. Its visits-to-sale ratio is {top_conv['Visits to Sale 2025']:.0f}, meaning it needs fewer visitors per passenger sale than competitors. Use this as the benchmark for digital retail efficiency.",
        f"{top_conv['ConversionPct_2025']:.2f}% conv",
        top_conv["OEM"]
    ))

    cards.append(make_insight(
        "intelligence",
        f"{top_traffic_growth['OEM']}: fastest visitor growth",
        f"{top_traffic_growth['OEM']} had the strongest visitor growth in the selected cohort. Visitor growth was {top_traffic_growth['Visitors YoY %']:+.1f}% YoY, while sales growth was {top_traffic_growth['Sales YoY %']:+.1f}%. This shows whether awareness is converting into actual demand.",
        f"{top_traffic_growth['Visitors YoY %']:+.1f}% visits YoY",
        top_traffic_growth["OEM"]
    ))

    cards.append(make_insight(
        "opportunity",
        f"{top_sales_growth['OEM']}: strongest sales growth",
        f"{top_sales_growth['OEM']} posted the strongest passenger sales growth in the selected cohort. Sales increased {top_sales_growth['Sales YoY %']:+.1f}% YoY against visitor growth of {top_sales_growth['Visitors YoY %']:+.1f}%. This is the clearest signal of demand momentum.",
        f"{top_sales_growth['Sales YoY %']:+.1f}% sales YoY",
        top_sales_growth["OEM"]
    ))

    cards.append(make_insight(
        "risk",
        f"{worst_conv_decline['OEM']}: conversion quality deteriorating",
        f"{worst_conv_decline['OEM']} saw the weakest conversion movement in the selected cohort. Conversion moved {worst_conv_decline['Conv Var pp']:+.2f}pp YoY. If visitor growth is positive but conversion falls, the brand is likely attracting lower-quality traffic or losing users deeper in the funnel.",
        f"{worst_conv_decline['Conv Var pp']:+.2f}pp conv",
        worst_conv_decline["OEM"]
    ))

    cards.append(make_insight(
        "risk",
        f"{weakest_conv['OEM']}: weakest 2025 conversion",
        f"{weakest_conv['OEM']} has the lowest 2025 conversion rate in the selected cohort. Visits-to-sale is {weakest_conv['Visits to Sale 2025']:.0f}, creating a clear efficiency gap versus the cohort leader.",
        f"{weakest_conv['ConversionPct_2025']:.2f}% conv",
        weakest_conv["OEM"]
    ))

    # Deduplicate by title and cap at 9 cards
    seen = set()
    unique_cards = []
    for card in cards:
        title_marker_start = card.find('<div class="insight-title">')
        title_marker_end = card.find("</div>", title_marker_start)
        title = card[title_marker_start:title_marker_end]
        if title not in seen:
            unique_cards.append(card)
            seen.add(title)
    return unique_cards[:9]


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



def render_footer_notes():
    st.markdown(
        """
        <div class="footer-note">
            <b>Definitions and caveats:</b> This is not a total market-share dashboard. It uses passenger car sales from Marklines and unique visitor website data from Similarweb. It does not include fleet, LCV or tactical registrations. Conversion rate = passenger car sales divided by unique visitors. Data period: 2024–2025.
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_hero():
    st.markdown(
        """
        <div class="hero">
            <div class="hero-logo">
                <img src="https://mma.prnewswire.com/media/2728124/Valtech_Logo.jpg" alt="Valtech logo">
            </div>
            <div class="hero-title">OEM Omnichannel Conversion Funnel</div>
            <div class="hero-subtitle">
                Website to customer contract conversion across MM5. Explores how unique visitors demand converts into passenger new car customer contracts, and where Toyota/Lexus under- or over-perform by market.
            </div>
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



def build_yoy_growth_chart(data, market, selected_oems, metric="Visitors YoY %", top_n=6):
    yoy = get_yoy_table(data, market, selected_oems)
    fig = go.Figure()

    if yoy.empty or metric not in yoy.columns:
        return fig

    yoy = yoy.replace([float("inf"), float("-inf")], pd.NA).dropna(subset=[metric])

    if yoy.empty:
        return fig

    top = yoy.sort_values(metric, ascending=False).head(top_n)
    bottom = yoy.sort_values(metric, ascending=True).head(top_n)
    chart_df = pd.concat([top, bottom]).drop_duplicates(subset=["OEM"])
    chart_df = chart_df.sort_values(metric, ascending=True)

    colors = chart_df[metric].apply(lambda v: "#5ED6AC" if v >= 0 else "#FF5C8A")

    fig.add_trace(
        go.Bar(
            x=chart_df[metric],
            y=chart_df["OEM"],
            orientation="h",
            marker=dict(
                color=colors,
                line=dict(color=colors, width=1.4),
            ),
            text=chart_df[metric].map(lambda v: f"{v:+.1f}%"),
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>YoY growth: %{x:+.1f}%<extra></extra>",
        )
    )

    fig.update_layout(
        title=dict(
            text=f"YoY visit growth — 2025 vs 2024 (top & bottom {top_n})",
            x=0.02,
            font=dict(size=15, color="#8D96A0"),
        ),
        height=560,
        margin=dict(l=120, r=70, t=70, b=45),
        plot_bgcolor=WHITE,
        paper_bgcolor=WHITE,
        font=dict(family="Helvetica Neue, Arial, sans-serif", color="#0A2342"),
        showlegend=False,
    )

    min_x = min(chart_df[metric].min() * 1.25, -5)
    max_x = max(chart_df[metric].max() * 1.20, 5)

    fig.update_xaxes(
        range=[min_x, max_x],
        ticksuffix="%",
        gridcolor="#E2E6EA",
        zeroline=True,
        zerolinecolor="#B8C0C8",
        linecolor="#E2E6EA",
    )
    fig.update_yaxes(
        gridcolor="#FFFFFF",
        linecolor="#E2E6EA",
    )

    return fig


def build_visit_volume_chart(data, market, selected_oems, top_n=12):
    yoy = get_yoy_table(data, market, selected_oems)
    fig = go.Figure()

    if yoy.empty:
        return fig

    chart_df = yoy.sort_values("UniqueVisitors_2025", ascending=False).head(top_n)
    chart_df = chart_df.sort_values("UniqueVisitors_2025", ascending=True)

    fig.add_trace(
        go.Bar(
            x=chart_df["UniqueVisitors_2024"],
            y=chart_df["OEM"],
            orientation="h",
            name="2024",
            marker=dict(color="#DDE3EA", line=dict(color="#9AA5B1", width=1.2)),
            hovertemplate="<b>%{y}</b><br>2024 visits: %{x:,.0f}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Bar(
            x=chart_df["UniqueVisitors_2025"],
            y=chart_df["OEM"],
            orientation="h",
            name="2025",
            marker=dict(color=VALTECH_BLUE, line=dict(color="#2563EB", width=1.2)),
            hovertemplate="<b>%{y}</b><br>2025 visits: %{x:,.0f}<extra></extra>",
        )
    )

    fig.update_layout(
        title=dict(
            text=f"Visit volume — 2025 vs 2024 (top {top_n})",
            x=0.02,
            font=dict(size=15, color="#8D96A0"),
        ),
        height=420,
        margin=dict(l=95, r=40, t=70, b=45),
        barmode="overlay",
        bargap=0.38,
        plot_bgcolor=WHITE,
        paper_bgcolor=WHITE,
        font=dict(family="Helvetica Neue, Arial, sans-serif", color="#0A2342"),
        legend=dict(orientation="h", y=1.08, x=0.70),
    )

    max_x = max(chart_df["UniqueVisitors_2025"].max(), chart_df["UniqueVisitors_2024"].max()) * 1.12

    fig.update_xaxes(
        range=[0, max_x],
        tickformat=",.2s",
        gridcolor="#E2E6EA",
        zeroline=False,
        linecolor="#E2E6EA",
    )
    fig.update_yaxes(
        gridcolor="#FFFFFF",
        linecolor="#E2E6EA",
    )

    return fig


def render_exec_visuals(data, market, selected_oems):
    st.plotly_chart(
        build_yoy_growth_chart(data, market, selected_oems, "Visitors YoY %", top_n=10),
        use_container_width=True,
    )



def render_exec_summary(data, market, selected_oems):
    st.markdown(f'<div class="section-kicker">Executive insights — data-driven narratives from the {market} OEM cohort</div>', unsafe_allow_html=True)

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

    sales_yoy = (total_sales / prev_sales - 1) * 100 if prev_sales else 0
    uv_yoy = (total_uv / prev_uv - 1) * 100 if prev_uv else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("2025 passenger sales", f"{total_sales:,.0f}", f"{sales_yoy:+.1f}% YoY")
    c2.metric("2025 unique visitors", f"{total_uv:,.0f}", f"{uv_yoy:+.1f}% YoY")
    c3.metric("2025 conversion", f"{conv_2025:.2f}%", f"{conv_2025 - conv_2024:+.2f}pp")
    c4.metric("Visits to sale", f"{(total_uv / total_sales):,.0f}" if total_sales else "n/a")

    render_market_top10_inline(data, market, selected_oems)

    st.markdown('<div class="section-kicker">Performance visuals — YoY growth and visitor scale</div>', unsafe_allow_html=True)
    render_exec_visuals(data, market, selected_oems)

    render_toyota_lexus_benchmarks(data, market)

    cards = generate_insight_cards(data, market, selected_oems)

    for i in range(0, len(cards), 3):
        cols = st.columns(3)
        for col, card in zip(cols, cards[i:i + 3]):
            with col:
                st.markdown(card, unsafe_allow_html=True)

    render_market_weakness_summary(data)
    render_footer_notes()





def benchmark_card_html(title, copy, metric):
    return f"""
    <div class="benchmark-card">
        <div class="benchmark-title">{title}</div>
        <div class="benchmark-copy">{copy}</div>
        <div class="benchmark-metric">{metric}</div>
    </div>
    """


def render_toyota_lexus_benchmarks(data, market):
    yoy_all = get_yoy_table(data, market, None)
    if yoy_all.empty:
        return

    st.markdown('<div class="section-kicker">Toyota / Lexus benchmark callouts</div>', unsafe_allow_html=True)

    cards = []

    for brand, cohort in [
        ("Toyota", TOYOTA_SET),
        ("Lexus", LEXUS_SET),
    ]:
        row = get_row(yoy_all, brand)
        cohort_df = yoy_all[yoy_all["OEM"].isin(cohort)].copy()

        if row is None or cohort_df.empty:
            continue

        cohort_df["Conv Rank"] = cohort_df["ConversionPct_2025"].rank(method="min", ascending=False)
        rank_series = cohort_df.loc[cohort_df["OEM"].str.lower() == brand.lower(), "Conv Rank"]

        if rank_series.empty:
            continue

        brand_rank = int(rank_series.iloc[0])
        leader = cohort_df.sort_values("ConversionPct_2025", ascending=False).iloc[0]
        gap = row["ConversionPct_2025"] - leader["ConversionPct_2025"]
        visits_to_sale_gap = row["Visits to Sale 2025"] - leader["Visits to Sale 2025"]

        copy = (
            f"{brand} ranks #{brand_rank} of {len(cohort_df)} in its benchmark set. "
            f"The conversion gap to {leader['OEM']} is {gap:+.2f}pp. "
            f"Visits-to-sale gap is {visits_to_sale_gap:+.0f}."
        )

        metric = f"{row['ConversionPct_2025']:.2f}% conv"
        cards.append(benchmark_card_html(f"{brand} benchmark", copy, metric))

    if cards:
        cols = st.columns(len(cards))
        for col, card in zip(cols, cards):
            with col:
                st.markdown(card, unsafe_allow_html=True)





def badge_html(value, suffix="%"):
    try:
        v = float(value)
    except Exception:
        return '<span class="badge-neutral">n/a</span>'

    cls = "badge-pos" if v > 0 else "badge-neg" if v < 0 else "badge-neutral"
    sign = "+" if v > 0 else ""

    if suffix == "pp":
        text = f"{sign}{v:.2f}pp"
    else:
        text = f"{sign}{v:.1f}%"

    return f'<span class="{cls}">{text}</span>'


def render_market_weakness_summary(data):
    st.markdown('<div class="section-kicker">Market weakness summary — Toyota & Lexus</div>', unsafe_allow_html=True)

    rows = []
    for market in ["UK", "France", "Germany", "Italy", "Spain"]:
        yoy = get_yoy_table(data, market, None)
        if yoy.empty:
            continue

        for brand, cohort in [("Toyota", TOYOTA_SET), ("Lexus", LEXUS_SET)]:
            row = get_row(yoy, brand)
            cohort_df = yoy[yoy["OEM"].isin(cohort)].copy()

            if row is None or cohort_df.empty:
                continue

            leader = cohort_df.sort_values("ConversionPct_2025", ascending=False).iloc[0]
            rows.append({
                "Brand": brand,
                "Market": market,
                "2025 conversion": row["ConversionPct_2025"],
                "Benchmark leader": leader["OEM"],
                "Leader conversion": leader["ConversionPct_2025"],
                "Gap to leader": row["ConversionPct_2025"] - leader["ConversionPct_2025"],
                "Sales YoY": row["Sales YoY %"],
                "Visitor YoY": row["Visitors YoY %"],
            })

    if not rows:
        st.info("No Toyota/Lexus market weakness data available.")
        return

    summary = pd.DataFrame(rows).sort_values("Gap to leader")

    row_html = ""
    for _, r in summary.iterrows():
        gap_badge = badge_html(r["Gap to leader"], suffix="pp")
        sales_badge = badge_html(r["Sales YoY"], suffix="%")
        visitor_badge = badge_html(r["Visitor YoY"], suffix="%")

        row_html += (
            "<tr>"
            f"<td><b>{r['Brand']}</b></td>"
            f"<td>{r['Market']}</td>"
            f"<td>{r['2025 conversion']:.2f}%</td>"
            f"<td>{r['Benchmark leader']}</td>"
            f"<td>{r['Leader conversion']:.2f}%</td>"
            f"<td>{gap_badge}</td>"
            f"<td>{sales_badge}</td>"
            f"<td>{visitor_badge}</td>"
            "</tr>"
        )

    table_html = (
        "<div class='weakness-table-wrap'>"
        "<table class='weakness-table'>"
        "<thead>"
        "<tr>"
        "<th>Brand</th>"
        "<th>Market</th>"
        "<th>2025 conversion</th>"
        "<th>Benchmark leader</th>"
        "<th>Leader conversion</th>"
        "<th>Gap to leader</th>"
        "<th>Sales YoY</th>"
        "<th>Visitor YoY</th>"
        "</tr>"
        "</thead>"
        f"<tbody>{row_html}</tbody>"
        "</table>"
        "</div>"
    )

    st.markdown(table_html, unsafe_allow_html=True)



def render_gap_analysis_page(data, market, selected_oems):
    st.markdown(f'<div class="section-kicker">Toyota & Lexus gap analysis — {market}</div>', unsafe_allow_html=True)

    focus_oems = [o for o in selected_oems if o in ["Toyota", "Lexus"]]
    if not focus_oems:
        focus_oems = ["Toyota", "Lexus"]

    latest = get_market_data(data, market, 2025, focus_oems)
    previous = get_market_data(data, market, 2024, focus_oems)

    if latest.empty:
        st.warning("No Toyota/Lexus data available for this market.")
        return

    total_sales = latest["Sales"].sum()
    total_uv = latest["UniqueVisitors"].sum()
    conv_2025 = total_sales / total_uv * 100 if total_uv else 0
    prev_sales = previous["Sales"].sum() if not previous.empty else 0
    prev_uv = previous["UniqueVisitors"].sum() if not previous.empty else 0
    conv_2024 = prev_sales / prev_uv * 100 if prev_uv else 0
    sales_yoy = (total_sales / prev_sales - 1) * 100 if prev_sales else 0
    uv_yoy = (total_uv / prev_uv - 1) * 100 if prev_uv else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Toyota/Lexus sales", f"{total_sales:,.0f}", f"{sales_yoy:+.1f}% YoY")
    c2.metric("Toyota/Lexus visitors", f"{total_uv:,.0f}", f"{uv_yoy:+.1f}% YoY")
    c3.metric("Weighted conversion", f"{conv_2025:.2f}%", f"{conv_2025 - conv_2024:+.2f}pp")
    c4.metric("Visits to sale", f"{(total_uv / total_sales):,.0f}" if total_sales else "n/a")

    render_toyota_lexus_benchmarks(data, market)
    render_market_weakness_summary(data)

    st.markdown(f'<div class="section-kicker">AI insights — Toyota & Lexus market recommendations</div>', unsafe_allow_html=True)
    cards = generate_toyota_lexus_market_cards(data)
    for i in range(0, len(cards), 3):
        cols = st.columns(3)
        for col, card in zip(cols, cards[i:i + 3]):
            with col:
                st.markdown(card, unsafe_allow_html=True)

    render_footer_notes()



def fmt_short_num(value):
    try:
        v = float(value)
    except Exception:
        return "n/a"
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.2f}M"
    if abs(v) >= 1_000:
        return f"{v/1_000:.1f}K"
    return f"{v:,.0f}"


def render_market_top10_inline(data, market, selected_oems):
    yoy = get_yoy_table(data, market, selected_oems)
    if yoy.empty:
        return

    top = yoy.sort_values("UniqueVisitors_2025", ascending=False).head(10).copy()
    top["Rank"] = range(1, len(top) + 1)

    colors = ["#2563EB", "#B544F4", "#FFB000", "#FF3B77", "#FF7A1A", "#27C59A", "#27C59A", "#4C78FF", "#B544F4", "#22B8CF"]

    row_html = ""
    for idx, (_, r) in enumerate(top.iterrows()):
        row_html += (
            "<tr>"
            f"<td class='rank-cell'>{int(r['Rank'])}</td>"
            f"<td><span class='brand-dot' style='background:{colors[idx % len(colors)]};'></span><b>{r['OEM']}</b></td>"
            f"<td>{fmt_short_num(r['UniqueVisitors_2025'])}</td>"
            f"<td>{fmt_short_num(r['UniqueVisitors_2024'])}</td>"
            f"<td>{badge_html(r['Visitors YoY %'])}</td>"
            f"<td>{r['Sales_2025']:,.0f}</td>"
            f"<td>{r['ConversionPct_2025']:.2f}%</td>"
            f"<td>{r['Conv Var pp']:+.2f}pp</td>"
            "</tr>"
        )

    table_html = (
        "<div class='section-kicker'>Top 10 brands at a glance</div>"
        "<div class='top-table-wrap'>"
        "<table class='top-table'>"
        "<thead>"
        "<tr>"
        "<th>#</th>"
        "<th>Brand</th>"
        "<th>Visits 2025</th>"
        "<th>Visits 2024</th>"
        "<th>YoY</th>"
        "<th>Passenger sales</th>"
        "<th>Conv rate</th>"
        "<th>Conv var</th>"
        "</tr>"
        "</thead>"
        f"<tbody>{row_html}</tbody>"
        "</table>"
        "</div>"
    )

    st.markdown(table_html, unsafe_allow_html=True)




def render_market_top10_inline(data, market, selected_oems):
    yoy = get_yoy_table(data, market, selected_oems)
    if yoy.empty:
        st.info("No Top 10 data available for this selection.")
        return

    top = yoy.sort_values("UniqueVisitors_2025", ascending=False).head(10).copy()
    top["Rank"] = range(1, len(top) + 1)

    def short_num(value):
        v = float(value)
        if abs(v) >= 1_000_000:
            return f"{v/1_000_000:.2f}M"
        if abs(v) >= 1_000:
            return f"{v/1_000:.1f}K"
        return f"{v:,.0f}"

    rows = []
    for _, r in top.iterrows():
        rows.append({
            "#": int(r["Rank"]),
            "Brand": r["OEM"],
            "Visits 2025": short_num(r["UniqueVisitors_2025"]),
            "Visits 2024": short_num(r["UniqueVisitors_2024"]),
            "Visitor YoY": f"{r['Visitors YoY %']:+.1f}%",
            "Passenger sales": f"{r['Sales_2025']:,.0f}",
            "Conv rate": f"{r['ConversionPct_2025']:.2f}%",
            "Conv var": f"{r['Conv Var pp']:+.2f}pp",
        })

    table = pd.DataFrame(rows)

    def style_yoy(val):
        text = str(val)
        if text.startswith("+"):
            return "background-color: #DDF8EC; color: #12C76B; font-weight: 700;"
        if text.startswith("-"):
            return "background-color: #FFE5EF; color: #FF2F6D; font-weight: 700;"
        return "background-color: #EEF2F6; color: #6F6F6F; font-weight: 700;"

    st.markdown('<div class="section-kicker">Top 10 brands at a glance</div>', unsafe_allow_html=True)
    st.dataframe(
        table.style.map(style_yoy, subset=["Visitor YoY", "Conv var"]),
        use_container_width=True,
        hide_index=True,
    )



def render_market_performance_page(data, market, selected_oems):
    st.markdown(f'<div class="section-kicker">Market performance — {market} OEM cohort</div>', unsafe_allow_html=True)

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
    sales_yoy = (total_sales / prev_sales - 1) * 100 if prev_sales else 0
    uv_yoy = (total_uv / prev_uv - 1) * 100 if prev_uv else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("2025 passenger sales", f"{total_sales:,.0f}", f"{sales_yoy:+.1f}% YoY")
    c2.metric("2025 unique visitors", f"{total_uv:,.0f}", f"{uv_yoy:+.1f}% YoY")
    c3.metric("2025 conversion", f"{conv_2025:.2f}%", f"{conv_2025 - conv_2024:+.2f}pp")
    c4.metric("Visits to sale", f"{(total_uv / total_sales):,.0f}" if total_sales else "n/a")

    render_market_top10_inline(data, market, selected_oems)

    st.markdown('<div class="section-kicker">Performance visuals — YoY growth and visitor scale</div>', unsafe_allow_html=True)
    render_exec_visuals(data, market, selected_oems)

    st.markdown(f'<div class="section-kicker">AI insights — market performance</div>', unsafe_allow_html=True)
    cards = generate_insight_cards(data, market, selected_oems)
    # Avoid duplicating Toyota/Lexus recommendation focus too heavily here; show broader market cards.
    broader_cards = [c for c in cards if "Toyota:" not in c and "Lexus:" not in c][:6]
    if not broader_cards:
        broader_cards = cards[:6]

    for i in range(0, len(broader_cards), 3):
        cols = st.columns(3)
        for col, card in zip(cols, broader_cards[i:i + 3]):
            with col:
                st.markdown(card, unsafe_allow_html=True)

    render_footer_notes()


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
    st.subheader("Leadership scorecard")

    market = st.selectbox("Scorecard market", ["MM5", "UK", "France", "Germany", "Italy", "Spain"], index=0)

    scorecard = calculate_scorecard(data, market, selected_oems)

    if scorecard.empty:
        st.warning("No scorecard data available for this selection.")
        return

    st.markdown(f'<div class="scorecard-title">{market} 2024–2025 conversion scorecard</div>', unsafe_allow_html=True)
    st.caption("Ranking is based on 2025 conversion rate. Variance shows 2025 versus 2024.")

    st.dataframe(style_scorecard(scorecard), use_container_width=True, hide_index=True)

    csv = scorecard.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download scorecard as CSV",
        data=csv,
        file_name=f"{market.lower()}_2024_2025_conversion_scorecard.csv",
        mime="text/csv",
    )


render_hero()

data = load_data()
available_oems = sorted(data["OEM"].unique())

st.sidebar.header("Filters")

page = st.sidebar.radio(
    "Dashboard page",
    ["Toyota & Lexus Gap Analysis", "Market Performance", "Bubble chart", "Scorecard"],
    index=0,
)

summary_market = st.sidebar.selectbox(
    "Summary market",
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
    index=3,
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

if page == "Toyota & Lexus Gap Analysis":
    render_gap_analysis_page(data, summary_market, selected_oems)
elif page == "Market Performance":
    render_market_performance_page(data, summary_market, selected_oems)
elif page == "Bubble chart":
    render_bubble_page(data, selected_oems, year_view, show_logos)
else:
    render_scorecard_page(data, selected_oems)
