from pathlib import Path
import math

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# =========================
# App config
# =========================

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
GREEN = "#12C76B"
PINK = "#FF5C8A"
AMBER = "#FFB000"
INTELLIGENCE = "#2563EB"

TOYOTA_LOGO = "https://brand.toyota.com/content/dam/brandhub/guidelines/logo/four-column/BHUB_Logo_LogoFamily_01.svg"
LEXUS_LOGO = "https://upload.wikimedia.org/wikipedia/commons/7/75/Lexus.svg"
VALTECH_LOGO = "https://mma.prnewswire.com/media/2728124/Valtech_Logo.jpg"

MARKETS = ["MM5", "UK", "France", "Germany", "Italy", "Spain"]

TOYOTA_SET = ["Toyota", "VW", "Ford", "Peugeot", "Renault", "Hyundai", "Kia", "Nissan", "Skoda", "SEAT", "Dacia"]
LEXUS_SET = ["Lexus", "BMW", "Mercedes-Benz", "Audi", "Volvo", "Tesla", "Jaguar", "Land Rover", "Porsche", "Polestar"]
CHINESE_SET = ["BYD Auto", "MG", "XPeng", "NIO", "Geely", "Omoda", "Jaecoo", "Leapmotor", "GWM Ora"]

LOGO_DOMAIN_MAP = {
    "Toyota": "toyota.com",
    "Lexus": "lexus.com",
    "VW": "volkswagen.com",
    "BMW": "bmw.com",
    "Mercedes-Benz": "mercedes-benz.com",
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
    "SEAT": "seat.com",
    "Dacia": "dacia.com",
    "Jaguar": "jaguar.com",
    "Land Rover": "landrover.com",
    "BYD Auto": "byd.com",
    "MG": "mg.co.uk",
    "XPeng": "xpeng.com",
    "Geely": "geely.com",
    "Omoda": "omodaauto.com",
    "Jaecoo": "jaecoo.com",
    "Leapmotor": "leapmotor.com",
    "GWM Ora": "gwm-global.com",
    "Porsche": "porsche.com",
    "Polestar": "polestar.com",
}


# =========================
# Styling
# =========================

st.markdown(
    """
<style>
html, body, [class*="css"] {
    font-family: "Helvetica Neue", Arial, sans-serif;
}
.stApp {
    background: #ffffff;
}
section[data-testid="stSidebar"] {
    background: #000000;
}
section[data-testid="stSidebar"] * {
    color: #ffffff;
}
section[data-testid="stSidebar"] .stMultiSelect [data-baseweb="select"] *,
section[data-testid="stSidebar"] .stSelectbox [data-baseweb="select"] *,
section[data-testid="stSidebar"] input {
    color: #111827 !important;
}
span[data-baseweb="tag"] {
    background-color: #D9DDE3 !important;
    color: #111827 !important;
    border-radius: 8px !important;
}
span[data-baseweb="tag"] svg {
    color: #111827 !important;
}
.hero {
    background: #000000;
    color: #ffffff;
    padding: 28px 32px;
    border-radius: 22px;
    margin-bottom: 24px;
    position: relative;
    overflow: hidden;
    border-bottom: 8px solid #009FE3;
}
.hero::after {
    content: "";
    position: absolute;
    width: 430px;
    height: 430px;
    border-radius: 50%;
    background: linear-gradient(135deg, #009FE3, #6F6F6F);
    right: -180px;
    top: -230px;
    opacity: 0.55;
}
.hero-logo {
    position: relative;
    z-index: 2;
    margin-bottom: 18px;
}
.hero-logo img {
    height: 42px;
    width: auto;
    background: #ffffff;
    padding: 8px 12px;
    border-radius: 8px;
}
.hero-title {
    position: relative;
    z-index: 2;
    font-size: 42px;
    line-height: 1.05;
    font-weight: 800;
    letter-spacing: -0.04em;
    max-width: 1100px;
}
.hero-subtitle {
    position: relative;
    z-index: 2;
    margin-top: 12px;
    color: #e6e6e6;
    font-size: 17px;
    max-width: 1150px;
}
.hero-meta {
    position: relative;
    z-index: 2;
    margin-top: 18px;
    font-size: 14px;
    color: #d8d8d8;
}
.section-kicker {
    color: #8D96A0;
    font-size: 13px;
    font-weight: 800;
    letter-spacing: .32em;
    text-transform: uppercase;
    margin: 28px 0 18px;
    display: flex;
    align-items: center;
    gap: 14px;
}
.section-kicker::after {
    content: "";
    height: 1px;
    background: #E2E6EA;
    flex: 1;
}
.brand-strip {
    display: flex;
    align-items: center;
    gap: 36px;
    margin: 8px 0 20px 0;
}
.brand-strip img {
    object-fit: contain;
}
.card {
    background: #ffffff;
    border: 1px solid #e6e9ed;
    border-radius: 16px;
    padding: 20px 22px;
    box-shadow: 0 1px 8px rgba(0,0,0,.035);
}
.benchmark-title {
    font-size: 18px;
    font-weight: 800;
    color: #0A2342;
    margin-bottom: 10px;
}
.benchmark-copy {
    color: #6F7782;
    line-height: 1.45;
    font-size: 14px;
}
.benchmark-metric {
    font-size: 28px;
    color: #0A2342;
    margin-top: 16px;
}
.insight-card {
    background: #ffffff;
    border: 1px solid #e6e9ed;
    border-radius: 14px;
    padding: 20px 22px;
    min-height: 235px;
    box-shadow: 0 1px 8px rgba(0,0,0,.035);
    overflow: hidden;
}
.insight-card.opportunity {
    border-left: 5px solid #12C76B;
}
.insight-card.risk {
    border-left: 5px solid #FFB000;
}
.insight-card.intelligence {
    border-left: 5px solid #2563EB;
}
.market-flag {
    display: inline-block;
    background: #F3F7FF;
    color: #005F9E;
    border: 1px solid #CFE8FF;
    border-radius: 999px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 800;
    margin-bottom: 10px;
}
.insight-label {
    font-size: 12px;
    letter-spacing: .22em;
    text-transform: uppercase;
    font-weight: 800;
    margin-bottom: 12px;
}
.insight-label.opportunity {
    color: #12C76B;
}
.insight-label.risk {
    color: #FFB000;
}
.insight-label.intelligence {
    color: #2563EB;
}
.insight-title {
    font-size: 18px;
    line-height: 1.25;
    font-weight: 800;
    color: #0A2342;
    margin-bottom: 10px;
}
.insight-copy {
    color: #8B95A1;
    font-size: 14px;
    line-height: 1.55;
    min-height: 90px;
}
.insight-metric {
    color: #0A2342;
    font-size: 27px;
    margin-top: 14px;
    font-weight: 600;
}
.tag {
    display: inline-block;
    background: #EEF3FF;
    color: #2563EB;
    border-radius: 5px;
    padding: 4px 10px;
    font-size: 12px;
    margin-top: 10px;
}
.footer-note {
    margin-top: 36px;
    padding: 18px 20px;
    background: #F3F3F3;
    border-left: 6px solid #009FE3;
    border-radius: 12px;
    font-size: 13px;
    color: #0A2342;
}
div[data-testid="stMetricValue"] {
    color: #000000;
}
div[data-testid="stPlotlyChart"] {
    background: #ffffff;
    border: 1px solid #e6e9ed;
    border-radius: 14px;
    padding: 10px;
    box-shadow: 0 1px 8px rgba(0,0,0,.035);
}

.tl-detail-card {
    background: #ffffff;
    border: 1px solid #e6e9ed;
    border-radius: 16px;
    padding: 22px;
    box-shadow: 0 1px 8px rgba(0,0,0,.035);
    min-height: 260px;
}
.tl-detail-logo {
    height: 46px;
    display: flex;
    align-items: center;
    margin-bottom: 18px;
}
.tl-detail-logo img {
    max-height: 42px;
    max-width: 230px;
    object-fit: contain;
}
.tl-detail-grid {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 14px;
}
.tl-detail-metric {
    background: #F7F9FC;
    border-radius: 12px;
    padding: 14px;
    min-height: 120px;
}
.tl-detail-label {
    color: #6F7782;
    font-size: 13px;
    font-weight: 700;
    margin-bottom: 8px;
}
.tl-detail-value {
    color: #000000;
    font-size: 30px;
    line-height: 1.1;
    font-weight: 800;
    margin-bottom: 8px;
}
.tl-detail-delta-pos {
    display: inline-block;
    background: #DDF8EC;
    color: #128A3A;
    padding: 4px 8px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 800;
}
.tl-detail-delta-neg {
    display: inline-block;
    background: #FFE5EF;
    color: #D33F49;
    padding: 4px 8px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 800;
}

</style>
""",
    unsafe_allow_html=True,
)


# =========================
# Data layer
# =========================

def find_column(columns, candidates):
    lower = {str(c).strip().lower(): c for c in columns}
    for candidate in candidates:
        for key, original in lower.items():
            if candidate in key:
                return original
    raise ValueError(f"Could not find a column matching: {candidates}")


@st.cache_data
def load_data():
    raw = pd.read_excel(DATA_FILE)
    raw.columns = [str(c).strip() for c in raw.columns]

    brand_col = find_column(raw.columns, ["brand", "oem"])
    market_col = find_column(raw.columns, ["market", "country"])
    year_col = find_column(raw.columns, ["year"])
    sales_col = find_column(raw.columns, ["sales"])
    uv_col = find_column(raw.columns, ["unique visitors", "unique", "uv"])

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

    mm5 = (
        market_data.groupby(["OEM", "Year"], as_index=False)
        .agg({"Sales": "sum", "UniqueVisitors": "sum"})
    )
    mm5["Market"] = "MM5"
    mm5["ConversionPct"] = mm5["Sales"] / mm5["UniqueVisitors"] * 100

    combined = pd.concat(
        [
            mm5[["OEM", "Market", "Year", "Sales", "UniqueVisitors", "ConversionPct"]],
            market_data[["OEM", "Market", "Year", "Sales", "UniqueVisitors", "ConversionPct"]],
        ],
        ignore_index=True,
    )

    return combined.sort_values(["Market", "OEM", "Year"]).reset_index(drop=True)


def market_year(data, market, year, oems=None):
    df = data[(data["Market"] == market) & (data["Year"] == year)].copy()
    if oems:
        df = df[df["OEM"].isin(oems)]
    return df


def yoy_table(data, market, oems=None):
    d24 = market_year(data, market, 2024, oems).copy()
    d25 = market_year(data, market, 2025, oems).copy()

    if d24.empty or d25.empty:
        return pd.DataFrame()

    d24 = d24.rename(
        columns={
            "Sales": "Sales_2024",
            "UniqueVisitors": "UniqueVisitors_2024",
            "ConversionPct": "ConversionPct_2024",
            "Year": "Year_2024",
        }
    )
    d25 = d25.rename(
        columns={
            "Sales": "Sales_2025",
            "UniqueVisitors": "UniqueVisitors_2025",
            "ConversionPct": "ConversionPct_2025",
            "Year": "Year_2025",
        }
    )

    d24 = d24[["OEM", "Market", "Sales_2024", "UniqueVisitors_2024", "ConversionPct_2024"]]
    d25 = d25[["OEM", "Market", "Sales_2025", "UniqueVisitors_2025", "ConversionPct_2025"]]

    merged = d25.merge(d24, on=["OEM", "Market"], how="inner")
    if merged.empty:
        return pd.DataFrame()

    merged["Sales YoY %"] = (merged["Sales_2025"] / merged["Sales_2024"] - 1) * 100
    merged["Visitors YoY %"] = (merged["UniqueVisitors_2025"] / merged["UniqueVisitors_2024"] - 1) * 100
    merged["Conv Var pp"] = merged["ConversionPct_2025"] - merged["ConversionPct_2024"]
    merged["Visits to Sale 2025"] = merged["UniqueVisitors_2025"] / merged["Sales_2025"]
    merged["Visits to Sale 2024"] = merged["UniqueVisitors_2024"] / merged["Sales_2024"]
    merged["Visits to Sale Var"] = merged["Visits to Sale 2025"] - merged["Visits to Sale 2024"]
    merged = merged.replace([float("inf"), float("-inf")], pd.NA)
    return merged


def selected_or_all(selected, all_oems):
    return selected if selected else all_oems


# Backward-compatible aliases in case any code path uses older names.
get_yoy_table = yoy_table
get_market_data = market_year


# =========================
# Formatting helpers
# =========================

def fmt_int(v):
    if pd.isna(v):
        return "n/a"
    return f"{float(v):,.0f}"


def fmt_pct(v, decimals=1):
    if pd.isna(v):
        return "n/a"
    sign = "+" if float(v) > 0 else ""
    return f"{sign}{float(v):.{decimals}f}%"


def fmt_pp(v):
    if pd.isna(v):
        return "n/a"
    sign = "+" if float(v) > 0 else ""
    return f"{sign}{float(v):.2f}pp"


def fmt_short(v):
    v = float(v)
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.2f}M"
    if abs(v) >= 1_000:
        return f"{v/1_000:.1f}K"
    return f"{v:,.0f}"


def badge_style(value):
    text = str(value)
    if text.startswith("+"):
        return "background-color:#DDF8EC;color:#12C76B;font-weight:700;border-radius:7px;"
    if text.startswith("-"):
        return "background-color:#FFE5EF;color:#FF2F6D;font-weight:700;border-radius:7px;"
    return "background-color:#EEF2F6;color:#6F6F6F;font-weight:700;border-radius:7px;"


def badge_html(value, suffix="%"):
    try:
        v = float(value)
    except Exception:
        return '<span style="background:#EEF2F6;color:#6F6F6F;padding:5px 9px;border-radius:6px;font-weight:700;">n/a</span>'
    cls_bg = "#DDF8EC" if v > 0 else "#FFE5EF" if v < 0 else "#EEF2F6"
    cls_col = "#12C76B" if v > 0 else "#FF2F6D" if v < 0 else "#6F6F6F"
    sign = "+" if v > 0 else ""
    text = f"{sign}{v:.2f}pp" if suffix == "pp" else f"{sign}{v:.1f}%"
    return f'<span style="background:{cls_bg};color:{cls_col};padding:5px 9px;border-radius:6px;font-weight:700;">{text}</span>'


def get_row(yoy, brand):
    if yoy.empty:
        return None
    row = yoy[yoy["OEM"].str.lower() == brand.lower()]
    return None if row.empty else row.iloc[0]


def preset_selection(preset_name, all_oems):
    if preset_name == "Toyota volume competitors":
        preset = TOYOTA_SET
    elif preset_name == "Lexus premium competitors":
        preset = LEXUS_SET
    elif preset_name == "Chinese disruptors":
        preset = CHINESE_SET
    else:
        preset = all_oems
    return [x for x in preset if x in all_oems]


# =========================
# Visual components
# =========================

def render_hero():
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-logo"><img src="{VALTECH_LOGO}" alt="Valtech logo"></div>
            <div class="hero-title">OEM Omnichannel Conversion Funnel</div>
            <div class="hero-subtitle">
                Website to customer contract conversion across MM5. Explores how unique visitors demand converts into passenger new car customer contracts, and where Toyota/Lexus under- or over-perform by market.
            </div>
            <div class="hero-meta">
                <b>Coverage:</b> 2024 &amp; 2025 &nbsp;|&nbsp;
                <b>Sales data:</b> www.marklines.com &nbsp;|&nbsp;
                <b>Unique visitors:</b> www.similarweb.com
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section(title):
    st.markdown(f'<div class="section-kicker">{title}</div>', unsafe_allow_html=True)


def render_footer():
    st.markdown(
        """
        <div class="footer-note">
        <b>Definitions and caveats:</b> This is not a total market-share dashboard. It uses passenger car sales from Marklines and unique visitor website data from Similarweb. It does not include fleet, LCV or tactical registrations. Conversion rate = passenger car sales divided by unique visitors. Data period: 2024–2025.
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_brand_strip():
    st.markdown(
        f"""
        <div class="brand-strip">
            <img src="{TOYOTA_LOGO}" style="height:46px;" alt="Toyota logo">
            <img src="{LEXUS_LOGO}" style="height:30px;" alt="Lexus logo">
        </div>
        """,
        unsafe_allow_html=True,
    )


def delta_class_for_html(value):
    try:
        return "tl-detail-delta-pos" if float(value) >= 0 else "tl-detail-delta-neg"
    except Exception:
        return "tl-detail-delta-pos"


def brand_detail_html(brand, row, logo, logo_height="38px"):
    if row is None:
        return (
            "<div class='tl-detail-card'>"
            f"<div class='tl-detail-logo'><img src='{logo}' style='height:{logo_height};'></div>"
            f"<p>No {brand} data available for this market.</p>"
            "</div>"
        )

    conv_delta_class = delta_class_for_html(row["Conv Var pp"])
    sales_delta_class = delta_class_for_html(row["Sales YoY %"])
    visitor_delta_class = delta_class_for_html(row["Visitors YoY %"])

    return (
        "<div class='tl-detail-card'>"
        f"<div class='tl-detail-logo'><img src='{logo}' style='height:{logo_height};' alt='{brand} logo'></div>"
        "<div class='tl-detail-grid'>"
        "<div class='tl-detail-metric'>"
        "<div class='tl-detail-label'>2025 conversion</div>"
        f"<div class='tl-detail-value'>{row['ConversionPct_2025']:.2f}%</div>"
        f"<div class='{conv_delta_class}'>{fmt_pp(row['Conv Var pp'])} vs 2024</div>"
        "</div>"
        "<div class='tl-detail-metric'>"
        "<div class='tl-detail-label'>2025 passenger sales</div>"
        f"<div class='tl-detail-value'>{fmt_int(row['Sales_2025'])}</div>"
        f"<div class='{sales_delta_class}'>{fmt_pct(row['Sales YoY %'])} vs 2024</div>"
        "</div>"
        "<div class='tl-detail-metric'>"
        "<div class='tl-detail-label'>2025 unique visitors</div>"
        f"<div class='tl-detail-value'>{fmt_int(row['UniqueVisitors_2025'])}</div>"
        f"<div class='{visitor_delta_class}'>{fmt_pct(row['Visitors YoY %'])} vs 2024</div>"
        "</div>"
        "</div>"
        "</div>"
    )


def render_brand_detail(data, market):
    section("2025 Toyota & Lexus performance vs 2024")
    yoy = yoy_table(data, market, None)
    toyota = get_row(yoy, "Toyota")
    lexus = get_row(yoy, "Lexus")

    cols = st.columns(2)
    with cols[0]:
        st.markdown(brand_detail_html("Toyota", toyota, TOYOTA_LOGO, "42px"), unsafe_allow_html=True)
    with cols[1]:
        st.markdown(brand_detail_html("Lexus", lexus, LEXUS_LOGO, "30px"), unsafe_allow_html=True)


def benchmark_card(title, copy, metric):
    html = (
        "<div class='card'>"
        f"<div class='benchmark-title'>{title}</div>"
        f"<div class='benchmark-copy'>{copy}</div>"
        f"<div class='benchmark-metric'>{metric}</div>"
        "</div>"
    )
    st.markdown(html, unsafe_allow_html=True)


def render_benchmark_cards(data, market):
    section("Toyota / Lexus benchmark callouts")
    yoy = yoy_table(data, market, None)
    if yoy.empty:
        st.info("No benchmark data available.")
        return

    cols = st.columns(2)
    for col, brand, cohort in [(cols[0], "Toyota", TOYOTA_SET), (cols[1], "Lexus", LEXUS_SET)]:
        with col:
            row = get_row(yoy, brand)
            cohort_df = yoy[yoy["OEM"].isin(cohort)].copy()
            if row is None or cohort_df.empty:
                st.info(f"No {brand} benchmark available.")
                continue
            cohort_df["Rank"] = cohort_df["ConversionPct_2025"].rank(method="min", ascending=False)
            rank = int(cohort_df.loc[cohort_df["OEM"] == brand, "Rank"].iloc[0])
            leader = cohort_df.sort_values("ConversionPct_2025", ascending=False).iloc[0]
            gap = row["ConversionPct_2025"] - leader["ConversionPct_2025"]
            vts_gap = row["Visits to Sale 2025"] - leader["Visits to Sale 2025"]
            benchmark_card(
                f"{brand} benchmark",
                f"{brand} ranks #{rank} of {len(cohort_df)} in its benchmark set. The 2025 conversion gap to {leader['OEM']} is {gap:+.2f}pp. Visits-to-sale gap is {vts_gap:+.0f}.",
                f"2025 conversion: {row['ConversionPct_2025']:.2f}%",
            )


def render_market_weakness(data):
    section("Market weakness summary — Toyota & Lexus")
    rows = []
    for market in ["UK", "France", "Germany", "Italy", "Spain"]:
        yoy = yoy_table(data, market, None)
        if yoy.empty:
            continue
        for brand, cohort in [("Toyota", TOYOTA_SET), ("Lexus", LEXUS_SET)]:
            row = get_row(yoy, brand)
            cohort_df = yoy[yoy["OEM"].isin(cohort)].copy()
            if row is None or cohort_df.empty:
                continue
            leader = cohort_df.sort_values("ConversionPct_2025", ascending=False).iloc[0]
            rows.append(
                {
                    "Brand": brand,
                    "Market": market,
                    "2025 conversion": f"{row['ConversionPct_2025']:.2f}%",
                    "Benchmark leader": leader["OEM"],
                    "Leader conversion": f"{leader['ConversionPct_2025']:.2f}%",
                    "Gap to leader": row["ConversionPct_2025"] - leader["ConversionPct_2025"],
                    "Sales YoY vs 2024": row["Sales YoY %"],
                    "Visitor YoY vs 2024": row["Visitors YoY %"],
                }
            )
    if not rows:
        st.info("No Toyota/Lexus market weakness data available.")
        return

    df = pd.DataFrame(rows).sort_values("Gap to leader")
    display = df.copy()
    display["Gap to leader"] = display["Gap to leader"].map(fmt_pp)
    display["Sales YoY vs 2024"] = display["Sales YoY vs 2024"].map(fmt_pct)
    display["Visitor YoY vs 2024"] = display["Visitor YoY vs 2024"].map(fmt_pct)

    styler = display.style.map(badge_style, subset=["Gap to leader", "Sales YoY vs 2024", "Visitor YoY vs 2024"])
    st.dataframe(styler, use_container_width=True, hide_index=True)


def make_insight(kind, title, copy, metric, tag, market=None):
    label = {"opportunity": "Opportunity", "risk": "Risk Alert", "intelligence": "Intelligence"}[kind]
    flag = f"<div class='market-flag'>{market}</div>" if market else ""
    return (
        f"<div class='insight-card {kind}'>"
        f"{flag}"
        f"<div class='insight-label {kind}'>{label}</div>"
        f"<div class='insight-title'>{title}</div>"
        f"<div class='insight-copy'>{copy}</div>"
        f"<div class='insight-metric'>{metric}</div>"
        f"<div class='tag'>{tag}</div>"
        f"</div>"
    )


def market_action_text(brand, market, row, leader, gap):
    if market == "Germany":
        return f"{brand} is materially behind {leader['OEM']} in Germany. Prioritise retailer hand-off quality, stock availability messaging and high-intent model pages; this market needs fewer generic visits and more purchase-ready journeys."
    if market == "UK":
        return f"In the UK, {brand} needs sharper lower-funnel progression. Focus on enquiry forms, test-drive booking, finance CTAs and stock routing where intent is most likely leaking before contract."
    if market == "France":
        return f"France needs stronger localisation of offers and dealer follow-up prompts. The opportunity is to turn browsing into configured, dealer-ready demand."
    if market == "Italy":
        return f"Italy looks like a journey-efficiency problem. Reduce friction between model discovery and contact request; simplify price, stock and finance visibility."
    if market == "Spain":
        return f"Spain is relatively more resilient. Protect momentum by scaling the highest-converting traffic sources and making offer, stock and lead response paths more prominent."
    return f"{brand} trails {leader['OEM']} by {gap:+.2f}pp. Prioritise highest-intent contract funnel steps and remove friction from enquiry, finance and stock discovery."


def render_toyota_lexus_recommendations(data):
    section("AI insights — Toyota & Lexus market recommendations")
    cards = []
    for market in ["UK", "France", "Germany", "Italy", "Spain"]:
        yoy = yoy_table(data, market, None)
        if yoy.empty:
            continue
        for brand, cohort in [("Toyota", TOYOTA_SET), ("Lexus", LEXUS_SET)]:
            row = get_row(yoy, brand)
            cohort_df = yoy[yoy["OEM"].isin(cohort)].copy()
            if row is None or cohort_df.empty:
                continue
            leader = cohort_df.sort_values("ConversionPct_2025", ascending=False).iloc[0]
            gap = row["ConversionPct_2025"] - leader["ConversionPct_2025"]
            severity = abs(gap) + (0.5 if row["Sales YoY %"] < 0 else 0)
            kind = "risk" if gap < -0.75 or row["Sales YoY %"] < 0 else "opportunity"
            copy = (
                f"{brand} converts at {row['ConversionPct_2025']:.2f}% versus {leader['OEM']} at {leader['ConversionPct_2025']:.2f}%. "
                f"Sales are {fmt_pct(row['Sales YoY %'])} vs 2024; visitors are {fmt_pct(row['Visitors YoY %'])} vs 2024. "
                + market_action_text(brand, market, row, leader, gap)
            )
            cards.append(
                (
                    severity,
                    make_insight(
                        kind,
                        f"{brand}: {market} conversion gap to {leader['OEM']}",
                        copy,
                        f"{gap:+.2f}pp gap",
                        f"{brand} recommendation",
                        market,
                    ),
                )
            )
    cards = [c for _, c in sorted(cards, key=lambda x: x[0], reverse=True)[:8]]
    for i in range(0, len(cards), 3):
        cols = st.columns(3)
        for col, card in zip(cols, cards[i:i + 3]):
            with col:
                st.markdown(card, unsafe_allow_html=True)


def render_top10_table(data, market, selected_oems):
    section("Top 10 brands at a glance")
    yoy = yoy_table(data, market, selected_oems)
    if yoy.empty:
        st.info("No Top 10 data available for this selection.")
        return
    top = yoy.sort_values("UniqueVisitors_2025", ascending=False).head(10).copy()
    rows = []
    for i, (_, r) in enumerate(top.iterrows(), start=1):
        rows.append(
            {
                "#": i,
                "Brand": r["OEM"],
                "Visits 2025": fmt_short(r["UniqueVisitors_2025"]),
                "Visits 2024": fmt_short(r["UniqueVisitors_2024"]),
                "Visitor YoY vs 2024": fmt_pct(r["Visitors YoY %"]),
                "Passenger sales": fmt_int(r["Sales_2025"]),
                "Conv rate": f"{r['ConversionPct_2025']:.2f}%",
                "Conv var vs 2024": fmt_pp(r["Conv Var pp"]),
            }
        )
    display = pd.DataFrame(rows)
    st.dataframe(
        display.style.map(badge_style, subset=["Visitor YoY vs 2024", "Conv var vs 2024"]),
        use_container_width=True,
        hide_index=True,
    )


def build_yoy_growth_chart(data, market, selected_oems, top_n=10):
    yoy = yoy_table(data, market, selected_oems)
    fig = go.Figure()
    if yoy.empty:
        return fig

    metric = "Visitors YoY %"
    yoy = yoy.dropna(subset=[metric])
    top = yoy.sort_values(metric, ascending=False).head(top_n)
    bottom = yoy.sort_values(metric, ascending=True).head(top_n)
    chart_df = pd.concat([top, bottom]).drop_duplicates("OEM").sort_values(metric, ascending=True)
    colors = chart_df[metric].apply(lambda x: "#5ED6AC" if x >= 0 else "#FF5C8A")

    fig.add_trace(
        go.Bar(
            x=chart_df[metric],
            y=chart_df["OEM"],
            orientation="h",
            marker=dict(color=colors),
            text=chart_df[metric].map(lambda x: f"{x:+.1f}%"),
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Visitor YoY: %{x:+.1f}%<extra></extra>",
        )
    )
    fig.update_layout(
        title=dict(text=f"YoY visit growth — 2025 vs 2024 (top & bottom {top_n})", x=0.02, font=dict(size=16, color="#8D96A0")),
        height=620,
        margin=dict(l=120, r=80, t=70, b=50),
        plot_bgcolor=WHITE,
        paper_bgcolor=WHITE,
        showlegend=False,
    )
    min_x = min(chart_df[metric].min() * 1.25, -5)
    max_x = max(chart_df[metric].max() * 1.25, 5)
    fig.update_xaxes(range=[min_x, max_x], ticksuffix="%", gridcolor="#E2E6EA", zeroline=True, zerolinecolor="#B8C0C8")
    fig.update_yaxes(gridcolor="#FFFFFF")
    return fig


def render_market_insights(data, market, selected_oems):
    section("AI insights — market performance")
    yoy = yoy_table(data, market, selected_oems)
    if yoy.empty:
        st.info("No insight data available.")
        return

    top_conv = yoy.sort_values("ConversionPct_2025", ascending=False).iloc[0]
    top_sales = yoy.sort_values("Sales YoY %", ascending=False).iloc[0]
    top_visits = yoy.sort_values("Visitors YoY %", ascending=False).iloc[0]
    worst_conv = yoy.sort_values("Conv Var pp", ascending=True).iloc[0]
    weakest = yoy.sort_values("ConversionPct_2025", ascending=True).iloc[0]

    cards = [
        make_insight("opportunity", f"{top_conv['OEM']}: strongest conversion performer", f"{top_conv['OEM']} leads this cohort on 2025 conversion. Visits-to-sale is {top_conv['Visits to Sale 2025']:.0f}, showing stronger website-to-contract efficiency than peers.", f"{top_conv['ConversionPct_2025']:.2f}% conv", top_conv["OEM"]),
        make_insight("opportunity", f"{top_sales['OEM']}: strongest sales growth", f"{top_sales['OEM']} posted the strongest passenger sales growth, with sales {fmt_pct(top_sales['Sales YoY %'])} vs 2024 and visitors {fmt_pct(top_sales['Visitors YoY %'])} vs 2024.", f"{fmt_pct(top_sales['Sales YoY %'])} sales", top_sales["OEM"]),
        make_insight("intelligence", f"{top_visits['OEM']}: fastest visitor growth", f"{top_visits['OEM']} generated the strongest visitor growth. The key question is whether this awareness converts into contracts at the same rate.", f"{fmt_pct(top_visits['Visitors YoY %'])} visits", top_visits["OEM"]),
        make_insight("risk", f"{worst_conv['OEM']}: conversion movement deteriorating", f"{worst_conv['OEM']} had the weakest conversion movement: {fmt_pp(worst_conv['Conv Var pp'])} vs 2024. This suggests lower-quality traffic or leakage deeper in the funnel.", f"{fmt_pp(worst_conv['Conv Var pp'])}", worst_conv["OEM"]),
        make_insight("risk", f"{weakest['OEM']}: weakest 2025 conversion", f"{weakest['OEM']} has the lowest 2025 conversion rate in this selection, creating a clear efficiency gap against the cohort leader.", f"{weakest['ConversionPct_2025']:.2f}% conv", weakest["OEM"]),
    ]
    for i in range(0, len(cards), 3):
        cols = st.columns(3)
        for col, card in zip(cols, cards[i:i + 3]):
            with col:
                st.markdown(card, unsafe_allow_html=True)


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


def build_bubble_chart(chart_df, selected_oems, market, year_view, show_logos):
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
        df = chart_df[chart_df["Year"] == year].copy()
        if df.empty:
            continue
        sizes = df["Sales"].apply(lambda s: max(18, math.sqrt(s / max_sales) * 74))
        fig.add_trace(
            go.Scatter(
                x=df["UniqueVisitors"],
                y=df["ConversionPct"],
                mode="markers" if show_logos else "markers+text",
                text=None if show_logos else df["OEM"],
                textposition="top center",
                name=str(year),
                marker=dict(size=sizes, color=colors[year], opacity=0.68, line=dict(width=1.3, color="black")),
                customdata=df[["OEM", "Market", "Year", "Sales", "UniqueVisitors", "ConversionPct"]],
                hovertemplate="<b>%{customdata[0]}</b><br>Market: %{customdata[1]}<br>Year: %{customdata[2]}<br>Sales: %{customdata[3]:,.0f}<br>Visitors: %{customdata[4]:,.0f}<br>Conversion: %{customdata[5]:.2f}%<extra></extra>",
            )
        )

    if show_logos:
        add_logo_images(fig, chart_df, max_x, max_y)

    fig.update_layout(
        title=f"{market} | Unique visitors vs conversion rate",
        xaxis_title="Unique visitors",
        yaxis_title="Conversion rate (%)",
        height=720,
        hovermode="closest",
        legend=dict(orientation="h", y=1.08, x=0),
        margin=dict(l=70, r=40, t=90, b=70),
        plot_bgcolor=WHITE,
        paper_bgcolor=WHITE,
    )
    fig.update_xaxes(range=[0, max_x], tickformat=",", gridcolor="#e6e6e6")
    fig.update_yaxes(range=[0, max_y], ticksuffix="%", gridcolor="#e6e6e6")
    return fig


# =========================
# Pages
# =========================

def render_gap_analysis_page(data, market):
    section(f"Toyota & Lexus gap analysis — {market}")
    st.caption("All headline metrics reflect 2025 performance; variance badges show movement versus 2024.")
    render_brand_strip()

    combined = market_year(data, market, 2025, ["Toyota", "Lexus"])
    combined_24 = market_year(data, market, 2024, ["Toyota", "Lexus"])

    if combined.empty:
        st.warning("No Toyota/Lexus data available.")
        return

    sales_25 = combined["Sales"].sum()
    visitors_25 = combined["UniqueVisitors"].sum()
    conv_25 = sales_25 / visitors_25 * 100
    sales_24 = combined_24["Sales"].sum()
    visitors_24 = combined_24["UniqueVisitors"].sum()
    conv_24 = sales_24 / visitors_24 * 100

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("2025 Toyota/Lexus passenger sales", fmt_int(sales_25), f"{fmt_pct((sales_25 / sales_24 - 1) * 100)} vs 2024")
    c2.metric("2025 Toyota/Lexus unique visitors", fmt_int(visitors_25), f"{fmt_pct((visitors_25 / visitors_24 - 1) * 100)} vs 2024")
    c3.metric("2025 weighted conversion", f"{conv_25:.2f}%", f"{fmt_pp(conv_25 - conv_24)} vs 2024")
    c4.metric("2025 visits-to-sale efficiency", fmt_int(visitors_25 / sales_25))

    render_brand_detail(data, market)
    render_benchmark_cards(data, market)
    render_market_weakness(data)
    render_toyota_lexus_recommendations(data)
    render_footer()


def render_market_performance_page(data, market, selected_oems):
    section(f"Market performance — {market} OEM cohort")
    current = market_year(data, market, 2025, selected_oems)
    previous = market_year(data, market, 2024, selected_oems)
    if current.empty:
        st.warning("No data available for this selection.")
        return

    sales_25 = current["Sales"].sum()
    visitors_25 = current["UniqueVisitors"].sum()
    conv_25 = sales_25 / visitors_25 * 100
    sales_24 = previous["Sales"].sum()
    visitors_24 = previous["UniqueVisitors"].sum()
    conv_24 = sales_24 / visitors_24 * 100 if visitors_24 else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("2025 passenger sales", fmt_int(sales_25), f"{fmt_pct((sales_25 / sales_24 - 1) * 100)} vs 2024" if sales_24 else "n/a")
    c2.metric("2025 unique visitors", fmt_int(visitors_25), f"{fmt_pct((visitors_25 / visitors_24 - 1) * 100)} vs 2024" if visitors_24 else "n/a")
    c3.metric("2025 conversion", f"{conv_25:.2f}%", f"{fmt_pp(conv_25 - conv_24)} vs 2024")
    c4.metric("2025 visits-to-sale efficiency", fmt_int(visitors_25 / sales_25))

    render_top10_table(data, market, selected_oems)
    section("Performance visuals — YoY visit growth")
    st.plotly_chart(build_yoy_growth_chart(data, market, selected_oems, top_n=10), use_container_width=True)
    render_market_insights(data, market, selected_oems)
    render_footer()


def render_bubble_page(data, selected_oems, year_view, show_logos):
    section("Bubble chart")
    tabs = st.tabs(MARKETS)
    for market, tab in zip(MARKETS, tabs):
        with tab:
            df = data[(data["Market"] == market) & (data["OEM"].isin(selected_oems))].copy()
            if year_view != "2024 and 2025 + shift":
                df = df[df["Year"] == int(year_view)]
            if df.empty:
                st.info("No bubble chart data for this selection.")
                continue
            fig = build_bubble_chart(df, selected_oems, market, year_view, show_logos)
            st.plotly_chart(fig, use_container_width=True)


def scorecard_table(data, market, selected_oems):
    yoy = yoy_table(data, market, selected_oems)
    if yoy.empty:
        return pd.DataFrame()
    out = yoy.copy()
    out["Conv Ranking"] = out["ConversionPct_2025"].rank(method="first", ascending=False).astype(int)
    out = out.sort_values("Conv Ranking")
    return out[
        [
            "OEM",
            "UniqueVisitors_2024",
            "UniqueVisitors_2025",
            "Visitors YoY %",
            "Sales_2024",
            "Sales_2025",
            "Sales YoY %",
            "ConversionPct_2024",
            "ConversionPct_2025",
            "Conv Var pp",
            "Visits to Sale 2024",
            "Visits to Sale 2025",
            "Visits to Sale Var",
            "Conv Ranking",
        ]
    ]


def render_scorecard_page(data, selected_oems):
    section("Leadership scorecard")
    market = st.selectbox("Scorecard market", MARKETS, index=0)
    score = scorecard_table(data, market, selected_oems)
    if score.empty:
        st.info("No scorecard data available.")
        return

    display = score.rename(
        columns={
            "OEM": "Brand",
            "UniqueVisitors_2024": "Website visitors 2024",
            "UniqueVisitors_2025": "Website visitors 2025",
            "Visitors YoY %": "Visitor YoY vs 2024",
            "Sales_2024": "Passenger sales 2024",
            "Sales_2025": "Passenger sales 2025",
            "Sales YoY %": "Sales YoY vs 2024",
            "ConversionPct_2024": "Conv rate 2024",
            "ConversionPct_2025": "Conv rate 2025",
            "Conv Var pp": "Conv var vs 2024",
        }
    ).copy()

    for col in ["Website visitors 2024", "Website visitors 2025", "Passenger sales 2024", "Passenger sales 2025"]:
        display[col] = display[col].map(fmt_int)
    for col in ["Visitor YoY vs 2024", "Sales YoY vs 2024"]:
        display[col] = display[col].map(fmt_pct)
    for col in ["Conv rate 2024", "Conv rate 2025"]:
        display[col] = display[col].map(lambda x: f"{x:.2f}%")
    display["Conv var vs 2024"] = display["Conv var vs 2024"].map(fmt_pp)
    for col in ["Visits to Sale 2024", "Visits to Sale 2025", "Visits to Sale Var"]:
        display[col] = display[col].map(fmt_int)

    st.dataframe(
        display.style.map(badge_style, subset=["Visitor YoY vs 2024", "Sales YoY vs 2024", "Conv var vs 2024"]),
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Download scorecard as CSV",
        data=score.to_csv(index=False).encode("utf-8"),
        file_name=f"{market.lower()}_scorecard.csv",
        mime="text/csv",
    )


# =========================
# App runtime
# =========================

render_hero()

data = load_data()
all_oems = sorted(data["OEM"].unique())

# Startup sanity check. If this fails, the app should stop immediately with one useful message.
for m in MARKETS:
    y = yoy_table(data, m, None)
    if not y.empty:
        required_cols = {"Sales_2025", "Sales_2024", "UniqueVisitors_2025", "UniqueVisitors_2024", "Sales YoY %", "Visitors YoY %", "Conv Var pp"}
        missing = required_cols - set(y.columns)
        if missing:
            st.error(f"Data validation failed for {m}. Missing: {sorted(missing)}")
            st.stop()

st.sidebar.header("Filters")

page = st.sidebar.radio(
    "Dashboard page",
    ["Toyota & Lexus Gap Analysis", "Market Performance", "Bubble chart", "Scorecard"],
    index=0,
)

summary_market = st.sidebar.selectbox("Summary market", MARKETS, index=0)

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

default_oems = preset_selection(preset, all_oems)

selected_oems = st.sidebar.multiselect(
    "OEMs",
    all_oems,
    default=default_oems,
)

show_logos = st.sidebar.toggle(
    "Show OEM logos",
    value=True,
    help="Logos require internet access and a Brandfetch Client ID in Streamlit Secrets.",
)

if show_logos and not get_brandfetch_client_id():
    st.sidebar.warning("Brandfetch Client ID not found. Add BRANDFETCH_CLIENT_ID in Streamlit Secrets or turn logos off.")

selected_oems = selected_or_all(selected_oems, all_oems)

if page == "Toyota & Lexus Gap Analysis":
    render_gap_analysis_page(data, summary_market)
elif page == "Market Performance":
    render_market_performance_page(data, summary_market, selected_oems)
elif page == "Bubble chart":
    render_bubble_page(data, selected_oems, year_view, show_logos)
else:
    render_scorecard_page(data, selected_oems)
