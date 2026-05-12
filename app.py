from pathlib import Path
import base64
import math
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st



def image_file_to_data_uri(path):
    """Return a data URI so local images can be rendered inside raw HTML."""
    image_path = Path(path)
    if not image_path.exists():
        return ""
    suffix = image_path.suffix.lower()
    mime = "image/png" if suffix == ".png" else "image/jpeg" if suffix in [".jpg", ".jpeg"] else "image/svg+xml"
    encoded = base64.b64encode(image_path.read_bytes()).decode("utf-8")
    return f"data:{mime};base64,{encoded}"


# =========================
# App config
# =========================

st.set_page_config(
    page_title="OEM Macro Conversion Funnel",
    page_icon="🚗",
    layout="wide",
)

DATA_FILE = Path(__file__).parent / "OEM Visit to Sales Data 2024-2026.xlsx"

VALTECH_BLUE = "#009FE3"
VALTECH_GREY = "#6F6F6F"
VALTECH_LIGHT_GREY = "#F3F3F3"
BLACK = "#000000"
WHITE = "#FFFFFF"
GREEN = "#12C76B"
PINK = "#FF5C8A"
AMBER = "#FFB000"
INTELLIGENCE = "#2563EB"

TOYOTA_LOGO = image_file_to_data_uri(Path(__file__).parent / "toyota_logo.png")
LEXUS_LOGO = "https://upload.wikimedia.org/wikipedia/commons/7/75/Lexus.svg"
VALTECH_LOGO = "https://mma.prnewswire.com/media/2728124/Valtech_Logo.jpg"

MARKETS = ["MM5", "UK", "France", "Germany", "Italy", "Spain"]

COMPARISON_OPTIONS = {
    "2024 vs 2025": {
        "previous_period": "2024",
        "current_period": "2025",
        "previous_label": "2024",
        "current_label": "2025",
        "coverage_label": "2024 & 2025",
    },
    "Jan-Apr 2025 vs Jan-Apr 2026": {
        "previous_period": "Jan - April 25",
        "current_period": "Jan - April 26",
        "previous_label": "Jan-Apr 2025",
        "current_label": "Jan-Apr 2026",
        "coverage_label": "2026 (Jan to April)",
    },
}

ACTIVE_COMPARISON = COMPARISON_OPTIONS["2024 vs 2025"]
PREVIOUS_LABEL = ACTIVE_COMPARISON["previous_label"]
CURRENT_LABEL = ACTIVE_COMPARISON["current_label"]


TOYOTA_SET = ["Toyota", "VW", "Ford", "Peugeot", "Renault", "Hyundai", "Kia", "Nissan", "Skoda", "SEAT", "Dacia"]
LEXUS_SET = ["Lexus", "BMW", "Mercedes-Benz", "Audi", "Volvo", "Tesla", "Jaguar", "Land Rover", "Porsche", "Polestar"]
CHINESE_SET = ["BYD Auto", "MG", "XPeng", "NIO", "Geely", "Omoda", "Jaecoo", "Leapmotor", "GWM Ora"]

OEM_CLUSTERS = {
    "Volume Leaders": [
        "BMW", "Toyota", "Mercedes-Benz", "Audi", "Ford", "Kia", "VW", "Volkswagen",
        "Nissan", "Skoda", "Honda"
    ],
    "EV Challengers": [
        "Tesla", "BYD", "BYD Auto", "Polestar", "NIO", "Zeekr", "Leapmotor", "AION"
    ],
    "European Mass Market": [
        "Vauxhall", "Opel", "Renault", "Peugeot", "Dacia", "Citroën", "Citroen",
        "SEAT", "Fiat", "DS Automobiles", "Alpine"
    ],
    "Luxury & Performance": [
        "Volvo", "Land Rover", "Porsche", "Lexus", "Ferrari", "Jaguar", "Aston Martin",
        "Bentley", "Lamborghini", "Maserati", "Rolls-Royce"
    ],
    "Asian Mainstream": [
        "Hyundai", "MINI", "Mazda", "CUPRA", "Suzuki", "Genesis", "Subaru", "Mitsubishi"
    ],
    "Chinese New Entrants": [
        "MG", "Smart", "Jaecoo", "KGM", "Omoda", "GWM", "GWM Ora", "Isuzu",
        "Chery", "SAIC Maxus", "XPeng", "Geely"
    ],
}

CLUSTER_DESCRIPTIONS = {
    "Volume Leaders": "Established mass-market brands with the largest digital footprint",
    "EV Challengers": "Electric-first brands disrupting the traditional OEM landscape",
    "European Mass Market": "European mainstream brands competing on value and heritage",
    "Luxury & Performance": "Premium and ultra-premium brands targeting high-intent affluent buyers",
    "Asian Mainstream": "Japanese and Korean mainstream brands with strong hybrid / EV portfolios",
    "Chinese New Entrants": "Chinese OEMs building digital presence ahead of market entry or expansion",
}

CLUSTER_COLORS = {
    "Volume Leaders": "#2563EB",
    "EV Challengers": "#12C76B",
    "European Mass Market": "#FFB000",
    "Luxury & Performance": "#B544F4",
    "Asian Mainstream": "#FF7A1A",
    "Chinese New Entrants": "#FF3B77",
}


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


.start-intro {
    background:#F7F9FC;
    border-left:6px solid #009FE3;
    border-radius:14px;
    padding:20px 22px;
    color:#0A2342;
    font-size:16px;
    line-height:1.6;
    margin-bottom:10px;
}
.start-shell {
    position: relative;
    min-height: 980px;
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-radius:18px;
    box-shadow: 0 1px 8px rgba(0,0,0,.035);
    overflow:hidden;
    margin-bottom: 20px;
}
.factor-box {
    position:absolute;
    width:240px;
    min-height:96px;
    background:transparent;
    border-left:8px solid #B8EC63;
    padding:18px 16px;
    display:flex;
    align-items:center;
    justify-content:center;
    text-align:center;
    color:#111827;
    font-size:15px;
    line-height:1.3;
    font-weight:600;
    box-shadow:0 1px 4px rgba(0,0,0,.04);
}
.funnel-wrap {
    position:absolute;
    left:50%;
    top:70px;
    transform:translateX(-50%);
    width:390px;
    z-index:2;
}
.funnel-label-top, .funnel-label-bottom {
    text-align:center;
    font-weight:800;
    color:#0A2342;
    margin-bottom:10px;
}
.funnel-seg {
    margin: 0 auto 10px auto;
    color:#ffffff;
    display:flex;
    align-items:center;
    justify-content:center;
    text-align:center;
    font-weight:700;
    border-radius:12px;
    padding:10px 14px;
    background:linear-gradient(135deg, #009FE3, #005F9E);
    box-shadow:0 1px 4px rgba(0,0,0,.08);
}
.funnel-top { width: 360px; height: 62px; background:linear-gradient(135deg, #12C76B, #009FE3); }
.funnel-mid1 { width: 300px; height: 58px; clip-path: polygon(6% 0, 94% 0, 100% 100%, 0 100%); }
.funnel-mid2 { width: 245px; height: 58px; clip-path: polygon(8% 0, 92% 0, 100% 100%, 0 100%); background:linear-gradient(135deg, #007EB5, #0A2342); }
.funnel-mid3 { width: 195px; height: 54px; clip-path: polygon(10% 0, 90% 0, 100% 100%, 0 100%); background:linear-gradient(135deg, #0A2342, #005F9E); }
.funnel-bottom { width: 138px; height: 62px; background:linear-gradient(135deg, #FFB000, #12C76B); }
.funnel-note {
    text-align:center;
    color:#6F7782;
    font-size:13px;
    line-height:1.45;
    margin-top:10px;
}
.method-card {
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-radius:14px;
    padding:18px 20px;
    box-shadow:0 1px 8px rgba(0,0,0,.035);
    height:100%;
}
.method-title {
    font-size:16px;
    font-weight:800;
    color:#0A2342;
    margin-bottom:8px;
}
.method-copy {
    color:#6F7782;
    font-size:14px;
    line-height:1.55;
}
.oem-pill-wrap {
    display:flex;
    flex-wrap:wrap;
    gap:10px;
}
.oem-pill {
    background:#EEF3F8;
    color:#0A2342;
    border:1px solid #D9E2EA;
    padding:8px 12px;
    border-radius:999px;
    font-size:13px;
    font-weight:700;
}


.start-factors-grid {
    display:grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap:16px;
    margin-bottom:20px;
}
.start-factor-card {
    background:#F5F6F7;
    border-left:8px solid #B8EC63;
    border-radius:8px;
    min-height:72px;
    padding:14px 16px;
    display:flex;
    align-items:center;
    justify-content:center;
    text-align:center;
    color:#111827;
    font-size:18px;
    line-height:1.3;
    font-weight:600;
    box-shadow:0 1px 4px rgba(0,0,0,.04);
}
@media (max-width: 900px) {
    .start-factors-grid { grid-template-columns: 1fr; }
}


.methodology-grid {
    display:grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap:18px;
    margin-bottom:20px;
}
.methodology-card {
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-radius:16px;
    padding:18px;
    box-shadow:0 1px 8px rgba(0,0,0,.035);
    height:100%;
}
.methodology-card h4 {
    margin:0 0 10px 0;
    font-size:20px;
    color:#0A2342;
}
.methodology-callout {
    display:inline-block;
    background:#EEF3F8;
    color:#0A2342;
    border-left:5px solid #009FE3;
    padding:8px 12px;
    border-radius:8px;
    font-size:13px;
    font-weight:700;
    margin-bottom:14px;
}
.methodology-formula {
    background:#F7F9FC;
    border:1px solid #E3E8EE;
    border-radius:10px;
    padding:12px 14px;
    font-size:14px;
    color:#111827;
    margin:10px 0 12px 0;
    line-height:1.45;
}
.methodology-card ul {
    margin: 10px 0 0 18px;
    color:#6F7782;
    line-height:1.6;
}
@media (max-width: 900px) {
    .methodology-grid { grid-template-columns: 1fr; }
}


.usecase-grid {
    display:grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap:18px;
    margin-bottom:14px;
}
.usecase-card {
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-radius:16px;
    padding:18px;
    box-shadow:0 1px 8px rgba(0,0,0,.035);
    min-height:255px;
}
.usecase-card.nmsc {
    border-left:7px solid #009FE3;
}
.usecase-card.tme {
    border-left:7px solid #6F6F6F;
}
.usecase-card.shared {
    border-left:7px solid #B8EC63;
}
.usecase-audience {
    display:inline-block;
    background:#EEF3F8;
    color:#0A2342;
    border-radius:999px;
    padding:5px 10px;
    font-size:12px;
    font-weight:800;
    margin-bottom:12px;
}
.usecase-title {
    font-size:20px;
    font-weight:800;
    color:#0A2342;
    margin-bottom:10px;
}
.usecase-copy {
    color:#6F7782;
    font-size:14px;
    line-height:1.55;
    margin-bottom:14px;
}
.usecase-report {
    background:#F7F9FC;
    border-radius:10px;
    padding:12px 14px;
    color:#111827;
    font-size:14px;
    line-height:1.5;
}
.usecase-report b {
    color:#0A2342;
}
.market-link-grid {
    display:grid;
    grid-template-columns: repeat(5, minmax(0, 1fr));
    gap:12px;
    margin:12px 0 22px 0;
}
.market-link-card {
    display:block;
    text-decoration:none !important;
    background:#000000;
    color:#ffffff !important;
    border-radius:14px;
    padding:16px;
    text-align:center;
    font-weight:800;
    box-shadow:0 1px 8px rgba(0,0,0,.08);
}
.market-link-card:hover {
    background:#009FE3;
    color:#000000 !important;
}
.usecase-note {
    background:#F7F9FC;
    border-left:6px solid #009FE3;
    border-radius:14px;
    padding:18px 20px;
    color:#0A2342;
    line-height:1.55;
    margin-bottom:10px;
}
@media (max-width: 900px) {
    .usecase-grid { grid-template-columns: 1fr; }
    .market-link-grid { grid-template-columns: 1fr; }
}


.cluster-legend {
    display:flex;
    flex-wrap:wrap;
    gap:10px;
    margin: 8px 0 18px 0;
}
.cluster-chip {
    display:flex;
    align-items:center;
    gap:8px;
    background:#F7F9FC;
    border:1px solid #E1E7EF;
    border-radius:999px;
    padding:8px 12px;
    color:#0A2342;
    font-size:13px;
}
.cluster-chip span:last-child {
    color:#8B95A1;
    font-weight:500;
}
.cluster-dot {
    display:inline-block;
    width:11px;
    height:11px;
    border-radius:4px;
}


.bubble-control-note {
    background:#F7F9FC;
    border-left:6px solid #009FE3;
    border-radius:14px;
    padding:16px 18px;
    color:#0A2342;
    line-height:1.5;
    margin: 8px 0 18px 0;
}
.bubble-side-card {
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-radius:16px;
    padding:18px;
    box-shadow:0 1px 8px rgba(0,0,0,.035);
    max-height:720px;
    overflow:auto;
}
.bubble-side-title {
    color:#0A2342;
    font-size:16px;
    font-weight:800;
    margin-bottom:12px;
}
.bubble-row {
    display:grid;
    grid-template-columns: 22px 1fr auto;
    gap:8px;
    align-items:center;
    border-bottom:1px solid #EEF2F6;
    padding:9px 0;
}
.bubble-row:last-child {
    border-bottom:none;
}
.bubble-row-dot {
    width:12px;
    height:12px;
    border-radius:4px;
}
.bubble-row-brand {
    color:#0A2342;
    font-weight:700;
    font-size:13px;
}
.bubble-row-meta {
    color:#8B95A1;
    font-size:12px;
}
.bubble-row-value {
    color:#0A2342;
    font-weight:800;
    font-size:13px;
    text-align:right;
}


.assistant-shell {
    background:#000000;
    color:#ffffff;
    border-radius:20px;
    padding:24px;
    border-bottom:7px solid #009FE3;
    margin: 10px 0 22px 0;
}
.assistant-title {
    font-size:26px;
    font-weight:800;
    letter-spacing:-0.02em;
    margin-bottom:8px;
}
.assistant-copy {
    color:#D9DDE3;
    font-size:15px;
    line-height:1.5;
    max-width:1100px;
}
.assistant-hint-row {
    display:flex;
    flex-wrap:wrap;
    gap:8px;
    margin-top:14px;
}
.assistant-hint {
    background:#ffffff;
    color:#0A2342;
    border-radius:999px;
    padding:7px 11px;
    font-size:12px;
    font-weight:800;
}
.assistant-result-grid {
    display:grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap:16px;
    margin: 16px 0 22px 0;
}
.assistant-card {
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-radius:16px;
    padding:20px;
    box-shadow:0 1px 8px rgba(0,0,0,.035);
    min-height:190px;
}
.assistant-card.primary {
    border-left:7px solid #009FE3;
}
.assistant-card.positive {
    border-left:7px solid #12C76B;
}
.assistant-card.risk {
    border-left:7px solid #FF5C8A;
}
.assistant-card.neutral {
    border-left:7px solid #6F6F6F;
}
.assistant-card-title {
    font-size:17px;
    line-height:1.25;
    color:#0A2342;
    font-weight:800;
    margin-bottom:8px;
}
.assistant-card-copy {
    color:#6F7782;
    font-size:14px;
    line-height:1.55;
    margin-bottom:12px;
}
.assistant-card-metric {
    color:#0A2342;
    font-size:28px;
    line-height:1.1;
    font-weight:800;
    margin-top:8px;
}
.assistant-card-tag {
    display:inline-block;
    background:#EEF3FF;
    color:#2563EB;
    border-radius:6px;
    padding:4px 9px;
    font-size:12px;
    font-weight:700;
    margin-top:10px;
}
@media (max-width: 900px) {
    .assistant-result-grid { grid-template-columns: 1fr; }
}


/* Data Assistant premium concierge treatment */
.da-hero {
    position: relative;
    min-height: 320px;
    border-radius: 26px;
    overflow: hidden;
    margin: 6px 0 28px 0;
    background:
        linear-gradient(90deg, rgba(0,0,0,0.72) 0%, rgba(0,0,0,0.48) 42%, rgba(10,35,66,0.30) 100%),
        radial-gradient(circle at 75% 20%, rgba(255,255,255,0.26), rgba(255,255,255,0) 32%),
        linear-gradient(135deg, #050505 0%, #2A2F36 48%, #C9B7AA 100%);
    border-bottom: 7px solid #009FE3;
    box-shadow: 0 12px 38px rgba(0,0,0,0.16);
}
.da-hero::before {
    content: "";
    position:absolute;
    left:0;
    top:0;
    width:38%;
    height:100%;
    background:
        repeating-linear-gradient(90deg, rgba(255,255,255,0.08) 0px, rgba(255,255,255,0.08) 4px, transparent 4px, transparent 18px);
    opacity: .72;
}
.da-hero-content {
    position: relative;
    z-index: 2;
    padding: 28px 38px 28px 38px;
    text-align: center;
    color:#ffffff;
}
.da-eyebrow {
    display:inline-block;
    background:rgba(255,255,255,0.14);
    border:1px solid rgba(255,255,255,0.24);
    border-radius:999px;
    padding:8px 14px;
    font-size:12px;
    font-weight:800;
    letter-spacing:.14em;
    text-transform:uppercase;
    margin-bottom:14px;
}
.da-title {
    font-size:46px;
    line-height:0.98;
    font-weight:800;
    letter-spacing:-0.055em;
    margin:0 auto 18px auto;
    max-width:900px;
}
.da-subtitle {
    font-size:17px;
    line-height:1.45;
    color:rgba(255,255,255,0.88);
    max-width:1040px;
    margin:0 auto 22px auto;
    font-weight:500;
}
.da-popular-label {
    color:#ffffff;
    font-size:15px;
    font-weight:800;
    letter-spacing:.16em;
    text-transform:uppercase;
    margin-bottom:10px;
}
.da-popular-grid {
    display:flex;
    justify-content:center;
    flex-wrap:wrap;
    gap:18px;
}
.da-popular-pill {
    background:#ffffff;
    color:#101820;
    border-radius:12px;
    padding:12px 18px;
    font-size:16px;
    font-weight:700;
    box-shadow:0 3px 14px rgba(0,0,0,.16);
    min-width:230px;
}
.da-privacy {
    position:absolute;
    bottom:14px;
    left:0;
    right:0;
    text-align:center;
    color:rgba(255,255,255,0.72);
    font-size:13px;
    letter-spacing:.02em;
    z-index:2;
}
.da-search-wrap {
    background:#ffffff;
    border:1.5px solid rgba(167, 100, 70, .55);
    border-radius:999px;
    padding:8px 18px 8px 20px;
    display:flex;
    align-items:center;
    gap:14px;
    min-height:68px;
    box-shadow:0 8px 30px rgba(0,0,0,.08);
    margin: 6px 0 26px 0;
}
.da-search-icon {
    color:#B87554;
    font-size:24px;
    font-weight:900;
}
.da-search-wrap div[data-testid="stTextInput"] {
    flex:1;
}
.da-search-wrap div[data-testid="stTextInput"] > label {
    display:none;
}
.da-search-wrap div[data-testid="stTextInput"] input {
    border:0 !important;
    box-shadow:none !important;
    font-size:21px !important;
    color:#101820 !important;
    background:#ffffff !important;
}
.da-search-wrap div[data-testid="stTextInput"] input:focus {
    border:0 !important;
    box-shadow:none !important;
}
.da-context-bar {
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:14px;
    margin: 6px 0 18px 0;
    color:#6F7782;
    font-size:14px;
}
.da-context-chip {
    background:#F7F9FC;
    border:1px solid #E6E9ED;
    color:#0A2342;
    border-radius:999px;
    padding:8px 12px;
    font-weight:800;
}
.assistant-result-grid {
    display:grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap:18px;
    margin: 16px 0 22px 0;
}
.assistant-card {
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-radius:18px;
    padding:18px;
    box-shadow:0 4px 18px rgba(10,35,66,.06);
    min-height:175px;
}
.assistant-card.primary {
    border-left:7px solid #009FE3;
}
.assistant-card.positive {
    border-left:7px solid #12C76B;
}
.assistant-card.risk {
    border-left:7px solid #FF5C8A;
}
.assistant-card.neutral {
    border-left:7px solid #B87554;
}
.assistant-card-title {
    font-size:18px;
    line-height:1.25;
    color:#0A2342;
    font-weight:850;
    margin-bottom:10px;
}
.assistant-card-copy {
    color:#6F7782;
    font-size:14px;
    line-height:1.58;
    margin-bottom:14px;
}
.assistant-card-metric {
    color:#000000;
    font-size:27px;
    line-height:1.1;
    font-weight:850;
    margin-top:8px;
}
.assistant-card-tag {
    display:inline-block;
    background:#F3F6FA;
    color:#0A2342;
    border-radius:8px;
    padding:5px 10px;
    font-size:12px;
    font-weight:800;
    margin-top:12px;
}
@media (max-width: 900px) {
    .da-hero { min-height: 620px; }
    .da-hero-content { padding: 40px 22px; }
    .da-title { font-size:46px; }
    .da-subtitle { font-size:18px; margin-bottom:40px; }
    .da-popular-pill { width:100%; min-width:0; }
    .assistant-result-grid { grid-template-columns: 1fr; }
}




/* Data Assistant input cleanup */
div[data-testid="stTextInput"] input {
    min-height: 48px !important;
    font-size: 18px !important;
    padding: 10px 14px !important;
}


/* Simplified Data Assistant */
.da-simple {
    margin: 8px 0 22px 0;
}
.da-simple-title {
    color:#0A2342;
    font-size:44px;
    line-height:1.05;
    font-weight:850;
    letter-spacing:-0.04em;
    margin-bottom:10px;
}
.da-simple-copy {
    color:#6F7782;
    font-size:18px;
    line-height:1.45;
    max-width:1120px;
    margin-bottom:22px;
}
.da-simple-label {
    color:#8B95A1;
    font-size:13px;
    font-weight:850;
    letter-spacing:.16em;
    text-transform:uppercase;
    margin-bottom:12px;
}
.da-simple-questions {
    display:flex;
    flex-wrap:wrap;
    gap:12px;
    margin-bottom:18px;
}
.da-simple-question {
    background:#ffffff;
    color:#0A2342;
    border:1px solid #E1E7EF;
    border-radius:14px;
    padding:13px 16px;
    font-size:15px;
    font-weight:800;
    box-shadow:0 2px 10px rgba(10,35,66,.05);
}
.da-simple-note {
    color:#8B95A1;
    font-size:13px;
    margin:8px 0 18px 0;
}
@media (max-width: 900px) {
    .da-simple-title { font-size:34px; }
    .da-simple-copy { font-size:16px; }
    .da-simple-question { width:100%; }
}


/* Toyota / Lexus logo sizing cleanup */
.tl-logo-strip img[src*="ToyotaLogo"],
.tl-logo-strip img[src*="BHUB_Logo_ToyotaLogo"],
.brand-strip img[src*="ToyotaLogo"],
.brand-strip img[src*="BHUB_Logo_ToyotaLogo"],
img[src*="BHUB_Logo_ToyotaLogo_01.svg"] {
    background: transparent !important;
    background-color: transparent !important;
    box-shadow: none !important;
    border: 0 !important;
    padding: 0 !important;
    width: 210px !important;
    max-width: 210px !important;
    height: auto !important;
    object-fit: contain !important;
}

.tl-logo-strip,
.brand-strip {
    background: transparent !important;
}


/* Executive Market Summary */
.exec-kpi-grid {
    display:grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap:16px;
    margin: 12px 0 22px 0;
}
.exec-kpi-card {
    background:#F7F9FC;
    border-radius:18px;
    padding:20px;
    border:1px solid #E6E9ED;
    min-height:145px;
}
.exec-kpi-label {
    color:#6F7782;
    font-size:14px;
    font-weight:800;
    margin-bottom:10px;
}
.exec-kpi-value {
    color:#000000;
    font-size:34px;
    line-height:1.05;
    font-weight:850;
}
.exec-kpi-delta {
    display:inline-block;
    margin-top:12px;
    border-radius:999px;
    padding:5px 10px;
    font-weight:800;
    font-size:13px;
    background:#EEF3F8;
    color:#0A2342;
}
.share-grid {
    display:grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap:18px;
    margin: 12px 0 24px 0;
}
.share-card {
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-radius:18px;
    padding:22px;
    box-shadow:0 2px 14px rgba(10,35,66,.04);
}
.share-title {
    font-size:20px;
    font-weight:850;
    color:#0A2342;
    margin-bottom:14px;
}
.share-metric-row {
    display:flex;
    justify-content:space-between;
    gap:12px;
    border-top:1px solid #EEF2F6;
    padding:12px 0;
}
.share-metric-row:first-of-type {
    border-top:0;
}
.share-label {
    color:#6F7782;
    font-weight:700;
}
.share-value {
    color:#000000;
    font-weight:850;
}
.takeaway-grid {
    display:grid;
    grid-template-columns: repeat(5, minmax(0, 1fr));
    gap:14px;
    margin: 12px 0 22px 0;
}
.takeaway-card {
    background:#ffffff;
    border:1px solid #E6E9ED;
    border-left:7px solid #009FE3;
    border-radius:16px;
    padding:18px;
    min-height:185px;
    box-shadow:0 2px 14px rgba(10,35,66,.04);
}
.takeaway-card.risk { border-left-color:#FF5C8A; }
.takeaway-card.positive { border-left-color:#12C76B; }
.takeaway-card.neutral { border-left-color:#B87554; }
.takeaway-label {
    font-size:11px;
    letter-spacing:.14em;
    text-transform:uppercase;
    font-weight:850;
    color:#8B95A1;
    margin-bottom:10px;
}
.takeaway-title {
    color:#0A2342;
    font-size:16px;
    font-weight:850;
    line-height:1.25;
    margin-bottom:9px;
}
.takeaway-copy {
    color:#6F7782;
    font-size:13px;
    line-height:1.5;
}
@media (max-width: 1100px) {
    .exec-kpi-grid, .takeaway-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
}
@media (max-width: 800px) {
    .exec-kpi-grid, .share-grid, .takeaway-grid { grid-template-columns: 1fr; }
}


/* Toyota logo render fix */
img[src^="data:image/png"][alt="Toyota logo"] {
    background: transparent !important;
    border: 0 !important;
    box-shadow: none !important;
    padding: 0 !important;
    object-fit: contain !important;
}
.brand-strip img[alt="Toyota logo"],
.tl-detail-logo img[src^="data:image/png"] {
    max-height: 64px !important;
    width: auto !important;
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
def normalise_period(value):
    text_value = str(value).strip()
    if text_value.endswith(".0"):
        text_value = text_value[:-2]
    return text_value


@st.cache_data
def load_data():
    raw = pd.read_excel(DATA_FILE, sheet_name="CORE DATA CUT")
    raw.columns = [str(c).strip() for c in raw.columns]

    brand_col = find_column(raw.columns, ["brand", "oem"])
    market_col = find_column(raw.columns, ["market", "country"])
    period_col = find_column(raw.columns, ["period", "year"])
    sales_col = find_column(raw.columns, ["sales"])
    uv_col = find_column(raw.columns, ["unique visitors (dedupe)", "unique visitors", "unique", "uv"])

    df = raw[[brand_col, market_col, period_col, sales_col, uv_col]].copy()
    df = df.rename(
        columns={
            brand_col: "OEM",
            market_col: "Market",
            period_col: "Period",
            sales_col: "Sales",
            uv_col: "UniqueVisitors",
        }
    )

    df["OEM"] = df["OEM"].astype(str).str.strip()
    df["Market"] = df["Market"].astype(str).str.strip()
    df["Period"] = df["Period"].map(normalise_period)
    df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")
    df["UniqueVisitors"] = pd.to_numeric(df["UniqueVisitors"], errors="coerce")

    df = df.dropna(subset=["OEM", "Market", "Period", "Sales", "UniqueVisitors"])
    df = df[(df["UniqueVisitors"] > 0) & (df["Sales"] >= 0)]

    market_data = (
        df.groupby(["OEM", "Market", "Period"], as_index=False)
        .agg({"Sales": "sum", "UniqueVisitors": "sum"})
    )
    market_data["ConversionPct"] = market_data["Sales"] / market_data["UniqueVisitors"] * 100

    mm5 = (
        market_data.groupby(["OEM", "Period"], as_index=False)
        .agg({"Sales": "sum", "UniqueVisitors": "sum"})
    )
    mm5["Market"] = "MM5"
    mm5["ConversionPct"] = mm5["Sales"] / mm5["UniqueVisitors"] * 100

    combined = pd.concat(
        [
            mm5[["OEM", "Market", "Period", "Sales", "UniqueVisitors", "ConversionPct"]],
            market_data[["OEM", "Market", "Period", "Sales", "UniqueVisitors", "ConversionPct"]],
        ],
        ignore_index=True,
    )

    return combined.sort_values(["Market", "OEM", "Period"]).reset_index(drop=True)


def apply_comparison_view(data, comparison_name):
    config = COMPARISON_OPTIONS[comparison_name]
    previous_period = config["previous_period"]
    current_period = config["current_period"]

    active = data[data["Period"].isin([previous_period, current_period])].copy()
    active["Year"] = active["Period"].map({previous_period: 2024, current_period: 2025})
    active["PeriodLabel"] = active["Period"].map(
        {
            previous_period: config["previous_label"],
            current_period: config["current_label"],
        }
    )

    return active




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


def fmt_metric_number(v):
    """Format KPI numbers so they do not wrap awkwardly in metric cards."""
    if pd.isna(v):
        return "n/a"
    v = float(v)
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:.1f}M"
    if abs(v) >= 1_000:
        return f"{v:,.0f}"
    return f"{v:.0f}"


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


def normalise_cluster_oem(oem):
    if oem == "VW":
        return "Volkswagen"
    if oem == "Citroen":
        return "Citroën"
    if oem == "BYD":
        return "BYD Auto"
    return oem


def cluster_for_oem(oem):
    normalised = normalise_cluster_oem(oem)
    for cluster, brands in OEM_CLUSTERS.items():
        normalised_brands = [normalise_cluster_oem(b) for b in brands]
        if normalised in normalised_brands:
            return cluster
    return "Uncategorised"


def oems_for_clusters(clusters, available_oems):
    if not clusters or "All categories" in clusters:
        return available_oems

    allowed = set()
    for cluster in clusters:
        for brand in OEM_CLUSTERS.get(cluster, []):
            normalised = normalise_cluster_oem(brand)
            for available in available_oems:
                if normalise_cluster_oem(available) == normalised:
                    allowed.add(available)

    return sorted(allowed)


def filter_oems_by_cluster(selected_oems, clusters, available_oems):
    allowed = oems_for_clusters(clusters, available_oems)
    if not selected_oems:
        return allowed
    return sorted([oem for oem in selected_oems if oem in allowed])


def render_cluster_legend(selected_clusters):
    active = [c for c in selected_clusters if c != "All categories"]
    if not active:
        active = list(OEM_CLUSTERS.keys())

    chips = ""
    for cluster in active:
        color = CLUSTER_COLORS.get(cluster, "#6F6F6F")
        desc = CLUSTER_DESCRIPTIONS.get(cluster, "")
        chips += (
            f"<div class='cluster-chip'>"
            f"<span class='cluster-dot' style='background:{color};'></span>"
            f"<b>{cluster}</b><span>{desc}</span>"
            f"</div>"
        )

    st.markdown(f"<div class='cluster-legend'>{chips}</div>", unsafe_allow_html=True)


# =========================
# Visual components
# =========================

def render_hero():
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-logo"><img src="{VALTECH_LOGO}" alt="Valtech logo"></div>
            <div class="hero-title">OEM Macro Conversion Funnel</div>
            <div class="hero-subtitle">
                Website to customer contract conversion across MM5. Explores how unique visitor demand converts into passenger new car customer contracts, and where Toyota/Lexus under- or over-perform by market.
            </div>
            <div class="hero-meta">
                <b>Coverage:</b> 2024, 2025 &amp; 2026 (Jan to April) &nbsp;|&nbsp;
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
        f"""
        <div class="footer-note">
        <b>Definitions and caveats:</b> This is not a total market-share dashboard. It uses passenger car sales from Marklines and <b>monthly deduplicated unique visitor</b> website data from Similarweb. It does not include fleet, LCV or tactical registrations. <b>Website-to-Contract Conversion Rate</b> = passenger car sales / monthly deduplicated unique website visitors. Active comparison: <b>{PREVIOUS_LABEL} vs {CURRENT_LABEL}</b>.
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_brand_strip():
    st.markdown(
        f"""
        <div class="brand-strip">
            <img src="{TOYOTA_LOGO}" style="height:58px; width:auto; object-fit:contain;" alt="Toyota logo">
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
        f"<div class='tl-detail-label'>Website-to-Contract Conversion Rate</div>"
        f"<div class='tl-detail-value'>{row['ConversionPct_2025']:.2f}%</div>"
        f"<div class='{conv_delta_class}'>{fmt_pp(row['Conv Var pp'])} vs previous period</div>"
        "</div>"
        "<div class='tl-detail-metric'>"
        f"<div class='tl-detail-label'>Passenger sales</div>"
        f"<div class='tl-detail-value'>{fmt_metric_number(row['Sales_2025'])}</div>"
        f"<div class='{sales_delta_class}'>{fmt_pct(row['Sales YoY %'])} vs previous period</div>"
        "</div>"
        "<div class='tl-detail-metric'>"
        f"<div class='tl-detail-label'>Unique visitors</div>"
        f"<div class='tl-detail-value'>{fmt_metric_number(row['UniqueVisitors_2025'])}</div>"
        f"<div class='{visitor_delta_class}'>{fmt_pct(row['Visitors YoY %'])} vs previous period</div>"
        "</div>"
        "</div>"
        "</div>"
    )


def render_brand_detail(data, market):
    section("2025 Toyota & Lexus performance vs previous period")
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
    section("Toyota / Lexus Website-to-Contract benchmark callouts")
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
                f"{brand} ranks #{rank} of {len(cohort_df)} in its benchmark set. The Website-to-Contract Conversion Rate gap to {leader['OEM']} is {gap:+.2f}pp. Visits-to-sale gap is {vts_gap:+.0f}.",
                f"Website-to-Contract Conversion Rate: {row['ConversionPct_2025']:.2f}%",
            )


def render_market_weakness(data):
    section("Market Gap Analysis — Toyota & Lexus Website-to-Contract performance")
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
                    "Visitor YoY vs Prev Period": row["Visitors YoY %"],
                    "Sales YoY vs Previous Period": row["Sales YoY %"],
                    "W2C Conv Rate": f"{row['ConversionPct_2025']:.2f}%",
                    "Benchmark Leader": leader["OEM"],
                    "Leader W2C Rate": f"{leader['ConversionPct_2025']:.2f}%",
                    "Gap to Leader": row["ConversionPct_2025"] - leader["ConversionPct_2025"],
                }
            )
    if not rows:
        st.info("No Toyota/Lexus market gap data available.")
        return

    df = pd.DataFrame(rows).sort_values("Gap to Leader")
    display = df[
        [
            "Brand",
            "Market",
            "Visitor YoY vs Prev Period",
            "Sales YoY vs Previous Period",
            "W2C Conv Rate",
            "Benchmark Leader",
            "Leader W2C Rate",
            "Gap to Leader",
        ]
    ].copy()

    display["Visitor YoY vs Prev Period"] = display["Visitor YoY vs Prev Period"].map(fmt_pct)
    display["Sales YoY vs Previous Period"] = display["Sales YoY vs Previous Period"].map(fmt_pct)
    display["Gap to Leader"] = display["Gap to Leader"].map(fmt_pp)

    badge_cols = ["Visitor YoY vs Prev Period", "Sales YoY vs Previous Period", "Gap to Leader"]
    styler = display.style.map(badge_style, subset=badge_cols)
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
                f"Sales are {fmt_pct(row['Sales YoY %'])} vs previous period; visitors are {fmt_pct(row['Visitors YoY %'])} vs previous period. "
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
                "Category": cluster_for_oem(r["OEM"]),
                f"Visits current period": fmt_short(r["UniqueVisitors_2025"]),
                f"Visits previous period": fmt_short(r["UniqueVisitors_2024"]),
                "Visitor YoY vs previous period": fmt_pct(r["Visitors YoY %"]),
                "Passenger sales": fmt_int(r["Sales_2025"]),
                "W2C rate": f"{r['ConversionPct_2025']:.2f}%",
                "W2C var vs previous period": fmt_pp(r["Conv Var pp"]),
            }
        )
    display = pd.DataFrame(rows)
    st.dataframe(
        display.style.map(badge_style, subset=["Visitor YoY vs previous period", "W2C var vs previous period"]),
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
        title=dict(text=f"YoY visit growth — 2025 vs previous period (top & bottom {top_n})", x=0.02, font=dict(size=16, color="#8D96A0")),
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
        make_insight("opportunity", f"{top_conv['OEM']}: strongest Website-to-Contract performer", f"{top_conv['OEM']} leads this cohort on Website-to-Contract Conversion Rate. Visits-to-sale is {top_conv['Visits to Sale 2025']:.0f}, showing stronger website-to-contract efficiency than peers.", f"{top_conv['ConversionPct_2025']:.2f}% conv", top_conv["OEM"]),
        make_insight("opportunity", f"{top_sales['OEM']}: strongest sales growth", f"{top_sales['OEM']} posted the strongest passenger sales growth, with sales {fmt_pct(top_sales['Sales YoY %'])} vs previous period and visitors {fmt_pct(top_sales['Visitors YoY %'])} vs previous period.", f"{fmt_pct(top_sales['Sales YoY %'])} sales", top_sales["OEM"]),
        make_insight("intelligence", f"{top_visits['OEM']}: fastest visitor growth", f"{top_visits['OEM']} generated the strongest visitor growth. The key question is whether this awareness converts into contracts at the same rate.", f"{fmt_pct(top_visits['Visitors YoY %'])} visits", top_visits["OEM"]),
        make_insight("risk", f"{worst_conv['OEM']}: Website-to-Contract movement deteriorating", f"{worst_conv['OEM']} had the weakest conversion movement: {fmt_pp(worst_conv['Conv Var pp'])} vs previous period. This suggests lower-quality traffic or leakage deeper in the funnel.", f"{fmt_pp(worst_conv['Conv Var pp'])}", worst_conv["OEM"]),
        make_insight("risk", f"{weakest['OEM']}: weakest Website-to-Contract Conversion Rate", f"{weakest['OEM']} has the lowest Website-to-Contract Conversion Rate rate in this selection, creating a clear efficiency gap against the cohort leader.", f"{weakest['ConversionPct_2025']:.2f}% conv", weakest["OEM"]),
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
    max_x = max(chart_df["UniqueVisitors"].max() * 1.18, 1)
    max_y = max(chart_df["ConversionPct"].max() * 1.25, 0.1)

    selected_2025 = chart_df[chart_df["Year"] == 2025].copy()
    x_mid = selected_2025["UniqueVisitors"].median() if not selected_2025.empty else chart_df["UniqueVisitors"].median()
    y_mid = selected_2025["ConversionPct"].median() if not selected_2025.empty else chart_df["ConversionPct"].median()

    # Quadrant background zones.
    fig.add_shape(type="rect", x0=0, x1=x_mid, y0=y_mid, y1=max_y, fillcolor="rgba(0,159,227,0.05)", line=dict(width=0), layer="below")
    fig.add_shape(type="rect", x0=x_mid, x1=max_x, y0=y_mid, y1=max_y, fillcolor="rgba(18,199,107,0.07)", line=dict(width=0), layer="below")
    fig.add_shape(type="rect", x0=0, x1=x_mid, y0=0, y1=y_mid, fillcolor="rgba(111,111,111,0.04)", line=dict(width=0), layer="below")
    fig.add_shape(type="rect", x0=x_mid, x1=max_x, y0=0, y1=y_mid, fillcolor="rgba(255,176,0,0.06)", line=dict(width=0), layer="below")

    # Median guide lines.
    fig.add_shape(type="line", x0=x_mid, x1=x_mid, y0=0, y1=max_y, line=dict(color="rgba(10,35,66,0.12)", width=1), layer="below")
    fig.add_shape(type="line", x0=0, x1=max_x, y0=y_mid, y1=y_mid, line=dict(color="rgba(10,35,66,0.12)", width=1), layer="below")

    # Quadrant labels.
    fig.add_annotation(x=max_x * 0.02, y=max_y * 0.95, text="Efficient challengers", showarrow=False, font=dict(color="#8D96A0", size=14), xanchor="left")
    fig.add_annotation(x=max_x * 0.98, y=max_y * 0.95, text="Leading players", showarrow=False, font=dict(color="#8D96A0", size=14), xanchor="right")
    fig.add_annotation(x=max_x * 0.02, y=max_y * 0.06, text="Developing players", showarrow=False, font=dict(color="#8D96A0", size=14), xanchor="left")
    fig.add_annotation(x=max_x * 0.98, y=max_y * 0.06, text="Scale without efficiency", showarrow=False, font=dict(color="#8D96A0", size=14), xanchor="right")

    # Movement arrows from 2024 to 2025.
    if year_view == "Previous and current + shift":
        for oem in selected_oems:
            d24 = chart_df[(chart_df["OEM"] == oem) & (chart_df["Year"] == 2024)]
            d25 = chart_df[(chart_df["OEM"] == oem) & (chart_df["Year"] == 2025)]
            if len(d24) and len(d25):
                r24 = d24.iloc[0]
                r25 = d25.iloc[0]
                cluster = cluster_for_oem(oem)
                color = CLUSTER_COLORS.get(cluster, "#6F6F6F")
                fig.add_annotation(
                    x=r25["UniqueVisitors"],
                    y=r25["ConversionPct"],
                    ax=r24["UniqueVisitors"],
                    ay=r24["ConversionPct"],
                    xref="x",
                    yref="y",
                    axref="x",
                    ayref="y",
                    showarrow=True,
                    arrowhead=3,
                    arrowsize=1.2,
                    arrowwidth=1.6,
                    arrowcolor=color,
                    opacity=0.78,
                )

    # 2024 faded markers.
    if year_view == "Previous and current + shift":
        df24 = chart_df[chart_df["Year"] == 2024].copy()
        if not df24.empty:
            fig.add_trace(
                go.Scatter(
                    x=df24["UniqueVisitors"],
                    y=df24["ConversionPct"],
                    mode="markers",
                    name="Previous period",
                    marker=dict(
                        size=df24["Sales"].apply(lambda s: max(18, math.sqrt(s / max_sales) * 72)),
                        color="rgba(190,198,208,0.45)",
                        line=dict(width=1.4, color="rgba(10,35,66,0.25)"),
                        symbol="circle",
                    ),
                    customdata=df24[["OEM", "Market", "Year", "Sales", "UniqueVisitors", "ConversionPct"]],
                    hovertemplate="<b>%{customdata[0]}</b><br>Year: %{customdata[2]}<br>Sales: %{customdata[3]:,.0f}<br>Visitors: %{customdata[4]:,.0f}<br>W2C Rate: %{customdata[5]:.2f}%<extra></extra>",
                )
            )

    primary_year = 2025 if year_view == "Previous and current + shift" else (2024 if year_view == "Previous period" else 2025)
    primary = chart_df[chart_df["Year"] == primary_year].copy()

    if not primary.empty:
        primary["Cluster"] = primary["OEM"].map(cluster_for_oem)
        for cluster in sorted(primary["Cluster"].unique()):
            dfc = primary[primary["Cluster"] == cluster].copy()
            color = CLUSTER_COLORS.get(cluster, "#6F6F6F")
            sizes = dfc["Sales"].apply(lambda s: max(24, math.sqrt(s / max_sales) * 88))
            fig.add_trace(
                go.Scatter(
                    x=dfc["UniqueVisitors"],
                    y=dfc["ConversionPct"],
                    mode="markers+text" if not show_logos else "markers",
                    text=None if show_logos else dfc["OEM"],
                    textposition="top center",
                    name=f"{cluster} · {primary_year}",
                    marker=dict(
                        size=sizes,
                        color=color,
                        opacity=0.78,
                        line=dict(width=2.2, color="white"),
                    ),
                    customdata=dfc[["OEM", "Market", "Year", "Sales", "UniqueVisitors", "ConversionPct", "Cluster"]],
                    hovertemplate="<b>%{customdata[0]}</b><br>Category: %{customdata[6]}<br>Year: %{customdata[2]}<br>Sales: %{customdata[3]:,.0f}<br>Visitors: %{customdata[4]:,.0f}<br>W2C Rate: %{customdata[5]:.2f}%<extra></extra>",
                )
            )

    if show_logos and not primary.empty:
        add_logo_images(fig, primary, max_x, max_y)

    fig.update_layout(
        title=dict(
            text=f"{market} | Audience scale vs Website-to-Contract Conversion Rate",
            x=0.02,
            y=0.98,
            font=dict(size=20, color="#0A2342"),
        ),
        xaxis_title="Audience size: unique website visitors",
        yaxis_title="Website-to-Contract Conversion Rate",
        height=760,
        hovermode="closest",
        legend=dict(orientation="h", y=1.08, x=0, font=dict(size=11)),
        margin=dict(l=75, r=45, t=105, b=75),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
    )
    fig.update_xaxes(
        range=[0, max_x],
        tickformat=",",
        gridcolor="#EDF1F5",
        zeroline=False,
        showline=True,
        linecolor="#E1E7EF",
    )
    fig.update_yaxes(
        range=[0, max_y],
        ticksuffix="%",
        gridcolor="#EDF1F5",
        zeroline=False,
        showline=True,
        linecolor="#E1E7EF",
    )

    return fig



def assistant_card(title, copy, metric, tag, style="primary"):
    return (
        f"<div class='assistant-card {style}'>"
        f"<div class='assistant-card-title'>{title}</div>"
        f"<div class='assistant-card-copy'>{copy}</div>"
        f"<div class='assistant-card-metric'>{metric}</div>"
        f"<div class='assistant-card-tag'>{tag}</div>"
        f"</div>"
    )


def extract_market_from_question(question):
    q = question.lower()
    market_aliases = {
        "mm5": "MM5",
        "uk": "UK",
        "united kingdom": "UK",
        "france": "France",
        "french": "France",
        "germany": "Germany",
        "german": "Germany",
        "italy": "Italy",
        "italian": "Italy",
        "spain": "Spain",
        "spanish": "Spain",
    }
    for key, market in market_aliases.items():
        if key in q:
            return market
    return "MM5"


def extract_brand_from_question(question, available_oems):
    q = question.lower()

    aliases = {
        "volkswagen": "VW",
        "vw": "VW",
        "byd": "BYD Auto",
        "byd auto": "BYD Auto",
        "ds automobiles": "DS",
        "citroen": "Citroen",
    }

    # Match aliases and brand names as whole tokens/phrases only.
    # This prevents short brands such as "DS" being detected inside words like "leads".
    candidates = []
    for brand in available_oems:
        candidates.append((brand.lower(), brand))
    for alias, brand in aliases.items():
        if brand in available_oems:
            candidates.append((alias.lower(), brand))

    for token, brand in sorted(candidates, key=lambda x: len(x[0]), reverse=True):
        pattern = r"(?<![a-z0-9])" + re.escape(token) + r"(?![a-z0-9])"
        if re.search(pattern, q):
            return brand

    return None


def assistant_result_html(cards):
    return "<div class='assistant-result-grid'>" + "".join(cards) + "</div>"


def top_brands_cards(data, market, selected_oems, question):
    q = question.lower()
    yoy = yoy_table(data, market, selected_oems)
    if yoy.empty:
        return [assistant_card("No data available", f"No comparable data was found for {market}.", "—", market, "neutral")]

    if "sales" in q:
        metric = "Sales_2025"
        label = "passenger sales"
        formatter = fmt_metric_number
    elif "visitor" in q or "traffic" in q or "audience" in q:
        metric = "UniqueVisitors_2025"
        label = "unique visitors"
        formatter = fmt_metric_number
    else:
        metric = "ConversionPct_2025"
        label = "Website-to-Contract Conversion Rate"
        formatter = lambda x: f"{x:.2f}%"

    ascending = any(term in q for term in ["worst", "weakest", "lowest", "underperform"])
    ranked = yoy.sort_values(metric, ascending=ascending).head(5)

    cards = []
    for idx, (_, row) in enumerate(ranked.iterrows(), start=1):
        style = "risk" if ascending else ("positive" if idx == 1 else "primary")
        direction = "lowest" if ascending else "highest"
        cards.append(
            assistant_card(
                f"#{idx}: {row['OEM']}",
                f"In {market}, {row['OEM']} is ranked #{idx} for {label} in the selected set. "
                f"Sales changed {fmt_pct(row['Sales YoY %'])}; unique visitors changed {fmt_pct(row['Visitors YoY %'])} versus {PREVIOUS_LABEL}.",
                formatter(row[metric]),
                f"{direction} {label}",
                style,
            )
        )

    leader = ranked.iloc[0]
    cards.append(
        assistant_card(
            "Read this correctly",
            f"This is not proof that {leader['OEM']} has the best website. W2C performance can reflect brand demand, product fit, pricing, stock, retailer follow-up and network effects.",
            "macro funnel",
            f"{CURRENT_LABEL} vs {PREVIOUS_LABEL}",
            "neutral",
        )
    )

    return cards


def brand_cards(data, market, brand):
    yoy = yoy_table(data, market, None)
    if yoy.empty:
        return [assistant_card("No data available", f"No comparable data was found for {market}.", "—", market, "neutral")]

    row = get_row(yoy, brand)
    if row is None:
        return [assistant_card("Brand not found", f"I could not find {brand} in {market} for the active comparison.", "—", brand, "neutral")]

    cohort = LEXUS_SET if brand == "Lexus" else TOYOTA_SET if brand == "Toyota" else sorted(yoy["OEM"].unique().tolist())
    cohort_df = yoy[yoy["OEM"].isin(cohort)].copy()
    leader = cohort_df.sort_values("ConversionPct_2025", ascending=False).iloc[0] if not cohort_df.empty else yoy.sort_values("ConversionPct_2025", ascending=False).iloc[0]
    gap = row["ConversionPct_2025"] - leader["ConversionPct_2025"]
    rank = int(cohort_df["ConversionPct_2025"].rank(method="min", ascending=False).loc[row.name]) if row.name in cohort_df.index else int(yoy["ConversionPct_2025"].rank(method="min", ascending=False).loc[row.name])

    visitor_style = "positive" if row["Visitors YoY %"] >= 0 else "risk"
    sales_style = "positive" if row["Sales YoY %"] >= 0 else "risk"
    conv_style = "positive" if row["Conv Var pp"] >= 0 else "risk"

    cards = [
        assistant_card(
            f"{brand} W2C performance in {market}",
            f"{brand}'s {CURRENT_LABEL} Website-to-Contract Conversion Rate is ranked #{rank} within the relevant comparison set.",
            f"{row['ConversionPct_2025']:.2f}%",
            "W2C rate",
            "primary",
        ),
        assistant_card(
            "Unique visitor movement",
            f"Unique visitor demand moved versus {PREVIOUS_LABEL}. This helps separate demand generation from conversion efficiency.",
            fmt_pct(row["Visitors YoY %"]),
            fmt_metric_number(row["UniqueVisitors_2025"]),
            visitor_style,
        ),
        assistant_card(
            "Passenger sales movement",
            f"Passenger sales moved versus {PREVIOUS_LABEL}. Read this alongside visitor movement to understand whether demand is translating into contracts.",
            fmt_pct(row["Sales YoY %"]),
            fmt_metric_number(row["Sales_2025"]),
            sales_style,
        ),
        assistant_card(
            "W2C rate movement",
            f"The W2C rate movement indicates whether conversion efficiency improved or deteriorated versus the prior period.",
            fmt_pp(row["Conv Var pp"]),
            f"vs {PREVIOUS_LABEL}",
            conv_style,
        ),
        assistant_card(
            "Benchmark leader gap",
            f"The benchmark leader is {leader['OEM']} at {leader['ConversionPct_2025']:.2f}%.",
            fmt_pp(gap),
            f"gap to {leader['OEM']}",
            "positive" if gap >= 0 else "risk",
        ),
        assistant_card(
            "Likely diagnostic route",
            "Use this as a prompt for deeper diagnosis across marketing quality, model-page content, offers, stock visibility, lead handling and retailer follow-up.",
            "next step",
            "diagnostic",
            "neutral",
        ),
    ]
    return cards


def toyota_lexus_cards(data, market):
    yoy = yoy_table(data, market, None)
    if yoy.empty:
        return [assistant_card("No data available", f"No comparable data was found for {market}.", "—", market, "neutral")]

    cards = []
    for brand in ["Toyota", "Lexus"]:
        row = get_row(yoy, brand)
        if row is None:
            continue
        cohort = TOYOTA_SET if brand == "Toyota" else LEXUS_SET
        cohort_df = yoy[yoy["OEM"].isin(cohort)].copy()
        leader = cohort_df.sort_values("ConversionPct_2025", ascending=False).iloc[0]
        gap = row["ConversionPct_2025"] - leader["ConversionPct_2025"]
        cards.extend(
            [
                assistant_card(
                    f"{brand}: W2C rate in {market}",
                    f"{brand}'s {CURRENT_LABEL} Website-to-Contract Conversion Rate is compared against its benchmark cohort.",
                    f"{row['ConversionPct_2025']:.2f}%",
                    brand,
                    "primary",
                ),
                assistant_card(
                    f"{brand}: benchmark gap",
                    f"The benchmark leader is {leader['OEM']} at {leader['ConversionPct_2025']:.2f}%.",
                    fmt_pp(gap),
                    f"leader: {leader['OEM']}",
                    "positive" if gap >= 0 else "risk",
                ),
                assistant_card(
                    f"{brand}: demand vs sales movement",
                    f"Unique visitors moved {fmt_pct(row['Visitors YoY %'])}; sales moved {fmt_pct(row['Sales YoY %'])} versus {PREVIOUS_LABEL}.",
                    f"{fmt_pct(row['Visitors YoY %'])} visitors",
                    f"{fmt_pct(row['Sales YoY %'])} sales",
                    "positive" if row["Conv Var pp"] >= 0 else "risk",
                ),
            ]
        )

    cards.append(
        assistant_card(
            "How to use this",
            "Where visitor growth is positive but sales or W2C decline, the likely issue is not just digital reach. Investigate traffic quality, offers, stock, retailer execution and proposition fit.",
            "diagnostic lens",
            "Toyota / Lexus",
            "neutral",
        )
    )

    return cards


def explain_metric_cards():
    return [
        assistant_card(
            "Website Conversion Rate",
            "This is the existing website effectiveness metric used by TME today. It measures visitors who complete a lead or online action divided by total website visitors.",
            "converted visitors / total visitors",
            "website effectiveness",
            "neutral",
        ),
        assistant_card(
            "Website-to-Contract Conversion Rate",
            "This dashboard uses a broader commercial efficiency metric: passenger new car customer contracts divided by unique website visitors.",
            "contracts / unique visitors",
            "macro conversion",
            "primary",
        ),
        assistant_card(
            "Why the distinction matters",
            "A low W2C rate does not automatically mean a weak website. It can reflect brand demand, pricing, product, stock, retailer follow-up, channel mix or market structure.",
            "not website-only",
            "interpretation",
            "positive",
        ),
        assistant_card(
            "Unique visitor source",
            "Unique visitors are sourced from Similarweb and should be used as a consistent market benchmark, not as a replacement for first-party web analytics.",
            "Similarweb",
            "data source",
            "neutral",
        ),
        assistant_card(
            "Sales source",
            "Passenger car sales are sourced from Marklines. The metric does not include every possible sales channel interpretation such as fleet or tactical registrations.",
            "Marklines",
            "sales source",
            "neutral",
        ),
        assistant_card(
            "Best use",
            "Use W2C to provoke bigger-picture questions about why demand does or does not translate into customer contracts.",
            "diagnostic",
            "use case",
            "primary",
        ),
    ]


def generate_assistant_cards(data, question, selected_oems):
    q = question.lower().strip()
    market = extract_market_from_question(q)
    brand = extract_brand_from_question(q, sorted(data["OEM"].unique()))

    if not q:
        return []

    if any(term in q for term in ["define", "definition", "methodology", "what is", "explain"]):
        if "conversion" in q or "w2c" in q or "website" in q:
            return explain_metric_cards()

    # Leaderboard-style questions must be answered before brand detection.
    # Example: "who leads W2C in Germany?" should return the #1 OEM, not detect "DS" inside "leads".
    leaderboard_terms = ["leader", "leads", "lead", "best", "top", "outperform", "highest", "worst", "weakest", "lowest", "underperform"]
    metric_terms = ["w2c", "conversion", "sales", "visitor", "traffic", "audience", "contract"]
    if any(term in q for term in leaderboard_terms) or any(term in q for term in metric_terms):
        # If the question specifically names Toyota and Lexus, handle that comparison first.
        if "toyota" in q and "lexus" in q:
            return toyota_lexus_cards(data, market)

        # If the question is about a named brand and asks "how is X performing", return brand detail.
        if brand and any(term in q for term in ["how is", "how are", "performance", "performing", "gap", "compare"]):
            return brand_cards(data, market, brand)

        # Otherwise answer as a ranking / top-bottom question.
        if any(term in q for term in leaderboard_terms) or any(term in q for term in ["highest", "lowest", "top", "best", "worst"]):
            return top_brands_cards(data, market, selected_oems, q)

    if "toyota" in q and "lexus" in q:
        return toyota_lexus_cards(data, market)

    if brand:
        return brand_cards(data, market, brand)

    if any(term in q for term in ["sales", "visitor", "traffic", "audience"]):
        return top_brands_cards(data, market, selected_oems, q)

    return [
        assistant_card(
            "Try asking a more specific data question",
            "I can answer questions about leaders, weakest performers, Toyota/Lexus gaps, brand performance, sales, visitors and W2C conversion by market.",
            "Example",
            "Who leads W2C in Germany?",
            "neutral",
        ),
        assistant_card(
            "Useful prompts",
            "Try: 'How is Toyota performing in Germany?', 'Who has the highest W2C rate in MM5?', or 'Explain the conversion metric'.",
            "3 prompts",
            "guided assistant",
            "primary",
        ),
    ]


def render_data_assistant(data, selected_oems):
    st.markdown(
        "<div class='da-simple'>"
        "<div class='da-simple-title'>Welcome. How can I help?</div>"
        "<div class='da-simple-copy'>Ask a question about markets, OEMs, Website-to-Contract Conversion Rate, sales, visitors or Toyota/Lexus gaps. Responses are generated from the dashboard dataset only.</div>"
        "<div class='da-simple-label'>Example questions</div>"
        "<div class='da-simple-questions'>"
        "<div class='da-simple-question'>Who leads W2C in Germany?</div>"
        "<div class='da-simple-question'>How is Toyota performing in France?</div>"
        "<div class='da-simple-question'>Compare Toyota and Lexus in MM5</div>"
                "</div>"
        "</div>",
        unsafe_allow_html=True,
    )

    st.markdown(
        f"<div class='da-context-bar'>"
        f"<div></div>"
        f"<div class='da-context-chip'>{CURRENT_LABEL} vs {PREVIOUS_LABEL}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )

    question = st.text_input(
        "Ask the data",
        placeholder="Search...",
        key="data_assistant_question",
    )

    if question:
        cards = generate_assistant_cards(data, question, selected_oems)
        st.markdown(assistant_result_html(cards), unsafe_allow_html=True)
    else:
        st.info("Start by typing a question above. Example: 'Which OEM has the highest W2C rate in the UK?'")


def render_data_assistant_page(data, selected_oems):
    render_data_assistant(data, selected_oems)
    render_footer()




def render_start_here_page(data):
    section("Start here — how to use this dashboard")

    intro_html = (
        "<div class='start-intro'>"
        "<b>This is not a dashboard that seeks to isolate which OEM has the best or worst website.</b><br><br>"
        "This dashboard looks to provoke bigger picture questions. It seeks to examine overall conversion efficiency from website visit to a customer buying the car. "
        "It is clear that the experience of the website including usability is a factor but this is only one part of a much broader set of factors.<br><br>"
        "Detailed below are all of the factors."
        "</div>"
    )
    st.markdown(intro_html, unsafe_allow_html=True)

    section("Factors that can influence Website-to-Contract Conversion Rate")
    factors = [
        "Brand Strategy",
        "Marketing Strategy",
        "Ecommerce Strategy",
        "Website Usability",
        "Website Experience",
        "Pricing & Promotions Strategy",
        "Ownership Options",
        "Product Strategy",
        "Stock Availability",
        "Network Strategy",
        "Powertrain Strategy",
        "Macro Economic factors",
    ]
    cards_html = "".join([f"<div class='start-factor-card'>{factor}</div>" for factor in factors])
    st.markdown(f"<div class='start-factors-grid'>{cards_html}</div>", unsafe_allow_html=True)

    section("Two conversion metrics for different purposes")
    methodology_html = (
        "<div class='methodology-grid'>"
        "<div class='methodology-card'>"
        "<div class='methodology-callout'>Focus: solely website effectiveness</div>"
        "<h4>Website Conversion Rate</h4>"
        "<div class='methodology-formula'><b>Overall Website Conversion Rate</b><br>Percentage of total website visitors who completed any kind of lead = converted visitors / total visitors.</div>"
        "<ul>"
        "<li><b>Used by TME today</b> to optimise website performance.</li>"
        "<li><b>Transactional leads:</b> brochure requests, keep-me-informed, newsletter, contest and similar forms.</li>"
        "<li><b>E-commerce leads:</b> TME OSB or local OSB forms for after-sales and online transactions for new / used / stock cars and pre-sales.</li>"
        "<li><b>Marketing leads:</b> forms classified as Marketing Lead.</li>"
        "<li><b>General Contact Leads:</b> general or non-transactional contact forms classified as Contact Lead.</li>"
        "</ul>"
        "</div>"
        "<div class='methodology-card'>"
        "<div class='methodology-callout'>Focus: bigger-picture conversion efficiency</div>"
        "<h4>Website-to-Contract Conversion Rate</h4>"
        "<div class='methodology-formula'><b>Website-to-Contract Conversion Rate</b><br>Passenger new car customer contracts / unique website visitors.</div>"
        "<ul>"
        "<li>Uses <b>Similarweb unique website visitors</b> as the demand base.</li>"
        "<li>Uses <b>passenger new car customer contracts / sales</b> as the outcome.</li>"
        "<li>Should be interpreted as a broader commercial efficiency metric, not purely a website UX metric.</li>"
        "<li>Captures the combined effect of digital demand, website experience, retailer execution, offer competitiveness and operating model.</li>"
        "</ul>"
        "</div>"
        "</div>"
    )
    st.markdown(methodology_html, unsafe_allow_html=True)

    section("How to interpret Similarweb unique visitors")
    c1, c2 = st.columns(2)

    with c1:
        method_html = (
            "<div class='method-card'>"
            "<div class='method-title'>Why Similarweb will not exactly match web analytics</div>"
            "<div class='method-copy'>"
            "Similarweb states that its traffic estimates are built from a blended methodology using <b>direct measurement</b>, "
            "a <b>contributory network</b>, <b>partnerships</b> and <b>public data extraction</b>, which are then processed and modelled through its "
            "machine-learning “Intelligence Engine”. Because this is an external estimation methodology, it should be treated as a "
            "<b>consistent market benchmark</b>, not as a direct replacement for Adobe or Google Analytics."
            "</div>"
            "</div>"
        )
        st.markdown(method_html, unsafe_allow_html=True)

    with c2:
        consent_html = (
            "<div class='method-card'>"
            "<div class='method-title'>Why consent management affects these sources differently</div>"
            "<div class='method-copy'>"
            "Similarweb also states that it <b>does not use cookies to collect behavioural data</b>. That means the impact of consent management platforms "
            "is <b>not the same</b> as in first-party analytics implementations, where tagging and consent choices can materially affect reported sessions or users. "
            "In practice, Similarweb and site analytics may move in similar directions over time, but the absolute numbers will not reconcile one-to-one."
            "</div>"
            "</div>"
        )
        st.markdown(consent_html, unsafe_allow_html=True)

    section("OEMs included in this report")
    oems = sorted(data["OEM"].dropna().astype(str).unique().tolist())
    pills = "".join([f"<span class='oem-pill'>{o}</span>" for o in oems])
    st.markdown(f"<div class='oem-pill-wrap'>{pills}</div>", unsafe_allow_html=True)

    render_footer()


def usecase_card(audience, title, copy, reports, audience_class="shared"):
    reports_html = "<br>".join([f"• {r}" for r in reports])
    return (
        f"<div class='usecase-card {audience_class}'>"
        f"<div class='usecase-audience'>{audience}</div>"
        f"<div class='usecase-title'>{title}</div>"
        f"<div class='usecase-copy'>{copy}</div>"
        f"<div class='usecase-report'><b>Recommended report view</b><br>{reports_html}</div>"
        f"</div>"
    )


def render_use_cases_page(data):
    section("Use cases — what to use this dashboard for")

    st.markdown(
        "<div class='usecase-note'>"
        "This dashboard is designed to help both <b>NMSCs</b> and <b>TME</b> ask sharper questions about Website-to-Contract Conversion Rate. "
        "The same metric can support different decisions: NMSCs will typically use it to improve marketing, content and local funnel execution, "
        "while TME will typically use it to identify product, powertrain, pricing and market-structure implications."
        "</div>",
        unsafe_allow_html=True,
    )

    cards = [
        usecase_card(
            "NMSC / TME",
            "Identify which OEMs are outperforming in your market",
            "Use this to understand which OEMs are converting website demand into customer contracts most efficiently in a specific market. This should trigger investigation into whether the driver is brand demand, offer strength, stock availability, retailer execution or website experience.",
            ["Market Summary", "Scorecard", "Bubble chart"],
            "shared",
        ),
        usecase_card(
            "NMSC / TME",
            "Assess different types of OEMs",
            "Compare volume OEMs, premium OEMs and Chinese disruptors to understand whether conversion efficiency is structurally different by business model, price point, product mix or market maturity.",
            ["OEM category filter", "Bubble chart", "Scorecard"],
            "shared",
        ),
        usecase_card(
            "NMSC",
            "Prioritise marketing and content optimisation",
            "Use market-level gaps to identify where Toyota or Lexus may need sharper content, clearer CTAs, better offer visibility, stronger model pages or more effective lower-funnel journeys.",
            ["Toyota & Lexus Gap Analysis", "Market Summary", "Start Here methodology"],
            "nmsc",
        ),
        usecase_card(
            "NMSC",
            "Diagnose whether traffic growth is quality traffic",
            "If unique visitors are growing but Website-to-Contract Conversion Rate is flat or declining, the issue may be traffic quality, audience targeting, message-market fit or weak lower-funnel content.",
            ["Market Summary", "YoY unique visitor growth chart", "Top 10 brands at a glance"],
            "nmsc",
        ),
        usecase_card(
            "NMSC",
            "Dig deeper into competitor marketing strategies",
            "Use this dashboard to identify which competitors need deeper investigation, then move into the market-specific marketing intelligence dashboards to review competitor activity, messaging and channel strategy.",
            ["Market Summary", "External market dashboards linked below"],
            "nmsc",
        ),
        usecase_card(
            "TME",
            "Identify product and powertrain pressure points",
            "Where Toyota or Lexus underperform consistently across markets, the cause may not be local digital execution. It may signal product-market fit, powertrain competitiveness, availability, pricing or customer proposition issues.",
            ["Toyota & Lexus Gap Analysis", "Market Gap Analysis", "Scorecard"],
            "tme",
        ),
        usecase_card(
            "TME",
            "Compare market readiness for future product strategy",
            "Use cross-market differences to understand where certain propositions appear to convert better. This can inform future model launches, powertrain emphasis, grade strategy and central product messaging.",
            ["Toyota & Lexus Gap Analysis", "Bubble chart", "Scorecard"],
            "tme",
        ),
        usecase_card(
            "TME",
            "Spot structural network or distribution constraints",
            "If a brand has strong visitor demand but weak contract conversion in multiple countries, investigate retailer footprint, sales model, availability, lead handling and distribution constraints before assuming a website problem.",
            ["Toyota & Lexus Gap Analysis", "Market Gap Analysis", "Start Here causal factors"],
            "tme",
        ),
        usecase_card(
            "NMSC / TME",
            "Separate website effectiveness from commercial conversion",
            "Use the Start Here page to align stakeholders that Website Conversion Rate and Website-to-Contract Conversion Rate are different. One is about lead capture; the other is about broader demand-to-contract efficiency.",
            ["Start Here", "Two conversion metrics for different purposes"],
            "shared",
        ),
        usecase_card(
            "NMSC",
            "Prepare local action plans",
            "For NMSCs, the dashboard should be used to identify where to run deeper diagnostics: landing page effectiveness, offer prominence, lead quality, retailer follow-up, campaign audience quality and local content gaps.",
            ["Toyota & Lexus Gap Analysis", "Market Summary", "Scorecard"],
            "nmsc",
        ),
    ]

    for i in range(0, len(cards), 2):
        cols = st.columns(2)
        for col, card in zip(cols, cards[i:i + 2]):
            with col:
                st.markdown(card, unsafe_allow_html=True)

    section("Market-specific competitor marketing dashboards")
    st.markdown(
        "<div class='usecase-note'>"
        "Separate market-level competitor marketing dashboards have been created to support deeper investigation into local marketing and content strategy."
        "</div>",
        unsafe_allow_html=True,
    )

    market_links = {
        "UK": "https://valtech-uk-auto.netlify.app/",
        "DE": "https://valtech-de-auto.netlify.app/",
        "FR": "https://valtech-fr-auto.netlify.app/",
        "IT": "https://valtech-it-auto.netlify.app/",
        "ES": "https://valtech-es-auto.netlify.app/",
    }
    links_html = "".join([f"<a class='market-link-card' href='{url}' target='_blank'>{market}</a>" for market, url in market_links.items()])
    st.markdown(f"<div class='market-link-grid'>{links_html}</div>", unsafe_allow_html=True)

    render_footer()

# =========================
# Pages
# =========================

def render_gap_analysis_page(data, market):
    section(f"Toyota & Lexus gap analysis — {market}")
    st.caption(f"Comparison selected: {CURRENT_LABEL} vs {PREVIOUS_LABEL}.")
    st.caption(f"All headline metrics reflect current period performance; variance badges show movement versus previous period.")
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
    conv_24 = sales_24 / visitors_24 * 100 if visitors_24 else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric(f"Toyota/Lexus passenger sales", fmt_int(sales_25), f"{fmt_pct((sales_25 / sales_24 - 1) * 100)} vs previous period" if sales_24 else "n/a")
    c2.metric(f"Toyota/Lexus Unique Visitors", fmt_int(visitors_25), f"{fmt_pct((visitors_25 / visitors_24 - 1) * 100)} vs previous period" if visitors_24 else "n/a")
    c3.metric(f"Website-to-Contract Conversion Rate", f"{conv_25:.2f}%", f"{fmt_pp(conv_25 - conv_24)} vs previous period")
    c4.metric(f"Visits per contract", fmt_int(visitors_25 / sales_25))

    render_brand_detail(data, market)
    render_benchmark_cards(data, market)
    render_market_weakness(data)
    render_toyota_lexus_recommendations(data)
    render_footer()



def exec_kpi_card(label, value, delta=None):
    delta_html = f"<div class='exec-kpi-delta'>{delta}</div>" if delta is not None else ""
    return (
        f"<div class='exec-kpi-card'>"
        f"<div class='exec-kpi-label'>{label}</div>"
        f"<div class='exec-kpi-value'>{value}</div>"
        f"{delta_html}"
        f"</div>"
    )


def render_exec_kpis(data, market, selected_oems):
    current = market_year(data, market, 2025, selected_oems)
    previous = market_year(data, market, 2024, selected_oems)

    if current.empty:
        st.warning("No data available for this selection.")
        return False

    sales_25 = current["Sales"].sum()
    visitors_25 = current["UniqueVisitors"].sum()
    conv_25 = sales_25 / visitors_25 * 100 if visitors_25 else 0

    sales_24 = previous["Sales"].sum() if not previous.empty else 0
    visitors_24 = previous["UniqueVisitors"].sum() if not previous.empty else 0
    conv_24 = sales_24 / visitors_24 * 100 if visitors_24 else 0

    cards = [
        exec_kpi_card(f"{CURRENT_LABEL} market contracts", fmt_metric_number(sales_25), f"{fmt_pct((sales_25 / sales_24 - 1) * 100)} vs {PREVIOUS_LABEL}" if sales_24 else "n/a"),
        exec_kpi_card(f"{CURRENT_LABEL} unique visitors", fmt_metric_number(visitors_25), f"{fmt_pct((visitors_25 / visitors_24 - 1) * 100)} vs {PREVIOUS_LABEL}" if visitors_24 else "n/a"),
        exec_kpi_card("Market W2C rate", f"{conv_25:.2f}%", f"{fmt_pp(conv_25 - conv_24)} vs {PREVIOUS_LABEL}"),
        exec_kpi_card("Visits per contract", fmt_int(visitors_25 / sales_25) if sales_25 else "n/a", "lower is better"),
    ]
    st.markdown("<div class='exec-kpi-grid'>" + "".join(cards) + "</div>", unsafe_allow_html=True)
    return True


def render_toyota_lexus_market_share(data, market):
    current_all = market_year(data, market, 2025, None)
    previous_all = market_year(data, market, 2024, None)
    current_tl = market_year(data, market, 2025, ["Toyota", "Lexus"])
    previous_tl = market_year(data, market, 2024, ["Toyota", "Lexus"])

    if current_all.empty or current_tl.empty:
        st.info("No Toyota/Lexus share data available for this market.")
        return

    all_visitors = current_all["UniqueVisitors"].sum()
    all_sales = current_all["Sales"].sum()
    tl_visitors = current_tl["UniqueVisitors"].sum()
    tl_sales = current_tl["Sales"].sum()

    prev_all_visitors = previous_all["UniqueVisitors"].sum() if not previous_all.empty else 0
    prev_all_sales = previous_all["Sales"].sum() if not previous_all.empty else 0
    prev_tl_visitors = previous_tl["UniqueVisitors"].sum() if not previous_tl.empty else 0
    prev_tl_sales = previous_tl["Sales"].sum() if not previous_tl.empty else 0

    visitor_share = tl_visitors / all_visitors * 100 if all_visitors else 0
    contract_share = tl_sales / all_sales * 100 if all_sales else 0
    prev_visitor_share = prev_tl_visitors / prev_all_visitors * 100 if prev_all_visitors else 0
    prev_contract_share = prev_tl_sales / prev_all_sales * 100 if prev_all_sales else 0

    html = (
        "<div class='share-grid'>"
        "<div class='share-card'>"
        "<div class='share-title'>Toyota & Lexus share of unique visitors</div>"
        f"<div class='share-metric-row'><span class='share-label'>Share of market visitors</span><span class='share-value'>{visitor_share:.1f}%</span></div>"
        f"<div class='share-metric-row'><span class='share-label'>Unique visitors</span><span class='share-value'>{fmt_metric_number(tl_visitors)}</span></div>"
        f"<div class='share-metric-row'><span class='share-label'>Movement vs {PREVIOUS_LABEL}</span><span class='share-value'>{fmt_pp(visitor_share - prev_visitor_share)}</span></div>"
        "</div>"
        "<div class='share-card'>"
        "<div class='share-title'>Toyota & Lexus share of customer contracts</div>"
        f"<div class='share-metric-row'><span class='share-label'>Share of market contracts</span><span class='share-value'>{contract_share:.1f}%</span></div>"
        f"<div class='share-metric-row'><span class='share-label'>Customer contracts</span><span class='share-value'>{fmt_metric_number(tl_sales)}</span></div>"
        f"<div class='share-metric-row'><span class='share-label'>Movement vs {PREVIOUS_LABEL}</span><span class='share-value'>{fmt_pp(contract_share - prev_contract_share)}</span></div>"
        "</div>"
        "</div>"
    )
    st.markdown(html, unsafe_allow_html=True)


def top10_leaders_table(data, market, selected_oems):
    yoy = yoy_table(data, market, selected_oems)
    if yoy.empty:
        st.info("No Top 10 data available for this selection.")
        return

    top = yoy.sort_values("ConversionPct_2025", ascending=False).head(10).copy()
    display = pd.DataFrame({
        "Rank": range(1, len(top) + 1),
        "Brand": top["OEM"],
        "Category": top["OEM"].map(cluster_for_oem),
        "W2C Conv Rate": top["ConversionPct_2025"].map(lambda x: f"{x:.2f}%"),
        "Unique Visitors": top["UniqueVisitors_2025"].map(fmt_metric_number),
        "Customer Contracts": top["Sales_2025"].map(fmt_metric_number),
        f"Visitor change vs {PREVIOUS_LABEL}": top["Visitors YoY %"].map(fmt_pct),
        f"Sales change vs {PREVIOUS_LABEL}": top["Sales YoY %"].map(fmt_pct),
    })
    st.dataframe(display, use_container_width=True, hide_index=True)


def render_market_takeaways(data, market, selected_oems):
    yoy = yoy_table(data, market, selected_oems)
    if yoy.empty:
        return

    top_w2c = yoy.sort_values("ConversionPct_2025", ascending=False).iloc[0]
    top_visitors = yoy.sort_values("UniqueVisitors_2025", ascending=False).iloc[0]
    fastest_visitors = yoy.sort_values("Visitors YoY %", ascending=False).iloc[0]
    weakest_w2c_move = yoy.sort_values("Conv Var pp", ascending=True).iloc[0]

    tl = yoy[yoy["OEM"].isin(["Toyota", "Lexus"])].copy()
    if not tl.empty:
        tl_avg = tl["ConversionPct_2025"].mean()
        market_avg = yoy["ConversionPct_2025"].mean()
        tl_gap_text = f"Toyota/Lexus average W2C is {tl_avg:.2f}% versus cohort average {market_avg:.2f}%."
        tl_metric = fmt_pp(tl_avg - market_avg)
        tl_style = "positive" if tl_avg >= market_avg else "risk"
    else:
        tl_gap_text = "Toyota/Lexus data is not available in this market selection."
        tl_metric = "n/a"
        tl_style = "neutral"

    cards = [
        ("Leader", f"{top_w2c['OEM']} leads W2C", f"{top_w2c['OEM']} records the highest Website-to-Contract Conversion Rate in {market}.", f"{top_w2c['ConversionPct_2025']:.2f}%", "positive"),
        ("Scale", f"{top_visitors['OEM']} leads visitor scale", f"{top_visitors['OEM']} has the largest unique visitor base in {market}.", fmt_metric_number(top_visitors["UniqueVisitors_2025"]), "primary"),
        ("Growth", f"{fastest_visitors['OEM']} has fastest visitor growth", f"Unique visitors grew {fmt_pct(fastest_visitors['Visitors YoY %'])} versus {PREVIOUS_LABEL}.", fmt_pct(fastest_visitors["Visitors YoY %"]), "positive"),
        ("Risk", f"{weakest_w2c_move['OEM']} has weakest W2C movement", f"W2C moved {fmt_pp(weakest_w2c_move['Conv Var pp'])} versus {PREVIOUS_LABEL}.", fmt_pp(weakest_w2c_move["Conv Var pp"]), "risk"),
        ("Toyota/Lexus", "Toyota & Lexus position", tl_gap_text, tl_metric, tl_style),
    ]

    html = "<div class='takeaway-grid'>"
    for label, title, copy, metric, style in cards:
        html += (
            f"<div class='takeaway-card {style}'>"
            f"<div class='takeaway-label'>{label}</div>"
            f"<div class='takeaway-title'>{title}</div>"
            f"<div class='takeaway-copy'>{copy}</div>"
            f"<div class='assistant-card-metric'>{metric}</div>"
            f"</div>"
        )
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)



def render_market_summary_page(data, market, selected_oems):
    section(f"Market Summary — {market}")
    st.caption(f"Executive view for {CURRENT_LABEL} vs {PREVIOUS_LABEL}: market growth, Toyota/Lexus share, leading competitors and key takeaways.")
    render_cluster_legend(selected_clusters)

    if not render_exec_kpis(data, market, selected_oems):
        return

    section("Toyota & Lexus share of market demand and contracts")
    render_toyota_lexus_market_share(data, market)

    section("Leading competitors — Top 10 by Website-to-Contract Conversion Rate")
    top10_leaders_table(data, market, selected_oems)

    section("Five executive takeaways")
    render_market_takeaways(data, market, selected_oems)

    render_footer()


def render_bubble_side_table(df, market):
    latest = df[df["Year"] == 2025].copy()
    if latest.empty:
        latest = df.copy()

    latest = latest.sort_values("UniqueVisitors", ascending=False).head(15)

    rows = ""
    for _, r in latest.iterrows():
        cluster = cluster_for_oem(r["OEM"])
        color = CLUSTER_COLORS.get(cluster, "#6F6F6F")
        rows += (
            "<div class='bubble-row'>"
            f"<span class='bubble-row-dot' style='background:{color};'></span>"
            "<div>"
            f"<div class='bubble-row-brand'>{r['OEM']}</div>"
            f"<div class='bubble-row-meta'>{r['ConversionPct']:.2f}% W2C · {fmt_int(r['Sales'])} sales</div>"
            "</div>"
            f"<div class='bubble-row-value'>{fmt_short(r['UniqueVisitors'])}</div>"
            "</div>"
        )

    st.markdown(
        f"<div class='bubble-side-card'>"
        f"<div class='bubble-side-title'>{market} selected OEMs</div>"
        f"{rows}"
        f"</div>",
        unsafe_allow_html=True,
    )


def render_bubble_page(data, selected_oems, year_view, show_logos):
    section("Bubble chart")
    st.caption(f"Comparison selected: {CURRENT_LABEL} vs {PREVIOUS_LABEL}.")
    render_cluster_legend(selected_clusters)
    st.markdown(
        "<div class='bubble-control-note'>"
        "<b>How to read this view:</b> bubble position shows audience scale versus Website-to-Contract Conversion Rate. "
        "Bubble size reflects passenger sales. In the 2024 and 2025 shift view, arrows show the movement from 2024 to 2025."
        "</div>",
        unsafe_allow_html=True,
    )

    tabs = st.tabs(MARKETS)
    for market, tab in zip(MARKETS, tabs):
        with tab:
            df = data[(data["Market"] == market) & (data["OEM"].isin(selected_oems))].copy()
            if year_view != "Previous and current + shift":
                selected_year = 2024 if year_view == "Previous period" else 2025
                df = df[df["Year"] == selected_year]
            if df.empty:
                st.info("No bubble chart data for this selection.")
                continue

            chart_col, side_col = st.columns([4.2, 1.15])
            with chart_col:
                fig = build_bubble_chart(df, selected_oems, market, year_view, show_logos)
                st.plotly_chart(fig, use_container_width=True)
            with side_col:
                render_bubble_side_table(df, market)



def scorecard_table(data, market, selected_oems):
    yoy = yoy_table(data, market, selected_oems)
    if yoy.empty:
        return pd.DataFrame()

    out = yoy.copy()
    out["Category"] = out["OEM"].map(cluster_for_oem)
    out["W2C Ranking"] = out["ConversionPct_2025"].rank(method="first", ascending=False).astype(int)
    out = out.sort_values("W2C Ranking")

    columns = [
        "OEM",
        "Category",
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
        "W2C Ranking",
    ]

    return out[[col for col in columns if col in out.columns]]


def render_scorecard_page(data, selected_oems):
    section("Leadership scorecard")
    st.caption(f"Comparison selected: {CURRENT_LABEL} vs {PREVIOUS_LABEL}.")
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
            "Visitors YoY %": "Visitor YoY vs previous period",
            "Sales_2024": "Passenger sales 2024",
            "Sales_2025": "Passenger sales 2025",
            "Sales YoY %": "Sales YoY vs previous period",
            "ConversionPct_2024": "W2C rate 2024",
            "ConversionPct_2025": "W2C rate 2025",
            "Conv Var pp": "W2C var vs previous period",
        }
    ).copy()

    for col in ["Website visitors 2024", "Website visitors 2025", "Passenger sales 2024", "Passenger sales 2025"]:
        if col in display.columns:
            display[col] = display[col].map(fmt_int)
    for col in ["Visitor YoY vs previous period", "Sales YoY vs previous period"]:
        if col in display.columns:
            display[col] = display[col].map(fmt_pct)
    for col in ["W2C rate 2024", "W2C rate 2025"]:
        if col in display.columns:
            display[col] = display[col].map(lambda x: f"{x:.2f}%")
    if "W2C var vs previous period" in display.columns:
        display["W2C var vs previous period"] = display["W2C var vs previous period"].map(fmt_pp)
    for col in ["Visits to Sale 2024", "Visits to Sale 2025", "Visits to Sale Var"]:
        if col in display.columns:
            display[col] = display[col].map(fmt_int)

    badge_cols = [c for c in ["Visitor YoY vs previous period", "Sales YoY vs previous period", "W2C var vs previous period"] if c in display.columns]
    styler = display.style.map(badge_style, subset=badge_cols) if badge_cols else display.style

    st.dataframe(
        styler,
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

raw_data = load_data()

st.sidebar.header("Comparison view")
comparison_view = st.sidebar.radio(
    "Choose the comparison period",
    list(COMPARISON_OPTIONS.keys()),
    index=0,
    help="Switch the full dashboard between annual 2024 vs 2025 and Jan-Apr 2025 vs Jan-Apr 2026.",
)

ACTIVE_COMPARISON = COMPARISON_OPTIONS[comparison_view]
PREVIOUS_LABEL = ACTIVE_COMPARISON["previous_label"]
CURRENT_LABEL = ACTIVE_COMPARISON["current_label"]

data = apply_comparison_view(raw_data, comparison_view)
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
    ["Start Here", "Use Cases", "Market Summary", "Toyota & Lexus Gap Analysis", "Bubble chart", "Scorecard", "Data Assistant"],
    index=0,
)

summary_market = st.sidebar.selectbox("Summary market", MARKETS, index=0)

year_view = st.sidebar.selectbox(
    "Bubble year view",
    ["Previous and current + shift", "Previous period", "Current period"],
    index=0,
)

preset = st.sidebar.selectbox(
    "Preset",
    ["Toyota volume competitors", "Lexus premium competitors", "Chinese disruptors", "All OEMs"],
    index=3,
)

cluster_options = ["All categories"] + list(OEM_CLUSTERS.keys())

selected_clusters = st.sidebar.multiselect(
    "OEM categories",
    cluster_options,
    default=["All categories"],
    help="Filter the dashboard by OEM grouping. These categories are based on the brand clusters used in the companion marketing dashboards.",
)

cluster_filtered_oems = oems_for_clusters(selected_clusters, all_oems)

default_oems = [o for o in preset_selection(preset, all_oems) if o in cluster_filtered_oems]
if not default_oems:
    default_oems = cluster_filtered_oems

selected_oems = st.sidebar.multiselect(
    "OEMs",
    cluster_filtered_oems,
    default=default_oems,
)

show_logos = st.sidebar.toggle(
    "Show OEM logos",
    value=True,
    help="Logos require internet access and a Brandfetch Client ID in Streamlit Secrets.",
)

if show_logos and not get_brandfetch_client_id():
    st.sidebar.warning("Brandfetch Client ID not found. Add BRANDFETCH_CLIENT_ID in Streamlit Secrets or turn logos off.")

selected_oems = selected_or_all(selected_oems, cluster_filtered_oems)

if page == "Start Here":
    render_start_here_page(data)
elif page == "Use Cases":
    render_use_cases_page(data)
elif page == "Market Summary":
    render_market_summary_page(data, summary_market, selected_oems)
elif page == "Toyota & Lexus Gap Analysis":
    render_gap_analysis_page(data, summary_market)
elif page == "Bubble chart":
    render_bubble_page(data, selected_oems, year_view, show_logos)
elif page == "Scorecard":
    render_scorecard_page(data, selected_oems)
else:
    render_data_assistant_page(data, selected_oems)
