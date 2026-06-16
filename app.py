import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import plotly.express as px
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import base64
import os
import requests as req
from openpyxl import load_workbook
import re

from pathlib import Path

def _load_img(path, mime):
    try:
        with open(path, "rb") as f:
            return f"data:{mime};base64," + base64.b64encode(f.read()).decode()
    except:
        return ""

BASE_DIR = Path(__file__).resolve().parent


def _load_img(path, mime):
    try:
        with open(path, "rb") as f:
            return f"data:{mime};base64," + base64.b64encode(f.read()).decode()
    except Exception:
        return ""


FIFA_LOGO    = _load_img(str(BASE_DIR / "FIFA_Logo_2026.webp"), "image/webp")
CONCEPT_LOGO = _load_img(str(BASE_DIR / "Concept Logo World Cup 2026.jpg"), "image/jpeg")

st.set_page_config(
    page_title="FIFA 2026 Dashboard",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================
# ✅ FIFA DARK THEME STYLE
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Roboto+Condensed:wght@300;400;700&display=swap');

/* ── Global background + host-nation ambient glow ── */
.stApp, html, body { background-color: #07091e; color: #e0e4f4; }
.stApp::before {
    content: '';
    position: fixed;
    inset: 0;
    background:
        radial-gradient(ellipse 55% 90% at 10% 50%, rgba(0,60,140,0.09) 0%, transparent 70%),
        radial-gradient(ellipse 55% 90% at 90% 50%, rgba(190,20,20,0.06) 0%, transparent 70%),
        radial-gradient(ellipse 80% 45% at 50% 105%, rgba(0,120,50,0.06) 0%, transparent 60%);
    pointer-events: none;
    z-index: 0;
}
/* ── Diagonal line tech texture ── */
.stApp::after {
    content: '';
    position: fixed;
    inset: 0;
    background-image: repeating-linear-gradient(
        -55deg,
        transparent 0px, transparent 28px,
        rgba(212,175,55,0.016) 28px, rgba(212,175,55,0.016) 29px
    );
    pointer-events: none;
    z-index: 0;
}
/* ── Full-width layout (no sidebar) ── */
.block-container {
    padding-top: 1.2rem !important;
    padding-left: 2.5rem !important;
    padding-right: 2.5rem !important;
    max-width: 100% !important;
}

/* ── FIFA header banner ── */
.fifa-banner {
    background:
        linear-gradient(160deg, rgba(0,23,74,0.94) 0%, rgba(0,13,46,0.91) 45%, rgba(0,13,46,0.91) 55%, rgba(0,26,82,0.94) 100%),
        url('https://images.unsplash.com/photo-1522778119026-d647f0596c20?w=1400&q=30') center/cover no-repeat;
    border: 1px solid #1a2f6b;
    border-bottom: 3px solid #d4af37;
    border-radius: 16px;
    padding: 20px 28px 18px;
    text-align: center;
    margin-bottom: 18px;
    position: relative;
    overflow: hidden;
}
.fifa-banner::before {
    content: '2026';
    font-family: 'Bebas Neue', Impact, sans-serif;
    position: absolute;
    font-size: 150px;
    letter-spacing: 24px;
    line-height: 1;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    color: transparent;
    -webkit-text-stroke: 1px rgba(212,175,55,0.11);
    pointer-events: none;
    z-index: 0;
    user-select: none;
    white-space: nowrap;
}
/* Sheen sweep across banner */
.fifa-banner::after {
    content: '';
    position: absolute;
    top: 0; left: -120%; width: 70%; height: 100%;
    background: linear-gradient(105deg, transparent 0%, rgba(255,255,255,0.025) 50%, transparent 100%);
    animation: bannerSheen 7s ease-in-out infinite;
    pointer-events: none;
    z-index: 0;
}
.fifa-banner > * { position: relative; z-index: 1; }
.fifa-banner-logo {
    font-size: 30px;
    font-weight: 900;
    letter-spacing: 6px;
    color: #4a90d9;
    font-family: Arial Black, sans-serif;
    text-transform: uppercase;
    margin-bottom: 2px;
}
.fifa-title {
    font-family: 'Bebas Neue', Impact, sans-serif;
    color: #ffffff;
    font-size: 50px;
    letter-spacing: 5px;
    margin: 0;
    line-height: 1;
    text-shadow: 0 0 50px rgba(255,255,255,0.25), 0 2px 14px rgba(0,0,0,0.95);
    animation: titleBreath 5s ease-in-out infinite;
}
.fifa-hosts {
    margin-top: 6px;
    font-size: 12px;
    letter-spacing: 2px;
    font-family: 'Roboto Condensed', Arial, sans-serif;
    font-weight: 700;
    text-transform: uppercase;
}
.fifa-hosts .canada { color: #ff6666; }
.fifa-hosts .mexico { color: #55cc77; }
.fifa-hosts .usa    { color: #66aaff; }
.fifa-hosts .sep    { color: #2a3a6a; margin: 0 10px; }
.fifa-dates {
    color: rgba(255,255,255,0.6);
    font-size: 11px;
    letter-spacing: 4px;
    text-transform: uppercase;
    margin-top: 6px;
}

/* ── Banner flex layout (logo | title | logo) ── */
.banner-flex { display:flex; align-items:center; justify-content:center; gap:12px; }
.banner-center { flex:0 1 auto; text-align:center; }
.banner-logo-left  { flex:0 0 auto; text-align:center; }
.banner-logo-right { flex:0 0 auto; text-align:center; }

/* ── Section titles ── */
.section-title {
    font-family: 'Bebas Neue', Impact, sans-serif;
    color: #ffffff;
    font-size: 20px;
    letter-spacing: 3px;
    text-transform: uppercase;
    border-left: 4px solid #d4af37;
    padding-left: 12px;
    margin: 30px 0 14px;
    animation: glowPulse 3s ease-in-out infinite, fadeInUp 0.5s ease both;
}

/* ── Match card ── */
.match-card {
    background: linear-gradient(135deg, #0c1132 0%, #111840 100%);
    border: 1px solid #1c2558;
    border-radius: 14px;
    margin-bottom: 9px;
    overflow: hidden;
    transition: border-color 0.25s, box-shadow 0.25s, transform 0.22s;
    animation: fadeInUp 0.4s ease both;
}
.match-card:hover { border-color: #d4af37; box-shadow: 0 4px 26px rgba(212,175,55,0.14); transform: translateY(-2px); }
.match-card.live {
    border: 2px solid #ff4444;
    background: linear-gradient(135deg, #1b0808 0%, #220e0e 100%);
    box-shadow: 0 0 24px rgba(255,60,60,0.3);
}
.match-card.finished { border-color: #162a16; background: linear-gradient(135deg, #0a130a 0%, #0e1a0e 100%); }

/* ── Match inner layout ── */
.match-inner {
    display: flex;
    align-items: center;
    padding: 13px 22px;
    gap: 10px;
}
.team-left  { display: flex; align-items: center; gap: 10px; flex: 2; }
.team-right { display: flex; align-items: center; gap: 10px; flex: 2; justify-content: flex-end; }
.team-name {
    font-family: 'Roboto Condensed', Arial, sans-serif;
    font-weight: 700;
    font-size: 16px;
    color: #c0c8e8;
    transition: color 0.2s;
}
.team-name.winner { color: #ffd700; text-shadow: 0 0 14px rgba(255,215,0,0.45); font-size: 17px; }
.team-name.loser  { color: #3a4560; }
.flag-img {
    width: 38px;
    height: 26px;
    object-fit: cover;
    border-radius: 3px;
    border: 1px solid rgba(255,255,255,0.18);
    flex-shrink: 0;
    display: inline-block;
}
.flag-blank { width: 38px; height: 26px; display: inline-block; }

/* ── VS / Result divider ── */
.vs-center { min-width: 88px; text-align: center; }
.vs-text {
    font-family: 'Bebas Neue', Impact, sans-serif;
    font-size: 20px;
    color: #d4af37;
    line-height: 1;
}
.result-win {
    font-family: 'Bebas Neue', Impact, sans-serif;
    font-size: 12px;
    color: #80e8a8;
    line-height: 1.35;
    letter-spacing: 1px;
}
.result-draw {
    font-family: 'Bebas Neue', Impact, sans-serif;
    font-size: 16px;
    color: #90c8ff;
    letter-spacing: 2px;
}
.live-pulse {
    font-family: 'Bebas Neue', Impact, sans-serif;
    font-size: 18px;
    color: #ff4444;
    animation: blinker 1.2s linear infinite;
    letter-spacing: 2px;
}

/* ── Match meta (right side) ── */
.match-meta { min-width: 175px; text-align: right; }
.match-date { font-size: 13px; font-weight: 600; color: #b8c4e8; }

/* ── Status badges ── */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1px;
    margin-top: 5px;
}
.badge-live     { background:#ff4444; color:#fff; animation: blinker 1.2s linear infinite, pulseRing 2s ease-out infinite; }
.badge-upcoming { background:#1a4f82; color:#90c8ff; }
.badge-finished { background:#1a5e30; color:#80e8a8; }
.badge-tbc      { background:#333; color:#aaa; }
@keyframes blinker { 50% { opacity: 0.25; } }

/* ── All animation keyframes ── */
@keyframes titleBreath {
    0%, 100% { text-shadow: 0 0 30px rgba(255,255,255,0.18), 0 2px 14px rgba(0,0,0,0.95); }
    50%       { text-shadow: 0 0 80px rgba(255,255,255,0.5), 0 0 30px rgba(180,210,255,0.3), 0 2px 14px rgba(0,0,0,0.95); }
}
@keyframes bannerSheen {
    0%    { left: -120%; }
    40%   { left: 130%;  }
    100%  { left: 130%;  }
}
@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(18px); }
    to   { opacity: 1; transform: translateY(0);    }
}
@keyframes glowPulse {
    0%, 100% { border-left-color: #d4af37; }
    50%       { border-left-color: #ffe97a; filter: drop-shadow(-2px 0 10px rgba(212,175,55,0.55)); }
}
@keyframes floatUp {
    0%, 100% { transform: translateY(0px);  }
    50%       { transform: translateY(-4px); }
}
@keyframes pulseRing {
    0%   { box-shadow: 0 0 0 0   rgba(255,68,68,0.8); }
    70%  { box-shadow: 0 0 0 10px rgba(255,68,68,0);   }
    100% { box-shadow: 0 0 0 0   rgba(255,68,68,0);   }
}
@keyframes barFill {
    from { width: 0; }
}
@keyframes rotateBall {
    from { transform: rotate(0deg); }
    to   { transform: rotate(360deg); }
}

/* ── HTML Leaderboard rows ── */
.lb-container { width: 100%; padding-bottom: 4px; }
.lb-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.lb-col  { display: flex; flex-direction: column; }
.lb-row {
    display: flex; align-items: center; gap: 8px;
    padding: 7px 10px;
    background: linear-gradient(135deg, #0c1132, #111840);
    border: 1px solid #1c2558;
    border-left: 3px solid var(--pc, #1c2558);
    border-radius: 8px; margin-bottom: 4px;
    transition: transform 0.18s, box-shadow 0.18s;
    animation: fadeInUp 0.35s ease both;
}
.lb-row:hover { transform: translateX(4px); box-shadow: -4px 0 18px rgba(0,0,0,0.5); }
.lb-row.rank-1 { border-left-color: #d4af37; background: linear-gradient(135deg, #1a1408 0%, #241c08 100%); }
.lb-row.rank-2 { border-left-color: #909090; background: linear-gradient(135deg, #131313 0%, #1c1c1e 100%); }
.lb-row.rank-3 { border-left-color: #a0522d; background: linear-gradient(135deg, #18100a 0%, #20130b 100%); }
.lb-num { font-family: 'Bebas Neue', Impact, sans-serif; font-size: 15px; min-width: 24px; text-align: center; flex-shrink: 0; }
.lb-num.p1 { color: #d4af37; } .lb-num.p2 { color: #a8a8a8; } .lb-num.p3 { color: #cd7f32; } .lb-num.pn { color: #2a3555; font-size: 12px; }
.lb-avatar { width: 26px; height: 26px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-family: 'Bebas Neue', Impact, sans-serif; font-size: 10px; color: #07091e; flex-shrink: 0; font-weight: 900; }
.lb-name { font-family: 'Roboto Condensed', Arial, sans-serif; font-weight: 700; font-size: 12px; color: #d8ddf0; flex: 1; min-width: 50px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.lb-bar-wrap { flex: 2; height: 3px; background: #0a0e20; border-radius: 3px; overflow: hidden; }
.lb-bar { height: 100%; border-radius: 3px; animation: barFill 1.2s ease-out both; }
.lb-score { font-family: 'Bebas Neue', Impact, sans-serif; font-size: 16px; min-width: 44px; text-align: right; color: #c8d0f0; flex-shrink: 0; }
.lb-score sup { font-size: 9px; color: #334560; font-family: 'Roboto Condensed'; vertical-align: super; }
.lb-acc { font-size: 9px; min-width: 34px; text-align: right; color: #445070; font-family: 'Roboto Condensed'; flex-shrink: 0; }

/* ── Football pitch section divider ── */
.pitch-divider {
    width: 100%; height: 54px;
    margin: 26px 0 6px;
    position: relative; overflow: hidden;
    border-radius: 10px;
    border: 1px solid rgba(30,80,30,0.6);
    background:
        linear-gradient(rgba(0,0,0,0.42), rgba(0,0,0,0.42)),
        url('https://images.unsplash.com/photo-1431324155629-1a6deb1dec8d?w=1000&q=30') center/cover no-repeat;
    display: flex; align-items: center; justify-content: center; gap: 14px;
}
.pitch-divider::before {
    content: '';
    position: absolute; inset: 0;
    background-image: linear-gradient(rgba(255,255,255,0.04) 1px, transparent 1px),
                      linear-gradient(90deg, rgba(255,255,255,0.04) 1px, transparent 1px);
    background-size: 44px 44px;
}
.pitch-ball { position: relative; z-index: 1; font-size: 24px; animation: rotateBall 4s linear infinite; }
.pitch-label {
    position: relative; z-index: 1;
    font-family: 'Bebas Neue', Impact, sans-serif;
    font-size: 22px; letter-spacing: 7px;
    color: rgba(255,255,255,0.92);
    text-shadow: 0 1px 10px rgba(0,0,0,0.95);
}

.stat-row { display: flex; gap: 14px; margin: 16px 0 22px; flex-wrap: wrap; }
.stat-card {
    flex: 1;
    min-width: 130px;
    background: linear-gradient(135deg, #0c1132, #111840);
    border: 1px solid #1c2558;
    border-radius: 14px;
    padding: 20px 14px 16px;
    text-align: center;
    position: relative;
    overflow: hidden;
    animation: fadeInUp 0.5s ease both, floatUp 4s ease-in-out infinite 0.6s;
    transition: transform 0.25s, box-shadow 0.25s, border-color 0.25s;
}
.stat-card:hover { border-color: #d4af37; transform: translateY(-5px); box-shadow: 0 10px 32px rgba(212,175,55,0.18); }
.stat-card::before {
    content: attr(data-icon);
    position: absolute;
    font-size: 80px;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    opacity: 0.055;
    pointer-events: none;
    line-height: 1;
}
.stat-label { font-size: 10px; font-weight: 700; letter-spacing: 2px; text-transform: uppercase; color: #7888b8; margin-bottom: 4px; }
.stat-value { font-family: 'Bebas Neue', Impact, sans-serif; font-size: 50px; line-height: 1.05; color: #e0e4f4; }
.stat-value.gold   { color: #ffd700; text-shadow: 0 0 22px rgba(255,215,0,0.4); }
.stat-value.blue   { color: #70b8ff; }
.stat-value.green  { color: #80e8a8; }
.stat-value.orange { color: #ffaa44; }
.stat-value.red    { color: #ff7070; }
.stat-sublabel { font-size: 10px; color: #445070; margin-top: 3px; }
.accuracy-track { height: 4px; background: #1c2558; border-radius: 4px; margin-top: 10px; overflow: hidden; }
.accuracy-fill  { height: 100%; border-radius: 4px; background: linear-gradient(90deg, #1a5e30, #40c060); animation: barFill 1.5s ease-out both; }

/* ── Input labels ── */
.stSelectbox label, .stDateInput label { color: #7888b8 !important; font-size: 13px !important; }

/* ── TM superscript ── */
.tm-sup { font-size: 28%; vertical-align: super; letter-spacing: 2px; opacity: 0.75; font-family: 'Roboto Condensed', Arial, sans-serif; font-weight: 400; }

/* ── Golden football cursor ── */
html, body, * { cursor: url("data:image/svg+xml;charset=utf-8,%3Csvg xmlns='http://www.w3.org/2000/svg' width='22' height='22'%3E%3Cdefs%3E%3CradialGradient id='rg' cx='38%25' cy='32%25' r='60%25'%3E%3Cstop offset='0%25' stop-color='%23ffe566'/%3E%3Cstop offset='55%25' stop-color='%23d4af37'/%3E%3Cstop offset='100%25' stop-color='%23a07800'/%3E%3C/radialGradient%3E%3C/defs%3E%3Ccircle cx='11' cy='11' r='10' fill='url(%23rg)' stroke='%23886000' stroke-width='0.8'/%3E%3Cpath d='M11 4.5l1.8 3.5 3.5.5-2.6 2.3 1 3.7-3.7-2-3.7 2 1-3.7-2.6-2.3 3.5-.5z' fill='%23704000' opacity='0.5'/%3E%3C/svg%3E") 11 11, auto !important; }

/* ── Floating transparent footballs ── */
.fb {
    position: fixed; font-size: 56px; opacity: 0; pointer-events: none;
    z-index: 1; user-select: none; line-height: 1;
    animation: floatBall var(--fdur,22s) ease-in-out infinite var(--fdel,0s);
}
@keyframes floatBall {
    0%   { transform: translateY(102vh) rotate(0deg);   opacity: 0;     }
    6%   { opacity: 0.04; }
    50%  { transform: translateY(45vh)  rotate(270deg); opacity: 0.028; }
    94%  { opacity: 0.04; }
    100% { transform: translateY(-8vh)  rotate(540deg); opacity: 0;     }
}

/* ── Hide: top header bar, deploy button, sidebar, toolbar ── */
header[data-testid="stHeader"]         { display: none !important; }
[data-testid="stToolbar"]               { display: none !important; }
[data-testid="stDecoration"]            { display: none !important; }
.stDeployButton                         { display: none !important; }
section[data-testid="stSidebar"]        { display: none !important; }
button[data-testid="collapsedControl"]  { display: none !important; }
#MainMenu { display: none !important; }
footer    { display: none !important; }
</style>
""", unsafe_allow_html=True)

# ── FIFA Banner ─────────────────────────────────────────────────────────────
_logo_l = f'<img src="{FIFA_LOGO}" style="height:114px;filter:drop-shadow(0 0 24px rgba(212,175,55,0.8));">' if FIFA_LOGO else '<span style="font-size:70px;">⚽</span>'
_logo_r = f'<img src="{CONCEPT_LOGO}" style="height:132px;opacity:0.94;filter:drop-shadow(0 0 24px rgba(40,120,255,0.65));">' if CONCEPT_LOGO else ''
st.markdown(f"""
<div class="fifa-banner">
    <div class="banner-flex">
        <div class="banner-logo-left">{_logo_l}</div>
        <div class="banner-center">
            <div class="fifa-banner-logo">F I F A</div>
            <div class="fifa-title">WORLD CUP 2026<sup class="tm-sup">™</sup></div>
            <div class="fifa-hosts">
                <span class="canada">🍁 CANADA</span>
                <span class="sep">·</span>
                <span class="mexico">🌵 MEXICO</span>
                <span class="sep">·</span>
                <span class="usa">🦅 USA</span>
            </div>
            <div class="fifa-dates">June 11 &nbsp;—&nbsp; July 19, 2026 &nbsp;·&nbsp; Predictions Dashboard</div>
        </div>
        <div class="banner-logo-right">{_logo_r}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Floating footballs (transparent, decorative) ────────────────────────
st.markdown("""
<div class="fb" style="left:4%;  --fdur:22s; --fdel:0s">⚽</div>
<div class="fb" style="left:18%; --fdur:31s; --fdel:-12s">⚽</div>
<div class="fb" style="left:37%; --fdur:19s; --fdel:-5s">⚽</div>
<div class="fb" style="left:56%; --fdur:26s; --fdel:-18s">⚽</div>
<div class="fb" style="left:74%; --fdur:23s; --fdel:-8s">⚽</div>
<div class="fb" style="left:89%; --fdur:28s; --fdel:-15s">⚽</div>
""", unsafe_allow_html=True)

FILE_PATH = str(BASE_DIR / "World Cup 2026 Comp.xlsx")


NZ_TZ = ZoneInfo("Pacific/Auckland")

def parse_time(val):
    try:
        if pd.isna(val): return None
        val = str(val).strip()
        if val in ["", "nan", "0"]: return None
        return pd.to_datetime(val).time()
    except:
        return None

def build_dt(row):
    if pd.isna(row["Date (NZDT)"]): return None
    base = row["Date (NZDT)"].date()
    t = row["ParsedTime"] or datetime.min.time()
    return datetime.combine(base, t).replace(tzinfo=NZ_TZ)

@st.cache_data
def load_data(_mtime=0):
    _df = pd.read_excel(FILE_PATH, sheet_name="Sheet1")
    _df["Date (NZDT)"] = pd.to_datetime(_df["Date (NZDT)"], errors="coerce")
    _df["ParsedTime"]  = _df["Time (NZDT)"].apply(parse_time)
    _df["DateTime"]    = _df.apply(build_dt, axis=1)
    return _df

# ── ESPN team-name normalisation map ─────────────────────────────────────────
_ESPN_MAP = {
    "united states":"United States","usa":"United States",
    "u.s.":"United States","u.s.a.":"United States",
    "türkiye":"Türkiye","turkey":"Türkiye",
    "korea republic":"South Korea","south korea":"South Korea",
    "republic of korea":"South Korea",
    "ir iran":"Iran","iran":"Iran",
    "côte d'ivoire":"Ivory Coast","ivory coast":"Ivory Coast",
    "cabo verde":"Cape Verde","cape verde":"Cape Verde",
    "bosnia & herzegovina":"Bosnia and Herzegovina",
    "bosnia and herzegovina":"Bosnia and Herzegovina",
    "north macedonia":"North Macedonia",
    "czechia":"Czech Republic","czech republic":"Czech Republic",
    "dr congo":"DR Congo","congo dr":"DR Congo",
    "saint kitts and nevis":"St Kitts and Nevis",
}

def _norm_team(name):
    raw = " ".join(str(name).split()).strip()
    return _ESPN_MAP.get(raw.casefold(), raw)

def _team_key(name):
    return _norm_team(name).casefold()


def _is_draw_result(value):
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass

    s = " ".join(str(value).replace("—", "-").replace("–", "-").replace(":", "-").split()).casefold()
    if s in {"draw", "tie", "0-0", "0:0", "nil nil", "nil-nil"}:
        return True

    m = re.search(r'(?<!\d)(\d+)\s*[-]\s*(\d+)(?!\d)', s)
    return bool(m and m.group(1) == m.group(2))



@st.cache_data(ttl=45)
def fetch_api_results():
    """Fetch only finished results from official / fallback providers."""
    _out = {}
    _
now = datetime.now(NZ_TZ)
today = now.date()
selected_date = st.date_input("Select Date", value=today)

matches = df[df["Date (NZDT)"].dt.date == selected_date]

if matches.empty:
    st.info("⚽ No matches scheduled for this date.")
else:
    for _, row in matches.iterrows():
        team1  = str(row["Team 1"]) if pd.notna(row["Team 1"]) else "TBD"
        team2  = str(row["Team 2"]) if pd.notna(row["Team 2"]) else "TBD"
        dt     = row["DateTime"]
        result = str(row["Result"]).strip() if pd.notna(row["Result"]) else None
        k1 = _team_key(team1)
        k2 = _team_key(team2)

        live = (_live_scores.get((k1, k2)) or _live_scores.get((k2, k1))) if '_live_scores' in globals() else None
        live_state = str((live or {}).get("status", "")).casefold()
        is_live_feed = bool(live and any(k in live_state for k in (
            "live", "in progress", "inprogress", "1st half", "2nd half",
            "half time", "halftime", "ht", "extra time", "penalty"
        )))
        is_finished_feed = bool(live and any(k in live_state for k in (
            "finished", "completed", "final", "post", "postperiod"
        )))
        api_result = _api_results.get((k1, k2)) or _api_results.get((k2, k1))
        has_result = result is not None and str(result).strip() != ""

        # If a live feed says the match is live, never let stale Result values override it.
        if is_live_feed:
            has_result = False
        elif not has_result and api_result:
            result = str(api_result).strip()
            has_result = True

        # If a live feed has clearly finished, map its score to the final result when needed.
        if is_finished_feed and not has_result:
            hs = live.get("home", None)
            as_ = live.get("away", None)
            if hs is not None and as_ is not None:
                if hs == as_:
                    result = "Draw"
                elif hs > as_:
                    result = team1
                else:
                    result = team2
                has_result = True

        f1 = flag_html(get_flag_url(team1))
        f2 = flag_html(get_flag_url(team2))

        # ── Status: live feed wins first; finished second; then scheduled/pending ──
        if is_live_feed:
            status = "LIVE"
        elif has_result or is_finished_feed:
            status = "Finished"
        elif dt:
            diff = (now - dt).total_seconds()
            if diff < 0:
                status = "Upcoming"
            else:
                status = "Result Pending"
        else:
            status = "TBC"

        # ── Centre display ───────────────────────────────────────────────────
        t1_cls = t2_cls = ""
        if status == "Finished":
            if result and _is_draw_result(result):
                score_html = '<span class="result-draw">⚖ DRAW</span>'
            elif result and _team_key(result) == _team_key(team1):
                t1_cls, t2_cls = "winner", "loser"
                score_html = f'<span class="result-win">✓ {team1}<br>WINS</span>'
            elif result and _team_key(result) == _team_key(team2):
                t1_cls, t2_cls = "loser", "winner"
                score_html = f'<span class="result-win">✓ {team2}<br>WINS</span>'
            else:
                score_html = f'<span class="result-win">{result}</span>'
        elif status == "LIVE":
            if live and live.get("home") is not None and live.get("away") is not None:
                hs = live.get("home", 0)
                as_ = live.get("away", 0)
                score_html = f'<span class="live-pulse">🔴 {hs} - {as_}</span>'
            else:
                score_html = '<span class="live-pulse">● LIVE</span>'
        else:
            score_html = '<span class="vs-text">VS</span>' if status == "Upcoming" else '<span class="result-win">WAITING<br>FOR RESULT</span>'

        time_str = dt.strftime("%d %b  ·  %I:%M %p NZT") if dt else "Time TBC"

        badge_map = {
            "LIVE":           '<span class="badge badge-live">🔴 LIVE</span>',
            "Upcoming":       '<span class="badge badge-upcoming">⏳ UPCOMING</span>',
            "Result Pending": '<span class="badge badge-finished">⏳ RESULT PENDING</span>',
            "Finished":       '<span class="badge badge-finished">✅ FINISHED</span>',
            "TBC":            '<span class="badge badge-tbc">🕐 TBC</span>',
        }
        if   status == "LIVE": card_cls = "match-card live"
        elif status in ("Finished", "Result Pending"): card_cls = "match-card finished"
        else: card_cls = "match-card"

        st.markdown(f"""
        <div class="{card_cls}">
            <div class="match-inner">
                <div class="team-left">
                    {f1}
                    <span class="team-name {t1_cls}">{team1}</span>
                </div>
                <div class="vs-center">
                    {score_html}
                </div>
                <div class="team-right">
                    <span class="team-name {t2_cls}">{team2}</span>
                    {f2}
                </div>
                <div class="match-meta">
                    <div class="match-date">{time_str}</div>
                    {badge_map[status]}
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
# =========================
# 👥 PLAYER SUMMARY
# =========================
st.markdown('<div class="section-title">👥 Player Summary</div>', unsafe_allow_html=True)
st.dataframe(leaderboard, use_container_width=True, hide_index=True)

# =========================
# 🔍 PLAYER DETAIL
# =========================
st.markdown('<div class="section-title">🔍 Player Detail</div>', unsafe_allow_html=True)

player_sel = st.selectbox("Select Player", leaderboard["Player"].tolist())

p_df     = data[data["Player"] == player_sel].copy()
p_played = p_df[p_df["Result"].notna()]

correct  = int(p_played["Correct"].sum())
total    = p_played.shape[0]
accuracy = round((correct / total) * 100, 1) if total else 0

acc_cls      = "green" if accuracy >= 60 else "orange" if accuracy >= 35 else "red"
p_rank_s     = leaderboard.loc[leaderboard["Player"] == player_sel, "Rank"]
p_rank       = p_rank_s.iloc[0] if not p_rank_s.empty else "—"
acc_bar_w    = min(int(accuracy), 100)

st.markdown(f"""
<div class="stat-row">
    <div class="stat-card" data-icon="⚽">
        <div class="stat-label">Points</div>
        <div class="stat-value gold">{correct}</div>
        <div class="stat-sublabel">Correct Predictions</div>
    </div>
    <div class="stat-card" data-icon="📋">
        <div class="stat-label">Matches Played</div>
        <div class="stat-value blue">{total}</div>
        <div class="stat-sublabel">Results Available</div>
    </div>
    <div class="stat-card" data-icon="🎯">
        <div class="stat-label">Accuracy</div>
        <div class="stat-value {acc_cls}">{accuracy}%</div>
        <div class="accuracy-track"><div class="accuracy-fill" style="width:{acc_bar_w}%"></div></div>
    </div>
    <div class="stat-card" data-icon="🏆">
        <div class="stat-label">Leaderboard Rank</div>
        <div class="stat-value" style="font-size:42px;line-height:1.2;">{p_rank}</div>
        <div class="stat-sublabel">Current Standing</div>
    </div>
</div>
""", unsafe_allow_html=True)

st.dataframe(
    p_df[["Match", "Pick", "Result", "Correct"]],
    use_container_width=True,
    hide_index=True,
)

# =========================
# ⚠️ MISSING PICKS
# =========================
st.markdown('<div class="section-title">⚠️ Missing Picks</div>', unsafe_allow_html=True)

stage1  = df.head(72)
missing = []

for col in pick_cols:
    p_name = col.replace(" Pick", "")
    filled = stage1[col].notna().sum()
    miss   = 72 - filled
    if miss > 0:
        missing.append({"Player": p_name, "Filled": filled, "Missing": miss})

if missing:
    st.dataframe(pd.DataFrame(missing), use_container_width=True, hide_index=True)
else:
    st.success("✅ All players have completed their picks!")
