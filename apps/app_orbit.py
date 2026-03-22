# ================================================================
# O.R.B.I.T — Optimised Research, Budget, Intelligence & Trajectory
# Tanaka Alex Mbendana  |  Beihang University 2025–2026
# Run: streamlit run app_orbit.py
# ================================================================
from __future__ import annotations
from pathlib import Path
from datetime import date, datetime, timedelta
import base64
import json

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from PIL import Image as _PILImage

# ── CONFIG ──────────────────────────────────────────────────────
_APP_DIR  = Path(__file__).parent          # apps/
_ROOT_DIR = _APP_DIR.parent                # Dashboard/
XLSX_PATH = _ROOT_DIR / "academic_command_centre_v3.xlsx"
if not XLSX_PATH.exists():                 # fallback: same dir as app
    XLSX_PATH = _APP_DIR / "academic_command_centre_v3.xlsx"

LOGS_DIR = Path("orbit_logs")
LOGS_DIR.mkdir(exist_ok=True)
LAST_VISIT_FILE  = LOGS_DIR / "last_visit.json"
SAVINGS_FILE     = LOGS_DIR / "savings.json"
EXPENSES_FILE    = LOGS_DIR / "expenses_log.csv"

ATTEND_THRESHOLD = 0.80
SEM_START  = date(2026, 3, 3)
SEM_END    = date(2026, 6, 26)
TODAY      = date.today()
NOW        = datetime.now()
TOTAL_DAYS = (SEM_END - SEM_START).days

# ── CLASS SCHEDULE (day=0Mon…6Sun, slots=50-min blocks) ──────────
# Slot start times (h, m)
_SLOT_START = {
    1:(8,0), 2:(8,50), 3:(10,0), 4:(10,50),
    5:(13,0),6:(13,50),7:(15,0), 8:(15,50),
    9:(17,0),10:(17,50),11:(19,0),12:(19,50),
}
CLASS_SCHEDULE = [
    {"name":"Chinese Language 2",  "code":"D253026002","color":"#2563eb",
     "room":"Teaching Bldg 2, Rm 3003",
     "days":[0,2], "slots":[1,2], "weeks":(1,16)},
    {"name":"Sci Paper Writing",   "code":"D253026011","color":"#7c3aed",
     "room":"Research Bldg 1, Rm 4108",
     "days":[0],   "slots":[3,4], "weeks":(1,8)},
    {"name":"RS Image Processing", "code":"D253051004","color":"#10b981",
     "room":"Research Bldg 1, Rm 1043",
     "days":[2,4], "slots":[3,4], "weeks":(1,8)},
    {"name":"UAV Remote Sensing",  "code":"D253052002","color":"#ea580c",
     "room":"Research Bldg 1, Rm 1043",
     "days":[0,3], "slots":[3,4], "weeks":(6,9)},
    {"name":"UAV Remote Sensing",  "code":"D253052002","color":"#ea580c",
     "room":"Research Bldg 1, Rm 1043",
     "days":[5,6], "slots":[1,2,3,4], "weeks":(10,11)},
    {"name":"RS Natural Disasters","code":"D253051005","color":"#f59e0b",
     "room":"Research Bldg 1, Rm 1045",
     "days":[3],   "slots":[8,9,10], "weeks":(9,16)},
    {"name":"RS Natural Disasters","code":"D253051005","color":"#f59e0b",
     "room":"Research Bldg 1, Rm 1045",
     "days":[4],   "slots":[8,9,10], "weeks":(15,16)},
    {"name":"AI & Large Models",   "code":"D253041002","color":"#6366f1",
     "room":"Research Bldg 1, Rm 4108",
     "days":[0,1], "slots":[11,12], "weeks":(1,8)},
]

def _current_class() -> dict | None:
    """Return the class happening right now, or None."""
    wk = max(1, (TODAY - SEM_START).days // 7 + 1)
    wd = NOW.weekday()
    now_t = NOW.time()
    for cls in CLASS_SCHEDULE:
        if wd not in cls["days"]:
            continue
        if not (cls["weeks"][0] <= wk <= cls["weeks"][1]):
            continue
        s_start = cls["slots"][0]
        s_end   = cls["slots"][-1]
        sh, sm = _SLOT_START[s_start]
        eh, em = _SLOT_START[s_end]
        eh_end_m = em + 50
        eh_real  = eh + eh_end_m // 60
        em_real  = eh_end_m % 60
        from datetime import time as _t
        if _t(sh, sm) <= now_t <= _t(eh_real, em_real):
            return cls
    return None

def _next_class() -> tuple[dict, str] | tuple[None, None]:
    """Return (class_dict, 'today HH:MM' or 'DayName HH:MM') for next class."""
    from datetime import time as _t
    wk = max(1, (TODAY - SEM_START).days // 7 + 1)
    # Check next 7 days
    DAY_NAMES = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
    for offset in range(7):
        check_date = TODAY + timedelta(days=offset)
        wd = check_date.weekday()
        for cls in CLASS_SCHEDULE:
            if wd not in cls["days"]:
                continue
            if not (cls["weeks"][0] <= wk <= cls["weeks"][1]):
                continue
            s_start = cls["slots"][0]
            sh, sm = _SLOT_START[s_start]
            label = ("Today" if offset == 0 else DAY_NAMES[wd]) + f" {sh:02d}:{sm:02d}"
            # skip if already passed today
            if offset == 0 and NOW.time() > _t(sh, sm):
                continue
            return cls, label
    return None, None

# ── PAGE ────────────────────────────────────────────────────────
_favicon_path = _ROOT_DIR / "Tanaka.jpg"
_favicon = _PILImage.open(str(_favicon_path)) if _favicon_path.exists() else "🛰️"

st.set_page_config(
    page_title="O.R.B.I.T — Tanaka Alex Mbendana",
    page_icon=_favicon,
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── PWA MANIFEST + INSTALL META TAGS ─────────────────────────────
import json as _json, base64 as _b64
_manifest = _json.dumps({
    "name": "O.R.B.I.T — Tanaka Alex Mbendana",
    "short_name": "O.R.B.I.T",
    "description": "Academic command centre — research, budget, trajectory",
    "start_url": "/",
    "display": "standalone",
    "background_color": "#030408",
    "theme_color": "#2563eb",
    "orientation": "any",
    "icons": [{"src": "https://raw.githubusercontent.com/lex565/orbit-dashboard/master/Tanaka.jpg",
               "sizes": "512x512", "type": "image/jpeg", "purpose": "any maskable"}]
})
_manifest_b64 = _b64.b64encode(_manifest.encode()).decode()
st.html(f"""
<meta name="theme-color" content="#2563eb">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="O.R.B.I.T">
<link rel="apple-touch-icon" href="https://raw.githubusercontent.com/lex565/orbit-dashboard/master/Tanaka.jpg">
<link rel="manifest" href="data:application/manifest+json;base64,{_manifest_b64}">
""")

# Auto-refresh every 5 minutes (300,000 ms) — picks up Excel changes & date rollovers
_refresh_count = st_autorefresh(interval=300_000, key="orbit_autorefresh")

# ── DARK / LIGHT MODE ────────────────────────────────────────────
if "light_mode" not in st.session_state:
    st.session_state.light_mode = False
_lm  = st.session_state.light_mode
_BG  = "#f0f2f6" if _lm else "#07080f"
_BG2 = "#ffffff" if _lm else "#0c1020"
_BG3 = "#e8eaf0" if _lm else "#07080f"
_BRD = "#d1d5db" if _lm else "#12192b"
_TXT = "#111827" if _lm else "#b8c4d0"
_TXT2= "#374151" if _lm else "#e2e8f0"
_TXT3= "#6b7280" if _lm else "#4a5568"
_TXT4= "#1f2937" if _lm else "#94b4d4"

# ── ALL CSS — use st.html() to prevent CSS leaking as visible text ─
st.html("""
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Space+Mono:wght@400;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
/* ── BASE ── */
*,*::before,*::after{box-sizing:border-box}
body{font-family:'Inter',sans-serif}
#MainMenu,footer,header{visibility:hidden}
section[data-testid="stSidebar"]{background:#030408!important;border-right:1px solid #0e1628}

/* ── ROOM BACKGROUND — applied directly to body/stApp, no pseudo-elements (avoids z-index clash with Streamlit loader) ── */
html,body,.stApp{
  background-color:#030408!important;
  background-image:
    repeating-linear-gradient(transparent,transparent 39px,rgba(37,99,235,.022) 39px,rgba(37,99,235,.022) 40px),
    repeating-linear-gradient(90deg,transparent,transparent 39px,rgba(37,99,235,.022) 39px,rgba(37,99,235,.022) 40px),
    linear-gradient(to right,rgba(0,0,0,.38) 0%,transparent 16%),
    linear-gradient(to left,rgba(0,0,0,.38) 0%,transparent 16%),
    linear-gradient(to bottom,rgba(0,0,0,.48) 0%,transparent 20%),
    linear-gradient(to top,rgba(0,0,0,.52) 0%,transparent 24%),
    radial-gradient(ellipse 65% 30% at 50% 108%,rgba(37,99,235,.14),transparent),
    radial-gradient(ellipse 36% 26% at 90% 3%,rgba(124,58,237,.08),transparent),
    linear-gradient(180deg,#020306 0%,#040a16 30%,#050c1a 50%,#040a16 70%,#020306 100%)!important;
  background-size:40px 40px,40px 40px,100% 100%,100% 100%,100% 100%,100% 100%,100% 100%,100% 100%,100% 100%!important;
  background-attachment:fixed!important;
  color:#b8c4d0}
.block-container{padding-top:1.2rem!important;padding-bottom:2rem!important;max-width:100%!important}

/* ── GLASSMORPHISM METRIC CARDS ── */
[data-testid="stMetric"]{
  background:rgba(8,12,28,0.75)!important;
  backdrop-filter:blur(14px)!important;-webkit-backdrop-filter:blur(14px)!important;
  border:1px solid rgba(37,99,235,.22)!important;
  border-radius:14px!important;padding:16px 14px 12px!important;
  position:relative;overflow:hidden;
  box-shadow:0 4px 28px rgba(0,0,0,.5),inset 0 1px 0 rgba(255,255,255,.04)!important;
  transition:transform .25s,border-color .25s,box-shadow .25s!important}
[data-testid="stMetric"]::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,#1d4ed8,#7c3aed,#3b82f6)}
[data-testid="stMetric"]:hover{transform:translateY(-3px)!important;
  border-color:rgba(37,99,235,.55)!important;
  box-shadow:0 10px 36px rgba(37,99,235,.25),inset 0 1px 0 rgba(255,255,255,.07)!important}
[data-testid="stMetricLabel"] p{color:#4a5568!important;font-size:.65rem!important;font-weight:700!important;letter-spacing:.1em!important;text-transform:uppercase;font-family:'Space Mono',monospace!important}
[data-testid="stMetricValue"] div{color:#e2e8f0!important;font-family:'Space Mono',monospace!important;font-size:1.4rem!important;font-weight:700!important}
[data-testid="stMetricDelta"]{display:none}

/* ── TABS ── */
[data-testid="stTabs"] [role="tablist"]{border-bottom:1px solid #0e1628;gap:2px;background:transparent}
button[role="tab"]{color:#4a5568!important;font-size:.68rem!important;font-weight:700!important;letter-spacing:.1em!important;text-transform:uppercase;padding:9px 18px!important;border-radius:6px 6px 0 0!important;border:1px solid transparent!important;border-bottom:none!important;font-family:'Space Mono',monospace!important;background:transparent!important;transition:all .25s}
button[role="tab"][aria-selected="true"]{color:#60a5fa!important;background:rgba(8,12,28,.85)!important;border-color:#0e1628!important}
button[role="tab"]:hover:not([aria-selected="true"]){color:#94b4d4!important;background:rgba(29,78,216,.08)!important}
[data-testid="stTabContent"]{animation:tab-fade-in .35s ease forwards}

/* ── INPUTS & BUTTONS ── */
[data-testid="stDataFrame"]{display:none!important}
.stAlert{border-radius:10px!important}
.stSelectbox [data-baseweb="select"]>div{background:#0c1020!important;border:1px solid #12192b!important;color:#e2e8f0!important}
.stNumberInput input,.stTextArea textarea,.stDateInput input,.stTextInput input{background:#08101e!important;border:1px solid #1a2540!important;color:#e2e8f0!important;border-radius:8px!important}
.stFormSubmitButton button,.stButton button{
  background:linear-gradient(135deg,#1d4ed8,#2563eb)!important;
  color:#fff!important;border:none!important;border-radius:8px!important;
  font-weight:700!important;font-family:'Space Mono',monospace!important;
  font-size:.75rem!important;letter-spacing:.06em!important;
  box-shadow:0 4px 18px rgba(37,99,235,.4)!important;
  transition:transform .2s,box-shadow .2s!important}
.stFormSubmitButton button:hover,.stButton button:hover{transform:translateY(-2px)!important;box-shadow:0 8px 24px rgba(37,99,235,.55)!important}

/* ── SCROLLBAR ── */
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:#030408}
::-webkit-scrollbar-thumb{background:rgba(37,99,235,.4);border-radius:2px}
::-webkit-scrollbar-thumb:hover{background:rgba(37,99,235,.7)}
.stCaption{color:#4a5568!important;font-size:.7rem!important}
div[data-testid="stVerticalBlock"] div[style*="overflow: hidden"]{overflow:visible!important}

/* ── GLASSMORPHISM HELPER ── */
.glass-card{background:rgba(8,12,28,0.7)!important;backdrop-filter:blur(16px)!important;-webkit-backdrop-filter:blur(16px)!important;border:1px solid rgba(37,99,235,.2)!important;box-shadow:0 8px 32px rgba(0,0,0,.45),inset 0 1px 0 rgba(255,255,255,.05)!important}

/* ── ANIMATION CLASSES ── */
.orbit-ring-1{animation:orbit-spin 8s linear infinite}
.orbit-ring-2{animation:orbit-spin-rev 14s linear infinite}
.orbit-ring-3{animation:orbit-spin 22s linear infinite}
.orbit-ring-4{animation:orbit-spin-rev 30s linear infinite}
.profile-avatar{animation:avatar-pulse 3s ease-in-out infinite}
.status-dot{animation:status-blink 2s ease-in-out infinite;border-radius:50%}
.orbit-title{animation:title-glow 4s ease-in-out infinite}
.float-satellite{animation:float-icon 3s ease-in-out infinite}
.fade-card{animation:card-in .5s ease forwards}
.shimmer-line{background:linear-gradient(90deg,#1d4ed8 0%,#7c3aed 25%,#d97706 50%,#2563eb 75%,#1d4ed8 100%);background-size:400px 100%;animation:shimmer-bar 3s linear infinite}

/* ── KEYFRAMES ── */
@keyframes orbit-spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
@keyframes orbit-spin-rev{from{transform:rotate(0deg)}to{transform:rotate(-360deg)}}
@keyframes pulse-glow{0%,100%{opacity:1;box-shadow:0 0 0 0 rgba(16,185,129,0)}50%{opacity:.7;box-shadow:0 0 12px 4px rgba(16,185,129,.18)}}
@keyframes avatar-pulse{0%,100%{box-shadow:0 0 0 0 rgba(37,99,235,.45),0 0 30px rgba(37,99,235,.2)}50%{box-shadow:0 0 0 14px rgba(37,99,235,0),0 0 40px rgba(37,99,235,.3)}}
@keyframes shimmer-bar{0%{background-position:-400px 0}100%{background-position:400px 0}}
@keyframes fade-in-up{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}
@keyframes float-icon{0%,100%{transform:translateY(0) rotate(-5deg)}50%{transform:translateY(-7px) rotate(-5deg)}}
@keyframes scan-line{0%{top:-2px;opacity:.6}100%{top:100%;opacity:0}}
@keyframes status-blink{0%,100%{background:#10b981;box-shadow:0 0 0 0 rgba(16,185,129,.7)}50%{background:#10b981;box-shadow:0 0 0 6px rgba(16,185,129,0)}}
@keyframes title-glow{0%,100%{text-shadow:0 0 20px rgba(37,99,235,.0)}50%{text-shadow:0 0 40px rgba(37,99,235,.3),0 0 80px rgba(124,58,237,.15)}}
@keyframes progress-anim{from{width:0%}to{width:inherit}}
@keyframes card-in{from{opacity:0;transform:translateY(12px) scale(.97)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes orbit-dot{0%{transform:rotate(0deg) translateX(70px) rotate(0deg)}100%{transform:rotate(360deg) translateX(70px) rotate(-360deg)}}
@keyframes star-twinkle{0%,100%{opacity:.15}50%{opacity:.7}}
@keyframes tab-fade-in{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes pulse-border{0%,100%{box-shadow:0 0 0 0 rgba(37,99,235,0)}50%{box-shadow:0 0 0 5px rgba(37,99,235,.2)}}
@keyframes count-up{from{opacity:0;transform:scale(.85)}to{opacity:1;transform:scale(1)}}
@keyframes ambient-glow{0%,100%{opacity:.4}50%{opacity:.7}}
</style>
""")

# ── DYNAMIC THEME + MOBILE CSS (st.html prevents CSS leak) ───────
st.html(f"""<style>
html,body,.stApp{{background-color:{_BG}!important;color:{_TXT}!important}}
section[data-testid="stSidebar"]{{background:{_BG}!important}}
[data-testid="stTabs"] [role="tablist"]{{border-color:{_BRD}}}
button[role="tab"]{{color:{_TXT3}!important}}
button[role="tab"][aria-selected="true"]{{color:#60a5fa!important;background:rgba(8,12,28,.88)!important;border-color:{_BRD}!important}}
.stNumberInput input,.stTextArea textarea,.stDateInput input,.stTextInput input{{background:{_BG2}!important;border-color:{_BRD}!important;color:{_TXT2}!important}}
.block-container{{background:transparent!important;position:relative;z-index:1}}

/* ── PORTRAIT & LANDSCAPE MOBILE ── */
@media screen and (max-width:900px){{
  .block-container{{padding:0.4rem 0.4rem 2rem!important}}
  [data-testid="stTabs"] [role="tablist"]{{gap:1px;overflow-x:auto;flex-wrap:nowrap;padding-bottom:4px;-webkit-overflow-scrolling:touch}}
  button[role="tab"]{{padding:5px 8px!important;font-size:.52rem!important;white-space:nowrap}}
  .orbit-title{{font-size:1.6rem!important}}
}}
@media screen and (max-width:900px) and (orientation:portrait){{
  div[data-testid="stColumns"]>div{{min-width:100%!important;flex:1 1 100%!important}}
  .stButton button{{font-size:.65rem!important;padding:10px!important}}
}}
@media screen and (max-width:900px) and (orientation:landscape){{
  div[data-testid="stColumns"]>div{{min-width:44%!important;flex:1 1 44%!important}}
  .block-container{{padding:0.3rem 0.5rem 1rem!important}}
  [data-testid="stMetricValue"] div{{font-size:1rem!important}}
}}
@media screen and (max-width:480px){{
  div[data-testid="stColumns"]>div{{min-width:100%!important;flex:1 1 100%!important}}
}}
</style>""")

# ── PLOTLY BASE THEME ────────────────────────────────────────────
PD = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(12,16,32,0.6)",
    font_color="#94b4d4",
    font_family="Inter",
    margin=dict(l=10, r=10, t=30, b=10),
    xaxis=dict(gridcolor="#12192b", linecolor="#12192b", tickfont_color="#4a5568", zeroline=False),
    yaxis=dict(gridcolor="#12192b", linecolor="#12192b", tickfont_color="#4a5568", zeroline=False),
)

# ── HTML HELPERS ─────────────────────────────────────────────────
def _bar(val: float, color: str = "#2563eb", height: int = 5) -> str:
    w = min(max(val * 100, 0), 100)
    return (f'<div style="background:#12192b;border-radius:3px;height:{height}px;overflow:hidden;margin-top:4px">'
            f'<div style="width:{w:.1f}%;height:100%;background:{color};border-radius:3px"></div></div>')

def _kpi_card(label: str, value: str, sub: str = "", accent: str = "#2563eb",
              icon: str = "", bg: str = "#0c1020") -> str:
    return (f'<div style="background:{bg};border:1px solid #12192b;border-radius:12px;'
            f'padding:16px 14px;position:relative;overflow:hidden">'
            f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{accent}"></div>'
            f'<div style="font-size:.62rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.1em">{icon} {label}</div>'
            f'<div style="font-size:1.5rem;font-family:\'Space Mono\',monospace;font-weight:700;'
            f'color:#e2e8f0;margin:6px 0 2px">{value}</div>'
            f'{"<div style=font-size:.68rem;color:#4a5568>" + sub + "</div>" if sub else ""}'
            f'</div>')

def _section_badge(label: str, value: str, color: str = "#2563eb") -> str:
    return (f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:8px;'
            f'padding:10px 12px;text-align:center">'
            f'<div style="font-size:.58rem;color:{color};font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.1em">{label}</div>'
            f'<div style="font-size:1.1rem;font-family:\'Space Mono\',monospace;font-weight:700;'
            f'color:#e2e8f0;margin-top:3px">{value}</div></div>')

def _pct(v: float) -> str:
    return f"{v*100:.1f}%"

def _pill(txt: str, color: str, bg: str) -> str:
    return (f'<span style="display:inline-block;padding:3px 10px;border-radius:20px;'
            f'background:{bg};color:{color};font-size:.6rem;font-weight:700;'
            f'font-family:\'Space Mono\',monospace;letter-spacing:.08em;text-transform:uppercase">{txt}</span>')

# ── DATA LOADING ─────────────────────────────────────────────────
@st.cache_data(ttl=30)
def _sheets(path: str) -> dict:
    xls = pd.ExcelFile(path, engine="openpyxl")
    return {sh: pd.read_excel(path, sheet_name=sh, engine="openpyxl", header=None)
            for sh in xls.sheet_names}

if not XLSX_PATH.exists():
    st.markdown("### 📂 Upload Data File")
    st.caption("The Excel data file was not found. Upload it to continue — it is only used for this session.")
    _uploaded = st.file_uploader(
        "Upload `academic_command_centre_v3.xlsx`",
        type=["xlsx"],
        key="xlsx_upload",
    )
    if _uploaded is not None:
        _tmp_path = Path(__file__).parent / "academic_command_centre_v3.xlsx"
        _tmp_path.write_bytes(_uploaded.read())
        XLSX_PATH = _tmp_path
        st.success("File loaded — refreshing…")
        st.rerun()
    else:
        st.info("Tabs that depend on your Excel data (Obligations, Budget, Research) will be empty until the file is uploaded.")
        # Provide empty fallback so the rest of the app (Expenses, Profile) still renders
        RAW = {}
else:
    RAW = _sheets(str(XLSX_PATH.resolve()))


def _analytics() -> pd.DataFrame:
    try:
        df = RAW.get("📊 Analytics", pd.DataFrame())
        sub = df.iloc[2:9, :5].copy()
        sub.columns = ["Course", "Attended", "Total", "Pct", "Remaining"]
        sub = sub[sub["Course"].notna() & (sub["Course"] != "COURSE")].reset_index(drop=True)
        for c in ["Attended", "Total", "Remaining"]:
            sub[c] = pd.to_numeric(sub[c], errors="coerce").fillna(0).astype(int)
        sub["Pct"] = pd.to_numeric(sub["Pct"], errors="coerce").fillna(0)
        return sub
    except Exception:
        return pd.DataFrame(columns=["Course","Attended","Total","Pct","Remaining"])


def _research():
    try:
        df = RAW.get("📄 Research Hub", pd.DataFrame())
        p1 = float(df.iloc[3, 3]) if pd.notna(df.iloc[3, 3]) else 0.6
        p2 = float(df.iloc[3, 9]) if pd.notna(df.iloc[3, 9]) else 0.0
        t1 = str(df.iloc[5, 0]).split("\n")[0].replace("📜  ", "").strip() if pd.notna(df.iloc[5, 0]) else "Paper 1"
        t2 = str(df.iloc[5, 6]).split("\n")[0].replace("📜  ", "").strip() if pd.notna(df.iloc[5, 6]) else "Paper 2"
        s1, s2 = [], []
        for i in range(7, 25):
            try:
                row = df.iloc[i]
                a = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
                if a and a not in ("nan","SECTION"):
                    s1.append({"Section": a,
                                "Status": str(row.iloc[2]) if pd.notna(row.iloc[2]) else "🔒",
                                "Target": str(row.iloc[3]) if pd.notna(row.iloc[3]) else "",
                                "Notes":  str(row.iloc[5]) if pd.notna(row.iloc[5]) else ""})
                b = str(row.iloc[7]) if pd.notna(row.iloc[7]) else ""
                if b and b not in ("nan","SECTION"):
                    s2.append({"Section": b,
                                "Status": str(row.iloc[8]) if pd.notna(row.iloc[8]) else "🔒",
                                "Target": str(row.iloc[9]) if pd.notna(row.iloc[9]) else "",
                                "Notes":  str(row.iloc[11]) if pd.notna(row.iloc[11]) else ""})
            except Exception:
                continue
        return p1, p2, pd.DataFrame(s1), pd.DataFrame(s2), t1, t2
    except Exception:
        return 0.6, 0.0, pd.DataFrame(), pd.DataFrame(), "Paper 1", "Paper 2"


# ── FINANCIAL MODEL ──────────────────────────────────────────────
YEAR1_MONTHLY   = 4500.0        # RMB/month stipend — Year 1
YEAR2_MONTHLY   = 3000.0        # RMB/month stipend — Year 2 (from Sep 2026)
YEAR1_START     = date(2025, 9, 1)
YEAR1_END       = date(2026, 8, 31)
YEAR2_START     = date(2026, 9, 1)
CURRENT_SAVINGS_DEFAULT = 5000.0  # RMB — starting savings this month (Mar 2026)
PENDING_BONUS   = 1600.0        # face-scan bonus — confirmed, not yet disbursed
FINANCE_BUDGET  = YEAR1_MONTHLY # current reference budget


def _auto_income() -> float:
    """Auto-calculate cumulative stipend received based on months elapsed."""
    today = date.today()
    if today <= YEAR1_END:
        months_y1 = (today.year - YEAR1_START.year) * 12 + (today.month - YEAR1_START.month)
        return max(0, months_y1) * YEAR1_MONTHLY
    else:
        months_y1 = (YEAR1_END.year - YEAR1_START.year) * 12 + (YEAR1_END.month - YEAR1_START.month) + 1
        months_y2 = (today.year - YEAR2_START.year) * 12 + (today.month - YEAR2_START.month)
        return months_y1 * YEAR1_MONTHLY + max(0, months_y2) * YEAR2_MONTHLY

def _months_to_year2() -> int:
    today = date.today()
    if today >= YEAR2_START:
        return 0
    return (YEAR2_START.year - today.year) * 12 + (YEAR2_START.month - today.month)

def _finance() -> dict:
    try:
        df = RAW.get("💰 Finance", pd.DataFrame())
        ts  = float(df.iloc[4, 0])  if pd.notna(df.iloc[4, 0])  else 0.0
        wa  = float(df.iloc[4, 6])  if pd.notna(df.iloc[4, 6])  else 0.0
        ssf = float(df.iloc[4, 12]) if pd.notna(df.iloc[4, 12]) else 0.0
        current_monthly = YEAR2_MONTHLY if date.today() >= YEAR2_START else YEAR1_MONTHLY
        tx_rows, cat_rows = [], []
        for i in range(6, 30):
            try:
                row = df.iloc[i]
                wk  = row.iloc[0]
                if pd.notna(wk) and str(wk) not in ("nan", "WK", "TOTALS"):
                    amt = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0.0
                    if amt > 0:
                        tx_rows.append({
                            "Wk":          int(float(wk)) if str(wk).replace(".","").isdigit() else "–",
                            "Description": str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
                            "Category":    str(row.iloc[2]) if pd.notna(row.iloc[2]) else "",
                            "Amount ¥":    amt,
                            "Balance":     float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0.0,
                        })
            except Exception:
                continue
        for i in range(7, 18):
            try:
                row = df.iloc[i]
                cat = row.iloc[8]
                if pd.notna(cat) and str(cat) not in ("nan", "CATEGORY", "TOTALS"):
                    cat_rows.append({
                        "Category": str(cat),
                        "Budget ¥": float(row.iloc[10]) if pd.notna(row.iloc[10]) else 0.0,
                        "Spent ¥":  float(row.iloc[12]) if pd.notna(row.iloc[12]) else 0.0,
                        "Left ¥":   float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0.0,
                    })
            except Exception:
                continue
        return dict(ts=ts, wa=wa, ssf=ssf,
                    current_monthly=current_monthly,
                    bonus_pending=PENDING_BONUS,
                    tx=pd.DataFrame(tx_rows), cats=pd.DataFrame(cat_rows))
    except Exception:
        return dict(ts=0, wa=0, ssf=0,
                    current_monthly=YEAR1_MONTHLY,
                    bonus_pending=PENDING_BONUS,
                    tx=pd.DataFrame(), cats=pd.DataFrame())


COURSE_SHORT = ["Chinese", "Sci Writing", "RS Image", "UAV", "RS Disasters", "AI Models"]
COURSE_FULL  = ["Chinese Language 2", "Sci Paper Writing", "RS Image Processing",
                "UAV Remote Sensing", "RS Natural Disasters", "AI & Large Models"]


def _weekly_att() -> pd.DataFrame:
    try:
        df = RAW.get("📅 Timetable", pd.DataFrame())
        rows = []
        for i in range(14, 32):
            row = df.iloc[i]
            wk = str(row.iloc[10]) if pd.notna(row.iloc[10]) else ""
            if "Wk" not in wk:
                continue
            entry = {"Week": wk}
            for ci, col in enumerate([11,12,13,14,15,16]):
                v = row.iloc[col] if pd.notna(row.iloc[col]) else 0
                try:
                    entry[COURSE_SHORT[ci]] = int(float(v))
                except Exception:
                    entry[COURSE_SHORT[ci]] = 0
            entry["Σ"] = sum(entry[c] for c in COURSE_SHORT)
            rows.append(entry)
        return pd.DataFrame(rows)
    except Exception:
        return pd.DataFrame()


def _schedule() -> list:
    try:
        df = RAW.get("📅 Timetable", pd.DataFrame())
        out = []
        for i in range(7, 20):
            row = df.iloc[i]
            slot = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
            time_ = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
            if not slot.startswith("S"):
                continue
            cells = []
            for col in [2, 3, 4, 5, 6, 7, 8]:
                raw = str(row.iloc[col]) if pd.notna(row.iloc[col]) else ""
                raw = raw.replace("\\n", "\n").strip()
                cells.append(raw if raw and raw != "nan" else "")
            out.append({"slot": slot, "time": time_, "cells": cells})
        return out
    except Exception:
        return []


# ── PARSE ────────────────────────────────────────────────────────
att_df   = _analytics()
p1p, p2p, p1s, p2s, p1t, p2t = _research()
fin      = _finance()
watt_df  = _weekly_att()
sched    = _schedule()

# ── LAST VISIT TRACKING ──────────────────────────────────────────
def _update_last_visit():
    LAST_VISIT_FILE.write_text(json.dumps({"last_visit": TODAY.isoformat()}))

def _days_since_visit() -> int:
    try:
        data = json.loads(LAST_VISIT_FILE.read_text())
        last = date.fromisoformat(data["last_visit"])
        return (TODAY - last).days
    except Exception:
        return 0

_days_absent = _days_since_visit()
_update_last_visit()

# ── SAVINGS (editable from app) ──────────────────────────────────
def _load_savings() -> float:
    try:
        return float(json.loads(SAVINGS_FILE.read_text())["savings"])
    except Exception:
        return CURRENT_SAVINGS_DEFAULT

def _save_savings(val: float):
    SAVINGS_FILE.write_text(json.dumps({"savings": val, "updated": TODAY.isoformat()}))

# ── MANUAL EXPENSE LOG (CSV) ─────────────────────────────────────
_EXP_COLS = ["date", "category", "description", "amount"]
_EXP_CATS = ["Food & Drink", "Transport", "Stationery", "Electronics", "Health",
             "Entertainment", "Clothing", "Utilities", "Other"]

def _load_expenses() -> "pd.DataFrame":
    try:
        if EXPENSES_FILE.exists():
            df = pd.read_csv(EXPENSES_FILE, encoding="utf-8")
            for c in _EXP_COLS:
                if c not in df.columns:
                    df[c] = "" if c != "amount" else 0.0
            return df[_EXP_COLS]
    except Exception:
        pass
    return pd.DataFrame(columns=_EXP_COLS)

def _append_expense(dt: str, cat: str, desc: str, amt: float):
    df = _load_expenses()
    new_row = pd.DataFrame([{"date": dt, "category": cat, "description": desc, "amount": amt}])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_csv(EXPENSES_FILE, index=False, encoding="utf-8")

def _delete_expense(idx: int):
    df = _load_expenses()
    df = df.drop(index=idx).reset_index(drop=True)
    df.to_csv(EXPENSES_FILE, index=False, encoding="utf-8")

current_savings = _load_savings()

# ── AUTO-CALCULATED INCOME & KPIs ────────────────────────────────
auto_income   = _auto_income()
months_to_y2  = _months_to_year2()

attended     = int(att_df["Attended"].sum()) if not att_df.empty else 0
total_sess   = int(att_df["Total"].sum())    if not att_df.empty else 126
sem_started  = attended > 0

att_frac     = attended / total_sess if (sem_started and total_sess > 0) else 0.0
res_avg      = (p1p + p2p) / 2

# current_savings IS the ground truth — what's actually in the bank right now.
# Do NOT reconstruct from income - expenses (expenses in Excel are partial).
net_balance  = current_savings
months_elapsed_total = max(1, (TODAY - YEAR1_START).days / 30.44)
monthly_burn = fin["ts"] / months_elapsed_total if fin["ts"] > 0 else 0.0
# Savings rate = savings as % of total income received to date
savings_rate = (net_balance / auto_income * 100) if auto_income > 0 else 0.0
spend_ratio  = 1.0 - (net_balance / auto_income) if auto_income > 0 else 0.0

elapsed      = max(0, (TODAY - SEM_START).days)
days_left    = max(0, (SEM_END - TODAY).days)
sem_frac     = min(max(elapsed / TOTAL_DAYS, 0.0), 1.0)

# Drift — ONLY red when not opened for 4+ consecutive days
drift = _days_absent >= 4

if not sem_started:
    sys_label = "PENDING START"
    sys_color = "#4a5568"
elif drift:
    sys_label = "⚠  THREATENED"
    sys_color = "#ef4444"
else:
    sys_label = "●  STABLE"
    sys_color = "#10b981"

# ── MODE TOGGLE ──────────────────────────────────────────────────
_tc1, _tc2, _tc3 = st.columns([7, 1, 1])
with _tc2:
    _btn_lbl = "☀️ Light" if not st.session_state.light_mode else "🌑 Dark"
    if st.button(_btn_lbl, use_container_width=True, key="mode_toggle"):
        st.session_state.light_mode = not st.session_state.light_mode
        st.rerun()
with _tc3:
    if st.button("📲 Install", use_container_width=True, key="install_btn"):
        st.session_state["show_install"] = not st.session_state.get("show_install", False)

if st.session_state.get("show_install", False):
    st.html("""
    <div style="background:linear-gradient(135deg,#07091a,#0a0f24);border:1px solid #1d4ed8;
      border-radius:14px;padding:18px 22px;margin-bottom:14px">
      <div style="font-size:.6rem;color:#2563eb;font-weight:700;font-family:'Space Mono',monospace;
        text-transform:uppercase;letter-spacing:.12em;margin-bottom:12px">📲 Install O.R.B.I.T as an App</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
        <div style="background:rgba(0,0,0,.3);border-radius:10px;padding:14px">
          <div style="font-size:.75rem;font-weight:700;color:#e2e8f0;margin-bottom:8px">🤖 Android (Chrome)</div>
          <div style="font-size:.65rem;color:#94b4d4;line-height:1.8">
            1. Open the app URL in Chrome<br>
            2. Tap the <b style="color:#60a5fa">⋮ menu</b> (top-right)<br>
            3. Tap <b style="color:#60a5fa">"Add to Home screen"</b><br>
            4. Tap <b style="color:#60a5fa">Add</b> — done ✓<br>
            <span style="color:#4a5568;font-size:.58rem">Opens full-screen like a native app</span>
          </div>
        </div>
        <div style="background:rgba(0,0,0,.3);border-radius:10px;padding:14px">
          <div style="font-size:.75rem;font-weight:700;color:#e2e8f0;margin-bottom:8px">🍎 iPhone (Safari)</div>
          <div style="font-size:.65rem;color:#94b4d4;line-height:1.8">
            1. Open the app URL in <b style="color:#60a5fa">Safari</b><br>
            2. Tap the <b style="color:#60a5fa">Share button</b> (□↑)<br>
            3. Scroll down → <b style="color:#60a5fa">"Add to Home Screen"</b><br>
            4. Tap <b style="color:#60a5fa">Add</b> — done ✓<br>
            <span style="color:#4a5568;font-size:.58rem">Must use Safari, not Chrome on iOS</span>
          </div>
        </div>
      </div>
      <div style="margin-top:12px;padding-top:10px;border-top:1px solid #1e3a5f;
        font-size:.6rem;color:#4a5568;line-height:1.7">
        💡 Once installed: opens full-screen, no browser bar, looks exactly like a native app.
        The app icon will be your profile photo. Works on both portrait and landscape.
      </div>
    </div>
    """)

# ── CLASS ALERT BANNER ───────────────────────────────────────────
_active_class = _current_class()
_next_cls, _next_time = _next_class()
if _active_class:
    _cls = _active_class
    _sh, _sm = _SLOT_START[_cls["slots"][0]]
    _eh_raw = _SLOT_START[_cls["slots"][-1]][1] + 50
    _eh = _SLOT_START[_cls["slots"][-1]][0] + _eh_raw // 60
    _em = _eh_raw % 60
    st.html(f"""
    <div style="background:linear-gradient(135deg,{_cls['color']}22,{_cls['color']}11);
      border:2px solid {_cls['color']};border-radius:14px;padding:14px 20px;margin-bottom:14px;
      display:flex;align-items:center;gap:14px;animation:pulse-border 2s ease-in-out infinite">
      <div style="font-size:1.6rem">🔔</div>
      <div>
        <div style="font-size:.6rem;color:{_cls['color']};font-family:'Space Mono',monospace;
          font-weight:700;text-transform:uppercase;letter-spacing:.12em">CLASS IN SESSION NOW</div>
        <div style="font-size:.95rem;font-weight:700;color:#e2e8f0;margin:3px 0">{_cls['name']}</div>
        <div style="font-size:.62rem;color:#94b4d4">{_cls['room']}  ·  {_sh:02d}:{_sm:02d} – {_eh:02d}:{_em:02d}</div>
      </div>
      <div style="margin-left:auto;text-align:right">
        <div style="font-size:.58rem;color:{_cls['color']};font-family:'Space Mono',monospace">{_cls['code']}</div>
        <div style="font-size:1.2rem;font-family:'Space Mono',monospace;font-weight:700;color:{_cls['color']}">
          {NOW.strftime('%H:%M')}</div>
      </div>
    </div>""")
elif _next_cls:
    st.html(f"""
    <div style="background:#07080f;border:1px solid #1e3a5f;border-radius:10px;
      padding:10px 18px;margin-bottom:10px;display:flex;align-items:center;gap:12px">
      <div style="font-size:1.1rem">📅</div>
      <div style="font-size:.62rem;color:#4a5568;font-family:'Space Mono',monospace">
        Next: <span style="color:{_next_cls['color']};font-weight:700">{_next_cls['name']}</span>
        · {_next_time} · {_next_cls['room']}
      </div>
    </div>""")

# ════════════════════════════════════════════════════════════════
#  HEADER
# ════════════════════════════════════════════════════════════════
st.markdown(
    '<div style="background:linear-gradient(135deg,#07091a 0%,#0a0f20 50%,#07091a 100%);'
    'border:1px solid #12192b;border-radius:16px;padding:22px 28px 18px;margin-bottom:18px;'
    'position:relative;overflow:hidden">'

    # top gradient bar (animated shimmer)
    '<div class="shimmer-line" style="position:absolute;top:0;left:0;right:0;height:3px"></div>'

    # scan line effect
    '<div style="position:absolute;left:0;right:0;height:1px;background:linear-gradient(90deg,transparent,rgba(37,99,235,.15),transparent);pointer-events:none;animation:scan-line 6s linear infinite"></div>'

    # glow orb
    '<div style="position:absolute;right:-60px;top:-60px;width:200px;height:200px;'
    'border-radius:50%;background:radial-gradient(circle,rgba(29,78,216,.07),transparent 70%);'
    'pointer-events:none"></div>'
    '<div style="position:absolute;left:-40px;bottom:-40px;width:160px;height:160px;'
    'border-radius:50%;background:radial-gradient(circle,rgba(124,58,237,.05),transparent 70%);'
    'pointer-events:none"></div>'

    '<div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px">'
    '<div>'
    '<div style="font-size:.6rem;color:#2563eb;font-weight:700;letter-spacing:.2em;'
    'text-transform:uppercase;font-family:\'Space Mono\',monospace;margin-bottom:8px">'
    '<span class="float-satellite">🛰️</span>  ACADEMIC COMMAND CENTRE  ·  SPRING 2026</div>'
    '<div class="orbit-title" style="font-family:\'Syne\',sans-serif;font-size:2.8rem;font-weight:800;'
    'color:#e2e8f0;line-height:.9;letter-spacing:.02em">O.R.B.I.T.</div>'
    '<div style="font-size:.68rem;color:#4a5568;margin-top:8px;'
    'font-family:\'Space Mono\',monospace;letter-spacing:.05em">'
    'TANAKA ALEX MBENDANA  ·  BEIHANG UNIVERSITY  ·  MSc REMOTE SENSING</div>'
    '</div>'

    f'<div style="display:flex;gap:10px;flex-wrap:wrap;align-items:flex-start">'
    # status box
    f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:10px;'
    f'padding:12px 16px;min-width:140px;text-align:right">'
    f'<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
    f'text-transform:uppercase;letter-spacing:.1em">System Status</div>'
    f'<div style="font-size:.9rem;font-weight:800;font-family:\'Space Mono\',monospace;'
    f'color:{sys_color};margin-top:4px">{sys_label}</div>'
    f'<div style="font-size:.58rem;color:#4a5568;margin-top:4px">{TODAY.strftime("%d %b %Y")}</div>'
    f'</div>'
    # date pills
    f'<div style="display:flex;flex-direction:column;gap:6px">'
    f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:8px;'
    f'padding:7px 12px;font-size:.62rem;font-family:\'Space Mono\',monospace">'
    f'<span style="color:#4a5568">START</span> '
    f'<span style="color:#94b4d4">{SEM_START.strftime("%d %b %Y")}</span></div>'
    f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:8px;'
    f'padding:7px 12px;font-size:.62rem;font-family:\'Space Mono\',monospace">'
    f'<span style="color:#4a5568">END</span> '
    f'<span style="color:#94b4d4">{SEM_END.strftime("%d %b %Y")}</span></div>'
    f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:8px;'
    f'padding:7px 12px;font-size:.62rem;font-family:\'Space Mono\',monospace">'
    f'<span style="color:#4a5568">DAYS LEFT</span> '
    f'<span style="color:#f59e0b;font-weight:700">{days_left}</span></div>'
    f'</div>'
    f'</div>'
    '</div></div>',
    unsafe_allow_html=True
)

# ── TOP KPI BAR ──────────────────────────────────────────────────
k1, k2, k3, k4, k5, k6 = st.columns(6)

def _m(col, lbl, val, accent="#2563eb"):
    with col:
        st.markdown(
            f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:12px;'
            f'padding:14px 12px;position:relative;overflow:hidden">'
            f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{accent}"></div>'
            f'<div style="font-size:.6rem;color:{_TXT3};font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.1em">{lbl}</div>'
            f'<div style="font-size:1.35rem;font-family:\'Space Mono\',monospace;font-weight:700;'
            f'color:{_TXT2};margin-top:5px">{val}</div>'
            f'</div>',
            unsafe_allow_html=True
        )

_m(k1, "📋 Attendance",   _pct(att_frac) if sem_started else "—",  "#2563eb")
_m(k2, "📄 Research",     _pct(res_avg),                            "#7c3aed")
_m(k3, "💰 Balance",      f"¥{net_balance:,.0f}",                   "#10b981" if net_balance >= 0 else "#ef4444")
_m(k4, "💹 Savings Rate", f"{savings_rate:.1f}%",                   "#ea580c")
_m(k5, "📅 Semester",     _pct(sem_frac),                           "#6366f1")
_m(k6, "⏳ Days Left",    str(days_left),                            "#f59e0b")

st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

# ── TABS ─────────────────────────────────────────────────────────
t0, t1, t2, t3, t4, t5, t6 = st.tabs([
    "⬡  OVERVIEW",
    "📋  OBLIGATIONS",
    "📄  RESEARCH",
    "💰  BUDGET",
    "🚀  TRAJECTORY",
    "👤  PROFILE",
    "📱  EXPENSES",
])

# ╔══════════════════════════════════════════════════════════════╗
#  OVERVIEW  —  Dashboard-style layout
# ╚══════════════════════════════════════════════════════════════╝
with t0:
    # ─ Row 1: sidebar stats + main charts ─────────────────────
    L, R = st.columns([1, 3], gap="medium")

    with L:
        # Stat cards stacked like dashboard sidebar
        ACCENT_MAP = [
            ("📋 Attendance",  _pct(att_frac) if sem_started else "—", "#2563eb"),
            ("📄 Research",    _pct(res_avg),                           "#7c3aed"),
            ("💰 Balance",     f"¥{net_balance:,.0f}",                  "#10b981"),
            ("💹 Savings",     f"{savings_rate:.1f}%",                  "#ea580c"),
            ("📅 Semester",    _pct(sem_frac),                          "#6366f1"),
            ("⏳ Days Left",   str(days_left),                           "#f59e0b"),
        ]
        for lbl, val, acc in ACCENT_MAP:
            st.markdown(
                f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:10px;'
                f'padding:12px 14px;margin-bottom:8px;position:relative;overflow:hidden">'
                f'<div style="position:absolute;left:0;top:0;bottom:0;width:3px;background:{acc}"></div>'
                f'<div style="padding-left:10px">'
                f'<div style="font-size:.58rem;color:{_TXT3};font-family:\'Space Mono\',monospace;'
                f'text-transform:uppercase;letter-spacing:.09em">{lbl}</div>'
                f'<div style="font-size:1.2rem;font-family:\'Space Mono\',monospace;font-weight:700;'
                f'color:{_TXT2};margin-top:2px">{val}</div>'
                f'</div></div>',
                unsafe_allow_html=True
            )

        # ORBIT Status box
        if not sem_started:
            box_bg, box_brd, box_tc = "#0c1020", "#12192b", "#4a5568"
            box_msg = "PENDING START"
            box_sub = "Awaiting semester ignition"
        elif drift:
            box_bg, box_brd, box_tc = "#1a0808", "#7f1d1d", "#ef4444"
            box_msg = "⚠ THREATENED"
            box_sub = "Reduce complexity 7 days"
        else:
            box_bg, box_brd, box_tc = "#061208", "#064e3b", "#10b981"
            box_msg = "● ORBIT STABLE"
            box_sub = "Maintain trajectory"
        st.markdown(
            f'<div style="background:{box_bg};border:1px solid {box_brd};border-radius:10px;'
            f'padding:14px;text-align:center;margin-top:4px">'
            f'<div style="font-size:.6rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px">ORBIT STATUS</div>'
            f'<div style="font-size:.85rem;font-weight:800;font-family:\'Space Mono\',monospace;'
            f'color:{box_tc}">{box_msg}</div>'
            f'<div style="font-size:.6rem;color:#4a5568;margin-top:4px">{box_sub}</div>'
            f'</div>',
            unsafe_allow_html=True
        )

    with R:
        # ─ Row 1 of charts: Stability bars + Donut ──────────────
        rc1, rc2 = st.columns([3, 2], gap="medium")

        with rc1:
            st.markdown(
                '<div style="background:#0c1020;border:1px solid #12192b;border-radius:12px;padding:18px">'
                '<div style="font-size:.65rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
                'text-transform:uppercase;letter-spacing:.1em;margin-bottom:14px">📊 Stability Indicators</div>',
                unsafe_allow_html=True
            )
            bars = [
                ("Attendance",   att_frac if sem_started else 0, "#2563eb"),
                ("Research",     res_avg,                         "#7c3aed"),
                ("Budget Health",savings_rate / 100.0,            "#10b981"),
                ("Semester",     sem_frac,                        "#6366f1"),
            ]
            for lbl, val, col in bars:
                w = min(max(val * 100, 0), 100)
                st.markdown(
                    f'<div style="margin-bottom:12px">'
                    f'<div style="display:flex;justify-content:space-between;margin-bottom:3px">'
                    f'<span style="font-size:.68rem;color:#94b4d4">{lbl}</span>'
                    f'<span style="font-size:.68rem;color:{col};font-family:\'Space Mono\',monospace">'
                    f'{w:.1f}%</span></div>'
                    f'<div style="background:#12192b;border-radius:3px;height:6px;overflow:hidden">'
                    f'<div style="width:{w:.1f}%;height:100%;background:{col};border-radius:3px"></div>'
                    f'</div></div>',
                    unsafe_allow_html=True
                )
            st.markdown('</div>', unsafe_allow_html=True)

        with rc2:
            # Donut for research split
            fig_d = go.Figure(go.Pie(
                labels=["P1 Done","P1 Left","P2 Done","P2 Left"],
                values=[p1p*100, (1-p1p)*100, p2p*100, (1-p2p)*100],
                hole=0.62, textinfo="none",
                marker=dict(colors=["#2563eb","#12192b","#7c3aed","#12192b"]),
                hovertemplate="<b>%{label}</b>: %{value:.0f}%<extra></extra>",
            ))
            fig_d.update_layout(**PD, height=220,
                annotations=[dict(text=f"<b>{res_avg*100:.0f}%</b><br><span style='font-size:9px'>research</span>",
                                  x=0.5, y=0.5, font_size=14, font_color="#e2e8f0", showarrow=False)])
            st.plotly_chart(fig_d, use_container_width=True, config={"displayModeBar": False})

        # ─ Row 2 of charts: Budget donut + Drift signals ────────
        rc3, rc4 = st.columns([2, 3], gap="medium")

        with rc3:
            # Budget ring
            cats_v = fin["cats"][fin["cats"]["Budget ¥"] > 0] if not fin["cats"].empty else pd.DataFrame()
            if not cats_v.empty:
                fig_b = go.Figure(go.Pie(
                    labels=cats_v["Category"],
                    values=cats_v["Budget ¥"],
                    hole=0.65, textinfo="none",
                    marker=dict(colors=["#1d4ed8","#2563eb","#3b82f6","#60a5fa","#93c5fd",
                                        "#10b981","#f59e0b","#ea580c","#7c3aed","#6b7280"][:len(cats_v)]),
                    hovertemplate="<b>%{label}</b><br>¥%{value:,.0f}<extra></extra>",
                ))
                fig_b.update_layout(**PD, height=200,
                    annotations=[dict(text=f"<b>¥{net_balance:,.0f}</b><br><span style='font-size:9px'>balance</span>",
                                      x=0.5, y=0.5, font_size=12, font_color="#e2e8f0", showarrow=False)])
                st.plotly_chart(fig_b, use_container_width=True, config={"displayModeBar": False})
            else:
                st.markdown(_kpi_card("Net Balance", f"¥{net_balance:,.0f}", accent="#10b981"), unsafe_allow_html=True)

        with rc4:
            st.markdown(
                '<div style="background:#0c1020;border:1px solid #12192b;border-radius:12px;padding:18px">'
                '<div style="font-size:.65rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
                'text-transform:uppercase;letter-spacing:.1em;margin-bottom:14px">🚨 Drift Detection</div>',
                unsafe_allow_html=True
            )
            if not sem_started:
                st.markdown(
                    '<div style="text-align:center;padding:20px 0">'
                    '<div style="font-size:.75rem;color:#4a5568;font-family:\'Space Mono\',monospace">'
                    'AWAITING SEMESTER IGNITION</div>'
                    '<div style="font-size:.65rem;color:#12192b;margin-top:6px">'
                    'Activates once attendance is entered</div></div>',
                    unsafe_allow_html=True
                )
            else:
                is_att_stale = sem_started and attended == 0
                dr1, dr2 = st.columns(2)
                for dcol, dlbl, is_d, icon in [
                    (dr1, "Attendance Log", is_att_stale,  "📋"),
                    (dr2, "Last Visit",     drift,         "👁"),
                ]:
                    bg_  = "#1a0808" if is_d else "#061208"
                    brd_ = "#7f1d1d" if is_d else "#064e3b"
                    tc_  = "#ef4444" if is_d else "#10b981"
                    txt_ = "STALE"  if is_d else "UPDATED"
                    with dcol:
                        st.markdown(
                            f'<div style="background:{bg_};border:1px solid {brd_};'
                            f'border-radius:10px;padding:12px 8px;text-align:center">'
                            f'<div style="font-size:1rem">{icon}</div>'
                            f'<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
                            f'text-transform:uppercase;letter-spacing:.08em;margin-top:5px">{dlbl}</div>'
                            f'<div style="font-size:.72rem;font-weight:700;color:{tc_};margin-top:2px;'
                            f'font-family:\'Space Mono\',monospace">{txt_}</div></div>',
                            unsafe_allow_html=True
                        )
                if drift:
                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                    st.warning(f"**O.R.B.I.T Alert:** {_days_absent} days since last check-in. Return to dashboard to maintain orbit stability.")
            st.markdown('</div>', unsafe_allow_html=True)

        # ─ Row 3: Budget stat cards ──────────────────────────────
        bc1, bc2, bc3, bc4 = st.columns(4)
        for col, lbl, val, acc in [
            (bc1, "Auto Income",     f"¥{auto_income:,.0f}",  "#2563eb"),
            (bc2, "Spent",           f"¥{fin['ts']:,.0f}",    "#ea580c"),
            (bc3, "Net Balance",     f"¥{net_balance:,.0f}",  "#10b981" if net_balance >= 0 else "#ef4444"),
            (bc4, "Savings Rate",    f"{savings_rate:.1f}%",  "#7c3aed"),
        ]:
            with col:
                st.markdown(
                    f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:10px;'
                    f'padding:12px 14px;text-align:center;margin-top:8px">'
                    f'<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
                    f'text-transform:uppercase;letter-spacing:.09em">{lbl}</div>'
                    f'<div style="font-size:1.1rem;font-family:\'Space Mono\',monospace;font-weight:700;'
                    f'color:{acc};margin-top:4px">{val}</div>'
                    f'</div>',
                    unsafe_allow_html=True
                )


# ╔══════════════════════════════════════════════════════════════╗
#  O — OBLIGATIONS
# ╚══════════════════════════════════════════════════════════════╝
with t1:
    st.markdown(
        '<div style="font-family:\'Syne\',sans-serif;font-size:1.3rem;font-weight:800;'
        'color:#e2e8f0;margin-bottom:4px">📋 Obligations</div>'
        '<div style="font-size:.68rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'margin-bottom:18px">ATTENDANCE COMPLIANCE  ·  SPRING 2026  ·  MIN THRESHOLD: 80%</div>',
        unsafe_allow_html=True
    )

    # ─ Pre-start banner ─────────────────────────────────────────
    if not sem_started:
        st.markdown(
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
            'padding:36px;text-align:center;margin-bottom:20px">'
            '<div style="font-size:2.5rem;margin-bottom:10px">📅</div>'
            '<div style="font-size:.85rem;font-weight:700;color:#e2e8f0;'
            'font-family:\'Space Mono\',monospace">SEMESTER HAS NOT STARTED</div>'
            '<div style="font-size:.7rem;color:#4a5568;margin-top:10px;line-height:1.7;max-width:420px;margin-left:auto;margin-right:auto">'
            'Classes begin <b style="color:#60a5fa">03 March 2026</b>.<br>'
            'As you attend each class, update the <b style="color:#60a5fa">ATTENDED</b> column '
            'in the <b style="color:#60a5fa">📊 Analytics</b> sheet.<br>'
            'All metrics will populate automatically.</div></div>',
            unsafe_allow_html=True
        )

    # ─ Course cards ─────────────────────────────────────────────
    if not att_df.empty:
        cols_c = st.columns(len(att_df))
        for i, (_, row) in enumerate(att_df.iterrows()):
            pv = float(row["Pct"])
            pv = pv / 100 if pv > 1 else pv
            if not sem_started:
                cc, cb, cd = "#4a5568", "#0c1020", "#12192b"
            else:
                cc = "#10b981" if pv >= ATTEND_THRESHOLD else "#ef4444"
                cb = "#061208" if pv >= ATTEND_THRESHOLD else "#1a0808"
                cd = "#064e3b" if pv >= ATTEND_THRESHOLD else "#7f1d1d"
            with cols_c[i]:
                st.markdown(
                    f'<div style="background:{cb};border:1px solid {cd};border-radius:12px;'
                    f'padding:14px 10px;text-align:center">'
                    f'<div style="font-size:.58rem;color:#4a5568;font-weight:700;text-transform:uppercase;'
                    f'letter-spacing:.06em;font-family:\'Space Mono\',monospace;min-height:24px;'
                    f'display:flex;align-items:center;justify-content:center">{row["Course"]}</div>'
                    f'<div style="font-size:1.5rem;font-family:\'Space Mono\',monospace;font-weight:700;'
                    f'color:{cc};margin:8px 0 2px">{int(row["Attended"])}/{int(row["Total"])}</div>'
                    f'<div style="font-size:.65rem;color:{cc}">{pv*100:.0f}%</div>'
                    f'<div style="background:#12192b;border-radius:3px;height:5px;overflow:hidden;margin-top:8px">'
                    f'<div style="width:{min(pv*100,100):.1f}%;height:100%;background:{cc};border-radius:3px"></div></div>'
                    f'<div style="font-size:.58rem;color:#4a5568;margin-top:5px">{int(row["Remaining"])} left</div>'
                    f'</div>',
                    unsafe_allow_html=True
                )

    # ─ Attendance bar chart (only when active) ───────────────────
    if sem_started and not att_df.empty:
        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
        fig_a = go.Figure()
        fig_a.add_trace(go.Bar(x=att_df["Course"], y=att_df["Total"],
                               name="Total", marker_color="#12192b",
                               hovertemplate="%{x}<br>Total: %{y}<extra></extra>"))
        fig_a.add_trace(go.Bar(x=att_df["Course"], y=att_df["Attended"],
                               name="Attended", marker_color="#2563eb",
                               hovertemplate="%{x}<br>Attended: %{y}<extra></extra>"))
        if att_df["Total"].mean() > 0:
            fig_a.add_hline(y=att_df["Total"].mean() * ATTEND_THRESHOLD,
                            line_dash="dot", line_color="#ef4444",
                            annotation_text="80% min", annotation_font_color="#ef4444",
                            annotation_position="right")
        fig_a.update_layout(**PD, height=280, barmode="overlay", xaxis_tickangle=-20,
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        st.plotly_chart(fig_a, use_container_width=True, config={"displayModeBar": False})

    # ─ Timetable ─────────────────────────────────────────────────
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:.65rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px">🗓️ Course Timetable — Spring 2026</div>',
        unsafe_allow_html=True
    )
    DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    hdr = ('<div style="display:grid;grid-template-columns:70px repeat(7,1fr);'
           'gap:4px;margin-bottom:4px">'
           '<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
           'text-transform:uppercase;text-align:center;padding:5px 0">SLOT</div>')
    for d in DAYS:
        hdr += (f'<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
                f'text-transform:uppercase;text-align:center;padding:5px 0">{d[:3].upper()}</div>')
    hdr += '</div>'
    st.markdown(hdr, unsafe_allow_html=True)

    if sched:
        for sd in sched:
            rh = ('<div style="display:grid;grid-template-columns:70px repeat(7,1fr);'
                  'gap:4px;margin-bottom:4px">'
                  f'<div style="display:flex;flex-direction:column;justify-content:center;'
                  f'padding:4px 6px;background:#0c1020;border:1px solid #12192b;border-radius:7px">'
                  f'<span style="font-size:.58rem;color:#2563eb;font-family:\'Space Mono\',monospace;'
                  f'font-weight:700">{sd["slot"]}</span>'
                  f'<span style="font-size:.52rem;color:#4a5568">{sd["time"]}</span></div>')
            for cell in sd["cells"]:
                if cell and cell != "nan":
                    lines = cell.split("\n")
                    name   = lines[0][:28] if lines else cell[:28]
                    detail = lines[1][:22] if len(lines) > 1 else ""
                    rh += (f'<div style="background:#0d1829;border:1px solid #1e3a5f;'
                           f'border-radius:7px;padding:7px 8px;font-size:.65rem;'
                           f'color:#b8c4d0;min-height:44px;line-height:1.35">'
                           f'<b style="color:#60a5fa">{name}</b>')
                    if detail:
                        rh += f'<br><span style="color:#4a5568;font-size:.56rem">{detail}</span>'
                    rh += '</div>'
                else:
                    rh += '<div style="background:transparent;border:1px solid transparent;border-radius:7px;min-height:44px"></div>'
            rh += '</div>'
            st.markdown(rh, unsafe_allow_html=True)

    # ─ Weekly attendance log ─────────────────────────────────────
    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:.65rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px">'
        '📆 Week-by-Week Attendance Log</div>',
        unsafe_allow_html=True
    )
    if not watt_df.empty and sem_started:
        # render as styled HTML table, not st.dataframe
        rows_html = ('<div style="background:#0c1020;border:1px solid #12192b;'
                     'border-radius:12px;overflow:hidden">'
                     '<div style="display:grid;grid-template-columns:130px repeat(7,1fr);'
                     'border-bottom:1px solid #12192b">')
        for h in ["Week"] + COURSE_SHORT + ["Σ"]:
            rows_html += (f'<div style="padding:8px 10px;font-size:.58rem;color:#4a5568;'
                          f'font-family:\'Space Mono\',monospace;text-transform:uppercase;'
                          f'letter-spacing:.07em">{h}</div>')
        rows_html += '</div>'
        for _, row in watt_df.iterrows():
            rows_html += '<div style="display:grid;grid-template-columns:130px repeat(7,1fr);border-bottom:1px solid #0d1117">'
            rows_html += (f'<div style="padding:7px 10px;font-size:.65rem;color:#94b4d4;'
                          f'font-family:\'Space Mono\',monospace">{row["Week"]}</div>')
            for c in COURSE_SHORT:
                v = row.get(c, 0)
                col_c = "#10b981" if v == 1 else "#4a5568"
                rows_html += (f'<div style="padding:7px 10px;font-size:.65rem;color:{col_c};'
                               f'font-family:\'Space Mono\',monospace;text-align:center">{v}</div>')
            tot = row.get("Σ", 0)
            rows_html += (f'<div style="padding:7px 10px;font-size:.65rem;color:#60a5fa;'
                          f'font-family:\'Space Mono\',monospace;text-align:center;font-weight:700">{tot}</div>')
            rows_html += '</div>'
        rows_html += '</div>'
        st.markdown(rows_html, unsafe_allow_html=True)
    else:
        st.markdown(
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:12px;'
            'padding:20px;text-align:center;font-size:.7rem;color:#4a5568;'
            'font-family:\'Space Mono\',monospace">No attendance data yet · Update Excel to populate</div>',
            unsafe_allow_html=True
        )


# ╔══════════════════════════════════════════════════════════════╗
#  R — RESEARCH
# ╚══════════════════════════════════════════════════════════════╝
with t2:
    st.markdown(
        '<div style="font-family:\'Syne\',sans-serif;font-size:1.3rem;font-weight:800;'
        'color:#e2e8f0;margin-bottom:4px">📄 Research Hub</div>'
        '<div style="font-size:.68rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'margin-bottom:18px">MSc REMOTE SENSING  ·  BEIHANG 2025–2026  ·  TARGET: RSE / GLOBAL CHANGE BIOLOGY</div>',
        unsafe_allow_html=True
    )

    rp1, rp2 = st.columns(2, gap="large")
    for col, title, pv, secs, idx, color, icon in [
        (rp1, p1t, p1p, p1s, "PAPER 1", "#2563eb", "🌀"),
        (rp2, p2t, p2p, p2s, "PAPER 2", "#7c3aed", "🌧️"),
    ]:
        with col:
            nd = len(secs[secs["Status"]=="✅"]) if not secs.empty else 0
            nw = len(secs[secs["Status"]=="⏳"]) if not secs.empty else 0
            nl = len(secs[secs["Status"]=="🔒"]) if not secs.empty else 0
            w  = min(max(pv * 100, 0), 100)
            st.markdown(
                f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
                f'padding:20px;position:relative;overflow:hidden;margin-bottom:14px">'
                f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{color}"></div>'
                f'<div style="font-size:.6rem;color:{color};font-weight:700;text-transform:uppercase;'
                f'letter-spacing:.14em;font-family:\'Space Mono\',monospace">{icon} {idx}</div>'
                f'<div style="font-size:.85rem;font-weight:600;color:#e2e8f0;margin:8px 0 5px;line-height:1.35">'
                f'{title[:65]}{"…" if len(title)>65 else ""}</div>'
                f'<div style="font-size:2rem;font-family:\'Space Mono\',monospace;font-weight:700;'
                f'color:{color};line-height:1">{pv*100:.0f}%</div>'
                f'<div style="background:#12192b;border-radius:3px;height:6px;overflow:hidden;margin-top:8px">'
                f'<div style="width:{w:.1f}%;height:100%;background:{color};border-radius:3px"></div></div>'
                f'<div style="display:flex;gap:10px;margin-top:10px">'
                f'<span style="font-size:.62rem;color:#10b981">✅ {nd} done</span>'
                f'<span style="font-size:.62rem;color:#f59e0b">⏳ {nw} in progress</span>'
                f'<span style="font-size:.62rem;color:#4a5568">🔒 {nl} locked</span>'
                f'</div></div>',
                unsafe_allow_html=True
            )
            # Section pipeline
            if not secs.empty:
                pipe = ('<div style="background:#0c1020;border:1px solid #12192b;'
                        'border-radius:12px;overflow:hidden">'
                        '<div style="padding:9px 12px;font-size:.58rem;color:#4a5568;'
                        'font-family:\'Space Mono\',monospace;text-transform:uppercase;'
                        'letter-spacing:.1em;border-bottom:1px solid #12192b">'
                        'Section Pipeline</div>')
                for _, row in secs.iterrows():
                    sv   = str(row.get("Status","🔒"))
                    icn  = "✅" if sv=="✅" else ("⏳" if sv=="⏳" else "🔒")
                    ic   = "#10b981" if icn=="✅" else ("#f59e0b" if icn=="⏳" else "#4a5568")
                    tgt  = str(row.get("Target",""))
                    name = str(row.get("Section",""))[:40]
                    pipe += (f'<div style="display:flex;align-items:center;padding:8px 12px;'
                             f'border-bottom:1px solid #0d1117;font-size:.75rem;'
                             f'transition:background .15s">'
                             f'<span style="flex:0 0 22px;font-size:.85rem">{icn}</span>'
                             f'<span style="flex:1;color:#b8c4d0;padding:0 8px">{name}</span>'
                             f'<span style="flex:0 0 76px;text-align:right;font-size:.62rem;'
                             f'color:{ic};font-family:\'Space Mono\',monospace">{tgt}</span></div>')
                pipe += '</div>'
                st.markdown(pipe, unsafe_allow_html=True)

    # Publication scatter
    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:.65rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px">'
        '📅 Publication Timeline</div>',
        unsafe_allow_html=True
    )
    all_s = []
    for sdf, pname in [(p1s,"Paper 1"),(p2s,"Paper 2")]:
        if not sdf.empty:
            for _, row in sdf.iterrows():
                all_s.append({"Paper":pname,
                               "Section":str(row.get("Section",""))[:36],
                               "Target": str(row.get("Target","")),
                               "Status": str(row.get("Status","🔒"))})
    if all_s:
        tdf = pd.DataFrame(all_s)
        fig_tl = go.Figure()
        cmap = {"Paper 1":"#2563eb","Paper 2":"#7c3aed"}
        smap = {"✅":"#10b981","⏳":"#f59e0b","🔒":"#4a5568"}
        for paper, grp in tdf.groupby("Paper"):
            fig_tl.add_trace(go.Scatter(
                x=grp["Target"], y=grp["Section"],
                mode="markers",
                name=paper,
                marker=dict(
                    size=12,
                    color=[smap.get(s,"#4a5568") for s in grp["Status"]],
                    line=dict(width=2, color=cmap[paper]),
                    symbol="circle",
                ),
                hovertemplate="<b>%{y}</b><br>%{x}<extra></extra>",
            ))
        fig_tl.update_layout(**PD, height=420,
                             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        st.plotly_chart(fig_tl, use_container_width=True, config={"displayModeBar": False})


# ╔══════════════════════════════════════════════════════════════╗
#  B — BUDGET  (Financial Command)
# ╚══════════════════════════════════════════════════════════════╝
with t3:
    current_monthly = fin["current_monthly"]
    is_year2 = date.today() >= YEAR2_START

    st.markdown(
        f'<div style="font-family:\'Syne\',sans-serif;font-size:1.3rem;font-weight:800;'
        f'color:{_TXT2};margin-bottom:4px">💰 Financial Command</div>'
        f'<div style="font-size:.68rem;color:{_TXT3};font-family:\'Space Mono\',monospace;'
        f'margin-bottom:18px">{"YEAR 2" if is_year2 else "YEAR 1"}  ·  STIPEND: ¥{current_monthly:,.0f}/mo  ·  '
        f'TOTAL RECEIVED: ¥{auto_income:,.0f}  ·  NET BALANCE: ¥{net_balance:,.0f}</div>',
        unsafe_allow_html=True
    )

    if not is_year2:
        y2_days = (YEAR2_START - date.today()).days
        st.markdown(
            f'<div style="background:linear-gradient(135deg,#0a0820,#0c0a00);border:1px solid #78350f;'
            f'border-radius:10px;padding:12px 18px;margin-bottom:14px;display:flex;align-items:center;gap:12px">'
            f'<span style="font-size:1.1rem">📅</span>'
            f'<div><span style="font-size:.62rem;color:#f59e0b;font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.1em;font-weight:700">'
            f'YEAR 2 BEGINS IN {y2_days} DAYS — STIPEND DROPS TO ¥{YEAR2_MONTHLY:,.0f}/mo</span>'
            f'<div style="font-size:.65rem;color:#92400e;margin-top:2px">'
            f'¥{YEAR1_MONTHLY - YEAR2_MONTHLY:,.0f}/mo reduction · Maximise savings before Sep 2026</div>'
            f'</div></div>',
            unsafe_allow_html=True
        )

    if fin["bonus_pending"] > 0:
        st.markdown(
            f'<div style="background:linear-gradient(135deg,#1a1200,#0c0a00);border:1px solid #78350f;'
            f'border-radius:10px;padding:12px 18px;margin-bottom:14px;display:flex;align-items:center;gap:12px">'
            f'<span style="font-size:1.1rem">⏳</span>'
            f'<div><span style="font-size:.62rem;color:#f59e0b;font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.1em;font-weight:700">'
            f'PENDING BONUS — ¥{fin["bonus_pending"]:,.0f}</span>'
            f'<div style="font-size:.65rem;color:#92400e;margin-top:2px">'
            f'Face-scan registration bonus · Confirmed · Awaiting disbursement</div>'
            f'</div></div>',
            unsafe_allow_html=True
        )

    # ── Current Savings Editor ────────────────────────────────────
    with st.expander("✏️ Update Current Savings Balance", expanded=False):
        st.caption("Enter what's actually in your bank/wallet right now. This is your ground truth — the app projects everything forward from this number.")
        _sc1, _sc2 = st.columns([3, 1])
        with _sc1:
            new_sav = st.number_input("Current savings (¥)", min_value=0.0, step=100.0,
                                      value=float(current_savings), label_visibility="collapsed")
        with _sc2:
            if st.button("💾 Save", use_container_width=True, key="save_savings_btn"):
                _save_savings(new_sav)
                current_savings = new_sav
                st.success("Saved!")
                st.rerun()

    bm1, bm2, bm3, bm4, bm5 = st.columns(5)
    months_runway = net_balance / monthly_burn if monthly_burn > 0 else 99
    total_with_bonus = net_balance + fin["bonus_pending"]

    for col, lbl, val, acc in [
        (bm1, "📥 Auto Income",  f"¥{auto_income:,.0f}",      "#2563eb"),
        (bm2, "💸 Total Spent",  f"¥{fin['ts']:,.0f}",         "#ea580c"),
        (bm3, "✅ Net Balance",       f"¥{net_balance:,.0f}",      "#10b981" if net_balance >= 0 else "#ef4444"),
        (bm4, "📈 Monthly Burn", f"¥{monthly_burn:,.0f}/mo",  "#f59e0b"),
        (bm5, "⏳ Bonus Pending",    f"¥{fin['bonus_pending']:,.0f}", "#7c3aed"),
    ]:
        with col:
            st.markdown(_kpi_card(lbl, val, accent=acc), unsafe_allow_html=True)

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    if spend_ratio < 0.70:
        hbg, hbrd, htc = "#061208", "#064e3b", "#10b981"
        hmsg = "💚 EXCELLENT — STRONG SAVINGS DISCIPLINE"
    elif spend_ratio < 0.85:
        hbg, hbrd, htc = "#1a1200", "#78350f", "#f59e0b"
        hmsg = "⚠️  ON TRACK — MONITOR SPENDING"
    else:
        hbg, hbrd, htc = "#1a0808", "#7f1d1d", "#ef4444"
        hmsg = "🔴 CAUTION — REVIEW EXPENSES"
    st.markdown(
        f'<div style="background:{hbg};border:1px solid {hbrd};border-radius:12px;'
        f'padding:14px 20px;margin-bottom:18px">'
        f'<div style="font-size:.6rem;color:{htc};font-weight:700;text-transform:uppercase;'
        f'letter-spacing:.12em;font-family:\'Space Mono\',monospace">FINANCIAL HEALTH</div>'
        f'<div style="font-size:.9rem;font-weight:700;color:#e2e8f0;margin-top:5px">{hmsg}</div>'
        f'<div style="font-size:.68rem;color:#94b4d4;margin-top:4px">'
        f'{spend_ratio*100:.1f}% of income spent  ·  Savings rate: {savings_rate:.1f}%  ·  '
        f'Runway: {months_runway:.1f} months  ·  With bonus: ¥{total_with_bonus:,.0f}</div>'
        f'</div>',
        unsafe_allow_html=True
    )

    # Savings projection — starts from current_savings (ground truth) going forward 18 months
    proj_months, proj_savings_line, proj_income_line = [], [], []
    bal = current_savings
    est_monthly_exp = monthly_burn if monthly_burn > 0 else 1500.0  # fallback ¥1500/mo
    for m in range(0, 19):  # 18 months forward from today
        yr = TODAY.year + (TODAY.month - 1 + m) // 12
        mn = (TODAY.month - 1 + m) % 12 + 1
        mo_d = date(yr, mn, 1)
        stipend = YEAR2_MONTHLY if mo_d >= YEAR2_START else YEAR1_MONTHLY
        if m > 0:
            bal += stipend - est_monthly_exp
        proj_months.append(mo_d.strftime("%b %Y"))
        proj_savings_line.append(bal)
        proj_income_line.append(stipend)

    fig_proj = go.Figure()
    fig_proj.add_trace(go.Scatter(
        x=proj_months, y=proj_savings_line, mode="lines+markers", name="Projected Savings",
        line=dict(color="#10b981", width=2.5), marker=dict(size=6, color="#10b981"),
        hovertemplate="%{x}<br>¥%{y:,.0f}<extra></extra>",
        fill="tozeroy", fillcolor="rgba(16,185,129,.06)",
    ))
    fig_proj.add_trace(go.Bar(
        x=proj_months, y=proj_income_line, name="Monthly Stipend",
        marker_color=["#7c3aed" if date(int(m.split()[1]), ["Jan","Feb","Mar","Apr","May","Jun",
                      "Jul","Aug","Sep","Oct","Nov","Dec"].index(m.split()[0])+1, 1) >= YEAR2_START
                      else "#2563eb" for m in proj_months],
        opacity=0.35, yaxis="y2",
        hovertemplate="%{x}<br>Stipend: ¥%{y:,.0f}<extra></extra>",
    ))
    fig_proj.add_hline(y=0, line_color="#ef4444", line_width=1,
                       annotation_text="Zero", annotation_font_color="#ef4444")
    if "Sep 2026" in proj_months:
        y2i = proj_months.index("Sep 2026")
        fig_proj.add_vline(x=y2i, line_dash="dash", line_color="#7c3aed",
                           annotation_text="Year 2 (¥3k/mo)",
                           annotation_font_color="#7c3aed", annotation_position="top left")
    fig_proj.update_layout(
        **PD, height=280, barmode="overlay",
        yaxis2=dict(title="Stipend (¥)", overlaying="y", side="right",
                    showgrid=False, tickprefix="¥"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        title=dict(text=f"18-Month Savings Projection from ¥{current_savings:,.0f} · Est. monthly spend: ¥{est_monthly_exp:,.0f}",
                   font_color="#4a5568", font_size=11),
    )
    fig_proj.update_yaxes(title_text="Savings (¥)", tickprefix="¥", selector=dict(type="linear"), row=None, col=None)
    st.plotly_chart(fig_proj, use_container_width=True, config={"displayModeBar": False})

    sa_col, sb_col = st.columns(2, gap="medium")
    with sa_col:
        # Donut: expenses from Excel vs current savings (what we know)
        _donut_exp = fin["ts"] if fin["ts"] > 0 else 1.0
        _donut_sav = max(0.01, net_balance)
        fig_donut = go.Figure(go.Pie(
            labels=["Logged Expenses", "Current Savings"],
            values=[_donut_exp, _donut_sav],
            hole=0.65, marker_colors=["#ef4444", "#10b981"],
            textinfo="none", hovertemplate="%{label}: ¥%{value:,.0f}<extra></extra>",
        ))
        fig_donut.update_layout(**PD, height=220,
                                 title=dict(text="Income Allocation", font_color="#4a5568", font_size=11),
                                 showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.15))
        st.plotly_chart(fig_donut, use_container_width=True, config={"displayModeBar": False})

    with sb_col:
        target_y2 = YEAR2_MONTHLY * 3
        proj_savings = net_balance + (months_to_y2 * max(0, YEAR1_MONTHLY - monthly_burn))
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=proj_savings,
            delta={"reference": target_y2, "valueformat": ",.0f"},
            number={"prefix": "¥", "valueformat": ",.0f", "font": {"size": 18, "color": "#e2e8f0"}},
            gauge={
                "axis": {"range": [0, target_y2 * 1.2], "tickcolor": "#4a5568"},
                "bar": {"color": "#10b981" if proj_savings >= target_y2 * 0.75 else "#f59e0b" if proj_savings >= target_y2 * 0.4 else "#ef4444"},
                "bgcolor": "#0c1020", "bordercolor": "#12192b",
                "steps": [
                    {"range": [0, target_y2 * 0.4], "color": "rgba(239,68,68,.1)"},
                    {"range": [target_y2 * 0.4, target_y2 * 0.75], "color": "rgba(245,158,11,.1)"},
                    {"range": [target_y2 * 0.75, target_y2 * 1.2], "color": "rgba(16,185,129,.1)"},
                ],
                "threshold": {"line": {"color": "#7c3aed", "width": 2}, "value": target_y2},
            },
            title={"text": "Year 2 Buffer Readiness", "font": {"color": "#4a5568", "size": 11}},
        ))
        fig_gauge.update_layout(**PD, height=220)
        st.plotly_chart(fig_gauge, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    bl, br = st.columns(2, gap="large")
    with bl:
        st.markdown(
            f'<div style="font-size:.65rem;color:{_TXT3};font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px">📜 Transactions</div>',
            unsafe_allow_html=True
        )
        tx = fin["tx"]
        if not tx.empty:
            th = (f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:12px;overflow:hidden">'
                  f'<div style="display:grid;grid-template-columns:40px 1fr 100px 80px 80px;border-bottom:1px solid {_BRD}">')
            for h_lbl in ["Wk", "Description", "Category", "Amount", "Balance"]:
                th += f'<div style="padding:8px 10px;font-size:.58rem;color:{_TXT3};font-family:\'Space Mono\',monospace;text-transform:uppercase;letter-spacing:.07em">{h_lbl}</div>'
            th += '</div>'
            for _, row in tx.iterrows():
                th += f'<div style="display:grid;grid-template-columns:40px 1fr 100px 80px 80px;border-bottom:1px solid #0d1117">'
                th += f'<div style="padding:7px 10px;font-size:.68rem;color:{_TXT3};font-family:\'Space Mono\',monospace">{row["Wk"]}</div>'
                th += f'<div style="padding:7px 10px;font-size:.68rem;color:{_TXT}">{str(row["Description"])[:26]}</div>'
                th += f'<div style="padding:7px 10px;font-size:.68rem;color:#60a5fa">{str(row["Category"])[:14]}</div>'
                th += f'<div style="padding:7px 10px;font-size:.68rem;color:#ea580c;font-family:\'Space Mono\',monospace;text-align:right">¥{row["Amount ¥"]:,.0f}</div>'
                th += f'<div style="padding:7px 10px;font-size:.68rem;color:#10b981;font-family:\'Space Mono\',monospace;text-align:right">¥{row["Balance"]:,.0f}</div>'
                th += '</div>'
            th += '</div>'
            st.markdown(th, unsafe_allow_html=True)
            st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
            fig_tx = go.Figure(go.Waterfall(
                x=[f"Wk{r['Wk']} {str(r['Description'])[:14]}" for _, r in tx.iterrows()],
                y=[-r["Amount ¥"] for _, r in tx.iterrows()],
                connector={"line": {"color": "#12192b"}},
                decreasing={"marker": {"color": "#ef4444"}},
                increasing={"marker": {"color": "#10b981"}},
                base=auto_income,
            ))
            fig_tx.update_layout(**PD, height=240, xaxis_tickangle=-30,
                                 title=dict(text="Balance Waterfall", font_color="#4a5568", font_size=11))
            st.plotly_chart(fig_tx, use_container_width=True, config={"displayModeBar": False})
        else:
            st.markdown(
                f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:12px;'
                f'padding:24px;text-align:center;font-size:.7rem;color:{_TXT3};'
                f'font-family:\'Space Mono\',monospace">No transactions · Add to Finance sheet</div>',
                unsafe_allow_html=True
            )

    with br:
        st.markdown(
            f'<div style="font-size:.65rem;color:{_TXT3};font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px">📊 Category Breakdown</div>',
            unsafe_allow_html=True
        )
        cats = fin["cats"]
        if not cats.empty:
            cf = cats[cats["Budget ¥"] > 0]
            if not cf.empty:
                fig_c = go.Figure()
                fig_c.add_trace(go.Bar(name="Budget", x=cf["Category"], y=cf["Budget ¥"], marker_color="#12192b"))
                fig_c.add_trace(go.Bar(name="Spent",  x=cf["Category"], y=cf["Spent ¥"],  marker_color="#2563eb"))
                fig_c.update_layout(**PD, barmode="overlay", height=240, xaxis_tickangle=-25,
                                    legend=dict(orientation="h", yanchor="bottom", y=1.02))
                st.plotly_chart(fig_c, use_container_width=True, config={"displayModeBar": False})
            ct = (f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:12px;overflow:hidden">'
                  f'<div style="display:grid;grid-template-columns:1fr 80px 70px 70px;border-bottom:1px solid {_BRD}">')
            for h_lbl in ["Category", "Budget", "Spent", "Left"]:
                ct += f'<div style="padding:8px 10px;font-size:.58rem;color:{_TXT3};font-family:\'Space Mono\',monospace;text-transform:uppercase;letter-spacing:.07em">{h_lbl}</div>'
            ct += '</div>'
            for _, row in cats.iterrows():
                lv = row.get("Left ¥", 0)
                lc = "#10b981" if lv > 0 else ("#ef4444" if lv < 0 else "#4a5568")
                ct += f'<div style="display:grid;grid-template-columns:1fr 80px 70px 70px;border-bottom:1px solid #0d1117">'
                ct += f'<div style="padding:7px 10px;font-size:.68rem;color:{_TXT}">{row["Category"]}</div>'
                ct += f'<div style="padding:7px 10px;font-size:.68rem;color:{_TXT3};font-family:\'Space Mono\',monospace;text-align:right">¥{row["Budget ¥"]:,.0f}</div>'
                ct += f'<div style="padding:7px 10px;font-size:.68rem;color:#ea580c;font-family:\'Space Mono\',monospace;text-align:right">¥{row["Spent ¥"]:,.0f}</div>'
                ct += f'<div style="padding:7px 10px;font-size:.68rem;color:{lc};font-family:\'Space Mono\',monospace;text-align:right">¥{lv:,.0f}</div>'
                ct += '</div>'
            ct += '</div>'
            st.markdown(ct, unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:12px;'
                f'padding:24px;text-align:center;font-size:.7rem;color:{_TXT3};'
                f'font-family:\'Space Mono\',monospace">No category data</div>',
                unsafe_allow_html=True
            )

    st.markdown(
        f'<div style="font-size:.55rem;color:{_TXT3};font-family:\'Space Mono\',monospace;'
        f'text-align:center;padding:10px">🔄 Income auto-updates monthly · Expenses from Excel · Refreshes every 5 min</div>',
        unsafe_allow_html=True
    )

# ╔══════════════════════════════════════════════════════════════╗
#  T — TRAJECTORY
# ╚══════════════════════════════════════════════════════════════╝
with t4:
    st.markdown(
        '<div style="font-family:\'Syne\',sans-serif;font-size:1.3rem;font-weight:800;'
        'color:#e2e8f0;margin-bottom:4px">🚀 Trajectory</div>'
        '<div style="font-size:.68rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'margin-bottom:18px">2026–2028 DOCTORAL ARC  ·  CONSISTENT CONTROLLED ASCENT</div>',
        unsafe_allow_html=True
    )

    # Semester overview card
    sem_disp = _pct(sem_frac) if sem_started else "NOT STARTED"
    st.markdown(
        f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
        f'padding:20px;margin-bottom:18px;position:relative;overflow:hidden">'
        f'<div style="position:absolute;top:0;left:0;right:0;height:2px;'
        f'background:linear-gradient(90deg,#1d4ed8,#7c3aed,#10b981)"></div>'
        f'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:16px">'
        f'<div style="text-align:center">'
        f'<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;text-transform:uppercase;letter-spacing:.1em">Semester</div>'
        f'<div style="font-size:1.4rem;font-family:\'Space Mono\',monospace;font-weight:700;color:#e2e8f0;margin-top:4px">{sem_disp}</div></div>'
        f'<div style="text-align:center">'
        f'<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;text-transform:uppercase;letter-spacing:.1em">Days Left</div>'
        f'<div style="font-size:1.4rem;font-family:\'Space Mono\',monospace;font-weight:700;color:#f59e0b;margin-top:4px">{days_left}</div></div>'
        f'<div style="text-align:center">'
        f'<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;text-transform:uppercase;letter-spacing:.1em">Sem Start</div>'
        f'<div style="font-size:1.2rem;font-family:\'Space Mono\',monospace;font-weight:700;color:#94b4d4;margin-top:4px">{SEM_START.strftime("%d %b")}</div></div>'
        f'<div style="text-align:center">'
        f'<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;text-transform:uppercase;letter-spacing:.1em">Sem End</div>'
        f'<div style="font-size:1.2rem;font-family:\'Space Mono\',monospace;font-weight:700;color:#94b4d4;margin-top:4px">{SEM_END.strftime("%d %b")}</div></div>'
        f'</div>'
        # Progress bars
        + "".join([
            f'<div style="margin-bottom:10px">'
            f'<div style="display:flex;justify-content:space-between;margin-bottom:3px">'
            f'<span style="font-size:.65rem;color:#94b4d4">{lbl}</span>'
            f'<span style="font-size:.65rem;color:{col};font-family:\'Space Mono\',monospace">{_pct(val)}</span>'
            f'</div>'
            f'<div style="background:#12192b;border-radius:3px;height:5px;overflow:hidden">'
            f'<div style="width:{min(val*100,100):.1f}%;height:100%;background:{col};border-radius:3px"></div>'
            f'</div></div>'
            for lbl,val,col in [
                ("📅 Semester Progress",    sem_frac,                              "#6366f1"),
                ("📄 Research Velocity",    res_avg,                               "#7c3aed"),
                ("📋 Academic Compliance",  att_frac if sem_started else 0,        "#2563eb"),
            ]
        ])
        + '</div>',
        unsafe_allow_html=True
    )

    # Strategic horizon cards
    st.markdown(
        '<div style="font-size:.65rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'text-transform:uppercase;letter-spacing:.1em;margin-bottom:12px">'
        '🗺️ Strategic Time Horizon</div>',
        unsafe_allow_html=True
    )
    mc1, mc2, mc3 = st.columns(3)
    CURRENT_YEAR = TODAY.year
    MILESTONES = [
        ("2026", "🔬 Stabilisation",
         "MSc completion · Paper 1 submission · System discipline established",
         "#2563eb"),
        ("2027", "📡 Consolidation",
         "Paper 2 submission · Conference positioning · PhD research preparation",
         "#7c3aed"),
        ("2028", "🎓 Doctoral Arc",
         "PhD application · RSE-level output · Research velocity at peak",
         "#10b981"),
    ]
    for col, (year, phase, desc, color) in zip([mc1, mc2, mc3], MILESTONES):
        is_cur = (str(CURRENT_YEAR) == year)
        with col:
            bg     = "linear-gradient(145deg,#0c1020,#0d1829)" if is_cur else "#0c1020"
            border = f"2px solid {color}" if is_cur else "1px solid #12192b"
            badge_html  = ""
            days_html   = ""
            if is_cur:
                badge_html = (
                    f'<div style="margin-top:10px;display:inline-block;padding:3px 10px;'
                    f'border-radius:20px;background:{color}22;border:1px solid {color}55;'
                    f'font-size:.58rem;font-weight:700;color:{color};'
                    f'font-family: Space Mono,monospace;letter-spacing:.1em">'
                    f'{CURRENT_YEAR} - ACTIVE</div>'
                )
            if is_cur and year == "2026":
                days_html = (
                    f'<div style="margin-top:6px;font-size:.62rem;color:#f59e0b;'
                    f'font-family: Space Mono,monospace">{days_left} days remaining in semester</div>'
                )
            st.markdown(
                f'<div style="background:{bg};border:{border};border-radius:14px;'
                f'padding:20px 16px;position:relative;overflow:hidden">'
                f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{color};opacity:.7"></div>'
                f'<div style="font-family:\'Syne\',sans-serif;font-size:2.2rem;font-weight:800;'
                f'color:{color};line-height:1">{year}</div>'
                f'<div style="font-size:.82rem;font-weight:700;color:#e2e8f0;margin:8px 0 5px">{phase}</div>'
                f'<div style="font-size:.68rem;color:#4a5568;line-height:1.6">{desc}</div>'
                f'{badge_html}{days_html}</div>',
                unsafe_allow_html=True
            )

    # Gantt — no add_vline (causes crash with string dates on timeline charts)
    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:.65rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px">'
        '📊 Research Milestones — Gantt View</div>',
        unsafe_allow_html=True
    )
    MM = {"Mar 2026":"2026-03","Apr 2026":"2026-04","May 2026":"2026-05",
          "Jun 2026":"2026-06","Jul 2026":"2026-07","Aug 2026":"2026-08",
          "Sep 2026":"2026-09","Oct 2026":"2026-10","Nov 2026":"2026-11",
          "Dec 2026":"2026-12","2027":"2027-06"}
    gantt = []
    for sdf, pname in [(p1s,"Paper 1"),(p2s,"Paper 2")]:
        if not sdf.empty:
            for _, row in sdf.iterrows():
                dm = MM.get(str(row.get("Target","")))
                if dm:
                    gantt.append({"Paper":pname,
                                  "Section":str(row.get("Section",""))[:34],
                                  "Start":dm+"-01","Finish":dm+"-28",
                                  "Status":str(row.get("Status","🔒"))})
    if gantt:
        gdf = pd.DataFrame(gantt)
        import plotly.express as px
        fig_g = px.timeline(gdf, x_start="Start", x_end="Finish", y="Section", color="Paper",
                            color_discrete_map={"Paper 1":"#2563eb","Paper 2":"#7c3aed"},
                            height=max(360, len(gantt)*20+80))
        fig_g.update_yaxes(autorange="reversed")
        fig_g.update_layout(**PD, xaxis_title="", yaxis_title="",
                            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1))
        # Add today marker as a shape — using numeric conversion avoided to prevent crash
        # Instead we just show it as an annotation using a shape that spans the full y-axis
        fig_g.add_vrect(
            x0=TODAY.isoformat(), x1=TODAY.isoformat(),
            line_width=2, line_color="#f59e0b", line_dash="dot",
            annotation_text="Today", annotation_position="top right",
            annotation_font_color="#f59e0b", annotation_font_size=10,
        )
        st.plotly_chart(fig_g, use_container_width=True, config={"displayModeBar": False})
    else:
        st.markdown(
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:12px;'
            'padding:24px;text-align:center;font-size:.7rem;color:#4a5568;'
            'font-family:\'Space Mono\',monospace">Research timeline populates from Research Hub sheet</div>',
            unsafe_allow_html=True
        )

    # System doctrine
    st.markdown(
        '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
        'padding:20px 24px;margin-top:16px">'
        '<div style="font-size:.58rem;color:#2563eb;font-weight:700;text-transform:uppercase;'
        'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:10px">⚙️ SYSTEM DOCTRINE</div>'
        '<div style="font-size:.78rem;color:#94b4d4;line-height:1.8">'
        '<b style="color:#e2e8f0">Stability over intensity.</b><br>'
        'This system assumes failure events will occur · Energy fluctuates · Motivation is unreliable.<br>'
        'Therefore: <b style="color:#e2e8f0">Quantification over emotion · Weekly drift detection · Controlled correction.</b><br>'
        'It is not built for perfection. It is built for <b style="color:#2563eb">recovery.</b>'
        '</div></div>',
        unsafe_allow_html=True
    )


# ╔══════════════════════════════════════════════════════════════╗
#  PROFILE  —  Researcher Identity Card
# ╚══════════════════════════════════════════════════════════════╝
with t5:
    week_num = max(1, (TODAY - SEM_START).days // 7 + 1)

    # ── Profile image loader ──────────────────────────────────
    def _b64_img(path: str) -> str:
        try:
            with open(path, "rb") as _f:
                _data = base64.b64encode(_f.read()).decode()
            _ext = path.rsplit(".", 1)[-1].lower()
            _mime = "image/jpeg" if _ext in ("jpg", "jpeg") else f"image/{_ext}"
            return f"data:{_mime};base64,{_data}"
        except Exception:
            return ""

    # Search multiple candidate locations for the profile image
    _img_candidates = [
        _ROOT_DIR / "Tanaka.jpg",
        _APP_DIR / "Tanaka.jpg",
        Path(__file__).parent.parent / "Tanaka.jpg",
        Path.home() / "Desktop" / "Dashboard" / "Tanaka.jpg",
    ]
    _img_src = ""
    for _cand in _img_candidates:
        _src = _b64_img(str(_cand))
        if _src:
            _img_src = _src
            break
    _avatar_inner = (
        f'<div class="profile-avatar" style="position:absolute;top:50%;left:50%;'
        f'transform:translate(-50%,-50%);width:110px;height:110px;border-radius:50%;'
        f'overflow:hidden;box-shadow:0 0 40px rgba(37,99,235,.45)">'
        + (f'<img src="{_img_src}" style="width:100%;height:100%;object-fit:cover;border-radius:50%">'
           if _img_src else
           '<div style="width:100%;height:100%;background:linear-gradient(135deg,#1d4ed8,#7c3aed);'
           'display:flex;align-items:center;justify-content:center;font-family:Syne,sans-serif;'
           'font-size:2.1rem;font-weight:800;color:#fff">TAM</div>')
        + '</div>'
    )

    # ── Row 1: Avatar + Identity ─────────────────────────────
    pa, pb = st.columns([1, 2], gap="large")

    with pa:
        _avatar_section = (
            '<div style="display:flex;flex-direction:column;align-items:center;padding:24px 10px 10px">'
            '<div style="position:relative;width:160px;height:160px;margin:0 auto 20px">'
            '<div class="orbit-ring-4" style="position:absolute;top:-36px;left:-36px;'
            'width:232px;height:232px;border:1px dashed rgba(16,185,129,.15);border-radius:50%"></div>'
            '<div class="orbit-ring-3" style="position:absolute;top:-22px;left:-22px;'
            'width:204px;height:204px;border:1px solid rgba(16,185,129,.25);border-radius:50%;'
            'border-bottom-color:transparent"></div>'
            '<div class="orbit-ring-2" style="position:absolute;top:-10px;left:-10px;'
            'width:180px;height:180px;border:1px solid rgba(124,58,237,.35);border-radius:50%;'
            'border-right-color:transparent"></div>'
            '<div class="orbit-ring-1" style="position:absolute;top:0;left:0;'
            'width:160px;height:160px;border:2px solid rgba(37,99,235,.55);border-radius:50%;'
            'border-top-color:transparent"></div>'
            + _avatar_inner +
            '<div style="position:absolute;top:50%;left:50%;'
            'width:8px;height:8px;border-radius:50%;background:#60a5fa;'
            'box-shadow:0 0 8px #60a5fa;margin:-4px;'
            'animation:orbit-dot 8s linear infinite"></div>'
            '<div style="position:absolute;top:50%;left:50%;'
            'width:5px;height:5px;border-radius:50%;background:#a78bfa;'
            'box-shadow:0 0 6px #a78bfa;margin:-2.5px;'
            'animation:orbit-dot 14s linear infinite reverse"></div>'
            '</div>'
            '<div style="display:inline-flex;align-items:center;gap:7px;'
            'background:#061208;border:1px solid #064e3b;border-radius:20px;padding:5px 14px;margin-bottom:12px">'
            '<div class="status-dot" style="width:7px;height:7px;background:#10b981"></div>'
            '<span style="font-size:.62rem;font-family:\'Space Mono\',monospace;color:#10b981;'
            'font-weight:700;letter-spacing:.12em">ACTIVE RESEARCHER</span>'
            '</div>'
            '<div style="font-size:.62rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
            'text-align:center;margin-bottom:4px">北京航空航天大学</div>'
            '<div style="font-size:.6rem;color:#2563eb;font-family:\'Space Mono\',monospace;'
            'text-align:center;font-weight:700;letter-spacing:.08em">BUAA · BEIJING, CHINA</div>'
            '</div>'
        )
        st.markdown(_avatar_section, unsafe_allow_html=True)

    with pb:
        profile_tags = [
            "Climate–Ecosystem Systems", "Solar-Induced Fluorescence",
            "Earth Observation Science", "Southern Africa · SADC",
            "TROPOMI / Sentinel-5P", "Extreme Event Diagnostics",
            "Google Earth Engine", "Python · R · MATLAB",
            "Early Warning Systems", "Ecosystem Stress Detection",
        ]
        st.markdown(
            '<div class="fade-card" style="padding:10px 0 0">'
            '<div style="font-size:.6rem;color:#7c3aed;font-weight:700;letter-spacing:.2em;'
            'text-transform:uppercase;font-family:\'Space Mono\',monospace;margin-bottom:10px">'
            '⬡  RESEARCHER IDENTITY  ·  SPRING 2026</div>'
            '<div class="orbit-title" style="font-family:\'Syne\',sans-serif;font-size:2.6rem;'
            'font-weight:800;color:#e2e8f0;line-height:1;letter-spacing:.01em;margin-bottom:6px">'
            'TANAKA ALEX<br><span style="color:#60a5fa">MBENDANA</span></div>'
            '<div style="font-size:.78rem;color:#7c3aed;font-family:\'Space Mono\',monospace;'
            'font-weight:700;margin-bottom:2px">Climate–Ecosystem Systems Researcher</div>'
            '<div style="font-size:.7rem;color:#2563eb;font-family:\'Space Mono\',monospace;'
            'font-weight:700;margin-bottom:3px">Earth Observation Scientist</div>'
            '<div style="font-size:.68rem;color:#4a5568;margin-bottom:16px">'
            'MSc Space Technology  ·  School of Astronautics  ·  Beihang University  ·  Cohort 2025–2026</div>'
            '<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:18px">'
            + "".join([
                f'<span style="display:inline-block;padding:4px 11px;border-radius:20px;'
                f'background:rgba(37,99,235,.12);border:1px solid rgba(37,99,235,.25);'
                f'font-size:.58rem;font-family:\'Space Mono\',monospace;color:#60a5fa;'
                f'letter-spacing:.05em">{tag}</span>'
                for tag in profile_tags
            ])
            + '</div>'
            '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px">'
            + "".join([
                f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:10px;'
                f'padding:11px 8px;text-align:center;position:relative;overflow:hidden">'
                f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{c}"></div>'
                f'<div style="font-size:.52rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
                f'text-transform:uppercase;letter-spacing:.08em;margin-bottom:3px">{l}</div>'
                f'<div style="font-size:.95rem;font-family:\'Space Mono\',monospace;font-weight:700;color:{c}">{v}</div>'
                f'</div>'
                for l, v, c in [
                    ("Semester", "Spring '26", "#6366f1"),
                    ("Courses", "6 Active", "#2563eb"),
                    ("Papers", "2 Running", "#7c3aed"),
                    ("Days Left", str(days_left), "#f59e0b"),
                ]
            ])
            + '</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── Research Vision Banner ────────────────────────────────
    st.markdown(
        '<div style="background:linear-gradient(135deg,#07091a,#0a0f24);'
        'border:1px solid #1d4ed8;border-radius:14px;padding:22px 26px;margin-bottom:18px;'
        'position:relative;overflow:hidden">'
        '<div class="shimmer-line" style="position:absolute;top:0;left:0;right:0;height:2px"></div>'
        '<div style="font-size:.6rem;color:#2563eb;font-weight:700;letter-spacing:.18em;'
        'text-transform:uppercase;font-family:\'Space Mono\',monospace;margin-bottom:10px">'
        '🔭  Research Vision</div>'
        '<div style="font-size:.75rem;color:#b8c4d0;line-height:1.9;margin-bottom:16px">'
        'The central motivation of Tanaka\'s work is to improve the scientific understanding '
        'of how <b style="color:#60a5fa">ecosystems respond to climate disturbances</b> and '
        'to develop <b style="color:#60a5fa">observation-based early warning systems</b> '
        'capable of detecting ecological stress before visible degradation occurs — enabling '
        'proactive environmental monitoring and more responsive climate adaptation strategies. '
        'His research sits at the intersection of '
        '<b style="color:#e2e8f0">climate science</b>, '
        '<b style="color:#e2e8f0">ecosystem ecology</b>, and '
        '<b style="color:#e2e8f0">Earth observation</b>.</div>'
        '<div style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px">Three Core Research Questions</div>'
        '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px">'
        + "".join([
            f'<div style="background:rgba(37,99,235,.06);border:1px solid rgba(37,99,235,.18);'
            f'border-radius:10px;padding:14px">'
            f'<div style="font-size:1.2rem;margin-bottom:7px">{icon}</div>'
            f'<div style="font-size:.63rem;color:#94b4d4;line-height:1.65">{q}</div>'
            f'</div>'
            for icon, q in [
                ("🌪️", "How do extreme climate events disrupt vegetation photosynthetic function?"),
                ("🌿", "Why do ecosystems respond differently to the same climatic disturbance?"),
                ("🛰️", "How can satellite observations build early detection systems for ecological stress?"),
            ]
        ])
        + '</div></div>',
        unsafe_allow_html=True
    )

    # ── Row 2: About + Regional Focus + Supervisors ──────────
    rc1, rc2, rc3 = st.columns([2, 2, 1], gap="medium")

    with rc1:
        st.markdown(
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
            'padding:20px;position:relative;overflow:hidden;height:100%">'
            '<div style="position:absolute;top:0;left:0;right:0;height:2px;'
            'background:linear-gradient(90deg,#2563eb,#7c3aed)"></div>'
            '<div style="font-size:.6rem;color:#2563eb;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:12px">'
            '📍 About</div>'
            '<div style="font-size:.72rem;color:#b8c4d0;line-height:1.9;margin-bottom:12px">'
            'Climate–ecosystem systems researcher focused on how '
            '<b style="color:#e2e8f0">climate variability and extreme weather events</b> '
            'alter the functional dynamics of terrestrial ecosystems. Research integrates '
            '<b style="color:#60a5fa">satellite observations</b>, climate diagnostics, and '
            'computational modelling to investigate how vegetation responds to '
            'hydroclimatic stress, disturbance events, and long-term climate change.</div>'
            '<div style="font-size:.72rem;color:#b8c4d0;line-height:1.9">'
            'Currently studying climate-driven ecosystem disruption across '
            '<b style="color:#e2e8f0">Southern Africa</b> using '
            '<b style="color:#60a5fa">solar-induced chlorophyll fluorescence (SIF)</b> '
            'via TROPOMI — a signal directly linked to photosynthetic activity, enabling '
            'detection of functional disruption before structural damage is visible.</div>'
            '<div style="margin-top:14px;padding-top:12px;border-top:1px solid #12192b;'
            'display:flex;gap:14px;flex-wrap:wrap">'
            '<div style="font-size:.6rem;color:#4a5568;font-family:\'Space Mono\',monospace">'
            '📧 tanakambendanata@buaa.edu.cn</div>'
            '<div style="font-size:.6rem;color:#4a5568;font-family:\'Space Mono\',monospace">'
            '🌏 Beijing, China</div>'
            '</div></div>',
            unsafe_allow_html=True
        )

    with rc2:
        cyclones = [
            ("🌀", "Cyclone Idai",    "#ef4444", "2019 · Mozambique, Zimbabwe, Malawi — catastrophic flooding"),
            ("🌀", "Cyclone Kenneth", "#f59e0b", "2019 · Northern Mozambique — rapid intensification"),
            ("🌀", "Cyclone Eloise",  "#ea580c", "2021 · Central Mozambique — coastal ecosystem impact"),
            ("🌀", "Cyclone Freddy",  "#7c3aed", "2023 · Record-duration event · SADC-wide disruption"),
        ]
        rf_html = (
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
            'padding:20px;height:100%">'
            '<div style="font-size:.6rem;color:#10b981;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:6px">'
            '🌍 Regional Focus  ·  Southern Africa</div>'
            '<div style="font-size:.62rem;color:#4a5568;line-height:1.7;margin-bottom:12px">'
            'One of the world\'s most climate-sensitive regions — high interannual rainfall '
            'variability, frequent drought, and increasing tropical cyclone exposure from '
            'the southwest Indian Ocean. Tanaka\'s work examines ecosystem responses '
            'across the <b style="color:#94b4d4">SADC region</b> using TROPOMI SIF.</div>'
            '<div style="font-size:.56rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
            'text-transform:uppercase;letter-spacing:.09em;margin-bottom:8px">'
            'Natural Experiment Case Studies</div>'
        )
        for icon, name, color, detail in cyclones:
            rf_html += (
                f'<div style="display:flex;gap:10px;align-items:flex-start;'
                f'padding:7px 0;border-bottom:1px solid #0d1117">'
                f'<span style="font-size:.9rem;flex:0 0 18px">{icon}</span>'
                f'<div><div style="font-size:.64rem;font-weight:700;color:{color};'
                f'font-family:\'Space Mono\',monospace">{name}</div>'
                f'<div style="font-size:.57rem;color:#4a5568;margin-top:2px;line-height:1.4">{detail}</div>'
                f'</div></div>'
            )
        rf_html += '</div>'
        st.markdown(rf_html, unsafe_allow_html=True)

    with rc3:
        advisors = [
            ("Prof. Zhao Feng", "UAV RS · Image Processing", "#2563eb"),
            ("Dr. A.D Gumbo", "Hydrologist · Sci Writing", "#7c3aed"),
        ]
        phase_label = (
            "Weeks 1–8: Full load" if week_num <= 8 else
            "Weeks 6–9: UAV intensive" if week_num <= 9 else
            "Weeks 10–11: Weekend intensive" if week_num <= 11 else
            "Weeks 9–16: RS Disasters" if week_num <= 16 else "Semester ended"
        )
        adv_html = (
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
            'padding:20px;height:100%">'
            '<div style="font-size:.6rem;color:#10b981;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:12px">'
            '🎓 Supervisors</div>'
        )
        for name, role, color in advisors:
            adv_html += (
                f'<div style="background:#07080f;border:1px solid #12192b;border-radius:8px;'
                f'padding:10px;margin-bottom:8px">'
                f'<div style="font-size:.64rem;font-weight:700;color:{color};'
                f'font-family:\'Space Mono\',monospace">{name}</div>'
                f'<div style="font-size:.57rem;color:#4a5568;margin-top:3px">{role}</div>'
                f'</div>'
            )
        adv_html += (
            '<div style="margin-top:10px;padding-top:10px;border-top:1px solid #12192b">'
            f'<div style="font-size:.56rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
            f'text-transform:uppercase;letter-spacing:.08em;margin-bottom:4px">Current Phase</div>'
            f'<div style="font-size:.62rem;color:#f59e0b;font-family:\'Space Mono\',monospace">'
            f'Week {week_num}  ·  {phase_label}</div>'
            '</div></div>'
        )
        st.markdown(adv_html, unsafe_allow_html=True)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Row 3: Research Framework — Three Pillars ─────────────
    st.markdown(
        '<div style="font-size:.6rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'text-transform:uppercase;letter-spacing:.12em;margin-bottom:10px">'
        '🔬  Research Framework  ·  Three Pillars</div>',
        unsafe_allow_html=True
    )
    fp1, fp2, fp3 = st.columns(3, gap="medium")
    for col, icon, title, color, points in [
        (fp1, "🌪️", "Climate Disturbance &amp; Ecosystem Function", "#2563eb", [
            "Extreme weather impacts on terrestrial vegetation",
            "Tropical cyclone effects across Southern Africa",
            "SIF as proxy for photosynthetic disruption",
            "Contrasting outcomes: structural damage vs moisture benefit",
            "TROPOMI satellite as primary observation instrument",
        ]),
        (fp2, "📊", "Ecosystem Stress Diagnostics", "#7c3aed", [
            "Metrics tracking photosynthetic deviation during disturbance",
            "Immediate disruption · persistence · recovery trajectory",
            "Functional classification: suppression, delayed recovery, enhancement",
            "Mechanistic understanding of disturbance-response pathways",
            "Statistical modelling of disturbance severity and ecosystem lag",
        ]),
        (fp3, "🚨", "Early Detection of Ecological Stress", "#10b981", [
            "Satellite early warning before visible degradation occurs",
            "Climate indicators: rainfall anomalies, drought indices",
            "Extreme event diagnostics at regional scale (SADC)",
            "Decision-support tools for climate adaptation planning",
            "Bridging Earth observation science and resilience policy",
        ]),
    ]:
        with col:
            pts_html = "".join([
                f'<div style="display:flex;gap:7px;padding:5px 0;'
                f'border-bottom:1px solid #0d1117;align-items:flex-start">'
                f'<span style="color:{color};opacity:.7;flex:0 0 10px;margin-top:2px;font-size:.65rem">›</span>'
                f'<span style="font-size:.62rem;color:#94b4d4;line-height:1.55">{p}</span></div>'
                for p in points
            ])
            st.markdown(
                f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
                f'padding:18px;height:100%;position:relative;overflow:hidden">'
                f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{color}"></div>'
                f'<div style="font-size:1.4rem;margin-bottom:8px">{icon}</div>'
                f'<div style="font-size:.67rem;font-weight:700;color:#e2e8f0;'
                f'font-family:\'Space Mono\',monospace;margin-bottom:10px;line-height:1.35">{title}</div>'
                f'{pts_html}'
                f'</div>',
                unsafe_allow_html=True
            )

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Row 4: Computational Stack + Data Instruments ─────────
    cs1, cs2 = st.columns(2, gap="medium")

    with cs1:
        stack = [
            ("Python",              "Primary — data pipelines, remote sensing, ML",   "#3b82f6", 92),
            ("R",                   "Statistical modelling, ecological time series",   "#10b981", 78),
            ("Google Earth Engine", "Cloud-based geospatial analysis at scale",        "#ea580c", 80),
            ("ArcGIS / QGIS",       "Geospatial processing and cartography",           "#7c3aed", 75),
            ("MATLAB",              "Signal processing, numerical analysis",            "#f59e0b", 65),
        ]
        st_html = (
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;padding:20px">'
            '<div style="font-size:.6rem;color:#ea580c;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:14px">'
            '⚡ Computational Stack</div>'
        )
        for tool, desc, color, pct in stack:
            st_html += (
                f'<div style="margin-bottom:11px">'
                f'<div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:3px">'
                f'<div><span style="font-size:.68rem;color:#e2e8f0;font-weight:700">{tool}</span>'
                f'<span style="font-size:.57rem;color:#4a5568;margin-left:8px">{desc}</span></div>'
                f'<span style="font-size:.6rem;color:{color};font-family:\'Space Mono\',monospace">{pct}%</span>'
                f'</div>'
                f'<div style="background:#12192b;border-radius:3px;height:4px;overflow:hidden">'
                f'<div style="width:{pct}%;height:100%;background:linear-gradient(90deg,{color},{color}99);'
                f'border-radius:3px"></div></div></div>'
            )
        st_html += '</div>'
        st.markdown(st_html, unsafe_allow_html=True)

    with cs2:
        instruments = [
            ("🛰️", "TROPOMI / Sentinel-5P",         "#2563eb", "Primary SIF source — photosynthetic activity at global scale"),
            ("🌱", "Solar-Induced Fluorescence (SIF)","#10b981", "Functional vegetation signal — physiological stress proxy"),
            ("🌦️", "Climate Reanalysis (ERA5 / CHIRPS)","#7c3aed","Hydroclimatic forcing data — rainfall, temperature, drought"),
            ("🌀", "IBTrACS Cyclone Tracks",          "#ea580c", "Historical tropical cyclone paths and intensity records"),
            ("📈", "Ecological Time Series",           "#f59e0b", "Multi-year vegetation dynamics before/during/after events"),
        ]
        ins_html = (
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;'
            'padding:20px;height:100%">'
            '<div style="font-size:.6rem;color:#2563eb;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:14px">'
            '📡 Data Sources &amp; Instruments</div>'
        )
        for icon, name, color, desc in instruments:
            ins_html += (
                f'<div style="display:flex;gap:10px;padding:8px 0;'
                f'border-bottom:1px solid #0d1117;align-items:flex-start">'
                f'<div style="font-size:1.1rem;flex:0 0 22px">{icon}</div>'
                f'<div><div style="font-size:.65rem;font-weight:700;color:{color};'
                f'font-family:\'Space Mono\',monospace">{name}</div>'
                f'<div style="font-size:.57rem;color:#4a5568;margin-top:2px;line-height:1.5">{desc}</div>'
                f'</div></div>'
            )
        ins_html += '</div>'
        st.markdown(ins_html, unsafe_allow_html=True)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Row 5: Course Load Visual ────────────────────────────
    st.markdown(
        '<div style="font-size:.6rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
        'text-transform:uppercase;letter-spacing:.12em;margin-bottom:10px">'
        '📚 Semester Course Load  ·  Spring 2026</div>',
        unsafe_allow_html=True
    )
    COURSE_DETAILS = [
        ("Chinese Language 2",   "D253026002", "Mon + Wed  S1–2  Weeks 1–16",          "Teaching Bldg 2, Rm 3003",   32, "#2563eb", "Zhang Yongqin"),
        ("Sci Paper Writing",    "D253026011", "Mon  S3–4  Weeks 1–8",                 "Research Bldg 1, Rm 4108",   8,  "#7c3aed", "Sheikh Tawhidul Islam"),
        ("RS Image Processing",  "D253051004", "Wed + Fri  S3–4  Weeks 1–8",           "Research Bldg 1, Rm 1043",   16, "#10b981", "Tan Yumin · Tariq Aqil"),
        ("UAV Remote Sensing",   "D253052002", "Mon/Thu Wks 6–9 · Sat/Sun Wks 10–11", "Research Bldg 1, Rm 1043",   24, "#ea580c", "Tan Yumin · He Lingfeng"),
        ("RS Natural Disasters", "D253051005", "Thu S8–10 Wks 9–16 + Fri 15–16",      "Research Bldg 1, Rm 1045",   18, "#f59e0b", "Sheikh Tawhidul Islam"),
        ("AI & Large Models",    "D253041002", "Mon + Tue  S11–12  Weeks 1–8",         "Research Bldg 1, Rm 4108",   16, "#6366f1", "Populus euphratica"),
    ]
    cc = st.columns(3, gap="medium")
    for idx, (name, code, schedule, room, total_m, color, instructor) in enumerate(COURSE_DETAILS):
        with cc[idx % 3]:
            st.markdown(
                f'<div class="fade-card" style="background:#0c1020;border:1px solid #12192b;'
                f'border-radius:12px;padding:14px;margin-bottom:10px;position:relative;overflow:hidden">'
                f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{color}"></div>'
                f'<div style="font-size:.56rem;color:{color};font-family:\'Space Mono\',monospace;'
                f'text-transform:uppercase;letter-spacing:.1em;margin-bottom:4px">{code}</div>'
                f'<div style="font-size:.72rem;font-weight:700;color:#e2e8f0;margin-bottom:6px;line-height:1.3">{name}</div>'
                f'<div style="font-size:.6rem;color:#4a5568;margin-bottom:2px">📅 {schedule}</div>'
                f'<div style="font-size:.6rem;color:#4a5568;margin-bottom:2px">📍 {room}</div>'
                f'<div style="font-size:.6rem;color:#4a5568;margin-bottom:8px">👤 {instructor}</div>'
                f'<div style="display:flex;justify-content:space-between;align-items:center">'
                f'<div style="font-size:.58rem;color:{color};font-family:\'Space Mono\',monospace">'
                f'{total_m} meetings</div>'
                f'<div style="background:#12192b;border-radius:3px;height:4px;width:60%;overflow:hidden">'
                f'<div style="width:0%;height:100%;background:{color};border-radius:3px"></div></div>'
                f'</div></div>',
                unsafe_allow_html=True
            )

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Row 6: Goals + Academic Roadmap ─────────────────────
    gl, gr = st.columns([3, 2], gap="medium")

    with gl:
        goals = [
            ("✅", "Complete MSc Space Technology — graduate June 2026",               "#10b981", True),
            ("⏳", "Submit Paper 1 (SIF + cyclone impacts) to RSE / Remote Sensing",   "#f59e0b", False),
            ("⏳", "Begin Paper 2 — Early Detection of Ecosystem Stress methodology",  "#f59e0b", False),
            ("🎯", "Develop reproducible climate–ecosystem analysis pipeline",          "#2563eb", False),
            ("🎯", "Maintain ≥85% attendance across all 6 courses",                    "#2563eb", False),
            ("🔒", "PhD in climate–ecosystem interactions — 2028 target",              "#4a5568", False),
            ("🔒", "Publish early warning system paper at ISPRS / RSE level",          "#4a5568", False),
            ("🔒", "Build satellite monitoring tool for SADC regional scale",          "#4a5568", False),
        ]
        goal_html = (
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;padding:20px">'
            '<div style="font-size:.6rem;color:#2563eb;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:14px">'
            '🎯 Goals &amp; Objectives</div>'
        )
        for icon, text, color, done in goals:
            opac = "1" if color != "#4a5568" else "0.5"
            tc   = "#10b981" if done else "#b8c4d0"
            td   = "line-through" if done else "none"
            goal_html += (
                f'<div style="display:flex;align-items:flex-start;gap:10px;padding:8px 0;'
                f'border-bottom:1px solid #0d1117;opacity:{opac}">'
                f'<span style="font-size:1rem;flex:0 0 20px">{icon}</span>'
                f'<span style="font-size:.68rem;color:{tc};text-decoration:{td};line-height:1.5">{text}</span>'
                f'</div>'
            )
        goal_html += '</div>'
        st.markdown(goal_html, unsafe_allow_html=True)

    with gr:
        MILESTONES_P = [
            ("2026", "🔬 Stabilisation", [
                "MSc completion — June 2026",
                "Paper 1 submission — SIF & cyclone impacts (RSE)",
                "O.R.B.I.T discipline & 85%+ attendance maintained",
                "Reproducible research pipeline established",
            ], "#2563eb", True),
            ("2027", "📡 Consolidation", [
                "Paper 2 — Early Detection system methodology",
                "Conference positioning (ISPRS / AGU)",
                "Network building — climate-ecosystem community",
                "PhD application groundwork",
            ], "#7c3aed", False),
            ("2028", "🎓 Doctoral Arc", [
                "PhD start — climate-ecosystem interactions",
                "Carbon cycle + SIF signal integration",
                "Ecological tipping point detection methods",
                "International collaboration — SADC monitoring",
            ], "#10b981", False),
        ]
        for year, phase, items, color, is_cur in MILESTONES_P:
            bg = "linear-gradient(145deg,#0c1020,#0d1829)" if is_cur else "#0c1020"
            border = f"2px solid {color}" if is_cur else "1px solid #12192b"
            st.markdown(
                f'<div style="background:{bg};border:{border};border-radius:12px;'
                f'padding:14px 16px;margin-bottom:10px;position:relative;overflow:hidden">'
                f'<div style="position:absolute;top:0;left:0;right:0;height:2px;background:{color};opacity:.7"></div>'
                f'<div style="display:flex;align-items:baseline;gap:8px;margin-bottom:8px">'
                f'<span style="font-family:\'Syne\',sans-serif;font-size:1.6rem;font-weight:800;color:{color}">{year}</span>'
                f'<span style="font-size:.7rem;font-weight:700;color:#e2e8f0">{phase}</span>'
                + (f'<span style="margin-left:auto;font-size:.55rem;background:{color}22;border:1px solid {color}44;'
                   f'border-radius:10px;padding:2px 8px;color:{color};font-family:\'Space Mono\',monospace">'
                   f'ACTIVE</span>' if is_cur else '')
                + '</div>'
                + "".join([
                    f'<div style="font-size:.62rem;color:#4a5568;padding:2px 0;display:flex;gap:6px">'
                    f'<span style="color:{color};opacity:.6">›</span><span>{item}</span></div>'
                    for item in items
                ])
                + '</div>',
                unsafe_allow_html=True
            )

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Publications + Achievements ──────────────────────────
    pub_col, ach_col = st.columns([3, 2], gap="medium")

    with pub_col:
        PUBLICATIONS = [
            {
                "title": "Application of change detection techniques driven by expert opinions "
                         "for small-area studies in developing countries",
                "authors": "Mbendana T.A., Gumbo A.D., Jombo S., Mugari E., Kapangaziwiri E.",
                "journal": "Scientific African",
                "year": "2025",
                "volume": "Vol. 27, e02594",
                "status": "Published",
                "doi": "10.1016/j.sciaf.2025.e02594",
                "tags": ["Change Detection", "Expert Opinion", "Developing Countries", "Small-Area"],
                "color": "#10b981",
            },
            {
                "title": "Spatially Aggregated Cyclone Vegetation Impacts: "
                         "A SIF-Based Diagnostic Framework for Southern Africa",
                "authors": "Mbendana T.A., et al.",
                "journal": "Remote Sensing of Environment  ·  Target Journal",
                "year": "2026",
                "volume": "",
                "status": "In Preparation",
                "doi": None,
                "tags": ["SIF", "TROPOMI", "Cyclones", "Southern Africa"],
                "color": "#f59e0b",
            },
        ]
        pub_html = (
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;padding:20px">'
            '<div style="font-size:.6rem;color:#7c3aed;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:14px">'
            '📄 Publications &amp; Research Output</div>'
        )
        for i, p in enumerate(PUBLICATIONS):
            doi_block = (
                f'<a href="https://doi.org/{p["doi"]}" target="_blank" '
                f'style="font-size:.58rem;color:#2563eb;font-family:\'Space Mono\',monospace;'
                f'text-decoration:none">DOI: {p["doi"]}</a>'
            ) if p["doi"] else (
                '<span style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
                'font-style:italic">DOI pending</span>'
            )
            tags_html = "".join([
                f'<span style="padding:2px 8px;border-radius:10px;background:rgba(124,58,237,.12);'
                f'border:1px solid rgba(124,58,237,.25);font-size:.55rem;color:#a78bfa;'
                f'font-family:\'Space Mono\',monospace;margin-right:4px">{t}</span>'
                for t in p["tags"]
            ])
            pub_html += (
                f'<div style="padding:14px 0;border-bottom:1px solid #12192b;'
                f'{"padding-top:0" if i==0 else ""}">'
                f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">'
                f'<span style="font-size:.55rem;font-family:\'Space Mono\',monospace;'
                f'padding:2px 9px;border-radius:10px;'
                f'background:{p["color"]}22;border:1px solid {p["color"]}55;color:{p["color"]}">'
                f'{p["status"]}</span>'
                f'<span style="font-size:.58rem;color:#4a5568;font-family:\'Space Mono\',monospace">'
                f'{p["year"]}</span></div>'
                f'<div style="font-size:.7rem;font-weight:700;color:#e2e8f0;line-height:1.4;margin-bottom:4px">'
                f'{p["title"]}</div>'
                f'<div style="font-size:.6rem;color:#4a5568;margin-bottom:4px">'
                f'{p["authors"]} · <i>{p["journal"]}</i>'
                f'{(" · " + p["volume"]) if p.get("volume") else ""}</div>'
                f'<div style="margin-bottom:8px">{doi_block}</div>'
                f'<div style="display:flex;flex-wrap:wrap;gap:4px">{tags_html}</div>'
                f'</div>'
            )
        pub_html += (
            '<div style="margin-top:12px;padding:10px 12px;background:#07080f;border-radius:8px;'
            'border:1px dashed #1e3a5f">'
            '<div style="font-size:.58rem;color:#2563eb;font-family:\'Space Mono\',monospace;'
            'margin-bottom:4px">🔗 Google Scholar</div>'
            '<a href="https://scholar.google.com/citations?user=slfDEvoAAAAJ&hl=en" target="_blank" '
            'style="font-size:.6rem;color:#60a5fa;text-decoration:none;font-family:\'Space Mono\',monospace">'
            'scholar.google.com/citations?user=slfDEvoAAAAJ</a>'
            '</div></div>'
        )
        st.markdown(pub_html, unsafe_allow_html=True)

    with ach_col:
        ACHIEVEMENTS = [
            ("🏅", "Beihang University Scholarship",
             "Full MSc scholarship — School of Astronautics · Beihang University 2025–2026", "#f59e0b"),
            ("🎓", "MSc Space Technology",
             "School of Astronautics · Beihang University · Expected June 2026", "#2563eb"),
            ("🌍", "SADC Regional Researcher",
             "Active researcher on climate–ecosystem dynamics across "
             "Mozambique, Zimbabwe, Malawi & regional Southern Africa",  "#10b981"),
            ("📡", "TROPOMI / Sentinel-5P Analyst",
             "Applied SIF remote sensing for photosynthetic stress "
             "detection — one of few researchers applying this in SADC",  "#7c3aed"),
            ("🌀", "Multi-Cyclone Attribution Study",
             "Idai · Kenneth · Eloise · Freddy — comparative cyclone "
             "vegetation impact analysis using spatially aggregated SIF", "#ea580c"),
            ("⚙️", "Open-Source Research Pipeline",
             "Reproducible Python pipeline: GEE + TROPOMI + IBTrACS "
             "— end-to-end cyclone ecosystem impact analysis",            "#6366f1"),
            ("🤖", "O.R.B.I.T System",
             "Built personal academic operating system for discipline "
             "tracking, finance management & research monitoring",        "#2563eb"),
        ]
        ach_html = (
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;padding:20px">'
            '<div style="font-size:.6rem;color:#f59e0b;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:14px">'
            '🏆 Achievements &amp; Highlights</div>'
        )
        for icon, title, desc, color in ACHIEVEMENTS:
            ach_html += (
                f'<div style="display:flex;gap:10px;padding:9px 0;border-bottom:1px solid #0d1117;'
                f'align-items:flex-start">'
                f'<div style="font-size:1.05rem;flex:0 0 22px;margin-top:1px">{icon}</div>'
                f'<div><div style="font-size:.65rem;font-weight:700;color:{color};'
                f'font-family:\'Space Mono\',monospace;margin-bottom:2px">{title}</div>'
                f'<div style="font-size:.58rem;color:#4a5568;line-height:1.5">{desc}</div>'
                f'</div></div>'
            )
        ach_html += '</div>'
        st.markdown(ach_html, unsafe_allow_html=True)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ── Row 7: Skills + Mantra ───────────────────────────────
    sk1, sk2 = st.columns([2, 3], gap="medium")

    with sk1:
        skills = [
            ("Remote Sensing & SIF Analysis",   90, "#2563eb"),
            ("Python / Data Science",            85, "#7c3aed"),
            ("Google Earth Engine",              80, "#10b981"),
            ("GIS & Spatial Analysis",           78, "#6366f1"),
            ("R / Statistical Modelling",        75, "#ea580c"),
            ("Scientific Writing",               72, "#f59e0b"),
            ("UAV Operations",                   68, "#10b981"),
            ("MATLAB / Signal Processing",       65, "#2563eb"),
            ("Chinese Language",                 40, "#4a5568"),
        ]
        sk_html = (
            '<div style="background:#0c1020;border:1px solid #12192b;border-radius:14px;padding:20px">'
            '<div style="font-size:.6rem;color:#ea580c;font-weight:700;text-transform:uppercase;'
            'letter-spacing:.14em;font-family:\'Space Mono\',monospace;margin-bottom:14px">'
            '⚡ Skills &amp; Expertise</div>'
        )
        for skill, pct, color in skills:
            sk_html += (
                f'<div style="margin-bottom:9px">'
                f'<div style="display:flex;justify-content:space-between;margin-bottom:3px">'
                f'<span style="font-size:.63rem;color:#94b4d4">{skill}</span>'
                f'<span style="font-size:.6rem;color:{color};font-family:\'Space Mono\',monospace">{pct}%</span>'
                f'</div>'
                f'<div style="background:#12192b;border-radius:3px;height:4px;overflow:hidden">'
                f'<div style="width:{pct}%;height:100%;background:linear-gradient(90deg,{color},{color}99);'
                f'border-radius:3px"></div></div></div>'
            )
        sk_html += '</div>'
        st.markdown(sk_html, unsafe_allow_html=True)

    with sk2:
        st.markdown(
            '<div style="background:linear-gradient(145deg,#07091a,#0a0f24);border:1px solid #1e3a5f;'
            'border-radius:14px;padding:28px;position:relative;overflow:hidden;height:100%">'
            '<div class="shimmer-line" style="position:absolute;top:0;left:0;right:0;height:2px"></div>'
            '<div style="position:absolute;right:-40px;bottom:-40px;width:180px;height:180px;'
            'border-radius:50%;background:radial-gradient(circle,rgba(37,99,235,.06),transparent 70%);'
            'pointer-events:none"></div>'
            '<div style="font-size:.6rem;color:#2563eb;font-weight:700;letter-spacing:.2em;'
            'text-transform:uppercase;font-family:\'Space Mono\',monospace;margin-bottom:20px">'
            '⬡  SYSTEM DOCTRINE  ·  PERSONAL MANTRA</div>'
            '<div style="font-family:\'Syne\',sans-serif;font-size:1.45rem;font-weight:800;'
            'color:#e2e8f0;line-height:1.3;margin-bottom:16px">'
            '"Detect the signal<br>before the damage.<br>'
            '<span style="color:#2563eb">Act before the visible."</span></div>'
            '<div style="font-size:.72rem;color:#4a5568;line-height:1.9;margin-bottom:16px">'
            'This system — O.R.B.I.T — exists to track drift before it becomes failure. '
            'My research does the same: detect ecosystem stress '
            '<b style="color:#60a5fa">before visible degradation</b> occurs. '
            'Early signals. Early action. Every dataset processed, every paper submitted '
            'is a step toward <b style="color:#60a5fa">climate resilience</b> that works in the field.</div>'
            '<div style="font-size:.62rem;color:#4a5568;line-height:1.8;font-style:italic;'
            'border-left:2px solid #1d4ed8;padding-left:12px;margin-bottom:20px">'
            '"Across many climate-sensitive regions, ecosystem monitoring remains reactive. '
            'Environmental degradation is detected only after structural damage has occurred. '
            'My work changes that."</div>'
            f'<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:8px">'
            + "".join([
                f'<div style="background:#0c1020;border:1px solid #12192b;border-radius:8px;'
                f'padding:10px;text-align:center">'
                f'<div style="font-size:.55rem;color:#4a5568;font-family:\'Space Mono\',monospace;'
                f'text-transform:uppercase;letter-spacing:.08em">{l}</div>'
                f'<div style="font-size:.82rem;font-family:\'Space Mono\',monospace;font-weight:700;'
                f'color:{c};margin-top:3px">{v}</div></div>'
                for l, v, c in [
                    ("Semester Start", SEM_START.strftime("%d %b %Y"), "#94b4d4"),
                    ("Semester End",   SEM_END.strftime("%d %b %Y"),   "#94b4d4"),
                    ("Days Remaining", str(days_left),                  "#f59e0b"),
                    ("Weeks Elapsed",  f"Wk {week_num}",               "#10b981"),
                ]
            ])
            + '</div></div>',
            unsafe_allow_html=True
        )


# ── TAB 6: MOBILE EXPENSE ENTRY ─────────────────────────────────
with t6:
    st.html(f"""<style>
    .exp-header{{font-size:.62rem;color:{_TXT3};font-family:'Space Mono',monospace;
      text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px}}
    .exp-total{{font-size:1.6rem;font-family:'Space Mono',monospace;font-weight:700;color:#ef4444}}
    .exp-card{{background:{_BG2};border:1px solid {_BRD};border-radius:10px;
      padding:12px 14px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center}}
    .exp-cat{{font-size:.6rem;color:{_TXT3};text-transform:uppercase;letter-spacing:.08em}}
    .exp-desc{{font-size:.85rem;color:{_TXT2};margin:2px 0}}
    .exp-date{{font-size:.6rem;color:{_TXT3}}}
    .exp-amt{{font-size:1rem;font-family:'Space Mono',monospace;font-weight:700;color:#ef4444}}
    </style>""")

    _exp_df = _load_expenses()
    _exp_total = float(_exp_df["amount"].sum()) if not _exp_df.empty else 0.0

    # ── KPI row ──
    ek1, ek2, ek3 = st.columns(3, gap="small")
    with ek1:
        st.html(f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:10px;padding:14px 12px">'
                f'<div class="exp-header">📝 Logged</div>'
                f'<div style="font-size:1.4rem;font-family:\'Space Mono\',monospace;font-weight:700;color:{_TXT2}">'
                f'{len(_exp_df)}</div></div>')
    with ek2:
        st.html(f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:10px;padding:14px 12px">'
                f'<div class="exp-header">💸 Total Spent</div>'
                f'<div class="exp-total">¥{_exp_total:,.0f}</div></div>')
    with ek3:
        _this_month = TODAY.strftime("%Y-%m")
        _mo_total = float(_exp_df[_exp_df["date"].str.startswith(_this_month)]["amount"].sum()) \
                    if not _exp_df.empty else 0.0
        st.html(f'<div style="background:{_BG2};border:1px solid {_BRD};border-radius:10px;padding:14px 12px">'
                f'<div class="exp-header">📅 This Month</div>'
                f'<div class="exp-total" style="color:#f59e0b">¥{_mo_total:,.0f}</div></div>')

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ── Entry form ──
    with st.expander("➕ Add New Expense", expanded=True):
        ef1, ef2 = st.columns([1, 1], gap="small")
        with ef1:
            _new_cat = st.selectbox("Category", _EXP_CATS, key="exp_cat")
            _new_amt = st.number_input("Amount (¥)", min_value=0.0, step=1.0,
                                       format="%.0f", key="exp_amt")
        with ef2:
            _new_desc = st.text_input("Description", placeholder="e.g. Lunch at canteen", key="exp_desc")
            _new_date = st.date_input("Date", value=TODAY, key="exp_date")
        if st.button("💾  Log Expense", use_container_width=True, key="exp_submit"):
            if _new_amt > 0:
                _append_expense(str(_new_date), _new_cat, _new_desc.strip(), float(_new_amt))
                st.success(f"Logged ¥{_new_amt:,.0f} — {_new_cat}")
                st.rerun()
            else:
                st.warning("Enter an amount greater than 0.")

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    # ── Expense log list ──
    if not _exp_df.empty:
        _exp_sorted = _exp_df.copy()
        _exp_sorted["_idx"] = _exp_sorted.index
        _exp_sorted = _exp_sorted.sort_values("date", ascending=False)

        st.html(f'<div class="exp-header" style="margin-top:8px">📜 Recent Expenses</div>')

        # Category breakdown chart
        if len(_exp_df) >= 2:
            _cat_grp = _exp_df.groupby("category")["amount"].sum().reset_index()
            _fig_exp = go.Figure(go.Pie(
                labels=_cat_grp["category"], values=_cat_grp["amount"],
                hole=0.55, textinfo="label+percent",
                marker_colors=["#2563eb","#7c3aed","#10b981","#f59e0b",
                               "#ef4444","#06b6d4","#ec4899","#84cc16","#94b4d4"],
                textfont_size=10,
                hovertemplate="%{label}: ¥%{value:,.0f}<extra></extra>",
            ))
            _fig_exp.update_layout(**PD, height=220,
                title=dict(text="Spending by Category", font_color=_TXT3, font_size=11),
                showlegend=False, margin=dict(l=10,r=10,t=30,b=10))
            st.plotly_chart(_fig_exp, use_container_width=True, config={"displayModeBar": False})

        # Row-by-row cards with delete
        for _, row in _exp_sorted.iterrows():
            orig_idx = int(row["_idx"])
            _dc1, _dc2 = st.columns([5, 1], gap="small")
            with _dc1:
                st.html(
                    f'<div class="exp-card">'
                    f'<div><div class="exp-cat">{row["category"]}</div>'
                    f'<div class="exp-desc">{row["description"] or "—"}</div>'
                    f'<div class="exp-date">{row["date"]}</div></div>'
                    f'<div class="exp-amt">¥{float(row["amount"]):,.0f}</div>'
                    f'</div>'
                )
            with _dc2:
                if st.button("🗑", key=f"del_exp_{orig_idx}", help="Delete this entry"):
                    _delete_expense(orig_idx)
                    st.rerun()
    else:
        st.html(f'<div style="text-align:center;padding:40px;color:{_TXT3};font-size:.8rem">'
                f'No expenses logged yet. Add your first one above.</div>')


# ── FOOTER ───────────────────────────────────────────────────────
st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
st.markdown(
    '<div style="text-align:center;padding:14px;border-top:1px solid #12192b">'
    '<span style="font-size:.58rem;color:#1e3a5f;font-family:\'Space Mono\',monospace;letter-spacing:.1em">'
    'O.R.B.I.T.  ·  TANAKA ALEX MBENDANA  ·  BEIHANG UNIVERSITY 2025–2026  ·  CONSISTENT CONTROLLED ASCENT'
    '</span></div>',
    unsafe_allow_html=True
)