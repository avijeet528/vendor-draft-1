# ============================================================
#  app.py — IT Procurement Service Dashboard (Streamlit)
#  Upload to GitHub + deploy on share.streamlit.io for free
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import defaultdict
import re
import os

# ── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="IT Procurement Dashboard",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* KPI cards */
    .kpi-box {
        border-radius: 12px;
        padding: 18px 10px;
        text-align: center;
        color: white;
    }
    .kpi-value {
        font-size: 2.2em;
        font-weight: 800;
        margin: 0;
        line-height: 1.1;
    }
    .kpi-label {
        font-size: 0.85em;
        opacity: 0.88;
        margin-top: 5px;
    }

    /* Expander header */
    div[data-testid="stExpander"] details summary p {
        font-size: 1em;
        font-weight: 600;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #2c3e50;
    }
    section[data-testid="stSidebar"] * {
        color: white !important;
    }
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stTextInput label,
    section[data-testid="stSidebar"] .stMultiSelect label {
        color: #ecf0f1 !important;
        font-weight: 600;
    }
    section[data-testid="stSidebar"] .stMarkdown p {
        color: #bdc3c7 !important;
    }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# CONSTANTS
# ════════════════════════════════════════════════════════════
COLORS = [
    "#2980b9","#27ae60","#8e44ad","#e67e22","#e74c3c",
    "#16a085","#2c3e50","#f39c12","#1abc9c","#d35400",
    "#7f8c8d","#c0392b","#3498db","#2ecc71","#9b59b6",
    "#e91e63","#00bcd4","#ff5722","#795548","#607d8b",
]
def get_color(i):
    return COLORS[i % len(COLORS)]

# ════════════════════════════════════════════════════════════
# DATA LOADING — cached so it only runs once
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    FILE_PATH = "Master Catalog.xlsx"

    if not os.path.exists(FILE_PATH):
        st.error(
            f"❌ File not found: '{FILE_PATH}'. "
            "Make sure Master Catalog.xlsx is in the same folder as app.py"
        )
        return None, None

    # ── Read raw to find header row ──────────────────────────
    raw = pd.read_excel(FILE_PATH, engine="openpyxl", header=None)

    header_row = None
    for i, row in raw.iterrows():
        row_vals = [
            str(v).strip().lower()
            for v in row.values if pd.notna(v)
        ]
        if (any("category" in v for v in row_vals) and
                any("file" in v for v in row_vals)):
            header_row = i
            break

    if header_row is None:
        st.error("❌ Could not detect header row in Excel file.")
        return None, None

    # ── Re-read with correct header ──────────────────────────
    df = pd.read_excel(FILE_PATH, engine="openpyxl", header=header_row)

    # Drop NaN column names & convert all to string
    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)

    # ── Map columns ──────────────────────────────────────────
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category":
            col_map["Category"]     = c
        elif "vendor" in cl or "type" in cl:
            col_map["Vendor"]       = c
        elif cl == "file name":
            col_map["File Name"]    = c
        elif cl == "file link":
            col_map["File Link"]    = c
        elif cl == "file url":
            col_map["File URL"]     = c
        elif "comment" in cl:
            col_map["Comments"]     = c
        elif "quoted" in cl or "price" in cl:
            col_map["Quoted Price"] = c

    df.rename(columns={v: k for k, v in col_map.items()}, inplace=True)

    # Keep available standard columns
    keep = ["Category", "Vendor", "File Name", "Comments"]
    for extra in ["File Link", "File URL", "Quoted Price"]:
        if extra in df.columns:
            keep.append(extra)
    df = df[[c for c in keep if c in df.columns]].copy()

    # Drop rows where both Category AND Vendor are empty
    df = df[
        ~(
            df["Category"].astype(str).str.strip().isin(["", "nan"]) &
            df["Vendor"].astype(str).str.strip().isin(["", "nan"])
        )
    ].copy()

    df.fillna("", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # ── Parse services ───────────────────────────────────────
    def parse_services(raw_val):
        if not raw_val or str(raw_val).strip() == "":
            return ["(unspecified)"]
        parts = [s.strip() for s in str(raw_val).split("\n") if s.strip()]
        return parts if parts else ["(unspecified)"]

    df["Services List"] = df["Comments"].apply(parse_services)

    # ── Explode one row per service ──────────────────────────
    df_exp = df.explode("Services List").copy()
    df_exp.rename(columns={"Services List": "Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].str.strip()
    df_exp = df_exp[
        (df_exp["Service"] != "") &
        (df_exp["Service"] != "(unspecified)")
    ].reset_index(drop=True)

    return df, df_exp


# ════════════════════════════════════════════════════════════
# LOAD DATA
# ════════════════════════════════════════════════════════════
df_master, df_exploded = load_data()

if df_master is None or df_exploded is None:
    st.stop()

# Build vendor → color map once
vendor_color_map = {
    v: get_color(i)
    for i, v in enumerate(sorted(df_master["Vendor"].unique()))
}

# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
with st.sidebar:

    # Logo / Title
    st.markdown("""
    <div style='text-align:center;padding:20px 0 14px'>
        <div style='font-size:2.5em'>📋</div>
        <div style='font-size:1.15em;font-weight:800;
                    color:white;margin:6px 0 2px'>
            IT Procurement
        </div>
        <div style='font-size:0.78em;color:#95a5a6'>
            Service & Vendor Dashboard
        </div>
    </div>
    <hr style='border-color:#3d5166;margin:0 0 16px'>
    """, unsafe_allow_html=True)

    # ── FILTER 1: Category ────────────────────────────────────
    st.markdown("**📂 Category**")
    all_cats = ["All"] + sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["", "nan"]
    ])
    selected_cat = st.selectbox(
        "Category", all_cats, label_visibility="collapsed"
    )

    # ── FILTER 2: Vendor (scoped to category) ─────────────────
    st.markdown("**🏢 Vendor**")
    if selected_cat == "All":
        vendor_pool = df_master
    else:
        vendor_pool = df_master[df_master["Category"] == selected_cat]

    all_vendors = ["All"] + sorted([
        v for v in vendor_pool["Vendor"].unique()
        if str(v).strip() not in ["", "nan"]
    ])
    selected_vendor = st.selectbox(
        "Vendor", all_vendors, label_visibility="collapsed"
    )

    st.markdown(
        "<hr style='border-color:#3d5166;margin:14px 0'>",
        unsafe_allow_html=True
    )

    # ── Build filtered exploded df ────────────────────────────
    d_filt = df_exploded.copy()
    if selected_cat != "All":
        d_filt = d_filt[d_filt["Category"] == selected_cat]
    if selected_vendor != "All":
        d_filt = d_filt[d_filt["Vendor"] == selected_vendor]

    # ── FILTER 3: Service search + select ─────────────────────
    st.markdown("**🔍 Search Services**")
    svc_search = st.text_input(
        "Search", placeholder="e.g. Cisco, Oracle, M365…",
        label_visibility="collapsed"
    )

    available_svcs = sorted([
        s for s in d_filt["Service"].unique()
        if str(s).strip() not in ["", "nan"]
    ])
    if svc_search:
        available_svcs = [
            s for s in available_svcs
            if svc_search.lower() in s.lower()
        ]

    st.markdown(f"**🛠 Select Services** `{len(available_svcs)} available`")
    selected_svcs = st.multiselect(
        "Services",
        options=available_svcs,
        default=[],
        label_visibility="collapsed",
        help="Select one or more services — results appear instantly",
    )

    st.markdown(
        "<hr style='border-color:#3d5166;margin:14px 0'>",
        unsafe_allow_html=True
    )
    st.caption(f"📄 {len(df_master)} total quotes")
    st.caption(f"🛠 {df_exploded['Service'].nunique()} unique services")
    st.caption(f"🏢 {df_master['Vendor'].nunique()} vendors")

# ════════════════════════════════════════════════════════════
# MAIN PAGE — Header
# ════════════════════════════════════════════════════════════
st.markdown("""
<div style='background:linear-gradient(135deg,#2c3e50 0%,#3d5a73 100%);
            color:white;padding:20px 28px;border-radius:14px;
            margin-bottom:22px;box-shadow:0 4px 12px rgba(0,0,0,0.15)'>
    <h1 style='margin:0;font-size:1.65em;font-weight:800'>
        📋 IT Procurement — Service & Vendor Dashboard
    </h1>
    <p style='margin:7px 0 0;opacity:0.72;font-size:0.9em'>
        Use the sidebar to filter by Category & Vendor →
        Service list auto-updates →
        Select any service to see vendor & quotation file instantly
    </p>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# KPI CARDS
# ════════════════════════════════════════════════════════════
c1, c2, c3, c4 = st.columns(4)

def kpi_card(col, value, label, color):
    col.markdown(
        f"<div class='kpi-box' style='background:{color}'>"
        f"<div class='kpi-value'>{value}</div>"
        f"<div class='kpi-label'>{label}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )

kpi_card(c1, d_filt["File Name"].nunique(),  "Total Quotes",    "#2980b9")
kpi_card(c2, d_filt["Service"].nunique(),    "Unique Services", "#27ae60")
kpi_card(c3, d_filt["Vendor"].nunique(),     "Vendors",         "#8e44ad")
kpi_card(c4, d_filt["Category"].nunique(),   "Categories",      "#e67e22")

st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TABS: Charts | Raw Table
# ════════════════════════════════════════════════════════════
tab_charts, tab_table = st.tabs(["📊 Charts", "📄 Data Table"])

with tab_charts:

    # Row 1
    col_l, col_r = st.columns(2)

    with col_l:
        shared = (
            d_filt.groupby("Service")["Vendor"].nunique()
            .sort_values(ascending=False).head
