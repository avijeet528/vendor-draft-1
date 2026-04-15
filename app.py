# ============================================================
#  app.py — IT Procurement Dashboard
#  PwC Brand | Collapsible Filters | Extensive Price Table
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import defaultdict
import openpyxl
import re, io, os, requests, zipfile

st.set_page_config(
    page_title="IT Procurement Dashboard",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",   # ← sidebar hidden by default
)

# ════════════════════════════════════════════════════════════
# PwC BRAND CONSTANTS
# Primary   : #E03B24  (PwC Red)
# Dark      : #2D2D2D  (Charcoal)
# Mid grey  : #7D7D7D
# Light grey: #F2F2F2
# White     : #FFFFFF
# Accent    : #FFB600  (PwC Yellow — used sparingly)
# Font      : ITC Charter / Helvetica Neue → web-safe fallback
# ════════════════════════════════════════════════════════════
PWC_RED    = "#E03B24"
PWC_DARK   = "#2D2D2D"
PWC_GREY   = "#7D7D7D"
PWC_LGREY  = "#F2F2F2"
PWC_YELLOW = "#FFB600"
PWC_WHITE  = "#FFFFFF"

CHART_COLORS = [
    "#E03B24","#2D2D2D","#FFB600","#7D7D7D",
    "#C0392B","#4A4A4A","#FFC933","#A0A0A0",
    "#E85D47","#1A1A1A","#E6A800","#5A5A5A",
]

def get_color(i): return CHART_COLORS[i % len(CHART_COLORS)]

# ════════════════════════════════════════════════════════════
# CSS — PwC Theme + Fixed Table + Collapsible Sidebar
# ════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
  /* ── Google Font — Helvetica Neue fallback (PwC uses ITC Charter) */
  @import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

  html, body, [class*="css"] {{
      font-family: 'Source Sans Pro', 'Helvetica Neue',
                   Helvetica, Arial, sans-serif !important;
      color: {PWC_DARK};
  }}

  /* ── Page background */
  .main {{ background-color: {PWC_LGREY}; }}
  #MainMenu {{visibility:hidden;}}
  footer     {{visibility:hidden;}}
  header     {{visibility:hidden;}}

  /* ── Sidebar */
  section[data-testid="stSidebar"] {{
      background-color: {PWC_DARK} !important;
      border-right: 3px solid {PWC_RED} !important;
  }}
  section[data-testid="stSidebar"] label,
  section[data-testid="stSidebar"] p,
  section[data-testid="stSidebar"] span,
  section[data-testid="stSidebar"] div {{
      color: #ECECEC !important;
      font-family: 'Source Sans Pro', sans-serif !important;
  }}
  section[data-testid="stSidebar"] div[data-baseweb="select"] {{
      background-color: #3D3D3D !important;
      border-radius: 4px !important;
      border: 1px solid #555 !important;
  }}
  section[data-testid="stSidebar"] div[data-baseweb="select"] * {{
      color: #ECECEC !important;
  }}
  section[data-testid="stSidebar"] div[data-baseweb="input"] {{
      background-color: #3D3D3D !important;
      border-radius: 4px !important;
  }}
  section[data-testid="stSidebar"] div[data-baseweb="input"] input {{
      color: #ECECEC !important;
  }}
  section[data-testid="stSidebar"] span[data-baseweb="tag"] {{
      background-color: {PWC_RED} !important;
  }}
  section[data-testid="stSidebar"] span[data-baseweb="tag"] span {{
      color: white !important;
  }}

  /* ── KPI cards */
  .kpi-box {{
      border-radius  : 0px;
      padding        : 20px 12px;
      text-align     : center;
      color          : white;
      border-top     : 4px solid rgba(255,255,255,0.3);
  }}
  .kpi-value {{
      font-size   : 2.4em;
      font-weight : 700;
      margin      : 0;
      line-height : 1.1;
  }}
  .kpi-label {{
      font-size   : 0.82em;
      opacity     : 0.88;
      margin-top  : 6px;
      font-weight : 600;
      letter-spacing: 0.04em;
      text-transform: uppercase;
  }}

  /* ── Section headers */
  .section-title {{
      font-size      : 1.1em;
      font-weight    : 700;
      color          : {PWC_DARK};
      border-left    : 4px solid {PWC_RED};
      padding-left   : 10px;
      margin         : 18px 0 10px;
      text-transform : uppercase;
      letter-spacing : 0.05em;
  }}

  /* ── Vendor badge */
  .vendor-badge {{
      display        : inline-block;
      padding        : 3px 10px;
      border-radius  : 2px;
      color          : white;
      font-size      : 0.78em;
      font-weight    : 700;
      white-space    : nowrap;
      overflow       : hidden;
      text-overflow  : ellipsis;
      max-width      : 100%;
      box-sizing     : border-box;
      letter-spacing : 0.02em;
  }}

  /* ── Price table */
  .price-table {{
      width           : 100%;
      border-collapse : collapse;
      table-layout    : fixed;
      font-size       : 0.83em;
      font-family     : 'Source Sans Pro', sans-serif;
  }}
  .price-table thead tr {{
      background    : {PWC_DARK};
      color         : white;
  }}
  .price-table thead th {{
      padding       : 10px 12px;
      text-align    : left;
      font-weight   : 700;
      font-size     : 0.82em;
      letter-spacing: 0.04em;
      text-transform: uppercase;
      border        : none;
      word-break    : break-word;
  }}
  .price-table tbody tr:nth-child(even) {{
      background: #F9F9F9;
  }}
  .price-table tbody tr:hover {{
      background: #FFF0EE;
  }}
  .price-table td {{
      padding       : 9px 12px;
      border-bottom : 1px solid #E8E8E8;
      vertical-align: middle;
      word-break    : break-word;
  }}
  /* Sub-item rows (indented line items) */
  .price-table tr.line-item td {{
      background    : #FAFAFA;
      font-size     : 0.95em;
      color         : {PWC_GREY};
      border-bottom : 1px dashed #EEEEEE;
  }}
  .price-table tr.line-item td:first-child {{
      padding-left  : 28px;
      font-style    : italic;
  }}
  /* Total row */
  .price-table tr.total-row td {{
      background    : {PWC_LGREY};
      font-weight   : 700;
      color         : {PWC_DARK};
      border-top    : 2px solid {PWC_RED};
      border-bottom : 2px solid {PWC_RED};
  }}
  /* Grand total row */
  .price-table tr.grand-total td {{
      background    : {PWC_DARK};
      color         : white;
      font-weight   : 700;
      font-size     : 1.0em;
      border        : none;
  }}

  /* Fixed column widths */
  .price-table th:nth-child(1),
  .price-table td:nth-child(1) {{ width:13%; }}
  .price-table th:nth-child(2),
  .price-table td:nth-child(2) {{ width:13%; }}
  .price-table th:nth-child(3),
  .price-table td:nth-child(3) {{ width:28%; }}
  .price-table th:nth-child(4),
  .price-table td:nth-child(4) {{ width:14%; }}
  .price-table th:nth-child(5),
  .price-table td:nth-child(5) {{ width:14%; }}
  .price-table th:nth-child(6),
  .price-table td:nth-child(6) {{ width:18%; }}

  /* Expander */
  div[data-testid="stExpander"] {{
      border        : 1px solid #E0E0E0 !important;
      border-radius : 2px !important;
      background    : white !important;
  }}
  div[data-testid="stExpander"] summary {{
      background    : white !important;
  }}
  div[data-testid="stExpander"] summary p {{
      font-weight   : 700 !important;
      font-size     : 0.95em !important;
      color         : {PWC_DARK} !important;
  }}

  /* Tab styling */
  button[data-baseweb="tab"] {{
      font-weight   : 600 !important;
      color         : {PWC_GREY} !important;
  }}
  button[data-baseweb="tab"][aria-selected="true"] {{
      color         : {PWC_RED} !important;
      border-bottom : 3px solid {PWC_RED} !important;
  }}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# PRICE EXTRACTION ENGINE
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r'(?:USD|EUR|GBP|SGD|MYR|JPY|INR|THB|AUD|CAD)?\s*'
    r'[\$\€\£\¥\₹]?\s*'
    r'\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?'
    r'\s*(?:USD|EUR|GBP|SGD|MYR|JPY|INR|THB|AUD|CAD)?',
    re.IGNORECASE,
)
TOTAL_KW = [
    "grand total","total amount","total price","amount due",
    "total cost","net total","quote total","total value",
    "estimated total","subtotal","total",
]
LINE_KW = [
    "unit price","unit cost","list price","extended price",
    "extended cost","line total","each","per unit","qty","quantity",
]

def _parse_num(s):
    try:    return float(re.sub(r"[^\d.]", "", s) or "0")
    except: return 0.0

def _find_prices_near(text, keywords, min_val=10):
    """Return list of (keyword, price_str) found near any keyword."""
    results = []
    tl = text.lower()
    for kw in keywords:
        idx = tl.find(kw)
        while idx != -1:
            snippet = text[max(0,idx-30): idx+250]
            for m in PRICE_RE.findall(snippet):
                m = m.strip()
                if _parse_num(m) >= min_val:
                    results.append((kw, m))
            idx = tl.find(kw, idx+1)
    return results

def extract_all_prices(url: str):
    """
    Downloads file, returns dict:
    {
      "line_items": [(label, price_str), ...],
      "total"     : price_str or "",
      "raw_text"  : first 3000 chars for debugging
    }
    """
    result = {"line_items": [], "total": "", "raw_text": ""}
    if not url or not url.startswith("http"):
        return result
    try:
        resp = requests.get(url, timeout=20,
                            headers={"User-Agent":"Mozilla/5.0"})
        if resp.status_code != 200:
            return result
        content = resp.content
        ext     = url.split("?")[0].lower().rsplit(".", 1)[-1]

        # ── Extract raw text ─────────────────────────────────
        if ext == "pdf":
            import pdfplumber
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                text = "\n".join(
                    p.extract_text() or "" for p in pdf.pages
                )
        elif ext in ["xlsx", "xls"]:
            wb   = openpyxl.load_workbook(
                io.BytesIO(content), data_only=True, read_only=True)
            rows_text = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    line = "\t".join(
                        str(c) for c in row if c is not None
                    )
                    if line.strip():
                        rows_text.append(line)
            text = "\n".join(rows_text)
            wb.close()
        elif ext == "docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                xml = z.read("word/document.xml").decode(
                    "utf-8", errors="ignore")
            text = re.sub(r"<[^>]+>", " ", xml)
        else:
            return result

        result["raw_text"] = text[:3000]

        # ── 1. Grand / sub total ─────────────────────────────
        total_hits = _find_prices_near(text, TOTAL_KW, min_val=100)
        if total_hits:
            best = max(total_hits, key=lambda x: _parse_num(x[1]))
            result["total"] = best[1]

        # ── 2. Line items near unit/line keywords ─────────────
        line_hits = _find_prices_near(text, LINE_KW, min_val=1)
        seen = set()
        for kw, price in line_hits:
            key = (kw, price)
            if key not in seen:
                seen.add(key)
                result["line_items"].append((kw.title(), price))

        # ── 3. Fallback: all prices if nothing found ──────────
        if not result["line_items"] and not result["total"]:
            all_p = [
                m.strip() for m in PRICE_RE.findall(text)
                if _parse_num(m.strip()) >= 100
            ]
            unique = list(dict.fromkeys(all_p))[:15]
            result["line_items"] = [("Price found", p) for p in unique]
            if unique:
                result["total"] = max(
                    unique, key=lambda x: _parse_num(x))

    except Exception as e:
        result["raw_text"] = f"Error: {e}"
    return result

# ════════════════════════════════════════════════════════════
# HYPERLINK EXTRACTOR
# ════════════════════════════════════════════════════════════
@st.cache_data
def extract_hyperlinks(file_path: str) -> dict:
    link_map = {}
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        fn_col = hdr_row = None
        for row in ws.iter_rows():
            for cell in row:
                if (cell.value and
                        str(cell.value).strip().lower() == "file name"):
                    fn_col  = cell.column
                    hdr_row = cell.row
                    break
            if fn_col: break
        if fn_col:
            for row in ws.iter_rows(
                    min_row=hdr_row+1, min_col=fn_col, max_col=fn_col):
                cell = row[0]
                if cell.value and cell.hyperlink:
                    link_map[str(cell.value).strip()] = \
                        str(cell.hyperlink.target).strip()
        wb.close()
    except Exception as e:
        st.warning(f"Hyperlink extraction: {e}")
    return link_map

# ════════════════════════════════════════════════════════════
# DATA LOADING
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    FILE_PATH = "Master Catalog.xlsx"
    if not os.path.exists(FILE_PATH):
        st.error(f"❌ File not found: '{FILE_PATH}'")
        return None, None

    raw = pd.read_excel(FILE_PATH, engine="openpyxl", header=None)
    hdr = None
    for i, row in raw.iterrows():
        rv = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if any("category" in v for v in rv) and any("file" in v for v in rv):
            hdr = i; break
    if hdr is None:
        st.error("❌ Header row not found"); return None, None

    df = pd.read_excel(FILE_PATH, engine="openpyxl", header=hdr)
    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)

    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category":                  col_map["Category"]     = c
        elif "vendor" in cl or "type" in cl:  col_map["Vendor"]       = c
        elif cl == "file name":               col_map["File Name"]    = c
        elif cl == "file link":               col_map["File Link"]    = c
        elif cl == "file url":                col_map["File URL"]     = c
        elif "comment" in cl:                 col_map["Comments"]     = c
        elif "quoted" in cl or "price" in cl: col_map["Quoted Price"] = c

    df.rename(columns={v: k for k, v in col_map.items()}, inplace=True)
    keep = ["Category","Vendor","File Name","Comments"]
    for ex in ["File Link","File URL","Quoted Price"]:
        if ex in df.columns: keep.append(ex)
    df = df[[c for c in keep if c in df.columns]].copy()
    df = df[~(
        df["Category"].astype(str).str.strip().isin(["","nan"]) &
        df["Vendor"].astype(str).str.strip().isin(["","nan"])
    )].copy()
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df.reset_index(drop=True, inplace=True)

    # Hyperlinks
    hmap = extract_hyperlinks(FILE_PATH)
    df["Hyperlink"] = df["File Name"].map(hmap).fillna("")
    for fb in ["File Link","File URL"]:
        if fb in df.columns:
            df["Hyperlink"] = df.apply(
                lambda r: r["Hyperlink"]
                if r["Hyperlink"] not in ["","nan"] else r[fb], axis=1)

    # Parse services
    def parse_svc(v):
        if not v or str(v).strip() in ["","nan"]: return ["(unspecified)"]
        p = [s.strip() for s in str(v).split("\n") if s.strip()]
        return p or ["(unspecified)"]
    df["Services List"] = df["Comments"].apply(parse_svc)

    df_exp = df.explode("Services List").copy()
    df_exp.rename(columns={"Services List":"Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].str.strip()
    df_exp = df_exp[
        ~df_exp["Service"].isin(["","(unspecified)","nan"])
    ].reset_index(drop=True)

    return df, df_exp

# ────────────────────────────────────────────────────────────
df_master, df_exploded = load_data()
if df_master is None or df_exploded is None:
    st.stop()

vendor_color_map = {
    v: get_color(i)
    for i, v in enumerate(sorted(df_master["Vendor"].unique()))
}

# ════════════════════════════════════════════════════════════
# SIDEBAR  (collapsed by default — user clicks ▶ to open)
# ════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style='text-align:center;padding:18px 0 12px'>
        <div style='font-size:2em'>📋</div>
        <div style='font-size:1.1em;font-weight:700;
                    color:white;margin:4px 0 2px;
                    letter-spacing:0.05em'>
            IT PROCUREMENT
        </div>
        <div style='font-size:0.75em;color:{PWC_GREY};
                    letter-spacing:0.08em;text-transform:uppercase'>
            Service &amp; Vendor Dashboard
        </div>
    </div>
    <div style='height:3px;background:{PWC_RED};
                margin:0 0 16px;border-radius:2px'></div>
    """, unsafe_allow_html=True)

    # ── Category ─────────────────────────────────────────────
    st.markdown(
        f"<p style='color:#ECECEC;font-weight:700;font-size:0.8em;"
        f"text-transform:uppercase;letter-spacing:0.06em;"
        f"margin-bottom:4px'>📂 Category</p>",
        unsafe_allow_html=True)
    all_cats = ["All"] + sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["","nan"]])
    selected_cat = st.selectbox(
        "Category", all_cats, label_visibility="collapsed")

    # ── Vendor (scoped) ───────────────────────────────────────
    st.markdown(
        f"<p style='color:#ECECEC;font-weight:700;font-size:0.8em;"
        f"text-transform:uppercase;letter-spacing:0.06em;"
        f"margin:12px 0 4px'>🏢 Vendor</p>",
        unsafe_allow_html=True)
    vpool = (df_master if selected_cat == "All"
             else df_master[df_master["Category"] == selected_cat])
    all_vendors = ["All"] + sorted([
        v for v in vpool["Vendor"].unique()
        if str(v).strip() not in ["","nan"]])
    selected_vendor = st.selectbox(
        "Vendor", all_vendors, label_visibility="collapsed")

    st.markdown(
        f"<div style='height:1px;background:#444;margin:14px 0'></div>",
        unsafe_allow_html=True)

    # ── Filtered data ─────────────────────────────────────────
    d_filt = df_exploded.copy()
    if selected_cat    != "All": d_filt = d_filt[d_filt["Category"] == selected_cat]
    if selected_vendor != "All": d_filt = d_filt[d_filt["Vendor"]   == selected_vendor]

    # ── Service search ────────────────────────────────────────
    st.markdown(
        f"<p style='color:#ECECEC;font-weight:700;font-size:0.8em;"
        f"text-transform:uppercase;letter-spacing:0.06em;"
        f"margin-bottom:4px'>🔍 Search Services</p>",
        unsafe_allow_html=True)
    svc_search = st.text_input(
        "Search", placeholder="e.g. Cisco, Oracle…",
        label_visibility="collapsed")

    avail_svcs = sorted([
        s for s in d_filt["Service"].unique()
        if str(s).strip() not in ["","nan"]])
    if svc_search:
        avail_svcs = [s for s in avail_svcs
                      if svc_search.lower() in s.lower()]

    st.markdown(
        f"<p style='color:#ECECEC;font-weight:700;font-size:0.8em;"
        f"text-transform:uppercase;letter-spacing:0.06em;"
        f"margin:12px 0 4px'>🛠 Select Services "
        f"<span style='font-weight:400;color:{PWC_GREY}'>"
        f"({len(avail_svcs)})</span></p>",
        unsafe_allow_html=True)
    selected_svcs = st.multiselect(
        "Services", options=avail_svcs, default=[],
        label_visibility="collapsed",
        help="Select services to see vendor & pricing details")

    st.markdown(
        f"<div style='height:1px;background:#444;margin:14px 0'></div>",
        unsafe_allow_html=True)
    st.markdown(
        f"<p style='color:{PWC_GREY};font-size:0.78em;margin:2px 0'>"
        f"📄 {len(df_master)} quotes &nbsp;|&nbsp; "
        f"🛠 {df_exploded['Service'].nunique()} services &nbsp;|&nbsp; "
        f"🏢 {df_master['Vendor'].nunique()} vendors</p>",
        unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
col_logo, col_title = st.columns([1, 11])
with col_logo:
    # PwC logo placeholder — red square mimicking PwC logo block
    st.markdown(f"""
    <div style='background:{PWC_RED};width:48px;height:48px;
                border-radius:2px;margin-top:4px;
                display:flex;align-items:center;justify-content:center'>
        <span style='color:white;font-weight:900;font-size:1.1em;
                     letter-spacing:-1px'>PwC</span>
    </div>""", unsafe_allow_html=True)
with col_title:
    st.markdown(f"""
    <div style='padding:4px 0 8px'>
        <h1 style='margin:0;font-size:1.55em;font-weight:700;
                   color:{PWC_DARK};letter-spacing:0.02em'>
            IT Procurement — Service &amp; Vendor Dashboard
        </h1>
        <p style='margin:3px 0 0;color:{PWC_GREY};font-size:0.88em'>
            Filter by Category → Vendor auto-updates →
            Select a service → Analyse pricing &amp; quotation files
        </p>
    </div>""", unsafe_allow_html=True)

st.markdown(
    f"<div style='height:3px;background:{PWC_RED};"
    f"margin:4px 0 20px;border-radius:2px'></div>",
    unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# KPI CARDS
# ════════════════════════════════════════════════════════════
k1,k2,k3,k4 = st.columns(4)
def kpi(col, val, label, color):
    col.markdown(
        f"<div class='kpi-box' style='background:{color}'>"
        f"<div class='kpi-value'>{val}</div>"
        f"<div class='kpi-label'>{label}</div></div>",
        unsafe_allow_html=True)
kpi(k1, d_filt["File Name"].nunique(),  "Quote Files",      PWC_RED)
kpi(k2, d_filt["Service"].nunique(),    "Unique Services",  PWC_DARK)
kpi(k3, d_filt["Vendor"].nunique(),     "Vendors",          "#4A4A4A")
kpi(k4, d_filt["Category"].nunique(),   "Categories",       PWC_GREY)

st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# CHARTS SECTION
# Three purposeful charts — no heatmap
# ════════════════════════════════════════════════════════════
st.markdown(
    "<div class='section-title'>📊 Portfolio Overview</div>",
    unsafe_allow_html=True)

# Row 1: full-width stacked bar — services per vendor per category
st.markdown("##### Service Coverage by Vendor")
pivot = (
    d_filt.drop_duplicates(subset=["Vendor","Service","Category"])
    .groupby(["Vendor","Category"])["Service"].count()
    .reset_index()
)
pivot.columns = ["Vendor","Category","Service Count"]
if not pivot.empty:
    fig_stacked = px.bar(
        pivot, x="Vendor", y="Service Count", color="Category",
        title="Services Offered per Vendor — broken down by Category",
        color_discrete_sequence=CHART_COLORS,
        text_auto=True,
    )
    fig_stacked.update_layout(
        height=380, plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Source Sans Pro, Helvetica Neue, sans-serif",
                  size=11, color=PWC_DARK),
        legend=dict(orientation="h", yanchor="bottom",
                    y=1.02, xanchor="right", x=1,
                    font=dict(size=10)),
        xaxis_tickangle=-30,
        xaxis_title="", yaxis_title="# Services",
        margin=dict(l=10,r=10,t=60,b=10),
        bargap=0.25,
    )
    fig_stacked.update_traces(
        textfont_size=9, textangle=0,
        textposition="inside", cliponaxis=False)
    st.plotly_chart(fig_stacked, use_container_width=True)

# Row 2: two side-by-side charts
c_left, c_right = st.columns(2)

with c_left:
    # Horizontal bar — top shared services (red = shared by 2+)
    shared = (
        d_filt.groupby("Service")["Vendor"].nunique()
        .sort_values(ascending=False).head(18).reset_index()
    )
    shared.columns = ["Service","Vendor Count"]
    shared["Color"] = shared["Vendor Count"].apply(
        lambda x: PWC_RED if x > 1 else "#C0C0C0")
    fig_shared = go.Figure(go.Bar(
        x=shared["Vendor Count"],
        y=shared["Service"].str[:40],
        orientation="h",
        marker_color=shared["Color"],
        text=shared["Vendor Count"],
        textposition="outside",
    ))
    fig_shared.update_layout(
        title="🔁 Services Offered by Multiple Vendors",
        xaxis_title="# Vendors", height=500,
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Source Sans Pro, sans-serif",
                  size=10, color=PWC_DARK),
        margin=dict(l=10,r=20,t=45,b=10),
        yaxis=dict(autorange="reversed"),
        annotations=[dict(
            x=0.98, y=-0.08, xref="paper", yref="paper",
            text="<b style='color:#E03B24'>■</b> Shared  "
                 "<b style='color:#C0C0C0'>■</b> Single vendor",
            showarrow=False, font=dict(size=9), xanchor="right")]
    )
    st.plotly_chart(fig_shared, use_container_width=True)

with c_right:
    # Donut — quotes by category
    cat_counts = (
        d_filt.drop_duplicates(subset=["Category","File Name"])
        .groupby("Category").size().reset_index()
    )
    cat_counts.columns = ["Category","Count"]
    if not cat_counts.empty:
        fig_pie = px.pie(
            cat_counts, names="Category", values="Count",
            title="🥧 Quote Distribution by Category",
            hole=0.50, color_discrete_sequence=CHART_COLORS,
        )
        fig_pie.update_layout(
            height=500,
            margin=dict(l=10,r=10,t=45,b=10),
            paper_bgcolor="white",
            font=dict(family="Source Sans Pro, sans-serif",
                      size=11, color=PWC_DARK),
            legend=dict(font=dict(size=10), orientation="v"),
        )
        fig_pie.update_traces(
            textposition="inside",
            textinfo="percent+label",
            textfont_size=10)
        st.plotly_chart(fig_pie, use_container_width=True)

# ════════════════════════════════════════════════════════════
# DATA TABLE TAB
# ════════════════════════════════════════════════════════════
with st.expander("📄 Raw Data Table", expanded=False):
    dm = df_master.copy()
    if selected_cat    != "All": dm = dm[dm["Category"] == selected_cat]
    if selected_vendor != "All": dm = dm[dm["Vendor"]   == selected_vendor]
    st.dataframe(
        dm.drop(columns=["Services List","Hyperlink"], errors="ignore"),
        use_container_width=True, height=380)

# ════════════════════════════════════════════════════════════
# SERVICE SELECTION RESULTS
# ════════════════════════════════════════════════════════════
st.markdown(
    f"<div style='height:3px;background:{PWC_RED};"
    f"margin:20px 0 12px;border-radius:2px'></div>",
    unsafe_allow_html=True)
st.markdown(
    "<div class='section-title'>🛠 Service → Vendor & Pricing Analysis</div>",
    unsafe_allow_html=True)

if not selected_svcs:
    st.markdown(f"""
    <div style='background:white;border:1px solid #E0E0E0;
                border-left:4px solid {PWC_RED};
                padding:16px 20px;border-radius:2px;color:{PWC_GREY}'>
        <b>👈 Open the sidebar (▶ arrow on the left) to select services.</b><br>
        <span style='font-size:0.9em'>
        Filter by Category → Vendor updates automatically →
        Pick one or more services → See full vendor &amp; pricing breakdown.
        </span>
    </div>""", unsafe_allow_html=True)
else:
    d_sel = d_filt[d_filt["Service"].isin(selected_svcs)].copy()

    if d_sel.empty:
        st.warning("⚠️ No results for selected service(s) under current filters.")
    else:
        # ── Vendor coverage ───────────────────────────────────
        vsmap = defaultdict(set)
        for _, r in d_sel.iterrows():
            vsmap[r["Vendor"]].add(r["Service"])

        vendors_all  = sorted([v for v,s in vsmap.items()
                                if set(selected_svcs).issubset(s)])
        vendors_some = sorted([v for v,s in vsmap.items()
                                if not set(selected_svcs).issubset(s)])

        if len(selected_svcs) > 1:
            if vendors_all:
                names = " · ".join([f"**{v}**" for v in vendors_all])
                st.success(
                    f"✅ **{len(vendors_all)} vendor(s) offer ALL "
                    f"{len(selected_svcs)} services:** {names}")
            else:
                st.warning(
                    "⚠️ No single vendor covers all selected services — "
                    "see breakdown below.")
            if vendors_some:
                with st.expander("🔵 Vendors with partial coverage",
                                 expanded=False):
                    for v in vendors_some:
                        cov = vsmap[v].intersection(set(selected_svcs))
                        vc  = vendor_color_map.get(v,"#666")
                        st.markdown(
                            f"<span class='vendor-badge' "
                            f"style='background:{vc}'>{v}</span>"
                            f" &nbsp; covers **{len(cov)}/{len(selected_svcs)}**: "
                            f"_{', '.join(sorted(cov))}_",
                            unsafe_allow_html=True)

        has_price = "Quoted Price" in d_sel.columns

        # ════════════════════════════════════════════════════
        # PER-SERVICE EXPANDERS WITH EXTENSIVE PRICE TABLE
        # ════════════════════════════════════════════════════
        for svc in selected_svcs:
            d_svc = (
                d_sel[d_sel["Service"] == svc]
                .drop_duplicates(subset=["Vendor","File Name"])
                .sort_values("Vendor")
            )
            vc   = d_svc["Vendor"].nunique()
            stag = ("⚠️ SHARED" if vc > 1 else "✅ SINGLE VENDOR")

            with st.expander(
                f"🛠  {svc}  ·  {vc} vendor(s) · {len(d_svc)} file(s) [{stag}]",
                expanded=True):

                # Vendor badge row
                badges = " ".join([
                    f"<span class='vendor-badge' "
                    f"style='background:{vendor_color_map.get(v,\"#666\")}'>"
                    f"{v}</span>"
                    for v in sorted(d_svc["Vendor"].unique())])
                st.markdown(
                    f"<div style='margin-bottom:14px'>"
                    f"<b>Vendors offering this service:</b>"
                    f"&nbsp;{badges}</div>",
                    unsafe_allow_html=True)

                # ── Build extensive price table ───────────────
                # Columns:
                # Vendor | Category | File Name | Quoted Price
                # | Line Items & Values | File Link

                tbl = [
                    "<table class='price-table'>"
                    "<thead><tr>"
                    "<th>Vendor</th>"
                    "<th>Category</th>"
                    "<th>📄 File Name</th>"
                    "<th>💰 Quoted Price</th>"
                    "<th>🔢 Line Item</th>"
                    "<th>💵 Item Value &nbsp;|&nbsp; 🔗 Link</th>"
                    "</tr></thead><tbody>"
                ]

                grand_total_num = 0.0
                all_totals      = []

                for i, (_, row) in enumerate(d_svc.iterrows()):
                    bg  = "#ffffff" if i % 2 == 0 else "#FAFAFA"
                    vc2 = vendor_color_map.get(row["Vendor"],"#666")

                    fname = str(row.get("File Name","")).strip()
                    url   = str(row.get("Hyperlink","")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File Link","")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File URL","")).strip()
                    valid_url = (url and url not in ["","nan"]
                                 and url.startswith("http"))

                    # Quoted price from Excel column
                    qprice = str(row.get("Quoted Price","")).strip()
                    qprice = "" if qprice in ["","nan","0"] else qprice

                    # File link cell
                    label = fname[:38]+"…" if len(fname)>38 else fname
                    link_html = (
                        f"<a href='{url}' target='_blank' "
                        f"style='color:{PWC_RED};"
                        f"text-decoration:underline;"
                        f"font-size:0.82em;word-break:break-all'>"
                        f"🔗 {label}</a>"
                        if valid_url else
                        f"<span style='color:#C0C0C0;"
                        f"font-size:0.82em;font-style:italic'>"
                        f"No link</span>"
                    )

                    # Price from Excel
                    qprice_html = (
                        f"<span style='color:#27ae60;font-weight:700'>"
                        f"{qprice}</span>"
                        if qprice else
                        f"<span style='color:#C0C0C0'>—</span>"
                    )

                    # ── Main file row ─────────────────────────
                    tbl.append(
                        f"<tr style='background:{bg}'>"
                        f"<td><span class='vendor-badge' "
                        f"style='background:{vc2}'>"
                        f"{row['Vendor']}</span></td>"
                        f"<td style='color:{PWC_GREY}'>"
                        f"{row['Category']}</td>"
                        f"<td style='font-family:monospace;"
                        f"font-size:0.81em;color:{PWC_DARK};"
                        f"word-break:break-all'>{fname}</td>"
                        f"<td>{qprice_html}</td>"
                        f"<td colspan='2'>{link_html}</td>"
                        f"</tr>"
                    )

                    # ── Extracted line items (if URL available) ─
                    if valid_url:
                        extract_key = f"extracted_{fname[:40]}"
                        if st.session_state.get(f"run_{extract_key}"):
                            pr = st.session_state.get(
                                extract_key, {"line_items":[],"total":""})
                            items  = pr.get("line_items", [])
                            etotal = pr.get("total", "")

                            if items:
                                for j, (lbl, val) in enumerate(items[:20]):
                                    tbl.append(
                                        f"<tr class='line-item'>"
                                        f"<td></td><td></td>"
                                        f"<td style='color:{PWC_GREY}'>"
                                        f"↳ {fname[:30]}</td>"
                                        f"<td></td>"
                                        f"<td style='color:{PWC_GREY};"
                                        f"font-size:0.88em'>{lbl}</td>"
                                        f"<td style='color:{PWC_DARK};"
                                        f"font-weight:600'>{val}</td>"
                                        f"</tr>"
                                    )
                            if etotal:
                                tval = _parse_num(etotal)
                                if tval > 0:
                                    all_totals.append(tval)
                                    grand_total_num += tval
                                tbl.append(
                                    f"<tr class='total-row'>"
                                    f"<td colspan='4'></td>"
                                    f"<td><b>File Total</b></td>"
                                    f"<td style='color:{PWC_RED};"
                                    f"font-size:1.05em;font-weight:700'>"
                                    f"{etotal}</td>"
                                    f"</tr>"
                                )

                # Grand total row across all files in this service
                if grand_total_num > 0:
                    tbl.append(
                        f"<tr class='grand-total'>"
                        f"<td colspan='4'>"
                        f"<b>GRAND TOTAL</b> — all files for '{svc[:40]}'</td>"
                        f"<td></td>"
                        f"<td style='font-size:1.1em;"
                        f"letter-spacing:0.02em'>"
                        f"≈ {grand_total_num:,.2f}</td>"
                        f"</tr>"
                    )

                tbl.append("</tbody></table>")
                st.markdown("".join(tbl), unsafe_allow_html=True)

                # ── Extract prices buttons ────────────────────
                st.markdown(
                    f"<div style='margin-top:14px;padding:10px;"
                    f"background:{PWC_LGREY};border-radius:2px;"
                    f"border-left:3px solid {PWC_RED}'>"
                    f"<b style='font-size:0.85em;"
                    f"text-transform:uppercase;letter-spacing:0.05em'>"
                    f"💰 Extract Prices from Files</b></div>",
                    unsafe_allow_html=True)

                btn_cols = st.columns(
                    min(len(d_svc), 4))
                for bi, (_, brow) in enumerate(d_svc.iterrows()):
                    burl  = str(brow.get("Hyperlink","")).strip()
                    if not burl or burl=="nan":
                        burl = str(brow.get("File Link","")).strip()
                    bfname = str(brow.get("File Name","")).strip()
                    bkey   = f"extracted_{bfname[:40]}"
                    rkey   = f"run_{bkey}"

                    if burl and burl.startswith("http"):
                        with btn_cols[bi % 4]:
                            if st.button(
                                f"📥 {bfname[:28]}…"
                                if len(bfname)>28 else f"📥 {bfname}",
                                key=f"btn_{bkey}_{svc[:20]}",
                                help=f"Extract prices from {bfname}",
                                type="primary",
                            ):
                                with st.spinner(
                                        f"Extracting from {bfname[:40]}…"):
                                    result = extract_all_prices(burl)
                                    st.session_state[bkey]  = result
                                    st.session_state[rkey]  = True
                                st.rerun()
                    else:
                        with btn_cols[bi % 4]:
                            st.button(
                                f"🚫 {bfname[:24]}",
                                key=f"btn_na_{bkey}_{svc[:20]}",
                                disabled=True,
                                help="No URL available for this file")

        # ── Shared services summary ───────────────────────────
        shared_svcs = [
            s for s in selected_svcs
            if d_sel[d_sel["Service"]==s]["Vendor"].nunique() > 1]
        if shared_svcs:
            with st.expander(
                "🔁 Shared Services — same service, multiple vendors",
                    expanded=False):
                for s in shared_svcs:
                    vlist  = sorted(
                        d_sel[d_sel["Service"]==s]["Vendor"].unique())
                    bgs = " ".join([
                        f"<span class='vendor-badge' "
                        f"style='background:{vendor_color_map.get(v,\"#666\")}'>"
                        f"{v}</span>"
                        for v in vlist])
                    st.markdown(
                        f"<div style='margin-bottom:8px'>"
                        f"<b>{s}</b> → {bgs}</div>",
                        unsafe_allow_html=True)
