# ============================================================
#  app.py — IT Procurement Service Dashboard
#  PwC Brand: Colours #D04A02 (orange) #2D2D2D (dark)
#             #F3F3F3 (light grey)  #FFFFFF (white)
#  Font: ITC Charter / Georgia fallback (PwC editorial font)
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import defaultdict
import openpyxl
import os, re, io, zipfile

# ── Optional price-extraction libraries ──────────────────────
try:
    import requests
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False

try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ════════════════════════════════════════════════════════════
# PAGE CONFIG
# ════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="IT Procurement Dashboard",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ════════════════════════════════════════════════════════════
# PwC BRAND PALETTE
# ════════════════════════════════════════════════════════════
PWC = {
    "orange"      : "#D04A02",
    "orange_light": "#EB8C00",
    "dark"        : "#2D2D2D",
    "mid"         : "#7D7D7D",
    "light"       : "#F3F3F3",
    "white"       : "#FFFFFF",
    "red"         : "#E0301E",
    "teal"        : "#299D8F",
    "blue"        : "#295477",
    "yellow"      : "#FFB600",
    "green"       : "#22992E",
}

# Chart colour sequence (PwC approved palette)
PWC_CHART = [
    PWC["orange"], PWC["blue"],  PWC["teal"],
    PWC["yellow"], PWC["green"], PWC["red"],
    PWC["orange_light"], "#6E2585", "#8C8C8C", "#004F9F",
]

# ════════════════════════════════════════════════════════════
# GLOBAL CSS  — PwC font + colour scheme
# ════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
/* ── PwC font stack ── */
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

html, body, [class*="css"], .stMarkdown, .stText,
div, p, span, td, th, label, button {{
    font-family: 'Source Sans Pro', 'Helvetica Neue',
                 Arial, sans-serif !important;
    color: {PWC["dark"]};
}}

/* ── Page background ── */
.main .block-container {{
    background-color: {PWC["light"]};
    padding-top: 1.5rem;
}}

/* ── Hide Streamlit chrome ── */
#MainMenu, footer, header {{ visibility: hidden; }}

/* ── KPI cards ── */
.kpi-box {{
    border-radius   : 4px;
    padding         : 20px 12px;
    text-align      : center;
    color           : {PWC["white"]};
    border-left     : 5px solid rgba(255,255,255,0.35);
}}
.kpi-value {{
    font-size   : 2.3em;
    font-weight : 700;
    margin      : 0;
    line-height : 1.1;
    letter-spacing: -0.5px;
}}
.kpi-label {{
    font-size  : 0.82em;
    font-weight: 600;
    opacity    : 0.92;
    margin-top : 6px;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {{
    background-color: {PWC["dark"]} !important;
    border-right: 3px solid {PWC["orange"]};
}}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div {{
    color: #F0F0F0 !important;
    font-family: 'Source Sans Pro', sans-serif !important;
}}
section[data-testid="stSidebar"] div[data-baseweb="select"] {{
    background-color: {PWC["white"]} !important;
    border-radius: 2px !important;
    border: 1px solid #999 !important;
}}
section[data-testid="stSidebar"] div[data-baseweb="select"] * {{
    color: {PWC["dark"]} !important;
}}
section[data-testid="stSidebar"] div[data-baseweb="input"] {{
    background-color: {PWC["white"]} !important;
    border-radius: 2px !important;
}}
section[data-testid="stSidebar"] div[data-baseweb="input"] input {{
    color: {PWC["dark"]} !important;
}}
section[data-testid="stSidebar"] span[data-baseweb="tag"] {{
    background-color: {PWC["orange"]} !important;
    border-radius: 2px !important;
}}
section[data-testid="stSidebar"] span[data-baseweb="tag"] span {{
    color: white !important;
}}

/* ── Tabs ── */
button[data-baseweb="tab"] {{
    font-weight : 600 !important;
    font-size   : 0.92em !important;
    color       : {PWC["mid"]} !important;
}}
button[data-baseweb="tab"][aria-selected="true"] {{
    color       : {PWC["orange"]} !important;
    border-bottom: 3px solid {PWC["orange"]} !important;
}}

/* ── Expander ── */
div[data-testid="stExpander"] details summary p {{
    font-weight : 700;
    font-size   : 0.95em;
    color       : {PWC["dark"]} !important;
}}
div[data-testid="stExpander"] details {{
    border     : 1px solid #ddd;
    border-radius: 4px;
    margin-bottom: 10px;
}}

/* ── Vendor badge ── */
.vendor-badge {{
    display        : inline-block;
    padding        : 3px 10px;
    border-radius  : 2px;
    color          : white;
    font-size      : 0.79em;
    font-weight    : 700;
    white-space    : nowrap;
    overflow       : hidden;
    text-overflow  : ellipsis;
    max-width      : 100%;
    box-sizing     : border-box;
    letter-spacing : 0.3px;
}}

/* ── Price table ── */
.price-table {{
    width           : 100%;
    border-collapse : collapse;
    table-layout    : fixed;
    font-size       : 0.84em;
    margin-top      : 8px;
    border          : 1px solid #e0e0e0;
    border-radius   : 4px;
    overflow        : hidden;
}}
.price-table thead tr {{
    background : {PWC["dark"]};
    color      : white;
}}
.price-table thead th {{
    padding    : 10px 12px;
    text-align : left;
    font-weight: 700;
    font-size  : 0.83em;
    letter-spacing: 0.4px;
    text-transform: uppercase;
    border     : none;
}}
.price-table tbody tr:nth-child(even) {{
    background: {PWC["light"]};
}}
.price-table tbody tr:hover {{
    background: #FCE8DC;
}}
.price-table tbody td {{
    padding        : 9px 12px;
    border-bottom  : 1px solid #e8e8e8;
    vertical-align : middle;
    word-break     : break-word;
}}
/* column widths */
.price-table th:nth-child(1),
.price-table td:nth-child(1) {{ width:12%; }}  /* Vendor      */
.price-table th:nth-child(2),
.price-table td:nth-child(2) {{ width:13%; }}  /* Category    */
.price-table th:nth-child(3),
.price-table td:nth-child(3) {{ width:22%; }}  /* File Name   */
.price-table th:nth-child(4),
.price-table td:nth-child(4) {{ width:22%; }}  /* Line Items  */
.price-table th:nth-child(5),
.price-table td:nth-child(5) {{ width:12%; }}  /* Unit Price  */
.price-table th:nth-child(6),
.price-table td:nth-child(6) {{ width:10%; }}  /* Total       */
.price-table th:nth-child(7),
.price-table td:nth-child(7) {{ width:9%;  }}  /* File Link   */

/* total row */
.total-row {{
    background : #FCE8DC !important;
    font-weight: 700;
}}
.total-row td {{
    border-top  : 2px solid {PWC["orange"]} !important;
    color       : {PWC["orange"]} !important;
}}

/* info/success/warning boxes */
.stAlert > div {{
    border-radius: 2px !important;
    font-family: 'Source Sans Pro', sans-serif !important;
}}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════
def vendor_color(vendor_name: str, vmap: dict) -> str:
    return vmap.get(vendor_name, PWC["mid"])


def fmt_currency(val) -> str:
    """Format a numeric value as currency string."""
    try:
        f = float(str(val).replace(",", "").replace("$", "")
                  .replace("€", "").replace("£", "").strip())
        return f"${f:,.2f}"
    except Exception:
        return str(val)


# ════════════════════════════════════════════════════════════
# PRICE EXTRACTION ENGINE
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r"""
    (?:[\$\€\£\¥][\s]?)?
    (?:USD|EUR|GBP|JPY|SGD|MYR|THB|AUD|CAD|INR)[\s]?
    \d{1,3}(?:[,\s]\d{3})*
    (?:\.\d{1,2})?
    |
    (?:[\$\€\£\¥][\s]?)
    \d{1,3}(?:[,\s]\d{3})*
    (?:\.\d{1,2})?
    |
    \d{1,3}(?:[,]\d{3})+
    (?:\.\d{1,2})?
    """,
    re.VERBOSE | re.IGNORECASE,
)

TOTAL_KW = [
    "grand total", "total amount", "total price",
    "amount due",  "net total",    "total cost",
    "total value", "quote total",  "estimated total",
    "total",
]

LINE_KW = [
    r"unit\s*price", r"unit\s*cost", r"each", r"per\s*unit",
    r"list\s*price", r"extended\s*price", r"ext\.?\s*price",
    r"subtotal", r"line\s*total", r"amount",
]

def _parse_num(s: str) -> float:
    try:
        return float(re.sub(r"[^\d.]", "", s) or "0")
    except ValueError:
        return 0.0


def _best_price_near_keyword(text: str) -> str:
    tl = text.lower()
    for kw in TOTAL_KW:
        idx = tl.find(kw)
        if idx == -1:
            continue
        snippet = text[max(0, idx - 30): idx + 350]
        hits    = PRICE_RE.findall(snippet)
        valid   = [h.strip() for h in hits if _parse_num(h) >= 100]
        if valid:
            return max(valid, key=_parse_num)
    return ""


def _extract_line_items(text: str) -> list[dict]:
    """
    Try to pull individual line-item rows from text.
    Returns list of { description, unit_price, quantity, line_total }
    """
    items = []
    lines = text.split("\n")
    for line in lines:
        prices = PRICE_RE.findall(line)
        nums   = [p for p in prices if _parse_num(p) >= 1]
        if not nums:
            continue
        desc = re.sub(r"[\$\€\£\¥\d,\.]+", "", line).strip()
        desc = re.sub(r"\s{2,}", " ", desc)
        if len(desc) < 3:
            continue
        if len(nums) >= 2:
            unit  = min(nums, key=_parse_num)
            total = max(nums, key=_parse_num)
        else:
            unit  = nums[0]
            total = nums[0]
        items.append({
            "description": desc[:80],
            "unit_price" : unit,
            "line_total" : total,
        })
    return items[:30]          # cap at 30 items per file


def extract_prices_from_bytes(content: bytes, ext: str) -> dict:
    """
    Given file bytes + extension, returns:
      {
        'grand_total'  : str,
        'line_items'   : [ {description, unit_price, line_total}, ... ],
        'raw_text'     : str   (first 3000 chars for debugging)
      }
    """
    text = ""
    ext  = ext.lower().strip(".")

    try:
        if ext == "pdf":
            if not PDF_OK:
                return {"grand_total": "pdfplumber not installed",
                        "line_items": [], "raw_text": ""}
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"

        elif ext in ("xlsx", "xls"):
            wb = openpyxl.load_workbook(
                io.BytesIO(content), data_only=True, read_only=True)
            rows_text = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    row_str = "  ".join(
                        str(c) for c in row if c is not None
                    )
                    if row_str.strip():
                        rows_text.append(row_str)
            text = "\n".join(rows_text)
            wb.close()

        elif ext == "docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml  = z.read("word/document.xml").decode(
                        "utf-8", errors="ignore")
                    text = re.sub(r"<[^>]+>", " ", xml)
                    text = re.sub(r"\s{2,}", "\n", text)

        elif ext == "pptx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                for name in z.namelist():
                    if name.startswith("ppt/slides/slide"):
                        xml  = z.read(name).decode("utf-8", errors="ignore")
                        text += re.sub(r"<[^>]+>", " ", xml) + "\n"

    except Exception as e:
        return {"grand_total": f"Parse error: {e}",
                "line_items": [], "raw_text": ""}

    grand   = _best_price_near_keyword(text)
    items   = _extract_line_items(text)

    # If no grand total found, sum line items
    if not grand and items:
        total_sum = sum(_parse_num(i["line_total"]) for i in items)
        if total_sum > 0:
            grand = fmt_currency(total_sum) + " (calculated)"

    return {
        "grand_total": grand or "Not found",
        "line_items" : items,
        "raw_text"   : text[:3000],
    }


def fetch_and_extract(url: str) -> dict:
    if not REQUESTS_OK:
        return {"grand_total": "requests not installed",
                "line_items": [], "raw_text": ""}
    if not url or not url.startswith("http"):
        return {"grand_total": "No URL", "line_items": [], "raw_text": ""}
    try:
        resp = requests.get(url, timeout=20)
        if resp.status_code != 200:
            return {"grand_total": f"HTTP {resp.status_code}",
                    "line_items": [], "raw_text": ""}
        ext = url.split("?")[0].rsplit(".", 1)[-1].lower()
        return extract_prices_from_bytes(resp.content, ext)
    except Exception as e:
        return {"grand_total": f"Error: {e}",
                "line_items": [], "raw_text": ""}


# ════════════════════════════════════════════════════════════
# EXTRACT EMBEDDED HYPERLINKS FROM EXCEL
# ════════════════════════════════════════════════════════════
@st.cache_data
def extract_hyperlinks(file_path: str) -> dict:
    link_map = {}
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        fn_col = hr = None
        for row in ws.iter_rows():
            for cell in row:
                if (cell.value and
                        str(cell.value).strip().lower() == "file name"):
                    fn_col = cell.column
                    hr     = cell.row
                    break
            if fn_col:
                break
        if fn_col:
            for row in ws.iter_rows(
                    min_row=hr + 1, min_col=fn_col, max_col=fn_col):
                cell = row[0]
                if cell.value and cell.hyperlink:
                    link_map[str(cell.value).strip()] = \
                        str(cell.hyperlink.target).strip()
        wb.close()
    except Exception as e:
        st.warning(f"⚠️ Hyperlink extraction warning: {e}")
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
    header_row = None
    for i, row in raw.iterrows():
        vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if (any("category" in v for v in vals) and
                any("file" in v for v in vals)):
            header_row = i
            break
    if header_row is None:
        st.error("❌ Could not detect header row.")
        return None, None

    df = pd.read_excel(FILE_PATH, engine="openpyxl", header=header_row)
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

    keep = ["Category", "Vendor", "File Name", "Comments"]
    for e in ["File Link", "File URL", "Quoted Price"]:
        if e in df.columns:
            keep.append(e)
    df = df[[c for c in keep if c in df.columns]].copy()

    df = df[
        ~(df["Category"].astype(str).str.strip().isin(["", "nan"]) &
          df["Vendor"].astype(str).str.strip().isin(["", "nan"]))
    ].copy()

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df.reset_index(drop=True, inplace=True)

    # Embedded hyperlinks
    hmap = extract_hyperlinks(FILE_PATH)
    df["Hyperlink"] = df["File Name"].map(hmap).fillna("")
    for fb in ["File Link", "File URL"]:
        if fb in df.columns:
            df["Hyperlink"] = df.apply(
                lambda r: r["Hyperlink"]
                if r["Hyperlink"] not in ["", "nan"]
                else r[fb], axis=1)

    def parse_svc(raw_val):
        if not raw_val or str(raw_val).strip() in ["", "nan"]:
            return ["(unspecified)"]
        parts = [s.strip() for s in str(raw_val).split("\n") if s.strip()]
        return parts or ["(unspecified)"]

    df["Services List"] = df["Comments"].apply(parse_svc)

    df_exp = df.explode("Services List").copy()
    df_exp.rename(columns={"Services List": "Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].str.strip()
    df_exp = df_exp[
        ~df_exp["Service"].isin(["", "(unspecified)", "nan"])
    ].reset_index(drop=True)

    return df, df_exp


# ════════════════════════════════════════════════════════════
# LOAD DATA
# ════════════════════════════════════════════════════════════
df_master, df_exploded = load_data()
if df_master is None or df_exploded is None:
    st.stop()

# Vendor → PwC colour map
_sorted_vendors = sorted(df_master["Vendor"].unique())
vendor_color_map = {
    v: PWC_CHART[i % len(PWC_CHART)]
    for i, v in enumerate(_sorted_vendors)
}


# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style='text-align:center;padding:22px 0 16px'>
        <div style='font-size:2.2em'>📋</div>
        <div style='font-size:1.1em;font-weight:700;
                    color:white;margin:6px 0 2px;
                    letter-spacing:0.5px'>IT Procurement</div>
        <div style='font-size:0.75em;color:#aaa;
                    letter-spacing:1px;text-transform:uppercase'>
            Service &amp; Vendor Dashboard
        </div>
    </div>
    <hr style='border-color:{PWC["orange"]};
               border-width:2px;margin:0 0 18px'>
    """, unsafe_allow_html=True)

    def sb_label(txt):
        st.markdown(
            f"<p style='color:#F0F0F0;font-weight:700;"
            f"font-size:0.88em;margin:12px 0 4px;"
            f"letter-spacing:0.5px;text-transform:uppercase'>"
            f"{txt}</p>",
            unsafe_allow_html=True)

    sb_label("📂 Category")
    all_cats = ["All"] + sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["", "nan"]
    ])
    selected_cat = st.selectbox(
        "Category", all_cats, label_visibility="collapsed")

    sb_label("🏢 Vendor")
    vpool = (df_master if selected_cat == "All"
             else df_master[df_master["Category"] == selected_cat])
    all_vendors = ["All"] + sorted([
        v for v in vpool["Vendor"].unique()
        if str(v).strip() not in ["", "nan"]
    ])
    selected_vendor = st.selectbox(
        "Vendor", all_vendors, label_visibility="collapsed")

    st.markdown(
        f"<hr style='border-color:#555;margin:16px 0'>",
        unsafe_allow_html=True)

    d_filt = df_exploded.copy()
    if selected_cat    != "All":
        d_filt = d_filt[d_filt["Category"] == selected_cat]
    if selected_vendor != "All":
        d_filt = d_filt[d_filt["Vendor"]   == selected_vendor]

    sb_label("🔍 Search Services")
    svc_search = st.text_input(
        "Search", placeholder="e.g. Cisco, Oracle…",
        label_visibility="collapsed")

    avail = sorted([s for s in d_filt["Service"].unique()
                    if str(s).strip() not in ["", "nan"]])
    if svc_search:
        avail = [s for s in avail if svc_search.lower() in s.lower()]

    sb_label(f"🛠 Select Services ({len(avail)} available)")
    selected_svcs = st.multiselect(
        "Services", options=avail, default=[],
        label_visibility="collapsed")

    st.markdown(
        f"<hr style='border-color:#555;margin:16px 0'>",
        unsafe_allow_html=True)
    st.markdown(
        f"<p style='color:#888;font-size:0.80em;margin:3px 0'>"
        f"📄 {len(df_master)} quotes &nbsp;|&nbsp; "
        f"🛠 {df_exploded['Service'].nunique()} services &nbsp;|&nbsp; "
        f"🏢 {df_master['Vendor'].nunique()} vendors</p>",
        unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# MAIN HEADER — PwC style
# ════════════════════════════════════════════════════════════
st.markdown(f"""
<div style='background:{PWC["dark"]};color:white;
            padding:22px 30px;border-radius:4px;
            border-left:6px solid {PWC["orange"]};
            margin-bottom:24px'>
    <div style='font-size:0.75em;font-weight:700;
                letter-spacing:2px;text-transform:uppercase;
                color:{PWC["orange"]};margin-bottom:6px'>
        IT Procurement Analytics
    </div>
    <h1 style='margin:0;font-size:1.6em;font-weight:700;
               color:white;letter-spacing:-0.3px'>
        Service &amp; Vendor Dashboard
    </h1>
    <p style='margin:7px 0 0;opacity:0.65;font-size:0.88em;
              font-weight:300'>
        Filter by Category → Vendor auto-updates →
        Select a service → View vendor, quotation &amp; price breakdown
    </p>
</div>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# KPI CARDS
# ════════════════════════════════════════════════════════════
k1, k2, k3, k4 = st.columns(4)

def kpi(col, val, lbl, bg):
    col.markdown(
        f"<div class='kpi-box' style='background:{bg}'>"
        f"<div class='kpi-value'>{val}</div>"
        f"<div class='kpi-label'>{lbl}</div></div>",
        unsafe_allow_html=True)

kpi(k1, d_filt["File Name"].nunique(),  "Total Quotes",    PWC["orange"])
kpi(k2, d_filt["Service"].nunique(),    "Unique Services", PWC["blue"])
kpi(k3, d_filt["Vendor"].nunique(),     "Vendors",         PWC["teal"])
kpi(k4, d_filt["Category"].nunique(),   "Categories",      PWC["dark"])

st.markdown("<br>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# CHART LAYOUT
# Purpose-driven 3-panel layout:
#   Row 1 LEFT  → "Who has overlap?" (shared services bar)
#   Row 1 RIGHT → "How many services per vendor?" (bar)
#   Row 2 FULL  → "How is spend distributed?" (category donut)
# ════════════════════════════════════════════════════════════
tab_charts, tab_table = st.tabs(["📊 Analytics", "📄 Data Table"])

CHART_BG   = PWC["light"]
CHART_FONT = dict(family="Source Sans Pro, Helvetica Neue, Arial",
                  size=11, color=PWC["dark"])

with tab_charts:

    # ── Row 1 ─────────────────────────────────────────────────
    cl, cr = st.columns([1, 1], gap="large")

    # LEFT — Services shared across vendors
    with cl:
        st.markdown(
            f"<div style='font-size:0.78em;font-weight:700;"
            f"letter-spacing:1px;text-transform:uppercase;"
            f"color:{PWC['orange']};margin-bottom:4px'>"
            f"SERVICE OVERLAP ANALYSIS</div>",
            unsafe_allow_html=True)
        st.caption(
            "Red bars = same service quoted by multiple vendors. "
            "Use to identify competitive procurement opportunities.")

        shared = (
            d_filt.groupby("Service")["Vendor"].nunique()
            .sort_values(ascending=False).head(20).reset_index()
        )
        shared.columns = ["Service", "Vendor Count"]
        shared["Colour"] = shared["Vendor Count"].apply(
            lambda x: PWC["orange"] if x > 1 else "#C0C0C0")

        fig1 = go.Figure(go.Bar(
            x=shared["Vendor Count"],
            y=shared["Service"].str[:46],
            orientation="h",
            marker_color=shared["Colour"],
            marker_line_width=0,
            text=shared["Vendor Count"],
            textposition="outside",
            textfont=dict(size=10),
        ))
        fig1.update_layout(
            height=500,
            plot_bgcolor=CHART_BG,
            paper_bgcolor=CHART_BG,
            margin=dict(l=5, r=40, t=20, b=10),
            font=CHART_FONT,
            xaxis=dict(
                title="Number of Vendors",
                showgrid=True,
                gridcolor="#E0E0E0",
                zeroline=False,
            ),
            yaxis=dict(autorange="reversed", tickfont=dict(size=9.5)),
            bargap=0.35,
        )
        st.plotly_chart(fig1, use_container_width=True)

    # RIGHT — Services per vendor
    with cr:
        st.markdown(
            f"<div style='font-size:0.78em;font-weight:700;"
            f"letter-spacing:1px;text-transform:uppercase;"
            f"color:{PWC['orange']};margin-bottom:4px'>"
            f"VENDOR SERVICE COVERAGE</div>",
            unsafe_allow_html=True)
        st.caption(
            "Number of unique services each vendor provides. "
            "Higher = broader vendor capability.")

        spv = (
            d_filt.groupby("Vendor")["Service"].nunique()
            .sort_values(ascending=False).reset_index()
        )
        spv.columns = ["Vendor", "Count"]
        spv["Colour"] = [vendor_color_map.get(v, PWC["mid"])
                         for v in spv["Vendor"]]

        fig2 = go.Figure(go.Bar(
            x=spv["Vendor"],
            y=spv["Count"],
            marker_color=spv["Colour"],
            marker_line_width=0,
            text=spv["Count"],
            textposition="outside",
            textfont=dict(size=10),
        ))
        fig2.update_layout(
            height=500,
            plot_bgcolor=CHART_BG,
            paper_bgcolor=CHART_BG,
            margin=dict(l=5, r=10, t=20, b=10),
            font=CHART_FONT,
            yaxis=dict(
                title="Unique Services",
                showgrid=True,
                gridcolor="#E0E0E0",
                zeroline=False,
            ),
            xaxis=dict(tickangle=-35, tickfont=dict(size=9.5)),
            bargap=0.35,
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── Row 2 — Category distribution ─────────────────────────
    st.markdown(
        f"<div style='font-size:0.78em;font-weight:700;"
        f"letter-spacing:1px;text-transform:uppercase;"
        f"color:{PWC['orange']};margin-bottom:4px'>"
        f"PROCUREMENT CATEGORY DISTRIBUTION</div>",
        unsafe_allow_html=True)
    st.caption(
        "Share of total quote files per IT procurement category.")

    cat_counts = (
        d_filt.drop_duplicates(subset=["Category", "File Name"])
        .groupby("Category").size().reset_index()
    )
    cat_counts.columns = ["Category", "Count"]

    if not cat_counts.empty:
        fig3 = px.pie(
            cat_counts,
            names="Category", values="Count",
            hole=0.50,
            color_discrete_sequence=PWC_CHART,
        )
        fig3.update_traces(
            textposition="outside",
            textinfo="label+percent",
            textfont_size=11,
            pull=[0.03] * len(cat_counts),
        )
        fig3.update_layout(
            height=380,
            margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor=CHART_BG,
            font=CHART_FONT,
            showlegend=True,
            legend=dict(
                orientation="v",
                x=1.02, y=0.5,
                font=dict(size=10),
            ),
        )
        st.plotly_chart(fig3, use_container_width=True)

# ── Data table tab ────────────────────────────────────────────
with tab_table:
    dm = df_master.copy()
    if selected_cat    != "All":
        dm = dm[dm["Category"] == selected_cat]
    if selected_vendor != "All":
        dm = dm[dm["Vendor"]   == selected_vendor]
    st.dataframe(
        dm.drop(columns=["Services List", "Hyperlink"], errors="ignore"),
        use_container_width=True, height=460)


# ════════════════════════════════════════════════════════════
# SERVICE SELECTION RESULTS
# ════════════════════════════════════════════════════════════
st.markdown(
    f"<hr style='border:none;border-top:2px solid {PWC['orange']};"
    f"margin:24px 0 16px'>",
    unsafe_allow_html=True)

st.markdown(
    f"<div style='font-size:0.78em;font-weight:700;"
    f"letter-spacing:1px;text-transform:uppercase;"
    f"color:{PWC['orange']};margin-bottom:6px'>"
    f"SERVICE SELECTION &amp; QUOTATION ANALYSIS</div>",
    unsafe_allow_html=True)
st.markdown("### Select services from the sidebar to begin analysis")

if not selected_svcs:
    st.info(
        "👈 **Select one or more services** from the sidebar. "
        "Each selection shows the vendors, all quoted line items, "
        "unit prices, totals and a link to the quotation file.")
else:
    d_sel = d_filt[d_filt["Service"].isin(selected_svcs)].copy()

    if d_sel.empty:
        st.warning("⚠️ No results found under current filters.")
    else:
        # Vendor coverage
        vsmap = defaultdict(set)
        for _, r in d_sel.iterrows():
            vsmap[r["Vendor"]].add(r["Service"])

        vendors_all  = sorted([v for v, s in vsmap.items()
                                if set(selected_svcs).issubset(s)])
        vendors_some = sorted([v for v, s in vsmap.items()
                                if not set(selected_svcs).issubset(s)])

        # Summary banners
        if len(selected_svcs) > 1:
            if vendors_all:
                names = " · ".join([f"**{v}**" for v in vendors_all])
                st.success(
                    f"✅ **{len(vendors_all)} vendor(s) offer ALL "
                    f"{len(selected_svcs)} services:** {names}")
            else:
                st.warning(
                    f"⚠️ No single vendor covers all "
                    f"{len(selected_svcs)} selected services.")
            if vendors_some:
                with st.expander(
                        "🔵 Vendors with partial coverage",
                        expanded=False):
                    for v in vendors_some:
                        cov = vsmap[v].intersection(set(selected_svcs))
                        c   = vendor_color_map.get(v, PWC["mid"])
                        st.markdown(
                            f"<span class='vendor-badge' "
                            f"style='background:{c}'>{v}</span>"
                            f" &nbsp; covers **{len(cov)}/"
                            f"{len(selected_svcs)}**: "
                            f"_{', '.join(sorted(cov))}_",
                            unsafe_allow_html=True)

        # ── Per-service detail ─────────────────────────────────
        st.markdown(
            f"<div style='font-size:0.78em;font-weight:700;"
            f"letter-spacing:1px;text-transform:uppercase;"
            f"color:{PWC['orange']};margin:18px 0 8px'>"
            f"DETAILED QUOTATION BREAKDOWN — PER SERVICE</div>",
            unsafe_allow_html=True)

        has_price = "Quoted Price" in d_sel.columns

        for svc in selected_svcs:
            d_svc = (
                d_sel[d_sel["Service"] == svc]
                .drop_duplicates(subset=["Vendor", "File Name"])
                .sort_values("Vendor")
            )
            vc    = d_svc["Vendor"].nunique()
            s_tag = ("⚠️ SHARED" if vc > 1 else "✅ SINGLE VENDOR")

            with st.expander(
                f"🛠 {svc}  ·  {vc} vendor(s)  ·  "
                f"{len(d_svc)} file(s)  [{s_tag}]",
                expanded=True,
            ):
                # Vendor pill row
                pills = " ".join([
                    f"<span class='vendor-badge' "
                    f"style='background:"
                    f"{vendor_color_map.get(v, PWC[\"mid\"])}'>"
                    f"{v}</span>"
                    for v in sorted(d_svc["Vendor"].unique())
                ])
                st.markdown(
                    f"<div style='margin-bottom:14px'>"
                    f"<span style='font-weight:700;font-size:0.88em'>"
                    f"Vendors offering this service:</span>"
                    f"&nbsp;&nbsp;{pills}</div>",
                    unsafe_allow_html=True)

                # ── Build extensive price table ────────────────
                # Columns: Vendor | Category | File Name |
                #          Line Item / Description | Unit Price |
                #          Line Total | File Link
                # Also: a GRAND TOTAL row per file + overall total

                tbl  = ["<table class='price-table'>"]
                tbl += ["<thead><tr>"
                        "<th>Vendor</th>"
                        "<th>Category</th>"
                        "<th>📄 File Name</th>"
                        "<th>🔖 Line Item / Description</th>"
                        "<th>Unit Price</th>"
                        "<th>Line Total</th>"
                        "<th>🔗 Link</th>"
                        "</tr></thead><tbody>"]

                overall_total = 0.0
                has_any_price = False

                for i, (_, row) in enumerate(d_svc.iterrows()):
                    bg  = "#ffffff" if i % 2 == 0 else PWC["light"]
                    vc  = vendor_color_map.get(row["Vendor"], PWC["mid"])

                    fname = str(row.get("File Name", "")).strip()
                    url   = str(row.get("Hyperlink", "")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File Link", "")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File URL",  "")).strip()

                    # Vendor badge
                    v_cell = (
                        f"<span class='vendor-badge' "
                        f"style='background:{vc}'>"
                        f"{row['Vendor']}</span>"
                    )

                    # File name cell
                    fn_cell = (
                        f"<span style='font-family:monospace;"
                        f"font-size:0.80em;color:{PWC['dark']};"
                        f"word-break:break-all'>{fname}</span>"
                    )

                    # Link cell
                    if url and url.startswith("http"):
                        lbl   = fname[:35] + "…" if len(fname) > 35 else fname
                        l_cell = (
                            f"<a href='{url}' target='_blank' "
                            f"style='color:{PWC['orange']};"
                            f"font-weight:600;font-size:0.82em;"
                            f"text-decoration:none'>"
                            f"↗ Open</a>"
                        )
                    else:
                        l_cell = (
                            f"<span style='color:#bbb;"
                            f"font-size:0.80em'>—</span>"
                        )

                    # Quoted price from master sheet
                    q_price = str(row.get("Quoted Price", "")).strip()
                    q_num   = _parse_num(q_price) if q_price not in \
                        ["", "nan", "0"] else 0.0

                    # Try to parse line items from the file
                    # (uses session_state cache to avoid re-fetch)
                    cache_key = f"price_{fname}"
                    if cache_key not in st.session_state:
                        st.session_state[cache_key] = None

                    parsed = st.session_state.get(cache_key)

                    if parsed and parsed.get("line_items"):
                        items = parsed["line_items"]
                        gtotal = parsed.get("grand_total", "")

                        # One row per line item
                        for j, item in enumerate(items):
                            ibg  = bg
                            desc = item.get("description", "")
                            up   = item.get("unit_price", "—")
                            lt   = item.get("line_total", "—")
                            lt_n = _parse_num(lt)
                            has_any_price = True

                            if j == 0:
                                tbl.append(
                                    f"<tr style='background:{ibg}'>"
                                    f"<td>{v_cell}</td>"
                                    f"<td style='color:#555'>"
                                    f"{row['Category']}</td>"
                                    f"<td>{fn_cell}</td>"
                                    f"<td style='color:{PWC['dark']}'>"
                                    f"{desc}</td>"
                                    f"<td style='text-align:right;"
                                    f"font-family:monospace'>{up}</td>"
                                    f"<td style='text-align:right;"
                                    f"font-family:monospace;"
                                    f"font-weight:600'>{lt}</td>"
                                    f"<td>{l_cell}</td></tr>"
                                )
                            else:
                                tbl.append(
                                    f"<tr style='background:{ibg}'>"
                                    f"<td></td><td></td><td></td>"
                                    f"<td style='color:{PWC['dark']}'>"
                                    f"{desc}</td>"
                                    f"<td style='text-align:right;"
                                    f"font-family:monospace'>{up}</td>"
                                    f"<td style='text-align:right;"
                                    f"font-family:monospace;"
                                    f"font-weight:600'>{lt}</td>"
                                    f"<td></td></tr>"
                                )

                        # Grand total row for this file
                        file_total = (_parse_num(gtotal)
                                      if gtotal and
                                      "Not found" not in gtotal
                                      else sum(
                                          _parse_num(it["line_total"])
                                          for it in items))
                        overall_total += file_total
                        tbl.append(
                            f"<tr class='total-row'>"
                            f"<td colspan='5' style='text-align:right;"
                            f"text-transform:uppercase;"
                            f"letter-spacing:0.5px'>"
                            f"File Grand Total</td>"
                            f"<td style='text-align:right;"
                            f"font-family:monospace'>"
                            f"{fmt_currency(file_total)}</td>"
                            f"<td></td></tr>"
                        )

                    else:
                        # No parsed data yet — show master quoted price
                        qp_display = (
                            f"<span style='color:{PWC['green']};"
                            f"font-weight:700;font-family:monospace'>"
                            f"{fmt_currency(q_price)}</span>"
                            if q_num > 0
                            else f"<span style='color:#bbb'>"
                                 f"Click Extract below</span>"
                        )
                        if q_num > 0:
                            overall_total += q_num
                            has_any_price  = True

                        tbl.append(
                            f"<tr style='background:{bg}'>"
                            f"<td>{v_cell}</td>"
                            f"<td style='color:#555'>"
                            f"{row['Category']}</td>"
                            f"<td>{fn_cell}</td>"
                            f"<td style='color:#999;font-style:italic'>"
                            f"Not yet extracted</td>"
                            f"<td>—</td>"
                            f"<td style='text-align:right'>"
                            f"{qp_display}</td>"
                            f"<td>{l_cell}</td></tr>"
                        )

                # ── Overall total row ──────────────────────────
                if has_any_price and overall_total > 0:
                    tbl.append(
                        f"<tr style='background:{PWC['orange']};"
                        f"color:white'>"
                        f"<td colspan='5' style='text-align:right;"
                        f"font-weight:800;font-size:1em;"
                        f"letter-spacing:0.5px;"
                        f"text-transform:uppercase;color:white'>"
                        f"TOTAL VALUE (ALL FILES IN THIS SERVICE)"
                        f"</td>"
                        f"<td style='text-align:right;"
                        f"font-family:monospace;font-weight:800;"
                        f"font-size:1.05em;color:white'>"
                        f"{fmt_currency(overall_total)}</td>"
                        f"<td style='color:white'></td></tr>"
                    )

                tbl.append("</tbody></table>")
                st.markdown("".join(tbl), unsafe_allow_html=True)

                # ── Extract Price button ───────────────────────
                st.markdown("<br>", unsafe_allow_html=True)
                btn_cols = st.columns([2, 3])
                with btn_cols[0]:
                    if st.button(
                        f"💰 Extract Prices from Files — {svc[:40]}",
                        key=f"btn_{svc[:40]}",
                        type="primary",
                        use_container_width=True,
                    ):
                        prog = st.progress(0, text="Extracting…")
                        total_rows = len(d_svc)
                        for k, (_, row) in enumerate(d_svc.iterrows()):
                            fname     = str(row.get("File Name","")).strip()
                            url       = str(row.get("Hyperlink","")).strip()
                            cache_key = f"price_{fname}"
                            if (url and url.startswith("http") and
                                    st.session_state.get(cache_key) is None):
                                result = fetch_and_extract(url)
                                st.session_state[cache_key] = result
                            prog.progress(
                                (k + 1) / total_rows,
                                text=f"Processed {k+1}/{total_rows}")
                        prog.empty()
                        st.rerun()

        # Shared services summary
        shared_svcs = [
            s for s in selected_svcs
            if d_sel[d_sel["Service"] == s]["Vendor"].nunique() > 1
        ]
        if shared_svcs:
            with st.expander(
                "🔁 Shared Services — same service offered by "
                "multiple vendors (competitive opportunity)",
                expanded=False,
            ):
                for s in shared_svcs:
                    vlist  = sorted(
                        d_sel[d_sel["Service"] == s]["Vendor"].unique())
                    pills  = " ".join([
                        f"<span class='vendor-badge' "
                        f"style='background:"
                        f"{vendor_color_map.get(v, PWC[\"mid\"])}'>"
                        f"{v}</span>"
                        for v in vlist
                    ])
                    st.markdown(
                        f"<div style='margin-bottom:10px'>"
                        f"<span style='font-weight:700'>{s}</span>"
                        f" &nbsp;→&nbsp; {pills}</div>",
                        unsafe_allow_html=True)
