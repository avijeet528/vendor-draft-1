# ============================================================
#  app.py — IT Procurement Service Dashboard (Streamlit)
#  - Heatmap removed
#  - File Link extracted from Excel hyperlinks embedded
#    in the File Name column using openpyxl directly
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import defaultdict
import openpyxl
import os

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
    footer     {visibility: hidden;}
    header     {visibility: hidden;}
    [data-testid="collapsedControl"] {
        display: none !important;
    }
    /* ── FIX: Remove sidebar collapse/expand arrow button ── */
    button[data-testid="collapsedControl"],
    button[kind="header"][data-testid="baseButton-header"] {
        display: none !important;
    }
    section[data-testid="stSidebar"] {
        min-width: 350px !important;
        max-width: 350px !important;
        transform: none !important;
        visibility: visible !important;
    }
    section[data-testid="stSidebar"] > div {
        width: 350px !important;
    }

    /* KPI cards */
    .kpi-box   { border-radius:12px; padding:18px 10px;
                 text-align:center; color:white; }
    .kpi-value { font-size:2.2em; font-weight:800;
                 margin:0; line-height:1.1; }
    .kpi-label { font-size:0.85em; opacity:0.88; margin-top:5px; }

    /* Sidebar background */
    section[data-testid="stSidebar"] {
        background-color: #2c3e50 !important;
    }

    /* All sidebar text → light */
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] div {
        color: #ecf0f1 !important;
    }

    /* Dropdown / select box input area → white bg, dark text */
    section[data-testid="stSidebar"] div[data-baseweb="select"] {
        background-color: white !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"]
        div[data-baseweb="select"] * {
        color: #2c3e50 !important;
    }

    /* Text input area → white bg, dark text */
    section[data-testid="stSidebar"] div[data-baseweb="input"] {
        background-color: white !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"]
        div[data-baseweb="input"] input {
        color: #2c3e50 !important;
    }

    /* Multiselect tags */
    section[data-testid="stSidebar"]
        span[data-baseweb="tag"] {
        background-color: #2980b9 !important;
    }
    section[data-testid="stSidebar"]
        span[data-baseweb="tag"] span {
        color: white !important;
    }

    /* Caption / small text */
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] .stCaptionContainer p {
        color: #95a5a6 !important;
    }

    /* Expander */
    div[data-testid="stExpander"] details summary p {
        font-size:1em; font-weight:600;
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
# EXTRACT EMBEDDED HYPERLINKS FROM EXCEL FILE NAME COLUMN
# openpyxl reads the actual .hyperlink attribute per cell
# ════════════════════════════════════════════════════════════
@st.cache_data
def extract_hyperlinks(file_path: str) -> dict:
    """
    Opens the workbook with openpyxl and scans every cell
    in the 'File Name' column for embedded hyperlinks.
    Returns a dict: { display_text -> hyperlink_url }
    """
    link_map = {}
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        # Find the header row and the File Name column index
        file_name_col = None
        header_row_idx = None

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and str(cell.value).strip().lower() == "file name":
                    file_name_col  = cell.column        # 1-based int
                    header_row_idx = cell.row
                    break
            if file_name_col:
                break

        if file_name_col is None:
            return link_map

        # Scan every data cell in that column below the header
        for row in ws.iter_rows(
            min_row=header_row_idx + 1,
            min_col=file_name_col,
            max_col=file_name_col,
        ):
            cell = row[0]
            if cell.value and cell.hyperlink:
                display = str(cell.value).strip()
                url     = str(cell.hyperlink.target).strip()
                if url:
                    link_map[display] = url

        wb.close()
    except Exception as e:
        st.warning(f"⚠️ Could not extract hyperlinks: {e}")

    return link_map


# ════════════════════════════════════════════════════════════
# DATA LOADING
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    FILE_PATH = "Master Catalog.xlsx"

    if not os.path.exists(FILE_PATH):
        st.error(f"❌ File not found: '{FILE_PATH}'.")
        return None, None

    # ── Detect header row ────────────────────────────────────
    raw = pd.read_excel(FILE_PATH, engine="openpyxl", header=None)
    header_row = None
    for i, row in raw.iterrows():
        row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if (any("category" in v for v in row_vals) and
                any("file" in v for v in row_vals)):
            header_row = i
            break

    if header_row is None:
        st.error("❌ Could not detect header row.")
        return None, None

    # ── Read with correct header ─────────────────────────────
    df = pd.read_excel(FILE_PATH, engine="openpyxl", header=header_row)
    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)

    # ── Map columns ──────────────────────────────────────────
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category":
            col_map["Category"]      = c
        elif "vendor" in cl or "type" in cl:
            col_map["Vendor"]        = c
        elif cl == "file name":
            col_map["File Name"]     = c
        elif cl == "file link":
            col_map["File Link"]     = c
        elif cl == "file url":
            col_map["File URL"]      = c
        elif "comment" in cl:
            col_map["Comments"]      = c
        elif "quoted" in cl or "price" in cl:
            col_map["Quoted Price"]  = c

    df.rename(columns={v: k for k, v in col_map.items()}, inplace=True)

    keep = ["Category", "Vendor", "File Name", "Comments"]
    for extra in ["File Link", "File URL", "Quoted Price"]:
        if extra in df.columns:
            keep.append(extra)
    df = df[[c for c in keep if c in df.columns]].copy()

    # Drop fully empty rows
    df = df[
        ~(
            df["Category"].astype(str).str.strip().isin(["", "nan"]) &
            df["Vendor"].astype(str).str.strip().isin(["", "nan"])
        )
    ].copy()

    # Safe fillna per column
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df.reset_index(drop=True, inplace=True)

    # ── Extract embedded hyperlinks from File Name column ────
    hyperlink_map = extract_hyperlinks(FILE_PATH)

    # Add Hyperlink column using the embedded links
    df["Hyperlink"] = df["File Name"].map(hyperlink_map).fillna("")

    # If a separate File Link / File URL column exists, use it as fallback
    for fallback_col in ["File Link", "File URL"]:
        if fallback_col in df.columns:
            df["Hyperlink"] = df.apply(
                lambda r: r["Hyperlink"]
                if r["Hyperlink"] not in ["", "nan"]
                else r[fallback_col],
                axis=1,
            )

    # ── Parse services ───────────────────────────────────────
    def parse_services(raw_val):
        if not raw_val or str(raw_val).strip() in ["", "nan"]:
            return ["(unspecified)"]
        parts = [s.strip() for s in str(raw_val).split("\n") if s.strip()]
        return parts if parts else ["(unspecified)"]

    df["Services List"] = df["Comments"].apply(parse_services)

    # ── Explode: one row per service ─────────────────────────
    df_exp = df.explode("Services List").copy()
    df_exp.rename(columns={"Services List": "Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].str.strip()
    df_exp = df_exp[
        (df_exp["Service"] != "") &
        (df_exp["Service"] != "(unspecified)") &
        (df_exp["Service"] != "nan")
    ].reset_index(drop=True)

    return df, df_exp


# ════════════════════════════════════════════════════════════
# LOAD
# ════════════════════════════════════════════════════════════
df_master, df_exploded = load_data()
if df_master is None or df_exploded is None:
    st.stop()

vendor_color_map = {
    v: get_color(i)
    for i, v in enumerate(sorted(df_master["Vendor"].unique()))
}

# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
with st.sidebar:

    st.markdown("""
    <div style='text-align:center;padding:20px 0 14px'>
        <div style='font-size:2.5em'>📋</div>
        <div style='font-size:1.15em;font-weight:800;
                    color:white;margin:6px 0 2px'>
            IT Procurement
        </div>
        <div style='font-size:0.78em;color:#95a5a6'>
            Service &amp; Vendor Dashboard
        </div>
    </div>
    <hr style='border-color:#3d5166;margin:0 0 16px'>
    """, unsafe_allow_html=True)

    # Category
    st.markdown(
        "<p style='color:#ecf0f1;font-weight:700;margin-bottom:4px'>"
        "📂 Category</p>",
        unsafe_allow_html=True,
    )
    all_cats = ["All"] + sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["", "nan"]
    ])
    selected_cat = st.selectbox(
        "Category", all_cats, label_visibility="collapsed"
    )

    # Vendor — scoped to category
    st.markdown(
        "<p style='color:#ecf0f1;font-weight:700;margin:10px 0 4px'>"
        "🏢 Vendor</p>",
        unsafe_allow_html=True,
    )
    vendor_pool = (
        df_master if selected_cat == "All"
        else df_master[df_master["Category"] == selected_cat]
    )
    all_vendors = ["All"] + sorted([
        v for v in vendor_pool["Vendor"].unique()
        if str(v).strip() not in ["", "nan"]
    ])
    selected_vendor = st.selectbox(
        "Vendor", all_vendors, label_visibility="collapsed"
    )

    st.markdown(
        "<hr style='border-color:#3d5166;margin:14px 0'>",
        unsafe_allow_html=True,
    )

    # Filtered exploded df
    d_filt = df_exploded.copy()
    if selected_cat != "All":
        d_filt = d_filt[d_filt["Category"] == selected_cat]
    if selected_vendor != "All":
        d_filt = d_filt[d_filt["Vendor"] == selected_vendor]

    # Service search
    st.markdown(
        "<p style='color:#ecf0f1;font-weight:700;margin-bottom:4px'>"
        "🔍 Search Services</p>",
        unsafe_allow_html=True,
    )
    svc_search = st.text_input(
        "Search", placeholder="e.g. Cisco, Oracle, M365…",
        label_visibility="collapsed",
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

    st.markdown(
        f"<p style='color:#ecf0f1;font-weight:700;margin:10px 0 4px'>"
        f"🛠 Select Services "
        f"<span style='font-weight:400;color:#95a5a6'>"
        f"({len(available_svcs)} available)</span></p>",
        unsafe_allow_html=True,
    )
    selected_svcs = st.multiselect(
        "Services", options=available_svcs, default=[],
        label_visibility="collapsed",
        help="Select one or more services to see vendor & file link",
    )

    st.markdown(
        "<hr style='border-color:#3d5166;margin:14px 0'>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<p style='color:#95a5a6;font-size:0.82em;margin:2px 0'>"
        f"📄 {len(df_master)} total quotes</p>"
        f"<p style='color:#95a5a6;font-size:0.82em;margin:2px 0'>"
        f"🛠 {df_exploded['Service'].nunique()} unique services</p>"
        f"<p style='color:#95a5a6;font-size:0.82em;margin:2px 0'>"
        f"🏢 {df_master['Vendor'].nunique()} vendors</p>",
        unsafe_allow_html=True,
    )

# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
st.markdown("""
<div style='background:linear-gradient(135deg,#2c3e50 0%,#3d5a73 100%);
            color:white;padding:20px 28px;border-radius:14px;
            margin-bottom:22px;
            box-shadow:0 4px 12px rgba(0,0,0,0.15)'>
    <h1 style='margin:0;font-size:1.65em;font-weight:800'>
        📋 IT Procurement — Service &amp; Vendor Dashboard
    </h1>
    <p style='margin:7px 0 0;opacity:0.72;font-size:0.9em'>
        Filter by Category → Vendor auto-updates →
        Select a service → See vendor &amp; quotation file instantly
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
# TABS
# ════════════════════════════════════════════════════════════
tab_charts, tab_table = st.tabs(["📊 Charts", "📄 Data Table"])

# ────────────────────────────────────────────────────────────
with tab_charts:

    col_l, col_r = st.columns(2)

    # Chart 1 — Shared services bar
    with col_l:
        shared = (
            d_filt.groupby("Service")["Vendor"].nunique()
            .sort_values(ascending=False).head(20).reset_index()
        )
        shared.columns = ["Service", "Vendor Count"]
        shared["Color"] = shared["Vendor Count"].apply(
            lambda x: "#e74c3c" if x > 1 else "#bdc3c7"
        )
        fig1 = go.Figure(go.Bar(
            x=shared["Vendor Count"],
            y=shared["Service"].str[:44],
            orientation="h",
            marker_color=shared["Color"],
            text=shared["Vendor Count"],
            textposition="outside",
        ))
        fig1.update_layout(
            title="🔁 Services Shared Across Vendors (Top 20)",
            xaxis_title="# Vendors",
            height=540,
            plot_bgcolor="#f8f9fa",
            paper_bgcolor="#f8f9fa",
            margin=dict(l=10, r=30, t=45, b=10),
            font=dict(size=11),
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig1, use_container_width=True)

    # Chart 2 — Services per vendor bar
    with col_r:
        svc_per_vendor = (
            d_filt.groupby("Vendor")["Service"].nunique()
            .sort_values(ascending=False).reset_index()
        )
        svc_per_vendor.columns = ["Vendor", "Service Count"]
        svc_per_vendor["Color"] = [
            vendor_color_map.get(v, "#999")
            for v in svc_per_vendor["Vendor"]
        ]
        fig2 = go.Figure(go.Bar(
            x=svc_per_vendor["Vendor"],
            y=svc_per_vendor["Service Count"],
            marker_color=svc_per_vendor["Color"],
            text=svc_per_vendor["Service Count"],
            textposition="outside",
        ))
        fig2.update_layout(
            title="📦 Unique Services per Vendor",
            yaxis_title="# Unique Services",
            height=540,
            plot_bgcolor="#f8f9fa",
            paper_bgcolor="#f8f9fa",
            margin=dict(l=10, r=10, t=45, b=10),
            font=dict(size=11),
            xaxis_tickangle=-30,
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Chart 3 — Category pie (full width)
    cat_counts = (
        d_filt.drop_duplicates(subset=["Category", "File Name"])
        .groupby("Category").size().reset_index()
    )
    cat_counts.columns = ["Category", "Count"]

    if not cat_counts.empty:
        fig3 = px.pie(
            cat_counts,
            names="Category",
            values="Count",
            title="🥧 Quote Files by Category",
            hole=0.45,
            color_discrete_sequence=COLORS,
        )
        fig3.update_layout(
            height=480,
            margin=dict(l=10, r=10, t=45, b=10),
            paper_bgcolor="#f8f9fa",
            font=dict(size=12),
            legend=dict(font=dict(size=11)),
        )
        fig3.update_traces(
            textposition="inside",
            textinfo="percent+label",
        )
        st.plotly_chart(fig3, use_container_width=True)

# ────────────────────────────────────────────────────────────
with tab_table:
    dm_display = df_master.copy()
    if selected_cat != "All":
        dm_display = dm_display[dm_display["Category"] == selected_cat]
    if selected_vendor != "All":
        dm_display = dm_display[dm_display["Vendor"] == selected_vendor]
    st.dataframe(
        dm_display.drop(columns=["Services List","Hyperlink"],
                        errors="ignore"),
        use_container_width=True,
        height=450,
    )

# ════════════════════════════════════════════════════════════
# SERVICE SELECTION RESULTS
# ════════════════════════════════════════════════════════════
st.markdown(
    "<hr style='border-color:#ddd;margin:14px 0'>",
    unsafe_allow_html=True,
)
st.markdown("### 🛠 Service → Vendor & Quotation File Results")

if not selected_svcs:
    st.info(
        "👈 **Select one or more services** from the sidebar "
        "to see which vendor offers them and the quotation file link."
    )
else:
    d_sel = d_filt[d_filt["Service"].isin(selected_svcs)].copy()

    if d_sel.empty:
        st.warning("⚠️ No results found for selected service(s).")
    else:
        # Vendor coverage map
        vendor_svc_map = defaultdict(set)
        for _, row in d_sel.iterrows():
            vendor_svc_map[row["Vendor"]].add(row["Service"])

        vendors_all  = sorted([
            v for v, s in vendor_svc_map.items()
            if set(selected_svcs).issubset(s)
        ])
        vendors_some = sorted([
            v for v, s in vendor_svc_map.items()
            if not set(selected_svcs).issubset(s)
        ])

        # Summary banners (only meaningful for multi-select)
        if len(selected_svcs) > 1:
            if vendors_all:
                names = " · ".join([f"**{v}**" for v in vendors_all])
                st.success(
                    f"✅ **{len(vendors_all)} vendor(s) offer ALL "
                    f"{len(selected_svcs)} services:** {names}"
                )
            else:
                st.warning(
                    f"⚠️ No single vendor covers all "
                    f"{len(selected_svcs)} selected services."
                )
            if vendors_some:
                with st.expander(
                    "🔵 Vendors with partial coverage", expanded=False
                ):
                    for v in vendors_some:
                        covered = vendor_svc_map[v].intersection(
                            set(selected_svcs)
                        )
                        c = vendor_color_map.get(v, "#666")
                        st.markdown(
                            f"<span style='background:{c};color:white;"
                            f"padding:3px 10px;border-radius:10px;"
                            f"font-size:0.88em;font-weight:bold'>{v}</span>"
                            f" &nbsp; covers **{len(covered)}/"
                            f"{len(selected_svcs)}**: "
                            f"_{', '.join(sorted(covered))}_",
                            unsafe_allow_html=True,
                        )

        # ── Per-service expandable tables ─────────────────────
        st.markdown("#### 📄 Vendor & Quotation File — per Service")

        # Table styles
        TH = ("padding:8px 14px;text-align:left;"
              "border-bottom:2px solid #ddd;"
              "background:#f4f6f7;font-size:0.88em;color:#2c3e50")
        TD = ("padding:8px 14px;border-bottom:1px solid #eee;"
              "font-size:0.85em;vertical-align:middle")

        has_price = "Quoted Price" in d_sel.columns

        for svc in selected_svcs:
            d_svc = (
                d_sel[d_sel["Service"] == svc]
                .drop_duplicates(subset=["Vendor", "File Name"])
                .sort_values("Vendor")
            )
            vendor_count = d_svc["Vendor"].nunique()
            shared_tag   = (
                "⚠️ SHARED BY MULTIPLE VENDORS"
                if vendor_count > 1 else "✅ SINGLE VENDOR"
            )

            with st.expander(
                f"🛠 {svc}  |  {vendor_count} vendor(s) · "
                f"{len(d_svc)} file(s)  [{shared_tag}]",
                expanded=True,
            ):
                # Vendor badges row
                badges = " ".join([
                    f"<span style='background:{vendor_color_map.get(v,'#666')};"
                    f"color:white;padding:4px 13px;border-radius:12px;"
                    f"font-size:0.88em;font-weight:bold;margin:2px'>{v}</span>"
                    for v in sorted(d_svc["Vendor"].unique())
                ])
                st.markdown(
                    f"<div style='margin-bottom:12px'>"
                    f"<b>Vendors offering this service:</b>"
                    f"&nbsp;{badges}</div>",
                    unsafe_allow_html=True,
                )

                # Build HTML table
                rows_html = [
                    f"<table style='width:100%;border-collapse:collapse;"
                    f"border-radius:8px;overflow:hidden'><thead><tr>"
                    f"<th style='{TH}'>Vendor</th>"
                    f"<th style='{TH}'>Category</th>"
                    f"<th style='{TH}'>📄 File Name</th>"
                ]
                if has_price:
                    rows_html.append(f"<th style='{TH}'>💰 Quoted Price</th>")
                rows_html.append(
                    f"<th style='{TH}'>🔗 File Link</th>"
                    f"</tr></thead><tbody>"
                )

                for i, (_, row) in enumerate(d_svc.iterrows()):
                    bg  = "#ffffff" if i % 2 == 0 else "#f9f9f9"
                    vc  = vendor_color_map.get(row["Vendor"], "#666")

                    # Vendor badge
                    v_cell = (
                        f"<span style='background:{vc};color:white;"
                        f"padding:3px 10px;border-radius:10px;"
                        f"font-size:0.83em;font-weight:bold'>"
                        f"{row['Vendor']}</span>"
                    )

                    # File name — plain display
                    fname = str(row.get("File Name", "")).strip()
                    fn_cell = (
                        f"<span style='font-family:monospace;"
                        f"font-size:0.82em;color:#2c3e50'>{fname}</span>"
                    )

                    # ── File link — from embedded hyperlink ───
                    url = str(row.get("Hyperlink", "")).strip()

                    # Fallback: also check File Link / File URL columns
                    if not url or url == "nan":
                        url = str(row.get("File Link", "")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File URL", "")).strip()

                    if url and url not in ["", "nan"] and url.startswith("http"):
                        label = (
                            fname if len(fname) <= 45
                            else fname[:42] + "…"
                        )
                        link_cell = (
                            f"<a href='{url}' target='_blank' "
                            f"style='color:#2980b9;"
                            f"text-decoration:underline;"
                            f"font-family:monospace;"
                            f"font-size:0.82em'>"
                            f"🔗 {label}</a>"
                        )
                    else:
                        link_cell = (
                            f"<span style='color:#bdc3c7;"
                            f"font-size:0.82em;font-style:italic'>"
                            f"No link available</span>"
                        )

                    # Price
                    pval = str(row.get("Quoted Price","")).strip()
                    p_cell = (
                        f"<span style='color:#27ae60;font-weight:600'>"
                        f"{pval}</span>"
                        if pval and pval not in ["", "nan", "0"]
                        else "<span style='color:#bdc3c7'>—</span>"
                    )

                    rows_html.append(f"<tr style='background:{bg}'>")
                    rows_html.append(
                        f"<td style='{TD}'>{v_cell}</td>"
                        f"<td style='{TD};color:#555'>{row['Category']}</td>"
                        f"<td style='{TD}'>{fn_cell}</td>"
                    )
                    if has_price:
                        rows_html.append(f"<td style='{TD}'>{p_cell}</td>")
                    rows_html.append(
                        f"<td style='{TD}'>{link_cell}</td></tr>"
                    )

                rows_html.append("</tbody></table>")
                st.markdown("".join(rows_html), unsafe_allow_html=True)

        # Shared services summary
        shared_svcs = [
            s for s in selected_svcs
            if d_sel[d_sel["Service"] == s]["Vendor"].nunique() > 1
        ]
        if shared_svcs:
            with st.expander(
                "🔁 Shared Services — offered by multiple vendors",
                expanded=False,
            ):
                for s in shared_svcs:
                    vlist  = sorted(
                        d_sel[d_sel["Service"] == s]["Vendor"].unique()
                    )
                    badges = " ".join([
                        f"<span style='background:"
                        f"{vendor_color_map.get(v,'#666')};"
                        f"color:white;padding:3px 10px;"
                        f"border-radius:10px;font-size:0.85em;"
                        f"margin:2px'>{v}</span>"
                        for v in vlist
                    ])
                    st.markdown(
                        f"<div style='margin-bottom:8px'>"
                        f"<b>{s}</b> → {badges}</div>",
                        unsafe_allow_html=True,
                    )
