# ============================================================
#  app.py — IT Procurement Service Dashboard (Streamlit)
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

st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    #MainMenu {visibility: hidden;}
    footer     {visibility: hidden;}
    header     {visibility: hidden;}

    .kpi-box   { border-radius:12px; padding:18px 10px;
                 text-align:center; color:white; }
    .kpi-value { font-size:2.2em; font-weight:800;
                 margin:0; line-height:1.1; }
    .kpi-label { font-size:0.85em; opacity:0.88; margin-top:5px; }

    section[data-testid="stSidebar"] {
        background-color: #2c3e50 !important;
    }
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] div {
        color: #ecf0f1 !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="select"] {
        background-color: white !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="select"] * {
        color: #2c3e50 !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="input"] {
        background-color: white !important;
        border-radius: 6px !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="input"] input {
        color: #2c3e50 !important;
    }
    section[data-testid="stSidebar"] span[data-baseweb="tag"] {
        background-color: #2980b9 !important;
    }
    section[data-testid="stSidebar"] span[data-baseweb="tag"] span {
        color: white !important;
    }
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] .stCaptionContainer p {
        color: #95a5a6 !important;
    }
    div[data-testid="stExpander"] details summary p {
        font-size:1em; font-weight:600;
    }

    /* ── FIX: vendor badge always fits in one line ── */
    .vendor-badge {
        display          : inline-block;
        padding          : 4px 10px;
        border-radius    : 10px;
        color            : white;
        font-size        : 0.80em;
        font-weight      : 700;
        white-space      : nowrap;
        overflow         : hidden;
        text-overflow    : ellipsis;
        max-width        : 100%;
        box-sizing       : border-box;
    }
    /* ── service result table ── */
    .svc-table {
        width           : 100%;
        border-collapse : collapse;
        table-layout    : fixed;        /* fixed columns = no overflow */
    }
    .svc-table th {
        padding         : 8px 10px;
        text-align      : left;
        border-bottom   : 2px solid #ddd;
        background      : #f4f6f7;
        font-size       : 0.87em;
        color           : #2c3e50;
        word-break      : break-word;
    }
    .svc-table td {
        padding         : 8px 10px;
        border-bottom   : 1px solid #eee;
        font-size       : 0.84em;
        vertical-align  : middle;
        word-break      : break-word;
    }
    /* fixed column widths */
    .svc-table th:nth-child(1),
    .svc-table td:nth-child(1) { width: 14%; }   /* Vendor  */
    .svc-table th:nth-child(2),
    .svc-table td:nth-child(2) { width: 14%; }   /* Category */
    .svc-table th:nth-child(3),
    .svc-table td:nth-child(3) { width: 36%; }   /* File Name */
    .svc-table th:nth-child(4),
    .svc-table td:nth-child(4) { width: 14%; }   /* Quoted Price */
    .svc-table th:nth-child(5),
    .svc-table td:nth-child(5) { width: 22%; }   /* File Link */
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
# EXTRACT EMBEDDED HYPERLINKS
# ════════════════════════════════════════════════════════════
@st.cache_data
def extract_hyperlinks(file_path: str) -> dict:
    link_map = {}
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        file_name_col  = None
        header_row_idx = None
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and \
                        str(cell.value).strip().lower() == "file name":
                    file_name_col  = cell.column
                    header_row_idx = cell.row
                    break
            if file_name_col:
                break
        if file_name_col is None:
            return link_map
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

    df = pd.read_excel(FILE_PATH, engine="openpyxl", header=header_row)
    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)

    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category":                     col_map["Category"]     = c
        elif "vendor" in cl or "type" in cl:     col_map["Vendor"]       = c
        elif cl == "file name":                  col_map["File Name"]    = c
        elif cl == "file link":                  col_map["File Link"]    = c
        elif cl == "file url":                   col_map["File URL"]     = c
        elif "comment" in cl:                    col_map["Comments"]     = c
        elif "quoted" in cl or "price" in cl:    col_map["Quoted Price"] = c

    df.rename(columns={v: k for k, v in col_map.items()}, inplace=True)

    keep = ["Category", "Vendor", "File Name", "Comments"]
    for extra in ["File Link", "File URL", "Quoted Price"]:
        if extra in df.columns:
            keep.append(extra)
    df = df[[c for c in keep if c in df.columns]].copy()

    df = df[
        ~(
            df["Category"].astype(str).str.strip().isin(["", "nan"]) &
            df["Vendor"].astype(str).str.strip().isin(["", "nan"])
        )
    ].copy()

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df.reset_index(drop=True, inplace=True)

    # Embedded hyperlinks
    hyperlink_map  = extract_hyperlinks(FILE_PATH)
    df["Hyperlink"] = df["File Name"].map(hyperlink_map).fillna("")
    for fallback in ["File Link", "File URL"]:
        if fallback in df.columns:
            df["Hyperlink"] = df.apply(
                lambda r: r["Hyperlink"]
                if r["Hyperlink"] not in ["", "nan"]
                else r[fallback],
                axis=1,
            )

    def parse_services(raw_val):
        if not raw_val or str(raw_val).strip() in ["", "nan"]:
            return ["(unspecified)"]
        parts = [s.strip() for s in str(raw_val).split("\n") if s.strip()]
        return parts if parts else ["(unspecified)"]

    df["Services List"] = df["Comments"].apply(parse_services)

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

    st.markdown(
        "<p style='color:#ecf0f1;font-weight:700;margin-bottom:4px'>"
        "📂 Category</p>", unsafe_allow_html=True)
    all_cats = ["All"] + sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["", "nan"]
    ])
    selected_cat = st.selectbox(
        "Category", all_cats, label_visibility="collapsed")

    st.markdown(
        "<p style='color:#ecf0f1;font-weight:700;margin:10px 0 4px'>"
        "🏢 Vendor</p>", unsafe_allow_html=True)
    vendor_pool = (
        df_master if selected_cat == "All"
        else df_master[df_master["Category"] == selected_cat]
    )
    all_vendors = ["All"] + sorted([
        v for v in vendor_pool["Vendor"].unique()
        if str(v).strip() not in ["", "nan"]
    ])
    selected_vendor = st.selectbox(
        "Vendor", all_vendors, label_visibility="collapsed")

    st.markdown(
        "<hr style='border-color:#3d5166;margin:14px 0'>",
        unsafe_allow_html=True)

    d_filt = df_exploded.copy()
    if selected_cat != "All":
        d_filt = d_filt[d_filt["Category"] == selected_cat]
    if selected_vendor != "All":
        d_filt = d_filt[d_filt["Vendor"] == selected_vendor]

    st.markdown(
        "<p style='color:#ecf0f1;font-weight:700;margin-bottom:4px'>"
        "🔍 Search Services</p>", unsafe_allow_html=True)
    svc_search = st.text_input(
        "Search", placeholder="e.g. Cisco, Oracle, M365…",
        label_visibility="collapsed")

    available_svcs = sorted([
        s for s in d_filt["Service"].unique()
        if str(s).strip() not in ["", "nan"]
    ])
    if svc_search:
        available_svcs = [s for s in available_svcs
                          if svc_search.lower() in s.lower()]

    st.markdown(
        f"<p style='color:#ecf0f1;font-weight:700;margin:10px 0 4px'>"
        f"🛠 Select Services "
        f"<span style='font-weight:400;color:#95a5a6'>"
        f"({len(available_svcs)} available)</span></p>",
        unsafe_allow_html=True)
    selected_svcs = st.multiselect(
        "Services", options=available_svcs, default=[],
        label_visibility="collapsed",
        help="Select one or more services")

    st.markdown(
        "<hr style='border-color:#3d5166;margin:14px 0'>",
        unsafe_allow_html=True)
    st.markdown(
        f"<p style='color:#95a5a6;font-size:0.82em;margin:2px 0'>"
        f"📄 {len(df_master)} total quotes</p>"
        f"<p style='color:#95a5a6;font-size:0.82em;margin:2px 0'>"
        f"🛠 {df_exploded['Service'].nunique()} unique services</p>"
        f"<p style='color:#95a5a6;font-size:0.82em;margin:2px 0'>"
        f"🏢 {df_master['Vendor'].nunique()} vendors</p>",
        unsafe_allow_html=True)

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
        f"<div class='kpi-label'>{label}</div></div>",
        unsafe_allow_html=True)

kpi_card(c1, d_filt["File Name"].nunique(),  "Total Quotes",    "#2980b9")
kpi_card(c2, d_filt["Service"].nunique(),    "Unique Services", "#27ae60")
kpi_card(c3, d_filt["Vendor"].nunique(),     "Vendors",         "#8e44ad")
kpi_card(c4, d_filt["Category"].nunique(),   "Categories",      "#e67e22")
st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════════════════════
tab_charts, tab_table = st.tabs(["📊 Charts", "📄 Data Table"])

with tab_charts:
    col_l, col_r = st.columns(2)

    with col_l:
        shared = (
            d_filt.groupby("Service")["Vendor"].nunique()
            .sort_values(ascending=False).head(20).reset_index()
        )
        shared.columns = ["Service", "Vendor Count"]
        shared["Color"] = shared["Vendor Count"].apply(
            lambda x: "#e74c3c" if x > 1 else "#bdc3c7")
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
            xaxis_title="# Vendors", height=540,
            plot_bgcolor="#f8f9fa", paper_bgcolor="#f8f9fa",
            margin=dict(l=10, r=30, t=45, b=10),
            font=dict(size=11),
            yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig1, use_container_width=True)

    with col_r:
        spv = (d_filt.groupby("Vendor")["Service"].nunique()
               .sort_values(ascending=False).reset_index())
        spv.columns = ["Vendor", "Service Count"]
        spv["Color"] = [vendor_color_map.get(v, "#999") for v in spv["Vendor"]]
        fig2 = go.Figure(go.Bar(
            x=spv["Vendor"], y=spv["Service Count"],
            marker_color=spv["Color"],
            text=spv["Service Count"], textposition="outside"))
        fig2.update_layout(
            title="📦 Unique Services per Vendor",
            yaxis_title="# Unique Services", height=540,
            plot_bgcolor="#f8f9fa", paper_bgcolor="#f8f9fa",
            margin=dict(l=10, r=10, t=45, b=10),
            font=dict(size=11), xaxis_tickangle=-30)
        st.plotly_chart(fig2, use_container_width=True)

    cat_counts = (
        d_filt.drop_duplicates(subset=["Category", "File Name"])
        .groupby("Category").size().reset_index()
    )
    cat_counts.columns = ["Category", "Count"]
    if not cat_counts.empty:
        fig3 = px.pie(
            cat_counts, names="Category", values="Count",
            title="🥧 Quote Files by Category",
            hole=0.45, color_discrete_sequence=COLORS)
        fig3.update_layout(
            height=480,
            margin=dict(l=10, r=10, t=45, b=10),
            paper_bgcolor="#f8f9fa", font=dict(size=12))
        fig3.update_traces(
            textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig3, use_container_width=True)

with tab_table:
    dm_display = df_master.copy()
    if selected_cat != "All":
        dm_display = dm_display[dm_display["Category"] == selected_cat]
    if selected_vendor != "All":
        dm_display = dm_display[dm_display["Vendor"] == selected_vendor]
    st.dataframe(
        dm_display.drop(
            columns=["Services List", "Hyperlink"], errors="ignore"),
        use_container_width=True, height=450)

# ════════════════════════════════════════════════════════════
# SERVICE SELECTION RESULTS
# ════════════════════════════════════════════════════════════
st.markdown(
    "<hr style='border-color:#ddd;margin:14px 0'>",
    unsafe_allow_html=True)
st.markdown("### 🛠 Service → Vendor & Quotation File Results")

if not selected_svcs:
    st.info("👈 **Select one or more services** from the sidebar.")
else:
    d_sel = d_filt[d_filt["Service"].isin(selected_svcs)].copy()

    if d_sel.empty:
        st.warning("⚠️ No results found for selected service(s).")
    else:
        vendor_svc_map = defaultdict(set)
        for _, row in d_sel.iterrows():
            vendor_svc_map[row["Vendor"]].add(row["Service"])

        vendors_all  = sorted([v for v, s in vendor_svc_map.items()
                                if set(selected_svcs).issubset(s)])
        vendors_some = sorted([v for v, s in vendor_svc_map.items()
                                if not set(selected_svcs).issubset(s)])

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
                        "🔵 Vendors with partial coverage", expanded=False):
                    for v in vendors_some:
                        covered = vendor_svc_map[v].intersection(
                            set(selected_svcs))
                        c = vendor_color_map.get(v, "#666")
                        st.markdown(
                            f"<span style='background:{c};color:white;"
                            f"padding:3px 10px;border-radius:10px;"
                            f"font-size:0.88em;font-weight:bold'>{v}</span>"
                            f" &nbsp; covers **{len(covered)}/"
                            f"{len(selected_svcs)}**: "
                            f"_{', '.join(sorted(covered))}_",
                            unsafe_allow_html=True)

        st.markdown("#### 📄 Vendor & Quotation File — per Service")
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
                if vendor_count > 1 else "✅ SINGLE VENDOR")

            with st.expander(
                f"🛠 {svc}  |  {vendor_count} vendor(s) · "
                f"{len(d_svc)} file(s)  [{shared_tag}]",
                expanded=True):

                badges = " ".join([
                    f"<span class='vendor-badge' "
                    f"style='background:{vendor_color_map.get(v, \"#666\")}'>"
                    f"{v}</span>"
                    for v in sorted(d_svc["Vendor"].unique())
                ])
                st.markdown(
                    f"<div style='margin-bottom:12px'>"
                    f"<b>Vendors offering this service:</b>"
                    f"&nbsp;{badges}</div>",
                    unsafe_allow_html=True)

                # ── HTML table with fixed layout ──────────────
                rows = [
                    "<table class='svc-table'>"
                    "<thead><tr>"
                    "<th>Vendor</th>"
                    "<th>Category</th>"
                    "<th>📄 File Name</th>"
                ]
                if has_price:
                    rows.append("<th>💰 Quoted Price</th>")
                rows.append("<th>🔗 File Link</th></tr></thead><tbody>")

                for i, (_, row) in enumerate(d_svc.iterrows()):
                    bg  = "#ffffff" if i % 2 == 0 else "#f9f9f9"
                    vc  = vendor_color_map.get(row["Vendor"], "#666")

                    # ── Vendor cell — single-line badge ────────
                    v_cell = (
                        f"<span class='vendor-badge' "
                        f"style='background:{vc}'>"
                        f"{row['Vendor']}</span>"
                    )

                    # ── File name cell ─────────────────────────
                    fname   = str(row.get("File Name", "")).strip()
                    fn_cell = (
                        f"<span style='font-family:monospace;"
                        f"font-size:0.82em;color:#2c3e50;"
                        f"word-break:break-all'>{fname}</span>"
                    )

                    # ── File link cell — embedded hyperlink ────
                    url = str(row.get("Hyperlink", "")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File Link", "")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File URL",  "")).strip()

                    if url and url not in ["", "nan"] \
                            and url.startswith("http"):
                        label = (fname[:38] + "…"
                                 if len(fname) > 38 else fname)
                        link_cell = (
                            f"<a href='{url}' target='_blank' "
                            f"style='color:#2980b9;"
                            f"text-decoration:underline;"
                            f"font-family:monospace;"
                            f"font-size:0.82em;word-break:break-all'>"
                            f"🔗 {label}</a>"
                        )
                    else:
                        link_cell = (
                            "<span style='color:#bdc3c7;"
                            "font-size:0.82em;font-style:italic'>"
                            "No link available</span>"
                        )

                    # ── Price cell ─────────────────────────────
                    pval   = str(row.get("Quoted Price", "")).strip()
                    p_cell = (
                        f"<span style='color:#27ae60;font-weight:600'>"
                        f"{pval}</span>"
                        if pval and pval not in ["", "nan", "0"]
                        else "<span style='color:#bdc3c7'>—</span>"
                    )

                    rows.append(f"<tr style='background:{bg}'>")
                    rows.append(
                        f"<td>{v_cell}</td>"
                        f"<td style='color:#555'>{row['Category']}</td>"
                        f"<td>{fn_cell}</td>"
                    )
                    if has_price:
                        rows.append(f"<td>{p_cell}</td>")
                    rows.append(f"<td>{link_cell}</td></tr>")

                rows.append("</tbody></table>")
                st.markdown("".join(rows), unsafe_allow_html=True)

        shared_svcs = [
            s for s in selected_svcs
            if d_sel[d_sel["Service"] == s]["Vendor"].nunique() > 1
        ]
        if shared_svcs:
            with st.expander(
                    "🔁 Shared Services — offered by multiple vendors",
                    expanded=False):
                for s in shared_svcs:
                    vlist  = sorted(
                        d_sel[d_sel["Service"] == s]["Vendor"].unique())
                    badges = " ".join([
                        f"<span class='vendor-badge' "
                        f"style='background:"
                        f"{vendor_color_map.get(v, \"#666\")}'>"
                        f"{v}</span>"
                        for v in vlist
                    ])
                    st.markdown(
                        f"<div style='margin-bottom:8px'>"
                        f"<b>{s}</b> → {badges}</div>",
                        unsafe_allow_html=True)
