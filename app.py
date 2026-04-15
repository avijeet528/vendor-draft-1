# ============================================================
#  app.py — IT Procurement Service Dashboard (Streamlit)
#  PwC Brand | Source Sans Pro | Fixed table layout
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

# ════════════════════════════════════════════════════════════
# CSS
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

html, body, [class*="css"], div, p, span, td, th,
label, button, .stMarkdown {
    font-family: 'Source Sans Pro', 'Helvetica Neue',
                 Arial, sans-serif !important;
}

.main .block-container {
    background-color : #F3F3F3 !important;
    padding-top      : 1.5rem;
    max-width        : 100% !important;
    padding-left     : 2rem !important;
    padding-right    : 2rem !important;
}

#MainMenu {visibility: hidden;}
footer    {visibility: hidden;}
header    {visibility: hidden;}

[data-testid="collapsedControl"] {
    display: none !important;
}

section[data-testid="stSidebar"] {
    background-color : #2D2D2D !important;
    border-right     : 3px solid #D04A02;
    min-width        : 320px !important;
    max-width        : 320px !important;
}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div {
    color       : #F0F0F0 !important;
    font-family : 'Source Sans Pro', sans-serif !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] {
    background-color : #FFFFFF !important;
    border-radius    : 2px !important;
    border           : 1px solid #999 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] * {
    color: #2D2D2D !important;
}
section[data-testid="stSidebar"] div[data-baseweb="input"] {
    background-color : #FFFFFF !important;
    border-radius    : 2px !important;
}
section[data-testid="stSidebar"] div[data-baseweb="input"] input {
    color: #2D2D2D !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color : #D04A02 !important;
    border-radius    : 2px !important;
}
section[data-testid="stSidebar"] span[data-baseweb="tag"] span {
    color: white !important;
}

.kpi-box {
    border-radius : 4px;
    padding       : 20px 12px;
    text-align    : center;
    color         : white;
    border-left   : 5px solid rgba(255,255,255,0.3);
}
.kpi-value {
    font-size      : 2.3em;
    font-weight    : 700;
    margin         : 0;
    line-height    : 1.1;
    letter-spacing : -0.5px;
}
.kpi-label {
    font-size      : 0.80em;
    font-weight    : 700;
    opacity        : 0.92;
    margin-top     : 6px;
    letter-spacing : 0.8px;
    text-transform : uppercase;
}

button[data-baseweb="tab"] {
    font-weight : 600 !important;
    font-size   : 0.92em !important;
    color       : #7D7D7D !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color        : #D04A02 !important;
    border-bottom: 3px solid #D04A02 !important;
}

div[data-testid="stExpander"] details summary p {
    font-weight : 700;
    font-size   : 0.95em;
    color       : #2D2D2D !important;
}
div[data-testid="stExpander"] details {
    border        : 1px solid #ddd;
    border-radius : 4px;
    margin-bottom : 10px;
}

.svc-table {
    width           : 100%;
    border-collapse : collapse;
    table-layout    : fixed;
}
.svc-table thead tr {
    background: #2D2D2D;
}
.svc-table thead th {
    padding        : 10px 12px;
    text-align     : left;
    font-weight    : 700;
    font-size      : 0.82em;
    letter-spacing : 0.5px;
    text-transform : uppercase;
    color          : white !important;
    border         : none;
    word-break     : break-word;
}
.svc-table tbody tr:nth-child(even) { background: #F3F3F3; }
.svc-table tbody tr:hover           { background: #FCE8DC; }
.svc-table tbody td {
    padding        : 8px 12px;
    border-bottom  : 1px solid #e8e8e8;
    vertical-align : middle;
    word-break     : break-word;
    font-size      : 0.83em;
    color          : #2D2D2D;
}
.svc-table th:nth-child(1),
.svc-table td:nth-child(1) { width: 13%; }
.svc-table th:nth-child(2),
.svc-table td:nth-child(2) { width: 14%; }
.svc-table th:nth-child(3),
.svc-table td:nth-child(3) { width: 35%; }
.svc-table th:nth-child(4),
.svc-table td:nth-child(4) { width: 13%; }
.svc-table th:nth-child(5),
.svc-table td:nth-child(5) { width: 25%; }

.vendor-badge {
    display        : inline-block;
    padding        : 3px 8px;
    border-radius  : 2px;
    color          : white;
    font-size      : 0.78em;
    font-weight    : 700;
    white-space    : nowrap;
    overflow       : hidden;
    text-overflow  : ellipsis;
    max-width      : 100%;
    box-sizing     : border-box;
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# PwC COLOURS
# ════════════════════════════════════════════════════════════
COLORS = [
    "#D04A02","#295477","#299D8F",
    "#FFB600","#22992E","#E0301E",
    "#EB8C00","#6E2585","#8C8C8C","#004F9F",
]

def get_color(i):
    return COLORS[i % len(COLORS)]


# ════════════════════════════════════════════════════════════
# EXTRACT EMBEDDED HYPERLINKS
# ════════════════════════════════════════════════════════════
@st.cache_data
def extract_hyperlinks(file_path):
    link_map = {}
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        file_name_col  = None
        header_row_idx = None
        for row in ws.iter_rows():
            for cell in row:
                if (cell.value and
                        str(cell.value).strip().lower() == "file name"):
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
        st.warning("Could not extract hyperlinks: {}".format(e))
    return link_map


# ════════════════════════════════════════════════════════════
# DATA LOADING
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    FILE_PATH = "Master Catalog.xlsx"
    if not os.path.exists(FILE_PATH):
        st.error("File not found: {}".format(FILE_PATH))
        return None, None

    raw = pd.read_excel(FILE_PATH, engine="openpyxl", header=None)
    header_row = None
    for i, row in raw.iterrows():
        row_vals = [
            str(v).strip().lower() for v in row.values if pd.notna(v)
        ]
        if (any("category" in v for v in row_vals) and
                any("file" in v for v in row_vals)):
            header_row = i
            break
    if header_row is None:
        st.error("Could not detect header row.")
        return None, None

    df = pd.read_excel(FILE_PATH, engine="openpyxl", header=header_row)
    df = df.loc[:, df.columns.notna()]
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how="all", inplace=True)

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

    hyperlink_map   = extract_hyperlinks(FILE_PATH)
    df["Hyperlink"] = df["File Name"].map(hyperlink_map).fillna("")
    for fallback_col in ["File Link", "File URL"]:
        if fallback_col in df.columns:
            df["Hyperlink"] = df.apply(
                lambda r: r["Hyperlink"]
                if r["Hyperlink"] not in ["", "nan"]
                else r[fallback_col],
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
def sb_label(txt):
    st.markdown(
        "<p style='color:#F0F0F0;font-weight:700;font-size:0.88em;"
        "margin:12px 0 4px;letter-spacing:0.5px;"
        "text-transform:uppercase'>{}</p>".format(txt),
        unsafe_allow_html=True,
    )


with st.sidebar:

    st.markdown(
        "<div style='text-align:center;padding:22px 0 16px'>"
        "<div style='font-size:2.2em'>📋</div>"
        "<div style='font-size:1.1em;font-weight:700;color:white;"
        "margin:6px 0 2px;letter-spacing:0.5px'>IT Procurement</div>"
        "<div style='font-size:0.75em;color:#aaa;letter-spacing:1px;"
        "text-transform:uppercase'>Service &amp; Vendor Dashboard</div>"
        "</div>"
        "<hr style='border-color:#D04A02;border-width:2px;"
        "margin:0 0 18px'>",
        unsafe_allow_html=True,
    )

    # Category
    sb_label("📂 Category")
    all_cats = ["All"] + sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["", "nan"]
    ])
    selected_cat = st.selectbox(
        "Category", all_cats, label_visibility="collapsed")

    # Vendor — scoped to category
    sb_label("🏢 Vendor")
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
        "<hr style='border-color:#555;margin:16px 0'>",
        unsafe_allow_html=True)

    # Filtered data
    d_filt = df_exploded.copy()
    if selected_cat    != "All":
        d_filt = d_filt[d_filt["Category"] == selected_cat]
    if selected_vendor != "All":
        d_filt = d_filt[d_filt["Vendor"]   == selected_vendor]

    # Service search
    sb_label("🔍 Search Services")
    svc_search = st.text_input(
        "Search", placeholder="e.g. Cisco, Oracle, M365…",
        label_visibility="collapsed")

    available_svcs = sorted([
        s for s in d_filt["Service"].unique()
        if str(s).strip() not in ["", "nan"]
    ])
    if svc_search:
        available_svcs = [
            s for s in available_svcs
            if svc_search.lower() in s.lower()
        ]

    sb_label("🛠 Select Services ({} available)".format(
        len(available_svcs)))
    selected_svcs = st.multiselect(
        "Services",
        options=available_svcs,
        default=[],
        label_visibility="collapsed",
        help="Select one or more services",
    )

    st.markdown(
        "<hr style='border-color:#555;margin:16px 0'>",
        unsafe_allow_html=True)

    st.markdown(
        "<p style='color:#888;font-size:0.80em;margin:3px 0'>"
        "📄 {} total quotes</p>"
        "<p style='color:#888;font-size:0.80em;margin:3px 0'>"
        "🛠 {} unique services</p>"
        "<p style='color:#888;font-size:0.80em;margin:3px 0'>"
        "🏢 {} vendors</p>".format(
            len(df_master),
            df_exploded["Service"].nunique(),
            df_master["Vendor"].nunique()),
        unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
st.markdown(
    "<div style='background:#2D2D2D;color:white;"
    "padding:22px 30px;border-radius:4px;"
    "border-left:6px solid #D04A02;margin-bottom:24px'>"
    "<div style='font-size:0.75em;font-weight:700;"
    "letter-spacing:2px;text-transform:uppercase;"
    "color:#D04A02;margin-bottom:6px'>IT Procurement Analytics</div>"
    "<h1 style='margin:0;font-size:1.5em;font-weight:700;"
    "color:white;letter-spacing:-0.3px'>"
    "Service &amp; Vendor Dashboard</h1>"
    "<p style='margin:7px 0 0;opacity:0.65;font-size:0.88em;"
    "font-weight:300'>"
    "Filter by Category → Vendor auto-updates → "
    "Select a service → See vendor &amp; quotation file instantly"
    "</p></div>",
    unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# KPI CARDS
# ════════════════════════════════════════════════════════════
c1, c2, c3, c4 = st.columns(4)


def kpi_card(col, value, label, color):
    col.markdown(
        "<div class='kpi-box' style='background:{}'>"
        "<div class='kpi-value'>{}</div>"
        "<div class='kpi-label'>{}</div>"
        "</div>".format(color, value, label),
        unsafe_allow_html=True,
    )


kpi_card(c1, d_filt["File Name"].nunique(),  "Total Quotes",    "#D04A02")
kpi_card(c2, d_filt["Service"].nunique(),    "Unique Services", "#295477")
kpi_card(c3, d_filt["Vendor"].nunique(),     "Vendors",         "#299D8F")
kpi_card(c4, d_filt["Category"].nunique(),   "Categories",      "#2D2D2D")

st.markdown("<br>", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# CHART FONT
# ════════════════════════════════════════════════════════════
CFONT = dict(
    family="Source Sans Pro, Helvetica Neue, Arial",
    size=11,
    color="#2D2D2D",
)
CBG = "#F3F3F3"


def section_title(txt, caption=""):
    st.markdown(
        "<div style='font-size:0.78em;font-weight:700;"
        "letter-spacing:1px;text-transform:uppercase;"
        "color:#D04A02;margin-bottom:4px'>{}</div>".format(txt),
        unsafe_allow_html=True)
    if caption:
        st.caption(caption)


# ════════════════════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════════════════════
tab_charts, tab_table = st.tabs(["📊 Analytics", "📄 Data Table"])

with tab_charts:

    col_l, col_r = st.columns(2, gap="large")

    # Chart 1 — Service overlap
    with col_l:
        section_title(
            "SERVICE OVERLAP ANALYSIS",
            "Orange = same service quoted by multiple vendors — "
            "competitive procurement opportunity.")
        shared = (
            d_filt.groupby("Service")["Vendor"].nunique()
            .sort_values(ascending=False).head(20).reset_index()
        )
        shared.columns = ["Service", "Vendor Count"]
        shared["Color"] = shared["Vendor Count"].apply(
            lambda x: "#D04A02" if x > 1 else "#C0C0C0")
        fig1 = go.Figure(go.Bar(
            x=shared["Vendor Count"],
            y=shared["Service"].str[:44],
            orientation="h",
            marker_color=shared["Color"],
            marker_line_width=0,
            text=shared["Vendor Count"],
            textposition="outside",
            textfont=dict(size=10),
        ))
        fig1.update_layout(
            height=500,
            plot_bgcolor=CBG,
            paper_bgcolor=CBG,
            margin=dict(l=5, r=40, t=20, b=10),
            font=CFONT,
            xaxis=dict(
                title="Number of Vendors",
                showgrid=True,
                gridcolor="#E0E0E0",
                zeroline=False),
            yaxis=dict(
                autorange="reversed",
                tickfont=dict(size=9.5)),
            bargap=0.35,
        )
        st.plotly_chart(fig1, use_container_width=True)

    # Chart 2 — Vendor coverage
    with col_r:
        section_title(
            "VENDOR SERVICE COVERAGE",
            "Number of unique services each vendor provides. "
            "Higher = broader vendor capability.")
        spv = (
            d_filt.groupby("Vendor")["Service"].nunique()
            .sort_values(ascending=False).reset_index()
        )
        spv.columns = ["Vendor", "Count"]
        spv["Color"] = [
            vendor_color_map.get(v, "#8C8C8C") for v in spv["Vendor"]
        ]
        fig2 = go.Figure(go.Bar(
            x=spv["Vendor"],
            y=spv["Count"],
            marker_color=spv["Color"],
            marker_line_width=0,
            text=spv["Count"],
            textposition="outside",
            textfont=dict(size=10),
        ))
        fig2.update_layout(
            height=500,
            plot_bgcolor=CBG,
            paper_bgcolor=CBG,
            margin=dict(l=5, r=10, t=20, b=10),
            font=CFONT,
            yaxis=dict(
                title="Unique Services",
                showgrid=True,
                gridcolor="#E0E0E0",
                zeroline=False),
            xaxis=dict(
                tickangle=-35,
                tickfont=dict(size=9.5)),
            bargap=0.35,
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Chart 3 — Category donut
    section_title(
        "PROCUREMENT CATEGORY DISTRIBUTION",
        "Share of quote files across IT procurement categories.")
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
            hole=0.50,
            color_discrete_sequence=COLORS,
        )
        fig3.update_traces(
            textposition="outside",
            textinfo="label+percent",
            textfont_size=11,
            pull=[0.03] * len(cat_counts),
        )
        fig3.update_layout(
            height=420,
            margin=dict(l=20, r=20, t=20, b=20),
            paper_bgcolor=CBG,
            font=CFONT,
            legend=dict(
                orientation="v",
                x=1.02, y=0.5,
                font=dict(size=10)),
        )
        st.plotly_chart(fig3, use_container_width=True)

with tab_table:
    dm_display = df_master.copy()
    if selected_cat    != "All":
        dm_display = dm_display[dm_display["Category"] == selected_cat]
    if selected_vendor != "All":
        dm_display = dm_display[dm_display["Vendor"]   == selected_vendor]
    st.dataframe(
        dm_display.drop(
            columns=["Services List", "Hyperlink"], errors="ignore"),
        use_container_width=True,
        height=450,
    )


# ════════════════════════════════════════════════════════════
# SERVICE SELECTION RESULTS
# ════════════════════════════════════════════════════════════
st.markdown(
    "<hr style='border:none;border-top:2px solid #D04A02;"
    "margin:24px 0 16px'>",
    unsafe_allow_html=True)

section_title("SERVICE SELECTION & QUOTATION ANALYSIS")
st.markdown("### Select services from the sidebar to begin")

if not selected_svcs:
    st.info(
        "👈 **Select one or more services** from the sidebar. "
        "Results show vendor, quotation file and clickable file link.")
else:
    d_sel = d_filt[d_filt["Service"].isin(selected_svcs)].copy()

    if d_sel.empty:
        st.warning("⚠️ No results found under current filters.")
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

        # Summary banners
        if len(selected_svcs) > 1:
            if vendors_all:
                names = " · ".join(
                    ["**{}**".format(v) for v in vendors_all])
                st.success(
                    "✅ **{} vendor(s) offer ALL {} services:** {}".format(
                        len(vendors_all), len(selected_svcs), names))
            else:
                st.warning(
                    "⚠️ No single vendor covers all {} "
                    "selected services.".format(len(selected_svcs)))

            if vendors_some:
                with st.expander(
                        "🔵 Vendors with partial coverage",
                        expanded=False):
                    for v in vendors_some:
                        covered = vendor_svc_map[v].intersection(
                            set(selected_svcs))
                        color   = vendor_color_map.get(v, "#8C8C8C")
                        st.markdown(
                            "<span class='vendor-badge' "
                            "style='background:{}'>{}</span>"
                            " &nbsp; covers **{}/{}**: _{}_".format(
                                color, v,
                                len(covered), len(selected_svcs),
                                ", ".join(sorted(covered))),
                            unsafe_allow_html=True)

        # Per-service breakdown
        section_title(
            "DETAILED QUOTATION BREAKDOWN — PER SERVICE",
            "Each service shows the vendor, file name "
            "and a direct link to the quotation file.")

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
                "🛠 {}  |  {} vendor(s) · {} file(s)  [{}]".format(
                    svc, vendor_count, len(d_svc), shared_tag),
                expanded=True,
            ):
                # Vendor badges row
                badges = " ".join([
                    "<span class='vendor-badge' "
                    "style='background:{}'>{}</span>".format(
                        vendor_color_map.get(v, "#8C8C8C"), v)
                    for v in sorted(d_svc["Vendor"].unique())
                ])
                st.markdown(
                    "<div style='margin-bottom:14px'>"
                    "<b style='font-size:0.88em'>"
                    "Vendors offering this service:</b>"
                    "&nbsp;&nbsp;{}</div>".format(badges),
                    unsafe_allow_html=True)

                # Fixed-layout HTML table
                rows = [
                    "<table class='svc-table'>"
                    "<thead><tr>"
                    "<th>Vendor</th>"
                    "<th>Category</th>"
                    "<th>📄 File Name</th>"
                ]
                if has_price:
                    rows.append("<th>💰 Quoted Price</th>")
                rows.append(
                    "<th>🔗 File Link</th>"
                    "</tr></thead><tbody>"
                )

                for i, (_, row) in enumerate(d_svc.iterrows()):
                    bg    = "#ffffff" if i % 2 == 0 else "#F3F3F3"
                    color = vendor_color_map.get(row["Vendor"], "#8C8C8C")
                    fname = str(row.get("File Name", "")).strip()

                    # Vendor cell
                    v_cell = (
                        "<span class='vendor-badge' "
                        "style='background:{}'>{}</span>".format(
                            color, row["Vendor"])
                    )

                    # File name cell
                    fn_cell = (
                        "<span style='font-family:monospace;"
                        "font-size:0.80em;word-break:break-all;"
                        "color:#2D2D2D'>{}</span>".format(fname)
                    )

                    # URL
                    url = str(row.get("Hyperlink", "")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File Link", "")).strip()
                    if not url or url == "nan":
                        url = str(row.get("File URL",  "")).strip()
                    if url == "nan":
                        url = ""

                    # Link cell
                    if url and url.startswith("http"):
                        link_cell = (
                            "<a href='{}' target='_blank' "
                            "style='color:#D04A02;font-weight:600;"
                            "font-size:0.82em;text-decoration:none'>"
                            "↗ Open file</a>".format(url)
                        )
                    else:
                        link_cell = (
                            "<span style='color:#bbb;"
                            "font-size:0.80em'>—</span>"
                        )

                    # Price cell
                    pval   = str(row.get("Quoted Price", "")).strip()
                    p_cell = (
                        "<span style='color:#22992E;font-weight:700;"
                        "font-family:monospace'>{}</span>".format(pval)
                        if pval and pval not in ["", "nan", "0"]
                        else "<span style='color:#bbb'>—</span>"
                    )

                    rows.append(
                        "<tr style='background:{}'>"
                        "<td>{}</td>"
                        "<td style='color:#555'>{}</td>"
                        "<td>{}</td>".format(
                            bg, v_cell, row["Category"], fn_cell)
                    )
                    if has_price:
                        rows.append("<td>{}</td>".format(p_cell))
                    rows.append("<td>{}</td></tr>".format(link_cell))

                rows.append("</tbody></table>")
                st.markdown("".join(rows), unsafe_allow_html=True)

                # ── Mini charts per service ───────────────────
                st.markdown("<br>", unsafe_allow_html=True)
                section_title("ANALYSIS FOR THIS SERVICE")

                mc1, mc2 = st.columns(2, gap="large")

                # Mini chart 1 — Quote files per vendor
                with mc1:
                    vc_counts = (
                        d_svc.groupby("Vendor").size().reset_index()
                    )
                    vc_counts.columns = ["Vendor", "Files"]
                    vc_counts["Color"] = [
                        vendor_color_map.get(v, "#8C8C8C")
                        for v in vc_counts["Vendor"]
                    ]
                    mfig1 = go.Figure(go.Bar(
                        x=vc_counts["Vendor"],
                        y=vc_counts["Files"],
                        marker_color=vc_counts["Color"],
                        marker_line_width=0,
                        text=vc_counts["Files"],
                        textposition="outside",
                    ))
                    mfig1.update_layout(
                        title=dict(
                            text="Quote Files per Vendor",
                            font=dict(size=12, color="#2D2D2D")),
                        height=280,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5, r=10, t=40, b=10),
                        font=CFONT,
                        yaxis=dict(
                            showgrid=True,
                            gridcolor="#E0E0E0",
                            zeroline=False),
                        xaxis=dict(tickangle=-20),
                        bargap=0.4,
                    )
                    st.plotly_chart(mfig1, use_container_width=True)

                # Mini chart 2 — Category breakdown
                with mc2:
                    cat_svc = (
                        d_svc.groupby("Category").size().reset_index()
                    )
                    cat_svc.columns = ["Category", "Count"]
                    mfig2 = px.pie(
                        cat_svc,
                        names="Category",
                        values="Count",
                        hole=0.45,
                        color_discrete_sequence=COLORS,
                    )
                    mfig2.update_traces(
                        textposition="inside",
                        textinfo="label+percent",
                        textfont_size=10,
                    )
                    mfig2.update_layout(
                        title=dict(
                            text="Category Breakdown",
                            font=dict(size=12, color="#2D2D2D")),
                        height=280,
                        margin=dict(l=10, r=10, t=40, b=10),
                        paper_bgcolor=CBG,
                        font=CFONT,
                        showlegend=True,
                        legend=dict(font=dict(size=9)),
                    )
                    st.plotly_chart(mfig2, use_container_width=True)

        # Shared services summary
        shared_svcs = [
            s for s in selected_svcs
            if d_sel[d_sel["Service"] == s]["Vendor"].nunique() > 1
        ]
        if shared_svcs:
            with st.expander(
                "🔁 Shared Services — same service by multiple vendors",
                expanded=False,
            ):
                for s in shared_svcs:
                    vlist  = sorted(
                        d_sel[d_sel["Service"] == s]["Vendor"].unique())
                    badges = " ".join([
                        "<span class='vendor-badge' "
                        "style='background:{}'>{}</span>".format(
                            vendor_color_map.get(v, "#8C8C8C"), v)
                        for v in vlist
                    ])
                    st.markdown(
                        "<div style='margin-bottom:10px'>"
                        "<b>{}</b> &nbsp;→&nbsp; {}</div>".format(
                            s, badges),
                        unsafe_allow_html=True)
