# ============================================================
# app.py — IT Procurement Intelligence Dashboard (Cleaned)
# Part 1 of continuation
# ============================================================

import io
import os
import re
import zipfile
from collections import defaultdict

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

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


# ============================================================
# APP CONFIG
# ============================================================

st.set_page_config(
    page_title="IT Procurement Intelligence",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# THEME
# ============================================================

C_ORANGE = "#D04A02"
C_ORANGE_DARK = "#A33A00"
C_ORANGE_MID = "#E8703A"
C_ORANGE_LITE = "#FAD4C0"

C_BLACK = "#1A1A1A"
C_DARK = "#2D2D2D"
C_MID = "#4A4A4A"

C_GREY_DARK = "#7D7D7D"
C_GREY = "#B0B0B0"
C_GREY_LITE = "#E0E0E0"
C_GREY_BG = "#F3F3F3"
C_WHITE = "#FFFFFF"

CHART_SEQ = [
    C_ORANGE,
    C_DARK,
    C_ORANGE_MID,
    C_GREY_DARK,
    C_ORANGE_DARK,
    C_GREY,
    C_MID,
    C_BLACK,
]

CFONT = dict(
    family="Georgia,'Source Sans Pro',Arial",
    size=11,
    color=C_DARK,
)

CBG = C_GREY_BG
DEMO_DIR = "demo_quotes"
MASTER_CATALOG_FILE = "Master Catalog.xlsx"
DUMMY_CATALOG_FILE = "dummy_catalog.csv"


# ============================================================
# GLOBAL CSS
# ============================================================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

    html, body, [class*="css"], div, p, span, td, th, label, button, .stMarkdown {
        font-family: 'Source Sans Pro', 'Helvetica Neue', Arial, sans-serif !important;
    }

    h1, h2, h3 {
        font-family: Georgia, 'ITC Charter', serif !important;
        font-weight: 700 !important;
    }

    section[data-testid="stSidebar"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    header { visibility: hidden; }

    .main .block-container {
        background: #F3F3F3 !important;
        max-width: 100% !important;
        padding: 1.2rem 2.5rem !important;
    }

    button[data-baseweb="tab"] {
        font-weight: 600 !important;
        font-size: 0.90em !important;
        color: #7D7D7D !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color: #D04A02 !important;
        border-bottom: 3px solid #D04A02 !important;
        background: transparent !important;
    }

    .kpi-box {
        border-radius: 4px;
        padding: 20px 12px;
        text-align: center;
        color: white;
        border-left: 5px solid rgba(255,255,255,0.2);
    }

    .kpi-value {
        font-size: 2.4em;
        font-weight: 700;
        margin: 0;
        line-height: 1.1;
        font-family: Georgia, serif !important;
    }

    .kpi-label {
        font-size: 0.72em;
        font-weight: 700;
        opacity: 0.9;
        margin-top: 6px;
        letter-spacing: 1.2px;
        text-transform: uppercase;
    }

    .sec-head {
        font-size: 0.73em;
        font-weight: 700;
        letter-spacing: 1.4px;
        text-transform: uppercase;
        color: #D04A02;
        margin: 22px 0 10px;
        border-bottom: 2px solid #D04A02;
        padding-bottom: 5px;
        display: block;
    }

    .comp-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.81em;
        border: 1px solid #E0E0E0;
    }

    .comp-table thead tr { background: #2D2D2D; }

    .comp-table thead th {
        padding: 10px 12px;
        text-align: left;
        font-weight: 700;
        font-size: 0.78em;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        color: white !important;
        border: none;
    }

    .comp-table tbody tr:nth-child(even) { background: #F8F8F8; }
    .comp-table tbody tr:hover { background: #FAD4C0; }

    .comp-table tbody td {
        padding: 9px 12px;
        border-bottom: 1px solid #EBEBEB;
        vertical-align: middle;
        word-break: break-word;
        color: #2D2D2D;
    }

    .vbadge {
        display: inline-block;
        padding: 3px 9px;
        border-radius: 2px;
        color: white;
        font-size: 0.76em;
        font-weight: 700;
        white-space: nowrap;
    }

    .scard {
        border-radius: 4px;
        padding: 16px;
        margin-bottom: 10px;
        border-left: 5px solid #D04A02;
        background: white;
    }

    .scard-orange { border-color: #D04A02; background: #FFF5F0; }
    .scard-dark { border-color: #2D2D2D; background: #F5F5F5; }
    .scard-grey { border-color: #7D7D7D; background: #FAFAFA; }

    .verdict-good {
        background: #FFF5F0;
        border: 2px solid #D04A02;
        border-radius: 4px;
        padding: 14px 18px;
        color: #D04A02;
        font-weight: 700;
    }

    .verdict-mid {
        background: #F5F5F5;
        border: 2px solid #4A4A4A;
        border-radius: 4px;
        padding: 14px 18px;
        color: #4A4A4A;
        font-weight: 700;
    }

    .verdict-bad {
        background: #F0F0F0;
        border: 2px solid #7D7D7D;
        border-radius: 4px;
        padding: 14px 18px;
        color: #2D2D2D;
        font-weight: 700;
    }

    .chat-header {
        background: #2D2D2D;
        color: white;
        padding: 12px 18px;
        border-radius: 6px 6px 0 0;
    }

    .chat-outer {
        background: #F8F8F8;
        border: 1px solid #E0E0E0;
        border-top: none;
        border-radius: 0;
        padding: 14px 14px 6px;
        min-height: 320px;
        max-height: 420px;
        overflow-y: auto;
    }

    .msg-user {
        background: #D04A02;
        color: white;
        border-radius: 14px 14px 3px 14px;
        padding: 9px 14px;
        margin: 5px 0 5px auto;
        max-width: 74%;
        font-size: 0.86em;
        display: inline-block;
        float: right;
        clear: both;
    }

    .msg-bot {
        background: white;
        color: #2D2D2D;
        border: 1px solid #E0E0E0;
        border-left: 4px solid #D04A02;
        border-radius: 14px 14px 14px 3px;
        padding: 9px 14px;
        margin: 5px 0;
        max-width: 84%;
        font-size: 0.86em;
        display: inline-block;
        float: left;
        clear: both;
    }

    .chat-wrap {
        overflow: hidden;
        margin-bottom: 3px;
    }

    .filter-bar {
        background: #2D2D2D;
        padding: 14px 20px;
        border-radius: 4px;
        margin-bottom: 18px;
        border-left: 4px solid #D04A02;
    }

    .insight-box {
        background: #FFF5F0;
        border-left: 4px solid #D04A02;
        border-radius: 0 4px 4px 0;
        padding: 11px 16px;
        margin: 10px 0;
        font-size: 0.86em;
        color: #2D2D2D;
    }

    .bucket-header {
        background: #1A1A1A;
        color: white;
        padding: 10px 20px;
        border-radius: 4px;
        margin-bottom: 6px;
        border-left: 6px solid #D04A02;
        font-size: 0.85em;
        font-weight: 700;
        letter-spacing: 1px;
        text-transform: uppercase;
    }

    div[data-testid="stHorizontalBlock"] button[kind="secondary"] {
        background: white !important;
        border: 1.5px solid #D04A02 !important;
        color: #D04A02 !important;
        border-radius: 20px !important;
        font-size: 0.78em !important;
        font-weight: 600 !important;
        padding: 4px 8px !important;
    }

    div[data-testid="stHorizontalBlock"] button[kind="secondary"]:hover {
        background: #FFF5F0 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# REGEX / PRICE EXTRACTION
# ============================================================

PRICE_RE = re.compile(
    r"(?:USD|EUR|GBP|SGD|MYR|AUD|CAD)\s?\d{1,3}(?:[,]\d{3})*(?:\.\d{1,2})?"
    r"|(?:[\$\€\£]\s?)\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?"
    r"|\d{1,3}(?:[,]\d{3})+(?:\.\d{1,2})?",
    re.IGNORECASE,
)

TOTAL_KW = [
    "grand total",
    "total amount",
    "total price",
    "amount due",
    "net total",
    "total cost",
    "total value",
    "subtotal",
    "total",
]


# ============================================================
# GENERIC HELPERS
# ============================================================

def parse_num(value) -> float:
    try:
        return float(re.sub(r"[^\d.]", "", str(value)) or "0")
    except Exception:
        return 0.0


def fmt_currency(value) -> str:
    try:
        parsed = float(re.sub(r"[^\d.]", "", str(value)) or "0")
        return "—" if parsed <= 0 else "${:,.2f}".format(parsed)
    except Exception:
        return str(value)


def best_price_from_text(text: str) -> str:
    text_lower = text.lower()

    for keyword in TOTAL_KW:
        idx = text_lower.find(keyword)
        if idx == -1:
            continue

        snippet = text[max(0, idx - 20): idx + 300]
        hits = PRICE_RE.findall(snippet)
        valid = [hit.strip() for hit in hits if parse_num(hit) >= 50]
        if valid:
            return max(valid, key=parse_num)

    all_hits = PRICE_RE.findall(text)
    valid = [hit.strip() for hit in all_hits if parse_num(hit) >= 100]
    return max(valid, key=parse_num) if valid else ""


def text_from_bytes(content: bytes, ext: str) -> str:
    text = ""
    ext = ext.lower().strip(".")

    try:
        if ext == "pdf":
            if not PDF_OK:
                return ""
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"

        elif ext in ("xlsx", "xls"):
            workbook = openpyxl.load_workbook(
                io.BytesIO(content),
                data_only=True,
                read_only=True,
            )
            rows_text = []
            for worksheet in workbook.worksheets:
                for row in worksheet.iter_rows(values_only=True):
                    row_text = "  ".join(str(cell) for cell in row if cell is not None)
                    if row_text.strip():
                        rows_text.append(row_text)
            text = "\n".join(rows_text)
            workbook.close()

        elif ext == "docx":
            with zipfile.ZipFile(io.BytesIO(content)) as zipped:
                if "word/document.xml" in zipped.namelist():
                    xml = zipped.read("word/document.xml").decode("utf-8", errors="ignore")
                    text = re.sub(r"<[^>]+>", " ", xml)
                    text = re.sub(r"\s{2,}", "\n", text)

    except Exception:
        pass

    return text


def extract_price_from_bytes(content: bytes, ext: str) -> dict:
    text = text_from_bytes(content, ext)
    price = best_price_from_text(text)

    if not price or parse_num(price) <= 0:
        all_numbers = [
            hit.strip()
            for hit in PRICE_RE.findall(text)
            if parse_num(hit) >= 1000
        ]
        if all_numbers:
            price = max(all_numbers, key=parse_num)

    return {
        "price": price,
        "price_num": parse_num(price) if price else 0.0,
        "text": text[:5000],
    }


def extract_price_from_file(filepath: str) -> dict:
    try:
        with open(filepath, "rb") as file:
            content = file.read()
        ext = filepath.rsplit(".", 1)[-1]
        return extract_price_from_bytes(content, ext)
    except Exception:
        return {"price": "", "price_num": 0.0, "text": ""}


# ============================================================
# SCORING HELPERS
# ============================================================

def price_score(new_price: float, historical_prices: list[float]):
    valid = [price for price in historical_prices if price > 0]

    if not valid or new_price <= 0:
        return None, "No comparison data", 0, 0, 0

    minimum = min(valid)
    maximum = max(valid)
    average = sum(valid) / len(valid)

    if maximum == minimum:
        return 50, "Same as historical average", average, minimum, maximum

    score = round((1 - (new_price - minimum) / (maximum - minimum)) * 100, 1)
    score = max(0, min(100, score))

    pct = round((new_price - average) / average * 100, 1)
    label = (
        f"{abs(pct)}% BELOW average — COMPETITIVE"
        if new_price < average
        else f"{abs(pct)}% ABOVE average — REVIEW NEEDED"
    )

    return score, label, average, minimum, maximum


def score_color(score):
    if score is None:
        return C_GREY_DARK
    if score >= 70:
        return C_ORANGE
    if score >= 40:
        return C_GREY_DARK
    return C_MID


def get_verdict(score):
    if score is None:
        return "⚪ No Data", "No comparison data.", C_GREY_DARK
    if score >= 70:
        return "✅ COMPETITIVE", "Priced competitively.", C_ORANGE
    if score >= 40:
        return "🟡 AVERAGE", "Within range. Negotiate.", C_GREY_DARK
    return "🔴 HIGH", "Above average. Recommend negotiating.", C_MID


# ============================================================
# DATAFRAME HELPERS
# ============================================================

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    safe_columns = [col for col in df.columns if col != "Services List"]
    cleaned = df[safe_columns].copy()

    for col in cleaned.columns:
        try:
            cleaned[col] = cleaned[col].fillna("").apply(lambda x: str(x).strip())
        except Exception:
            cleaned[col] = ""

    mask = (
        cleaned["Category"].apply(lambda x: x in ["", "nan"])
        & cleaned["Vendor"].apply(lambda x: x in ["", "nan"])
    )
    cleaned = cleaned[~mask].copy()
    cleaned.reset_index(drop=True, inplace=True)
    return cleaned


def parse_services(value) -> list[str]:
    if not value or str(value).strip() in ["", "nan", "None"]:
        return ["(unspecified)"]

    text = str(value).replace("\\n", "\n").replace("\r\n", "\n").replace("\r", "\n")
    parts = [part.strip() for part in text.split("\n") if part.strip() and part.strip() != "nan"]

    if not parts:
        parts = [part.strip() for part in text.split(";") if part.strip()]

    if not parts and len(text) < 300:
        parts = [part.strip() for part in text.split(",") if part.strip()]

    return parts if parts else ["(unspecified)"]


def explode_services(df: pd.DataFrame):
    df_base = df.copy()
    df_base["Services List"] = df_base["Comments"].apply(parse_services)

    exploded = df_base.explode("Services List").copy()
    exploded.rename(columns={"Services List": "Service"}, inplace=True)
    exploded["Service"] = exploded["Service"].apply(lambda x: str(x).strip())

    exploded = exploded[
        ~exploded["Service"].isin(["", "(unspecified)", "nan", "None"])
    ].reset_index(drop=True)

    return df_base, exploded


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    col_map = {}

    for col in df.columns:
        normalized = str(col).lower().strip()

        if normalized == "category" and "Category" not in col_map:
            col_map["Category"] = col

        elif any(key in normalized for key in ["vendor", "supplier"]) and "Vendor" not in col_map:
            col_map["Vendor"] = col

        elif ("file name" in normalized or normalized == "filename") and "File Name" not in col_map:
            col_map["File Name"] = col

        elif any(key in normalized for key in ["file link", "file url"]) and "File Link" not in col_map:
            col_map["File Link"] = col

        elif any(key in normalized for key in ["comment", "service", "description", "scope"]) and "Comments" not in col_map:
            col_map["Comments"] = col

        elif any(key in normalized for key in ["price", "cost", "amount", "quoted"]) and "Quoted Price" not in col_map:
            col_map["Quoted Price"] = col

    df = df.rename(columns={v: k for k, v in col_map.items()})

    for required in ["Category", "Vendor", "File Name", "Comments"]:
        if required not in df.columns:
            df[required] = ""

    keep = ["Category", "Vendor", "File Name", "Comments"]
    for optional in ["File Link", "Quoted Price"]:
        if optional in df.columns:
            keep.append(optional)

    return df[[col for col in keep if col in df.columns]].copy()


def extract_hyperlink_map_from_excel(filepath: str) -> dict:
    hyperlink_map = {}

    try:
        workbook = openpyxl.load_workbook(filepath)
        worksheet = workbook.active

        file_col = None
        header_row = None

        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value and str(cell.value).strip().lower() == "file name":
                    file_col = cell.column
                    header_row = cell.row
                    break
            if file_col:
                break

        if file_col and header_row:
            for row in worksheet.iter_rows(min_row=header_row + 1):
                for cell in row:
                    if cell.column == file_col and cell.value and cell.hyperlink:
                        hyperlink_map[str(cell.value).strip()] = str(cell.hyperlink.target).strip()

        workbook.close()

    except Exception:
        pass

    return hyperlink_map


@st.cache_data
def load_master_catalog():
    if not os.path.exists(MASTER_CATALOG_FILE):
        return None, None

    try:
        raw = pd.read_excel(MASTER_CATALOG_FILE, engine="openpyxl", header=None)
        header_row = 0

        for idx, row in raw.iterrows():
            vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            if any("category" in v for v in vals) and any("vendor" in v for v in vals):
                header_row = idx
                break

        df = pd.read_excel(MASTER_CATALOG_FILE, engine="openpyxl", header=header_row)
        df.columns = [str(col).strip() for col in df.columns]

        hyperlink_map = extract_hyperlink_map_from_excel(MASTER_CATALOG_FILE)

        df = normalize_columns(df)
        df = clean_df(df)
        df["Hyperlink"] = df["File Name"].map(hyperlink_map).fillna("")

        df, exploded = explode_services(df)
        return df, exploded

    except Exception as exc:
        st.warning(f"Master catalog error: {exc}")
        return None, None


@st.cache_data
def load_dummy_data():
    if not os.path.exists(DUMMY_CATALOG_FILE):
        return None, None

    try:
        df = pd.read_csv(DUMMY_CATALOG_FILE)
        df.columns = [str(col).strip() for col in df.columns]

        df = normalize_columns(df)
        df = clean_df(df)
        df["Hyperlink"] = ""

        df, exploded = explode_services(df)
        return df, exploded

    except Exception as exc:
        st.warning(f"Dummy data error: {exc}")
        return None, None


def ensure_dummy_data():
    if os.path.exists(DUMMY_CATALOG_FILE):
        return

    import random

    random.seed(42)

    vendors = [
        "NTT Data",
        "Dimension Data",
        "Telstra",
        "Optus",
        "Vocus",
        "Datacom",
    ]

    categories = {
        "Cybersecurity": [
            "Endpoint Protection",
            "SIEM Monitoring",
            "Privileged Access Mgmt",
            "Security Awareness",
            "Network Access Control",
        ],
        "Network & Telecom": [
            "Cisco Catalyst 9200-L",
            "Cisco Catalyst 9400",
            "Palo Alto NGFW",
            "SD-WAN Solution",
            "Cisco Meraki MX",
        ],
        "Hosting": [
            "VMware vSphere",
            "NetApp Storage",
            "Oracle DB License",
            "Colocation Build",
            "Backup & Recovery",
        ],
        "M365 & Power Platform": [
            "M365 E3 License",
            "M365 E5 License",
            "Power BI Premium",
            "Teams Rooms",
        ],
    }

    rows = []
    for category, services in categories.items():
        for service in services:
            vendor_count = random.randint(2, 4)
            chosen_vendors = random.sample(vendors, vendor_count)

            for vendor in chosen_vendors:
                base = random.uniform(20000, 200000)
                price = round(base * random.uniform(0.85, 1.15), 2)
                filename = "{}_{}_{}.pdf".format(
                    vendor.replace(" ", "_"),
                    service.replace(" ", "_")[:15],
                    random.randint(1000, 9999),
                )

                rows.append(
                    {
                        "Category": category,
                        "Vendor": vendor,
                        "File Name": filename,
                        "Comments": service,
                        "Quoted Price": price,
                    }
                )

    pd.DataFrame(rows).to_csv(DUMMY_CATALOG_FILE, index=False)


# ============================================================
# DOMAIN HELPERS
# ============================================================

def infer_subcategory(category, comments, filename):
    text = f"{comments} {filename}".lower()
    category_text = str(category).lower().strip()

    if "cybersecurity" in category_text:
        if any(key in text for key in ["trendmicro", "endpoint", "antivirus"]):
            return "Endpoint Protection"
        if any(key in text for key in ["cyberark", "privileged", "pam"]):
            return "Privileged Access"
        if any(key in text for key in ["knowbe4", "awareness", "phishing"]):
            return "Security Awareness"
        if any(key in text for key in ["forescout", "nac"]):
            return "Network Access Control"
        if any(key in text for key in ["siem", "splunk", "monitor"]):
            return "SIEM / Monitoring"
        return "General Security"

    if "network" in category_text or "telecom" in category_text:
        if "meraki" in text:
            return "Cisco Meraki"
        if "palo alto" in text:
            return "Palo Alto NGFW"
        if "equinix" in text:
            return "Equinix"
        if "cisco" in text:
            return "Cisco Networking"
        return "General Network"

    if "hosting" in category_text:
        if any(key in text for key in ["vmware", "vcf"]):
            return "VMware"
        if "oracle" in text:
            return "Oracle DB"
        if "netapp" in text:
            return "NetApp"
        if any(key in text for key in ["colo", "colocation"]):
            return "Colocation"
        return "General Hosting"

    if "m365" in category_text:
        return "M365 Licensing"
    if "idam" in category_text:
        return "Identity Migration"
    if "snow" in category_text:
        return "ServiceNow ITSM"
    if "summary" in category_text:
        return "Reporting"

    return str(category).strip().title()


def resolve_url(row) -> str:
    for col in ["Hyperlink", "File Link"]:
        value = str(row.get(col, "")).strip()
        if value and value not in ["", "nan"] and value.startswith("http"):
            return value

    filename = str(row.get("File Name", "")).strip()
    if filename:
        local_path = os.path.join(DEMO_DIR, filename)
        if os.path.exists(local_path):
            return local_path

    return ""


def vendor_color_maps(df_master, df_dummy):
    master_map = {}
    dummy_map = {}

    if df_master is not None and not df_master.empty:
        for idx, vendor in enumerate(sorted(df_master["Vendor"].unique())):
            master_map[vendor] = CHART_SEQ[idx % len(CHART_SEQ)]

    if df_dummy is not None and not df_dummy.empty:
        for idx, vendor in enumerate(sorted(df_dummy["Vendor"].unique())):
            dummy_map[vendor] = CHART_SEQ[idx % len(CHART_SEQ)]

    return master_map, dummy_map
