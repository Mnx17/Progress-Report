# =============================================================================
# app.py  —  Fisheries Sector Progress-Report Dashboard
# =============================================================================
# Bilingual (Arabic / English), read-only dashboard.
#
# DUAL DATABASE BACKEND:
#   • Windows on-prem  →  Microsoft Access (.accdb) via pyodbc  [ReadOnly=1]
#   • Streamlit Cloud / Linux  →  SQLite (.db) via built-in sqlite3
#
# Backend is selected automatically from DB_PATH file extension:
#   .accdb  →  pyodbc / Access ODBC  (Windows only)
#   .db     →  sqlite3               (any platform)
#
# DB_PATH resolution order:
#   1. .streamlit/secrets.toml  [database] path = "..."
#   2. Environment variable  DB_PATH=...
#   3. Auto-detect: FisheriesQ1_2026.accdb (Windows) or fisheries.db (Linux)
#
# SECURITY MODEL (network-level):
#   • All SQL is SELECT-only — no INSERT / UPDATE / DELETE anywhere here.
#   • Access: ReadOnly=1 + readonly=True in connection.
#   • SQLite: opened in read-only URI mode (uri=True, ?mode=ro).
#   • DB_PATH is never hard-coded in source.
#
# RUN:
#   streamlit run app.py
# =============================================================================

import os
import sys
import sqlite3
import platform
from contextlib import contextmanager

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
from dotenv import load_dotenv

# Optional pyodbc — only needed on Windows with Access backend
try:
    import pyodbc
    _PYODBC_OK = True
except ImportError:
    _PYODBC_OK = False

# --------------------------------------------------------------------------- #
# 1. CONFIGURATION                                                             #
# --------------------------------------------------------------------------- #

load_dotenv()  # load DB_PATH from .env if present

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------- DB_PATH resolution (priority order) ----------
try:
    DB_PATH = st.secrets["database"]["path"]          # 1. secrets.toml
except Exception:
    DB_PATH = os.environ.get("DB_PATH", "")           # 2. env var

if not DB_PATH:
    # 3. Auto-detect: prefer .accdb on Windows, .db elsewhere
    _accdb = os.path.join(_SCRIPT_DIR, "FisheriesQ1_2026.accdb")
    _sqlite = os.path.join(_SCRIPT_DIR, "fisheries.db")
    if sys.platform == "win32" and os.path.exists(_accdb):
        DB_PATH = _accdb
    else:
        DB_PATH = _sqlite

# ---------- Backend selection ----------
_EXT = os.path.splitext(DB_PATH)[1].lower()
USE_SQLITE = (_EXT == ".db") or (not _PYODBC_OK)

# Access ODBC connection string (used only when USE_SQLITE is False)
_ODBC_CONN_STR = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={DB_PATH};"
    r"ReadOnly=1;"
)

# SQLite read-only URI (used when USE_SQLITE is True)
_SQLITE_URI = f"file:{DB_PATH}?mode=ro"

# Plotly font — supports Arabic Unicode rendering
CHART_FONT = dict(family="Arial Unicode MS, Segoe UI, Arial, sans-serif", size=13)

# Brand colour palette (marine theme)
C_NAVY    = "#003366"   # primary navy
C_BLUE    = "#0066CC"   # secondary blue
C_TEAL    = "#008B8B"   # teal (coastal)
C_ORANGE  = "#D35400"   # coral orange (commercial)
C_GREEN   = "#27AE60"   # sea green (aquaculture / growth positive)
C_RED     = "#C0392B"   # deep red (decline / negative)
C_LBLUE   = "#85C1E9"   # light blue (2025 comparison)

# Chart colours for sectors  [Artisanal, Coastal, Commercial, Aquaculture]
SECTOR_COLORS = [C_NAVY, C_TEAL, C_ORANGE, C_GREEN]

# Chart colours for governorates (6 entries)
GOV_COLORS = ["#003366", "#0066CC", "#008B8B", "#D35400", "#27AE60", "#8E44AD"]

# Month name lookup  {1: {"en": "January", "ar": "يناير"}, …}
MONTHS = {
    1: {"en": "January",  "ar": "يناير"},
    2: {"en": "February", "ar": "فبراير"},
    3: {"en": "March",    "ar": "مارس"},
    4: {"en": "April",    "ar": "أبريل"},
    5: {"en": "May",      "ar": "مايو"},
    6: {"en": "June",     "ar": "يونيو"},
    7: {"en": "July",     "ar": "يوليو"},
    8: {"en": "August",   "ar": "أغسطس"},
    9: {"en": "September","ar": "سبتمبر"},
   10: {"en": "October",  "ar": "أكتوبر"},
   11: {"en": "November", "ar": "نوفمبر"},
   12: {"en": "December", "ar": "ديسمبر"},
}


# --------------------------------------------------------------------------- #
# 2. TRANSLATION DICTIONARY                                                    #
# --------------------------------------------------------------------------- #
# All user-visible strings that are NOT pulled directly from the database.
# Key naming convention: snake_case.  Access with t(key) helper.
# --------------------------------------------------------------------------- #

T = {
    "en": {
        # --- Sidebar ---
        "language_label":        "Language / اللغة",
        "db_path_label":         "Database Path",
        "sidebar_title":         "Fisheries Dashboard",
        "data_refresh":          "Data refreshes every 10 minutes",

        # --- Page header ---
        "page_title":            "Fisheries Sector — Statistical Indicators",
        "quarter_badge":         "Q1 2026 · January – March",

        # --- KPI Cards ---
        "kpi_total_qty":         "Total Production",
        "kpi_total_val":         "Total Value",
        "kpi_artisanal_share":   "Artisanal Share",
        "kpi_aqua_growth":       "Aquaculture Growth",
        "unit_k_tons":           "Thousand Tons",
        "unit_m_omr":            "Million OMR",
        "unit_pct":              "%",
        "vs_2025":               "vs 2025",

        # --- Section headers ---
        "sec_overview":          "📊  Monthly Production Overview",
        "sec_sectors":           "🏭  Production by Fishing Sector",
        "sec_artisanal":         "⛵  Artisanal Fishing",
        "sec_aquaculture":       "🐟  Aquaculture",
        "sec_commercial":        "🚢  Commercial Fishing",
        "sec_raw":               "📋  Raw Data Tables",

        # --- Chart titles ---
        "chart_monthly_overall": "Overall Monthly Production — 2025 vs 2026 (Thousand Tons)",
        "chart_sector_donut":    "Production Share by Sector",
        "chart_sector_bar":      "Sector Production — 2025 vs 2026",
        "chart_gov_bar":         "Artisanal Production by Governorate (Thousand Tons)",
        "chart_gov_donut":       "Artisanal Share by Governorate",
        "chart_top_species":     "Top 5 Artisanal Species by Share",
        "chart_aqua_species":    "Aquaculture Production by Species",
        "chart_vessels":         "Commercial Production by Vessel (Top 15)",
        "chart_vessels_full":    "All Commercial Vessels — Q1 2026",

        # --- Axis / legend labels ---
        "axis_qty_tons":         "Quantity (Tons)",
        "axis_qty_ktons":        "Quantity (Thousand Tons)",
        "axis_val_komr":         "Value (Thousand OMR)",
        "axis_share_pct":        "Share (%)",
        "lbl_2025":              "2025",
        "lbl_2026":              "2026",
        "lbl_quantity":          "Quantity",
        "lbl_value":             "Value",
        "lbl_growth":            "Growth",
        "lbl_vessel":            "Vessel",
        "lbl_production":        "Production (Tons)",
        "lbl_species":           "Species",
        "lbl_share":             "Share",
        "lbl_governorate":       "Governorate",
        "lbl_sector":            "Sector",
        "lbl_month":             "Month",
        "lbl_jan":               "January",
        "lbl_feb":               "February",
        "lbl_mar":               "March",
        "lbl_total":             "Q1 Total",

        # --- Raw data expanders ---
        "exp_monthly_qty":       "Monthly Production by Sector — Quantity (Tons)",
        "exp_monthly_val":       "Monthly Production by Sector — Value (Thousand OMR)",
        "exp_gov_qty":           "Artisanal by Governorate — Quantity (Tons)",
        "exp_vessels":           "Commercial Fishing — All Vessels",

        # --- Misc ---
        "error_db":              "Cannot connect to the database. Check DB_PATH and ODBC driver.",
        "no_data":               "No data available.",
        "currency_note":         "Values in Thousand Omani Rials (OMR)",
        "data_source":           "Source: Q1 2026 Fisheries Progress Report",
        # Share panel
        "share_header":          "🔗 Share Dashboard",
        "share_note":            "View-only · opens in the same language",
        "share_copied":          "✅ Link ready — paste it anywhere",
    },

    "ar": {
        # --- Sidebar ---
        "language_label":        "اللغة / Language",
        "db_path_label":         "مسار قاعدة البيانات",
        "sidebar_title":         "لوحة القطاع السمكي",
        "data_refresh":          "تتجدد البيانات كل 10 دقائق",

        # --- Page header ---
        "page_title":            "البيانات والمؤشرات الإحصائية للقطاع السمكي",
        "quarter_badge":         "الربع الأول 2026 · يناير – مارس",

        # --- KPI Cards ---
        "kpi_total_qty":         "إجمالي الإنتاج",
        "kpi_total_val":         "إجمالي القيمة",
        "kpi_artisanal_share":   "نسبة الصيد الحرفي",
        "kpi_aqua_growth":       "نمو الاستزراع",
        "unit_k_tons":           "ألف طن",
        "unit_m_omr":            "مليون ريال عماني",
        "unit_pct":              "%",
        "vs_2025":               "مقارنة بـ 2025",

        # --- Section headers ---
        "sec_overview":          "📊  نظرة عامة على الإنتاج الشهري",
        "sec_sectors":           "🏭  الإنتاج حسب قطاع الصيد",
        "sec_artisanal":         "⛵  الصيد الحرفي",
        "sec_aquaculture":       "🐟  الاستزراع السمكي",
        "sec_commercial":        "🚢  الصيد التجاري",
        "sec_raw":               "📋  جداول البيانات التفصيلية",

        # --- Chart titles ---
        "chart_monthly_overall": "الإنتاج الشهري الإجمالي — 2025 مقابل 2026 (ألف طن)",
        "chart_sector_donut":    "نسب الإنتاج حسب قطاع الصيد",
        "chart_sector_bar":      "الإنتاج حسب القطاع — 2025 مقابل 2026",
        "chart_gov_bar":         "إنتاج الصيد الحرفي حسب المحافظات (ألف طن)",
        "chart_gov_donut":       "نسبة مساهمة المحافظات في الصيد الحرفي",
        "chart_top_species":     "أعلى 5 أنواع في الصيد الحرفي",
        "chart_aqua_species":    "إنتاج الاستزراع السمكي حسب الأنواع",
        "chart_vessels":         "إنتاج الصيد التجاري حسب السفن (أعلى 15)",
        "chart_vessels_full":    "جميع سفن الصيد التجاري — الربع الأول 2026",

        # --- Axis / legend labels ---
        "axis_qty_tons":         "الكمية (طن)",
        "axis_qty_ktons":        "الكمية (ألف طن)",
        "axis_val_komr":         "القيمة (ألف ريال عماني)",
        "axis_share_pct":        "النسبة (%)",
        "lbl_2025":              "2025",
        "lbl_2026":              "2026",
        "lbl_quantity":          "الكمية",
        "lbl_value":             "القيمة",
        "lbl_growth":            "النمو",
        "lbl_vessel":            "السفينة",
        "lbl_production":        "الإنتاج (طن)",
        "lbl_species":           "النوع",
        "lbl_share":             "النسبة",
        "lbl_governorate":       "المحافظة",
        "lbl_sector":            "القطاع",
        "lbl_month":             "الشهر",
        "lbl_jan":               "يناير",
        "lbl_feb":               "فبراير",
        "lbl_mar":               "مارس",
        "lbl_total":             "إجمالي الربع الأول",

        # --- Raw data expanders ---
        "exp_monthly_qty":       "الإنتاج الشهري حسب القطاع — الكمية (طن)",
        "exp_monthly_val":       "الإنتاج الشهري حسب القطاع — القيمة (ألف ريال عماني)",
        "exp_gov_qty":           "الصيد الحرفي حسب المحافظات — الكمية (طن)",
        "exp_vessels":           "الصيد التجاري — جميع السفن",

        # --- Misc ---
        "error_db":              "تعذّر الاتصال بقاعدة البيانات. تحقق من مسار الملف وبرنامج تشغيل ODBC.",
        "no_data":               "لا توجد بيانات.",
        "currency_note":         "القيم بألف ريال عماني",
        "data_source":           "المصدر: التقرير المرحلي للربع الأول 2026",
        # Share panel
        "share_header":          "🔗 مشاركة اللوحة",
        "share_note":            "للعرض فقط · يفتح بنفس اللغة",
        "share_copied":          "✅ الرابط جاهز — انسخه وشاركه",
    },
}


# --------------------------------------------------------------------------- #
# 3. HELPER FUNCTIONS                                                          #
# --------------------------------------------------------------------------- #

def t(key: str) -> str:
    """Return the translated string for the current language."""
    lang = st.session_state.get("lang", "en")
    return T.get(lang, T["en"]).get(key, key)


def month_label(month_no: int) -> str:
    """Return the month name in the current language."""
    lang = st.session_state.get("lang", "en")
    return MONTHS.get(month_no, {}).get(lang, str(month_no))


def fmt_num(value, decimals: int = 1) -> str:
    """Format a number with thousands separator."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    return f"{value:,.{decimals}f}"


def fmt_pct(value) -> str:
    """Format a growth value (decimal fraction) as a percentage string."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    sign = "+" if value >= 0 else ""
    return f"{sign}{value * 100:.0f}%"


def inject_css(lang: str) -> None:
    """
    Inject global CSS + direction switching.

    RTL strategy (Arabic mode):
      1. Set direction:rtl on .stApp — the outermost Streamlit wrapper.
         This cascades direction to every child automatically.
      2. Explicitly right-align text in all known Streamlit 1.5x containers
         (data-testid selectors are stable across emotion-cache renames).
      3. Pin Plotly chart containers back to ltr so chart axes stay correct.
      4. Custom HTML elements (.report-header, .section-header, .source-note)
         are written with an inline dir= attribute (see section_header /
         report header below), so they flip independently of the cascade.

    LTR mode (English): the direction block is simply absent — browser
    defaults to ltr so no explicit reset is needed.
    """
    is_rtl  = (lang == "ar")
    ta      = "right" if is_rtl else "left"       # text-align shorthand
    hdr_ta  = "right" if is_rtl else "center"     # report-header alignment
    note_ta = "right" if is_rtl else "center"     # footer alignment
    grad    = ("270deg" if is_rtl                  # gradient flips for RTL
               else "90deg")

    rtl_block = ""
    if is_rtl:
        rtl_block = f"""
        /* ============================================================
           GLOBAL RTL — Arabic mode
           All rules use !important to override Streamlit's own styles.
           ============================================================ */

        /* 1. Root app wrapper: set direction here and let it cascade */
        .stApp {{
            direction: rtl !important;
        }}

        /* 2. Main content containers */
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"],
        [data-testid="stMainBlockContainer"],
        .main, .block-container,
        .main .block-container {{
            direction: rtl !important;
        }}

        /* 3. Every vertical / horizontal block */
        [data-testid="stVerticalBlock"],
        [data-testid="stVerticalBlockBorderWrapper"],
        [data-testid="stHorizontalBlock"] {{
            direction: rtl !important;
        }}

        /* 4. Columns */
        [data-testid="column"] {{
            direction: rtl !important;
        }}

        /* 5. Sidebar */
        [data-testid="stSidebar"],
        [data-testid="stSidebarContent"],
        [data-testid="stSidebar"] .block-container {{
            direction: rtl !important;
            text-align: right !important;
        }}

        /* 6. All text / markdown */
        [data-testid="stMarkdownContainer"],
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li,
        [data-testid="stMarkdownContainer"] ul,
        [data-testid="stMarkdownContainer"] ol,
        p, li, ul, ol {{
            direction: rtl !important;
            text-align: right !important;
        }}

        /* 7. Headings (both Streamlit-wrapped and bare) */
        [data-testid="stHeading"],
        h1, h2, h3, h4, h5, h6 {{
            direction: rtl !important;
            text-align: right !important;
        }}

        /* 8. KPI metric cards — label, value, delta each get explicit RTL */
        [data-testid="metric-container"] {{
            direction: rtl !important;
            text-align: right !important;
        }}
        [data-testid="stMetricLabel"],
        [data-testid="metric-container"] label {{
            direction: rtl !important;
            text-align: right !important;
            display: block !important;
        }}
        [data-testid="stMetricValue"] {{
            direction: rtl !important;
            text-align: right !important;
            display: block !important;
        }}
        [data-testid="stMetricDelta"] {{
            direction: rtl !important;
            text-align: right !important;
            display: block !important;
        }}

        /* 9. Expanders */
        [data-testid="stExpander"],
        [data-testid="stExpander"] summary,
        [data-testid="stExpander"] > div {{
            direction: rtl !important;
            text-align: right !important;
        }}

        /* 10. Select box & radio */
        [data-testid="stSelectbox"] label,
        [data-testid="stRadio"] label,
        label {{
            direction: rtl !important;
            text-align: right !important;
        }}

        /* 11. Captions */
        [data-testid="stCaptionContainer"],
        [data-testid="stCaption"],
        small, caption {{
            direction: rtl !important;
            text-align: right !important;
        }}

        /* 12. DataFrames — swap header/cell alignment */
        [data-testid="stDataFrameResizable"] th,
        [data-testid="stDataFrameResizable"] td,
        .stDataFrame th, .stDataFrame td,
        table th, table td {{
            direction: rtl !important;
            text-align: right !important;
        }}

        /* 13. Code / pre blocks stay LTR (paths, filenames, numbers) */
        code, pre, [data-testid="stCode"] {{
            direction: ltr !important;
            text-align: left !important;
        }}

        /* 14. Plotly charts: pin to LTR — Plotly handles its own RTL
               via the layout settings set in each chart builder function */
        .js-plotly-plot, .plotly, [data-testid="stPlotlyChart"] {{
            direction: ltr !important;
        }}
        """

    st.markdown(f"""
    <style>
    /* ============================================================
       BASE STYLES — language-independent
       ============================================================ */

    /* KPI Cards */
    [data-testid="metric-container"] {{
        background: linear-gradient(135deg, #f8fbff 0%, #eaf3ff 100%);
        border-radius: 12px;
        border: 1px solid #d0e4f7;
        padding: 18px 20px;
        box-shadow: 0 2px 8px rgba(0,51,102,0.07);
    }}
    [data-testid="metric-container"] label {{
        color: #003366;
        font-weight: 600;
        font-size: 0.85rem;
        letter-spacing: 0.03em;
    }}
    [data-testid="stMetricValue"] {{
        color: #003366;
        font-size: 2rem !important;
        font-weight: 700;
    }}
    [data-testid="stMetricDelta"] {{
        font-size: 0.95rem;
        font-weight: 600;
    }}

    /* Section header bar — gradient direction aware */
    .section-header {{
        background: linear-gradient({grad}, #003366 0%, #0066CC 100%);
        color: white;
        padding: 10px 18px;
        border-radius: 8px;
        margin: 28px 0 16px 0;
        font-size: 1.05rem;
        font-weight: 700;
        text-align: {ta};
    }}

    /* Report header banner */
    .report-header {{
        background: linear-gradient(135deg, #002244 0%, #003366 50%, #0055AA 100%);
        color: white;
        padding: 28px 32px;
        border-radius: 14px;
        margin-bottom: 24px;
        text-align: {hdr_ta};
    }}
    .report-header h1 {{ color: white; margin: 0; font-size: 1.7rem; }}
    .report-header p  {{ color: #A8CFFF; margin: 6px 0 0 0; font-size: 1rem; }}

    /* Footer note */
    .source-note {{
        color: #7f8c8d;
        font-size: 0.78rem;
        text-align: {note_ta};
        margin-top: 32px;
        padding-top: 10px;
        border-top: 1px solid #ecf0f1;
    }}

    {rtl_block}
    </style>
    """, unsafe_allow_html=True)


def section_header(text: str) -> None:
    """Render a styled section header, with dir attribute set from current lang."""
    lang = st.session_state.get("lang", "en")
    direction = "rtl" if lang == "ar" else "ltr"
    st.markdown(
        f'<div class="section-header" dir="{direction}">{text}</div>',
        unsafe_allow_html=True,
    )


def growth_delta(pct_fraction) -> tuple:
    """Return (delta_string, delta_color_indicator) for st.metric."""
    if pct_fraction is None or (isinstance(pct_fraction, float) and pd.isna(pct_fraction)):
        return None, None
    label = fmt_pct(pct_fraction) + f" {t('vs_2025')}"
    return label, pct_fraction


# --------------------------------------------------------------------------- #
# 4. DATABASE QUERIES (all read-only, cached for 10 minutes)                  #
# --------------------------------------------------------------------------- #

@contextmanager
def _get_connection():
    """
    Context-manager that yields a read-only DB connection and always closes it.
    Uses sqlite3 on Streamlit Cloud / Linux; pyodbc on Windows with Access.
    """
    if USE_SQLITE:
        # sqlite3 read-only URI mode — no writes possible
        conn = sqlite3.connect(_SQLITE_URI, uri=True, check_same_thread=False)
    else:
        conn = pyodbc.connect(_ODBC_CONN_STR, readonly=True)
    try:
        yield conn
    finally:
        conn.close()


def _query(sql: str, conn) -> pd.DataFrame:
    """
    Execute a SELECT query and return a DataFrame.
    Uses cursor directly instead of pd.read_sql() to avoid the
    'non-SQLAlchemy connection' UserWarning from pandas.
    All queries are strictly SELECT-only (read-only enforcement).
    """
    cursor = conn.cursor()
    cursor.execute(sql)
    cols = [d[0] for d in cursor.description]
    rows = cursor.fetchall()
    cursor.close()
    return pd.DataFrame.from_records(rows, columns=cols)


@st.cache_data(ttl=600, show_spinner=False)
def load_report_meta() -> dict:
    """Load the single ReportMeta row."""
    sql = "SELECT * FROM [ReportMeta]"
    with _get_connection() as cx:
        df = _query(sql, cx)
    if df.empty:
        return {}
    return df.iloc[0].to_dict()


@st.cache_data(ttl=600, show_spinner=False)
def load_sector_summary() -> pd.DataFrame:
    """Load FactSectorSummary joined with DimSector."""
    sql = """
        SELECT
            s.[SectorCode],
            s.[NameAr],
            s.[NameEn],
            s.[SortOrder],
            f.[FiscalYear],
            f.[QuantityTons],
            f.[ValueThousandOMR],
            f.[QtyGrowthPct],
            f.[ValueGrowthPct]
        FROM [FactSectorSummary] AS f
        INNER JOIN [DimSector] AS s ON f.[SectorID] = s.[SectorID]
        ORDER BY s.[SortOrder], f.[FiscalYear]
    """
    with _get_connection() as cx:
        return _query(sql, cx)


@st.cache_data(ttl=600, show_spinner=False)
def load_gov_summary() -> pd.DataFrame:
    """Load FactGovernorateSummary joined with DimGovernorate."""
    sql = """
        SELECT
            g.[GovCode],
            g.[NameAr],
            g.[NameEn],
            g.[SortOrder],
            f.[FiscalYear],
            f.[QuantityTons],
            f.[ValueThousandOMR],
            f.[QtyGrowthPct],
            f.[ValueGrowthPct]
        FROM [FactGovernorateSummary] AS f
        INNER JOIN [DimGovernorate] AS g ON f.[GovernorateID] = g.[GovernorateID]
        ORDER BY g.[SortOrder], f.[FiscalYear]
    """
    with _get_connection() as cx:
        return _query(sql, cx)


@st.cache_data(ttl=600, show_spinner=False)
def load_overall_monthly() -> pd.DataFrame:
    """Load FactOverallMonthly (both years)."""
    sql = """
        SELECT [FiscalYear], [MonthNo], [QuantityThousandTons]
        FROM [FactOverallMonthly]
        ORDER BY [FiscalYear], [MonthNo]
    """
    with _get_connection() as cx:
        return _query(sql, cx)


@st.cache_data(ttl=600, show_spinner=False)
def load_monthly_production() -> pd.DataFrame:
    """Load FactMonthlyProduction (sector totals only — GovernorateID IS NULL)."""
    sql = """
        SELECT
            mp.[FiscalYear],
            mp.[MonthNo],
            s.[SectorCode],
            s.[NameAr],
            s.[NameEn],
            s.[SortOrder],
            mp.[QuantityTons],
            mp.[ValueThousandOMR]
        FROM [FactMonthlyProduction] AS mp
        INNER JOIN [DimSector] AS s ON mp.[SectorID] = s.[SectorID]
        WHERE mp.[GovernorateID] IS NULL
        ORDER BY s.[SortOrder], mp.[MonthNo]
    """
    with _get_connection() as cx:
        return _query(sql, cx)


@st.cache_data(ttl=600, show_spinner=False)
def load_gov_monthly() -> pd.DataFrame:
    """Load FactMonthlyProduction for artisanal-by-governorate rows."""
    sql = """
        SELECT
            mp.[FiscalYear],
            mp.[MonthNo],
            g.[GovCode],
            g.[NameAr],
            g.[NameEn],
            g.[SortOrder],
            mp.[QuantityTons],
            mp.[ValueThousandOMR]
        FROM [FactMonthlyProduction] AS mp
        INNER JOIN [DimGovernorate] AS g ON mp.[GovernorateID] = g.[GovernorateID]
        WHERE mp.[GovernorateID] IS NOT NULL
        ORDER BY g.[SortOrder], mp.[MonthNo]
    """
    with _get_connection() as cx:
        return _query(sql, cx)


@st.cache_data(ttl=600, show_spinner=False)
def load_aqua_species() -> pd.DataFrame:
    """Load FactAquacultureSpecies for the current report year."""
    sql = """
        SELECT [SpeciesNameAr], [SpeciesNameEn], [QuantityTons], [ValueThousandOMR]
        FROM [FactAquacultureSpecies]
        WHERE [FiscalYear] = 2026
        ORDER BY [QuantityTons] DESC
    """
    with _get_connection() as cx:
        return _query(sql, cx)


@st.cache_data(ttl=600, show_spinner=False)
def load_top_species() -> pd.DataFrame:
    """Load FactArtisanalTopSpecies."""
    sql = """
        SELECT [SpeciesNameAr], [SpeciesNameEn], [SharePct], [QuantityTons]
        FROM [FactArtisanalTopSpecies]
        WHERE [FiscalYear] = 2026
        ORDER BY [SharePct] DESC
    """
    with _get_connection() as cx:
        return _query(sql, cx)


@st.cache_data(ttl=600, show_spinner=False)
def load_vessels() -> pd.DataFrame:
    """Load FactCommercialVessel ordered by production descending."""
    sql = """
        SELECT [VesselName], [ProductionTons], [Trips]
        FROM [FactCommercialVessel]
        WHERE [FiscalYear] = 2026
        ORDER BY [ProductionTons] DESC
    """
    with _get_connection() as cx:
        return _query(sql, cx)


# --------------------------------------------------------------------------- #
# 5. CHART BUILDERS                                                            #
# --------------------------------------------------------------------------- #

def _fig_layout(fig, title: str, lang: str, height: int = 400) -> go.Figure:
    """Apply shared layout settings (font, margins, RTL title alignment)."""
    title_x = 0.98 if lang == "ar" else 0.02
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=14, color=C_NAVY, family=CHART_FONT["family"]),
            x=title_x,
            xanchor="right" if lang == "ar" else "left",
        ),
        font=CHART_FONT,
        height=height,
        margin=dict(l=40, r=40, t=50, b=40),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.25,
            xanchor="right" if lang == "ar" else "left",
            x=1 if lang == "ar" else 0,
        ),
    )
    return fig


def chart_overall_monthly(df_monthly: pd.DataFrame, lang: str) -> go.Figure:
    """Grouped bar chart: overall monthly production 2025 vs 2026."""
    df25 = df_monthly[df_monthly["FiscalYear"] == 2025].sort_values("MonthNo")
    df26 = df_monthly[df_monthly["FiscalYear"] == 2026].sort_values("MonthNo")
    months = [month_label(m) for m in df26["MonthNo"]]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name=t("lbl_2025"), x=months, y=df25["QuantityThousandTons"].values,
        marker_color=C_LBLUE, text=[f"{v:.1f}" for v in df25["QuantityThousandTons"].values],
        textposition="outside",
    ))
    fig.add_trace(go.Bar(
        name=t("lbl_2026"), x=months, y=df26["QuantityThousandTons"].values,
        marker_color=C_NAVY, text=[f"{v:.1f}" for v in df26["QuantityThousandTons"].values],
        textposition="outside",
    ))
    fig.update_layout(barmode="group", yaxis_title=t("axis_qty_ktons"))
    return _fig_layout(fig, t("chart_monthly_overall"), lang, height=380)


def chart_sector_donut(df_sector: pd.DataFrame, lang: str, year: int = 2026) -> go.Figure:
    """Donut chart: production share by sector for a given year."""
    name_col = "NameAr" if lang == "ar" else "NameEn"
    df = df_sector[df_sector["FiscalYear"] == year].copy()
    df = df.sort_values("SortOrder")
    total = df["QuantityTons"].sum()
    df["SharePct"] = df["QuantityTons"] / total * 100

    fig = go.Figure(go.Pie(
        labels=df[name_col].tolist(),
        values=df["QuantityTons"].tolist(),
        hole=0.48,
        marker=dict(colors=SECTOR_COLORS, line=dict(color="white", width=2)),
        textinfo="label+percent",
        hovertemplate="%{label}<br>%{value:,.0f} t<br>%{percent}<extra></extra>",
        textfont=dict(size=11),
        direction="clockwise",
    ))
    fig.add_annotation(
        text=f"<b>{year}</b>", x=0.5, y=0.5,
        font=dict(size=18, color=C_NAVY), showarrow=False,
    )
    return _fig_layout(fig, t("chart_sector_donut"), lang, height=380)


def chart_sector_bar(df_sector: pd.DataFrame, lang: str) -> go.Figure:
    """Grouped horizontal bar: sector production 2025 vs 2026."""
    name_col = "NameAr" if lang == "ar" else "NameEn"
    df25 = df_sector[df_sector["FiscalYear"] == 2025].sort_values("SortOrder")
    df26 = df_sector[df_sector["FiscalYear"] == 2026].sort_values("SortOrder")
    sectors = df26[name_col].tolist()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name=t("lbl_2025"), y=sectors, x=(df25["QuantityTons"] / 1000).values,
        orientation="h", marker_color=C_LBLUE,
    ))
    fig.add_trace(go.Bar(
        name=t("lbl_2026"), y=sectors, x=(df26["QuantityTons"] / 1000).values,
        orientation="h", marker_color=C_NAVY,
    ))
    fig.update_layout(barmode="group", xaxis_title=t("axis_qty_ktons"))
    return _fig_layout(fig, t("chart_sector_bar"), lang, height=320)


def chart_gov_bar(df_gov: pd.DataFrame, lang: str) -> go.Figure:
    """Grouped bar: artisanal production by governorate 2025 vs 2026."""
    name_col = "NameAr" if lang == "ar" else "NameEn"
    df25 = df_gov[df_gov["FiscalYear"] == 2025].sort_values("SortOrder")
    df26 = df_gov[df_gov["FiscalYear"] == 2026].sort_values("SortOrder")
    govs = df26[name_col].tolist()

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name=t("lbl_2025"), x=govs, y=(df25["QuantityTons"] / 1000).values,
        marker_color=C_LBLUE,
    ))
    fig.add_trace(go.Bar(
        name=t("lbl_2026"), x=govs, y=(df26["QuantityTons"] / 1000).values,
        marker_color=C_NAVY,
        text=[fmt_pct(g) for g in df26["QtyGrowthPct"].values],
        textposition="outside",
    ))
    fig.update_layout(barmode="group", yaxis_title=t("axis_qty_ktons"))
    return _fig_layout(fig, t("chart_gov_bar"), lang, height=380)


def chart_gov_donut(df_gov: pd.DataFrame, lang: str) -> go.Figure:
    """Donut chart: artisanal governorate share for 2026."""
    name_col = "NameAr" if lang == "ar" else "NameEn"
    df = df_gov[df_gov["FiscalYear"] == 2026].sort_values("SortOrder")

    fig = go.Figure(go.Pie(
        labels=df[name_col].tolist(),
        values=df["QuantityTons"].tolist(),
        hole=0.45,
        marker=dict(colors=GOV_COLORS, line=dict(color="white", width=2)),
        textinfo="label+percent",
        textfont=dict(size=10),
        hovertemplate="%{label}<br>%{value:,.0f} t<br>%{percent}<extra></extra>",
    ))
    return _fig_layout(fig, t("chart_gov_donut"), lang, height=380)


def chart_top_species(df_species: pd.DataFrame, lang: str) -> go.Figure:
    """Horizontal bar: top-5 artisanal species by share."""
    name_col = "SpeciesNameAr" if lang == "ar" else "SpeciesNameEn"
    df = df_species.sort_values("SharePct")

    fig = go.Figure(go.Bar(
        x=(df["SharePct"] * 100).values,
        y=df[name_col].tolist(),
        orientation="h",
        marker_color=C_BLUE,
        text=[f"{v*100:.1f}%" for v in df["SharePct"].values],
        textposition="outside",
    ))
    fig.update_layout(xaxis_title=t("axis_share_pct"))
    return _fig_layout(fig, t("chart_top_species"), lang, height=320)


def chart_aqua_species(df_aqua: pd.DataFrame, lang: str) -> go.Figure:
    """Bar chart: aquaculture species — quantity and value."""
    name_col = "SpeciesNameAr" if lang == "ar" else "SpeciesNameEn"
    df = df_aqua[df_aqua["QuantityTons"] > 0].sort_values("QuantityTons", ascending=False)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name=t("lbl_quantity") + " (t)",
        x=df[name_col].tolist(),
        y=df["QuantityTons"].values,
        marker_color=C_TEAL,
        yaxis="y",
    ))
    fig.add_trace(go.Bar(
        name=t("lbl_value") + " (k OMR)",
        x=df[name_col].tolist(),
        y=df["ValueThousandOMR"].values,
        marker_color=C_GREEN,
        yaxis="y2",
    ))
    fig.update_layout(
        barmode="group",
        yaxis=dict(title=t("axis_qty_tons")),
        yaxis2=dict(title=t("axis_val_komr"), overlaying="y", side="right"),
    )
    return _fig_layout(fig, t("chart_aqua_species"), lang, height=380)


def chart_vessels(df_vessels: pd.DataFrame, lang: str, top_n: int = 15) -> go.Figure:
    """Horizontal bar: top-N vessels by production."""
    df = df_vessels.dropna(subset=["ProductionTons"]).nlargest(top_n, "ProductionTons")
    df = df.sort_values("ProductionTons")  # ascending for horizontal bar readability

    fig = go.Figure(go.Bar(
        x=df["ProductionTons"].values,
        y=df["VesselName"].tolist(),
        orientation="h",
        marker=dict(
            color=df["ProductionTons"].values,
            colorscale=[[0, C_LBLUE], [1, C_NAVY]],
            showscale=False,
        ),
        text=[f"{v:,.1f}" for v in df["ProductionTons"].values],
        textposition="outside",
    ))
    fig.update_layout(xaxis_title=t("lbl_production"))
    return _fig_layout(fig, t("chart_vessels"), lang, height=480)


# --------------------------------------------------------------------------- #
# 6. SIDEBAR                                                                   #
# --------------------------------------------------------------------------- #

def _build_share_url(lang: str) -> str:
    """
    Build the full shareable URL including ?lang=.
    Uses the Referer header (most reliable source of the full URL in Streamlit)
    and falls back to the Host header, then a relative URL.
    """
    try:
        headers = st.context.headers
        # Referer gives us the full current URL including protocol + host
        referer = headers.get("Referer", "") or headers.get("referer", "")
        if referer:
            # Strip any existing query string, then append ?lang=
            base = referer.split("?")[0].rstrip("/")
            return f"{base}?lang={lang}"
        # Fallback: build from Host header
        host = headers.get("Host", "") or headers.get("host", "")
        if host:
            proto = "https" if ("streamlit.app" in host or ":443" in host) else "http"
            return f"{proto}://{host}?lang={lang}"
    except Exception:
        pass
    # Last resort — relative URL (user can prepend their host)
    return f"?lang={lang}"


def render_sidebar() -> str:
    """
    Render sidebar controls.  Returns the active language code ('en' or 'ar').

    View-only share link
    --------------------
    The current language is synced to the URL query parameter ?lang=.
    Anyone opening that URL gets the dashboard in the same language.
    The dashboard is inherently view-only: all SQL is SELECT-only and the
    DB connection is opened with ReadOnly=1 / SQLite read-only URI mode.
    """
    with st.sidebar:
        st.markdown(f"### {t('sidebar_title')}")
        st.divider()

        # ── Language toggle ──────────────────────────────────────────────────
        lang_choice = st.radio(
            t("language_label"),
            options=["English", "العربية"],
            index=0 if st.session_state.get("lang", "en") == "en" else 1,
            horizontal=True,
            key="lang_radio",
        )
        lang = "ar" if lang_choice == "العربية" else "en"
        st.session_state["lang"] = lang

        # Sync language into the browser URL (?lang=ar / ?lang=en).
        # This makes the address bar always reflect the current language
        # so copying the URL gives a ready-to-share link.
        try:
            if st.query_params.get("lang") != lang:
                st.query_params["lang"] = lang
        except Exception:
            pass  # query_params unavailable in some embedded environments

        st.divider()

        # ── Database / cache info ────────────────────────────────────────────
        st.caption(f"📁 {t('db_path_label')}")
        st.code(os.path.basename(DB_PATH), language=None)
        backend_label = "SQLite" if USE_SQLITE else "Access / pyodbc"
        st.caption(f"🔌 {backend_label}")
        st.caption(f"🔄 {t('data_refresh')}")

        st.divider()

        # ── View-only share panel ────────────────────────────────────────────
        st.markdown(f"**{t('share_header')}**")

        share_url = _build_share_url(lang)

        # Render URL in a plain HTML monospace box — no widget state, always
        # reflects the current lang.  User clicks → Ctrl/Cmd+A → Ctrl/Cmd+C.
        st.markdown(
            f'<div style="background:#eef2f7;border:1px solid #c9d9ee;'
            f'border-radius:6px;padding:8px 12px;font-family:monospace;'
            f'font-size:11.5px;word-break:break-all;direction:ltr;'
            f'color:#1a2a3a;user-select:all;cursor:text;">'
            f'{share_url}</div>',
            unsafe_allow_html=True,
        )
        st.caption(f"👁️ {t('share_note')}")

        st.divider()
        st.caption(t("data_source"))
        st.caption(t("currency_note"))

    return lang


# --------------------------------------------------------------------------- #
# 7. MAIN DASHBOARD                                                            #
# --------------------------------------------------------------------------- #

def main():
    # --- Page configuration -------------------------------------------------
    st.set_page_config(
        page_title="Fisheries Dashboard | لوحة القطاع السمكي",
        page_icon="🐠",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ── Language initialisation ──────────────────────────────────────────────
    # Priority order:
    #   1. ?lang=ar/en in the URL  ← shared links land here with the right lang
    #   2. st.session_state["lang"] preserved across rerenders
    #   3. Default: "en"
    if "lang" not in st.session_state:
        try:
            url_lang = st.query_params.get("lang", "en")
            st.session_state["lang"] = url_lang if url_lang in ("en", "ar") else "en"
        except Exception:
            st.session_state["lang"] = "en"

    # Render sidebar first (sets session_state lang)
    lang = render_sidebar()

    # Inject language-aware CSS
    inject_css(lang)

    # --- Database connectivity check ----------------------------------------
    try:
        with _get_connection() as _test:
            pass  # just verify the connection opens cleanly
    except Exception as e:
        st.error(f"⚠️  {t('error_db')}\n\n`{e}`")
        st.info(f"DB_PATH = `{DB_PATH}`  (backend: {'SQLite' if USE_SQLITE else 'Access/pyodbc'})")
        st.stop()

    # --- Load all data ------------------------------------------------------
    with st.spinner(""):
        meta         = load_report_meta()
        df_sector    = load_sector_summary()
        df_gov       = load_gov_summary()
        df_overall   = load_overall_monthly()
        df_monthly   = load_monthly_production()
        df_gov_mo    = load_gov_monthly()
        df_aqua      = load_aqua_species()
        df_top_sp    = load_top_species()
        df_vessels   = load_vessels()

    # --- Report header ------------------------------------------------------
    title_key = "ReportTitleAr" if lang == "ar" else "ReportTitleEn"
    period_key = "PeriodLabelAr" if lang == "ar" else "PeriodLabelEn"
    report_title  = meta.get(title_key, t("page_title"))
    period_label  = meta.get(period_key, t("quarter_badge"))

    _dir = "rtl" if lang == "ar" else "ltr"
    st.markdown(
        f'<div class="report-header" dir="{_dir}">'
        f'<h1>{report_title}</h1>'
        f'<p>{period_label}</p>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ======================================================================= #
    #  SECTION A — KPI CARDS                                                  #
    # ======================================================================= #
    # Source totals (2026):
    #   Total qty  = 259,264 t  → 259.3 k tons (▼3%)
    #   Total val  = 167,046 k OMR → 167.0 M OMR (▼9%)
    #   Artisanal share = 198,471 / 259,264 = 76.6%
    #   Aquaculture qty growth = +40%
    # These are derived from the loaded data to stay live.
    # ----------------------------------------------------------------------- #

    # Compute live KPIs from loaded DataFrames
    sec_2026 = df_sector[df_sector["FiscalYear"] == 2026]
    sec_2025 = df_sector[df_sector["FiscalYear"] == 2025]

    total_qty_2026 = sec_2026["QuantityTons"].sum()
    total_qty_2025 = sec_2025["QuantityTons"].sum()
    total_val_2026 = sec_2026["ValueThousandOMR"].sum()
    total_val_2025 = sec_2025["ValueThousandOMR"].sum()

    artisanal_share = (
        sec_2026.loc[sec_2026["SectorCode"] == "ARTISANAL", "QuantityTons"].sum()
        / total_qty_2026
        if total_qty_2026 > 0 else 0
    )
    aqua_row = sec_2026[sec_2026["SectorCode"] == "AQUACULTURE"]
    aqua_growth = aqua_row["QtyGrowthPct"].values[0] if not aqua_row.empty else None

    qty_growth  = (total_qty_2026 - total_qty_2025) / total_qty_2025 if total_qty_2025 else None
    val_growth  = (total_val_2026 - total_val_2025) / total_val_2025 if total_val_2025 else None

    kpi_cols = st.columns(4, gap="medium")

    with kpi_cols[0]:
        st.metric(
            label=t("kpi_total_qty"),
            value=f"{total_qty_2026 / 1000:,.1f} {t('unit_k_tons')}",
            delta=fmt_pct(qty_growth) + f" {t('vs_2025')}",
            delta_color="inverse",   # red when negative (production fell)
        )

    with kpi_cols[1]:
        st.metric(
            label=t("kpi_total_val"),
            value=f"{total_val_2026 / 1000:,.1f} {t('unit_m_omr')}",
            delta=fmt_pct(val_growth) + f" {t('vs_2025')}",
            delta_color="inverse",
        )

    with kpi_cols[2]:
        st.metric(
            label=t("kpi_artisanal_share"),
            value=f"{artisanal_share * 100:.1f} {t('unit_pct')}",
            delta=None,
        )

    with kpi_cols[3]:
        st.metric(
            label=t("kpi_aqua_growth"),
            value=fmt_pct(aqua_growth),
            delta=None,
        )

    # ======================================================================= #
    #  SECTION B — MONTHLY OVERVIEW                                           #
    # ======================================================================= #
    section_header(t("sec_overview"))

    st.plotly_chart(
        chart_overall_monthly(df_overall, lang),
        width="stretch",
    )

    # ======================================================================= #
    #  SECTION C — SECTOR BREAKDOWN                                           #
    # ======================================================================= #
    section_header(t("sec_sectors"))

    col_left, col_right = st.columns([1.1, 1], gap="large")

    with col_left:
        # Year toggle for the donut chart
        donut_year = st.selectbox(
            label=t("lbl_2026"),
            options=[2026, 2025],
            index=0,
            label_visibility="collapsed",
        )
        st.plotly_chart(
            chart_sector_donut(df_sector, lang, year=donut_year),
            width="stretch",
        )

    with col_right:
        st.plotly_chart(
            chart_sector_bar(df_sector, lang),
            width="stretch",
        )

    # Sector summary table
    name_col = "NameAr" if lang == "ar" else "NameEn"
    sec_table = df_sector[df_sector["FiscalYear"] == 2026][[
        name_col, "QuantityTons", "ValueThousandOMR", "QtyGrowthPct", "ValueGrowthPct"
    ]].copy()
    sec_table.columns = [
        t("lbl_sector"),
        t("lbl_quantity") + " (t)",
        t("lbl_value") + " (k OMR)",
        t("lbl_growth") + " (Qty)",
        t("lbl_growth") + " (Val)",
    ]
    sec_table[t("lbl_growth") + " (Qty)"] = sec_table[t("lbl_growth") + " (Qty)"].apply(fmt_pct)
    sec_table[t("lbl_growth") + " (Val)"] = sec_table[t("lbl_growth") + " (Val)"].apply(fmt_pct)
    st.dataframe(sec_table.set_index(t("lbl_sector")), width="stretch")

    # ======================================================================= #
    #  SECTION D — ARTISANAL FISHING                                          #
    # ======================================================================= #
    section_header(t("sec_artisanal"))

    st.plotly_chart(chart_gov_bar(df_gov, lang), width="stretch")

    col_d1, col_d2 = st.columns(2, gap="large")
    with col_d1:
        st.plotly_chart(chart_gov_donut(df_gov, lang), width="stretch")
    with col_d2:
        st.plotly_chart(chart_top_species(df_top_sp, lang), width="stretch")

    # ======================================================================= #
    #  SECTION E — AQUACULTURE                                                #
    # ======================================================================= #
    section_header(t("sec_aquaculture"))

    col_e1, col_e2 = st.columns([2, 1], gap="large")
    with col_e1:
        st.plotly_chart(chart_aqua_species(df_aqua, lang), width="stretch")
    with col_e2:
        # Species detail table
        sp_name_col = "SpeciesNameAr" if lang == "ar" else "SpeciesNameEn"
        sp_table = df_aqua[[sp_name_col, "QuantityTons", "ValueThousandOMR"]].copy()
        sp_table.columns = [t("lbl_species"), t("lbl_quantity") + " (t)", t("lbl_value") + " (k OMR)"]
        st.dataframe(sp_table.set_index(t("lbl_species")), width="stretch")

    # ======================================================================= #
    #  SECTION F — COMMERCIAL FISHING                                         #
    # ======================================================================= #
    section_header(t("sec_commercial"))

    st.plotly_chart(chart_vessels(df_vessels, lang, top_n=15), width="stretch")

    # Full vessels table in expander
    with st.expander(t("chart_vessels_full")):
        v_table = df_vessels[["VesselName", "ProductionTons"]].copy()
        v_table.columns = [t("lbl_vessel"), t("lbl_production")]
        st.dataframe(v_table, width="stretch")

    # ======================================================================= #
    #  SECTION G — RAW DATA EXPANDERS                                         #
    # ======================================================================= #
    section_header(t("sec_raw"))

    # Monthly quantity pivot (sector totals by month)
    with st.expander(t("exp_monthly_qty")):
        if not df_monthly.empty:
            nm = "NameAr" if lang == "ar" else "NameEn"
            pivot_qty = df_monthly[df_monthly["FiscalYear"] == 2026].pivot_table(
                index=nm, columns="MonthNo", values="QuantityTons", aggfunc="sum"
            )
            pivot_qty.columns = [month_label(m) for m in pivot_qty.columns]
            pivot_qty.index.name = t("lbl_sector")
            pivot_qty["Q1 Total"] = pivot_qty.sum(axis=1)
            st.dataframe(pivot_qty.style.format("{:,.0f}"), width="stretch")

    # Monthly value pivot
    with st.expander(t("exp_monthly_val")):
        if not df_monthly.empty:
            nm = "NameAr" if lang == "ar" else "NameEn"
            pivot_val = df_monthly[df_monthly["FiscalYear"] == 2026].pivot_table(
                index=nm, columns="MonthNo", values="ValueThousandOMR", aggfunc="sum"
            )
            pivot_val.columns = [month_label(m) for m in pivot_val.columns]
            pivot_val.index.name = t("lbl_sector")
            pivot_val["Q1 Total"] = pivot_val.sum(axis=1)
            st.dataframe(pivot_val.style.format("{:,.0f}"), width="stretch")

    # Artisanal governorate monthly quantity pivot
    with st.expander(t("exp_gov_qty")):
        if not df_gov_mo.empty:
            nm = "NameAr" if lang == "ar" else "NameEn"
            pivot_gov = df_gov_mo[df_gov_mo["FiscalYear"] == 2026].pivot_table(
                index=nm, columns="MonthNo", values="QuantityTons", aggfunc="sum"
            )
            pivot_gov.columns = [month_label(m) for m in pivot_gov.columns]
            pivot_gov.index.name = t("lbl_governorate")
            pivot_gov["Q1 Total"] = pivot_gov.sum(axis=1)
            st.dataframe(pivot_gov.style.format("{:,.0f}"), width="stretch")

    # Full vessels table
    with st.expander(t("exp_vessels")):
        v_full = df_vessels[["VesselName", "ProductionTons"]].copy()
        v_full.columns = [t("lbl_vessel"), t("lbl_production")]
        st.dataframe(v_full, width="stretch")

    # --- Footer ------------------------------------------------------------
    st.markdown(
        f'<p class="source-note" dir="{_dir}">{t("data_source")} &nbsp;|&nbsp; {t("currency_note")}</p>',
        unsafe_allow_html=True,
    )


# --------------------------------------------------------------------------- #
# ENTRY POINT                                                                  #
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    main()
