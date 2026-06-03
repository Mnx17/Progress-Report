# =============================================================================
# build_db.py
# =============================================================================
# Creates and fully seeds the local Microsoft Access database
# (FisheriesQ1_2026.accdb) from the Q1-2026 Progress Report data.
#
# PREREQUISITES (Windows only):
#   1. pip install pywin32 pyodbc
#   2. Microsoft Access Database Engine 2016 Redistributable (64-bit)
#      — must match your Python bitness (both 64-bit OR both 32-bit).
#      Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920
#
# USAGE:
#   python build_db.py
#
# The script is IDEMPOTENT — re-running it drops and recreates the database.
# =============================================================================

import os
import sys
import pyodbc
import datetime

# Force UTF-8 on Windows console so Arabic path characters in print() don't crash
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr.encoding and sys.stderr.encoding.lower() not in ("utf-8", "utf8"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# --------------------------------------------------------------------------- #
# 1. CONFIGURATION                                                             #
# --------------------------------------------------------------------------- #

# Output database path — placed next to this script by default.
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "FisheriesQ1_2026.accdb")

# Access ODBC connection string (READ-WRITE here; app.py uses ReadOnly=1).
CONN_STR = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={DB_PATH};"
)


# --------------------------------------------------------------------------- #
# 2. CREATE EMPTY .accdb FILE VIA ADOX (Windows COM)                          #
# --------------------------------------------------------------------------- #

def create_accdb(path: str) -> None:
    """Create a blank .accdb file using the ADOX COM object."""
    try:
        import win32com.client  # requires: pip install pywin32
    except ImportError:
        sys.exit(
            "\nERROR: pywin32 is not installed.\n"
            "Run:  pip install pywin32\n"
            "Then re-run this script.\n"
        )

    if os.path.exists(path):
        os.remove(path)
        print(f"  Removed existing: {path}")

    cat = win32com.client.Dispatch("ADOX.Catalog")
    cat.Create(f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};")
    del cat
    print(f"  Created blank database: {path}")


# --------------------------------------------------------------------------- #
# 3. DDL — CREATE ALL TABLES                                                  #
# --------------------------------------------------------------------------- #
# Access DDL notes:
#  • AUTOINCREMENT (not AUTO_INCREMENT)
#  • Use [BracketedNames] to avoid reserved-word conflicts
#  • No FOREIGN KEY constraint enforcement needed (just logical relationships)
# --------------------------------------------------------------------------- #

DDL_STATEMENTS = [
    # --- Dimension: Fishing Sector (4 rows) ---
    """
    CREATE TABLE [DimSector] (
        [SectorID]   AUTOINCREMENT CONSTRAINT [PK_DimSector] PRIMARY KEY,
        [SectorCode] TEXT(20)   NOT NULL,
        [NameAr]     TEXT(100)  NOT NULL,
        [NameEn]     TEXT(100)  NOT NULL,
        [SortOrder]  INTEGER    NOT NULL
    )
    """,

    # --- Dimension: Governorate (6 rows, artisanal breakdown) ---
    """
    CREATE TABLE [DimGovernorate] (
        [GovernorateID] AUTOINCREMENT CONSTRAINT [PK_DimGovernorate] PRIMARY KEY,
        [GovCode]       TEXT(20)  NOT NULL,
        [NameAr]        TEXT(100) NOT NULL,
        [NameEn]        TEXT(100) NOT NULL,
        [SortOrder]     INTEGER   NOT NULL
    )
    """,

    # --- Fact: Sector-level summary per fiscal year (slides 1,3,6,7,8) ---
    """
    CREATE TABLE [FactSectorSummary] (
        [SummaryID]        AUTOINCREMENT CONSTRAINT [PK_FactSectorSummary] PRIMARY KEY,
        [SectorID]         INTEGER NOT NULL,
        [FiscalYear]       INTEGER NOT NULL,
        [QuantityTons]     DOUBLE  NOT NULL,
        [ValueThousandOMR] DOUBLE  NOT NULL,
        [QtyGrowthPct]     DOUBLE,
        [ValueGrowthPct]   DOUBLE
    )
    """,

    # --- Fact: Governorate summary per year — artisanal only (slides 10/11) ---
    """
    CREATE TABLE [FactGovernorateSummary] (
        [GovSummaryID]     AUTOINCREMENT CONSTRAINT [PK_FactGovSummary] PRIMARY KEY,
        [GovernorateID]    INTEGER NOT NULL,
        [FiscalYear]       INTEGER NOT NULL,
        [QuantityTons]     DOUBLE  NOT NULL,
        [ValueThousandOMR] DOUBLE  NOT NULL,
        [QtyGrowthPct]     DOUBLE,
        [ValueGrowthPct]   DOUBLE
    )
    """,

    # --- Fact: Monthly production by sector (and by governorate for artisanal) ---
    # GovernorateID is NULL for non-artisanal sectors and for artisanal totals.
    """
    CREATE TABLE [FactMonthlyProduction] (
        [MonthlyID]        AUTOINCREMENT CONSTRAINT [PK_FactMonthly] PRIMARY KEY,
        [FiscalYear]       INTEGER NOT NULL,
        [MonthNo]          INTEGER NOT NULL,
        [SectorID]         INTEGER NOT NULL,
        [GovernorateID]    INTEGER,
        [QuantityTons]     DOUBLE  NOT NULL,
        [ValueThousandOMR] DOUBLE  NOT NULL
    )
    """,

    # --- Fact: Overall monthly totals (both years, from embedded chart 1) ---
    # Stored in THOUSAND TONS (matching the chart's y-axis unit).
    """
    CREATE TABLE [FactOverallMonthly] (
        [OverallMonthlyID]    AUTOINCREMENT CONSTRAINT [PK_FactOverallMonthly] PRIMARY KEY,
        [FiscalYear]          INTEGER NOT NULL,
        [MonthNo]             INTEGER NOT NULL,
        [QuantityThousandTons] DOUBLE NOT NULL
    )
    """,

    # --- Fact: Aquaculture production by species (slide 9) ---
    # ValueThousandOMR converted from the slide's million-OMR display.
    """
    CREATE TABLE [FactAquacultureSpecies] (
        [AquaID]           AUTOINCREMENT CONSTRAINT [PK_FactAquaSpecies] PRIMARY KEY,
        [FiscalYear]       INTEGER    NOT NULL,
        [SpeciesNameAr]    TEXT(100)  NOT NULL,
        [SpeciesNameEn]    TEXT(100)  NOT NULL,
        [QuantityTons]     DOUBLE     NOT NULL,
        [ValueThousandOMR] DOUBLE     NOT NULL
    )
    """,

    # --- Fact: Top-5 artisanal species by share (embedded chart 3) ---
    """
    CREATE TABLE [FactArtisanalTopSpecies] (
        [TopSpeciesID]  AUTOINCREMENT CONSTRAINT [PK_TopSpecies] PRIMARY KEY,
        [FiscalYear]    INTEGER    NOT NULL,
        [SpeciesNameAr] TEXT(100)  NOT NULL,
        [SpeciesNameEn] TEXT(100)  NOT NULL,
        [SharePct]      DOUBLE     NOT NULL,
        [QuantityTons]  DOUBLE
    )
    """,

    # --- Fact: Commercial fishing production by vessel (slides 12/13) ---
    """
    CREATE TABLE [FactCommercialVessel] (
        [VesselID]       AUTOINCREMENT CONSTRAINT [PK_FactCommercialVessel] PRIMARY KEY,
        [FiscalYear]     INTEGER   NOT NULL,
        [VesselName]     TEXT(150) NOT NULL,
        [ProductionTons] DOUBLE,
        [Trips]          INTEGER
    )
    """,

    # --- Report metadata (1 row per report edition) ---
    """
    CREATE TABLE [ReportMeta] (
        [MetaID]         AUTOINCREMENT CONSTRAINT [PK_ReportMeta] PRIMARY KEY,
        [ReportTitleAr]  TEXT(255) NOT NULL,
        [ReportTitleEn]  TEXT(255) NOT NULL,
        [PeriodLabelAr]  TEXT(100) NOT NULL,
        [PeriodLabelEn]  TEXT(100) NOT NULL,
        [FiscalYear]     INTEGER   NOT NULL,
        [QuarterNo]      INTEGER   NOT NULL,
        [CurrencyAr]     TEXT(50)  NOT NULL,
        [CurrencyEn]     TEXT(50)  NOT NULL,
        [GeneratedOn]    DATETIME  NOT NULL
    )
    """,
]


# --------------------------------------------------------------------------- #
# 4. SEED DATA — extracted from Q1 2026 Progress Report slides                #
# --------------------------------------------------------------------------- #
# All quantities: TONS (exact, from slides 10/11 authoritative tables).
# All values:     THOUSAND OMR (slides 10/11 detail tables).
# Growth:         decimal fraction (e.g. 0.05 = 5%), NULL for base year.
# --------------------------------------------------------------------------- #

SEED_DATA = {

    # ---- DimSector ---------------------------------------------------------
    "DimSector": {
        "sql": "INSERT INTO [DimSector] ([SectorCode],[NameAr],[NameEn],[SortOrder]) VALUES (?,?,?,?)",
        "rows": [
            ("ARTISANAL",    "الصيد الحرفي",       "Artisanal Fishing",  1),
            ("COASTAL",      "الصيد الساحلي",      "Coastal Fishing",    2),
            ("COMMERCIAL",   "الصيد التجاري",      "Commercial Fishing", 3),
            ("AQUACULTURE",  "الاستزراع السمكي",   "Aquaculture",        4),
        ],
    },

    # ---- DimGovernorate ----------------------------------------------------
    "DimGovernorate": {
        "sql": "INSERT INTO [DimGovernorate] ([GovCode],[NameAr],[NameEn],[SortOrder]) VALUES (?,?,?,?)",
        "rows": [
            ("MUSANDAM",     "مسندم",                 "Musandam",            1),
            ("BATINAH",      "شمال وجنوب الباطنة",   "N & S Batinah",       2),
            ("MUSCAT",       "مسقط",                  "Muscat",              3),
            ("S_SHARQIYAH",  "جنوب الشرقية",          "South Sharqiyah",     4),
            ("AL_WUSTA",     "الوسطى",                "Al Wusta",            5),
            ("DHOFAR",       "ظفار",                  "Dhofar",              6),
        ],
    },

    # ---- FactSectorSummary -------------------------------------------------
    # SectorID: 1=Artisanal, 2=Coastal, 3=Commercial, 4=Aquaculture
    # FiscalYear, QuantityTons, ValueThousandOMR, QtyGrowthPct, ValueGrowthPct
    # 2025 growth columns are NULL (base year; no prior-period data in deck).
    # Note: Aquaculture 2025 value = 4,200 k OMR (slide 8 columns were transposed;
    #       growth +50% confirms: (6400-4200)/4200 = +52% ≈ +50%).
    # Commercial 2025 value = 60,776 k OMR (from slide 11; slide 7 transposed).
    "FactSectorSummary": {
        "sql": """
            INSERT INTO [FactSectorSummary]
                ([SectorID],[FiscalYear],[QuantityTons],[ValueThousandOMR],[QtyGrowthPct],[ValueGrowthPct])
            VALUES (?,?,?,?,?,?)
        """,
        "rows": [
            # Artisanal
            (1, 2025, 189770.0, 103935.0, None,  None ),
            (1, 2026, 198471.0, 115610.0,  0.05,  0.11),
            # Coastal
            (2, 2025,  30835.0,  15298.0, None,  None ),
            (2, 2026,  31144.0,  15432.0,  0.01,  0.01),
            # Commercial
            (3, 2025,  45361.0,  60776.0, None,  None ),
            (3, 2026,  26968.0,  29595.0, -0.41, -0.51),
            # Aquaculture
            (4, 2025,   1912.0,   4200.0, None,  None ),
            (4, 2026,   2682.0,   6400.0,  0.40,  0.50),
        ],
    },

    # ---- FactGovernorateSummary --------------------------------------------
    # Artisanal only.  GovernorateID 1-6 per DimGovernorate order.
    # Growth derived from Q1 2026 vs Q1 2025 (slides 10/11).
    "FactGovernorateSummary": {
        "sql": """
            INSERT INTO [FactGovernorateSummary]
                ([GovernorateID],[FiscalYear],[QuantityTons],[ValueThousandOMR],[QtyGrowthPct],[ValueGrowthPct])
            VALUES (?,?,?,?,?,?)
        """,
        "rows": [
            # Musandam
            (1, 2025,  8755.0,  8129.0, None,  None ),
            (1, 2026,  6557.0,  7663.0, -0.25, -0.06),
            # N & S Batinah
            (2, 2025, 15927.0, 18215.0, None,  None ),
            (2, 2026, 18148.0, 20910.0,  0.14,  0.15),
            # Muscat
            (3, 2025,  8038.0,  9906.0, None,  None ),
            (3, 2026,  9950.0, 10599.0,  0.24,  0.07),
            # South Sharqiyah
            (4, 2025, 54348.0, 26012.0, None,  None ),
            (4, 2026, 57293.0, 30502.0,  0.05,  0.17),
            # Al Wusta
            (5, 2025, 78215.0, 30928.0, None,  None ),
            (5, 2026, 82426.0, 35113.0,  0.05,  0.14),
            # Dhofar
            (6, 2025, 24487.0, 10745.0, None,  None ),
            (6, 2026, 24097.0, 10822.0, -0.02,  0.01),
        ],
    },

    # ---- FactMonthlyProduction ---------------------------------------------
    # FiscalYear, MonthNo, SectorID, GovernorateID (None=total), Qty, Value
    # GovernorateID is populated for artisanal governorate rows;
    # NULL for artisanal totals and for coastal/commercial rows.
    # Only 2026 monthly detail is available (slides 10/11).
    "FactMonthlyProduction": {
        "sql": """
            INSERT INTO [FactMonthlyProduction]
                ([FiscalYear],[MonthNo],[SectorID],[GovernorateID],[QuantityTons],[ValueThousandOMR])
            VALUES (?,?,?,?,?,?)
        """,
        "rows": [
            # ---- Artisanal by Governorate (SectorID=1) ----
            # GovernorateID 1 = Musandam
            (2026, 1, 1, 1,  2663.0,  3611.0),
            (2026, 2, 1, 1,  3037.0,  3126.0),
            (2026, 3, 1, 1,   857.0,   926.0),
            # GovernorateID 2 = N & S Batinah
            (2026, 1, 1, 2,  4757.0,  7091.0),
            (2026, 2, 1, 2,  5295.0,  6589.0),
            (2026, 3, 1, 2,  8096.0,  7230.0),
            # GovernorateID 3 = Muscat
            (2026, 1, 1, 3,  3476.0,  4083.0),
            (2026, 2, 1, 3,  2864.0,  3100.0),
            (2026, 3, 1, 3,  3611.0,  3416.0),
            # GovernorateID 4 = South Sharqiyah
            (2026, 1, 1, 4, 22327.0, 10896.0),
            (2026, 2, 1, 4, 18415.0, 10408.0),
            (2026, 3, 1, 4, 16552.0,  9198.0),
            # GovernorateID 5 = Al Wusta
            (2026, 1, 1, 5, 31150.0, 12874.0),
            (2026, 2, 1, 5, 24847.0, 11458.0),
            (2026, 3, 1, 5, 26428.0, 10782.0),
            # GovernorateID 6 = Dhofar
            (2026, 1, 1, 6,  7626.0,  2831.0),
            (2026, 2, 1, 6,  6963.0,  3550.0),
            (2026, 3, 1, 6,  9507.0,  4441.0),

            # ---- Artisanal Totals (GovernorateID = NULL) ----
            (2026, 1, 1, None, 72000.0, 41385.0),
            (2026, 2, 1, None, 61421.0, 38231.0),
            (2026, 3, 1, None, 65050.0, 35993.0),

            # ---- Coastal (SectorID=2) ----
            (2026, 1, 2, None,  5070.0,  2099.0),
            (2026, 2, 2, None,  7376.0,  2474.0),
            (2026, 3, 2, None, 18697.0, 10858.0),

            # ---- Commercial (SectorID=3) ----
            (2026, 1, 3, None, 10611.0,  7971.0),
            (2026, 2, 3, None, 10542.0, 10759.0),
            (2026, 3, 3, None,  5815.0, 10864.0),
        ],
    },

    # ---- FactOverallMonthly ------------------------------------------------
    # From embedded chart 1.  Unit: THOUSAND TONS (matches chart axis label).
    "FactOverallMonthly": {
        "sql": """
            INSERT INTO [FactOverallMonthly]
                ([FiscalYear],[MonthNo],[QuantityThousandTons])
            VALUES (?,?,?)
        """,
        "rows": [
            (2025, 1, 85.972543540),
            (2025, 2, 77.054206360),
            (2025, 3, 104.851868520),
            (2026, 1, 87.681204105),
            (2026, 2, 79.339480834),
            (2026, 3, 92.243739973),
        ],
    },

    # ---- FactAquacultureSpecies -------------------------------------------
    # From slide 9.  ValueThousandOMR converted from the slide's M OMR values.
    # Species with zero production are included for completeness.
    "FactAquacultureSpecies": {
        "sql": """
            INSERT INTO [FactAquacultureSpecies]
                ([FiscalYear],[SpeciesNameAr],[SpeciesNameEn],[QuantityTons],[ValueThousandOMR])
            VALUES (?,?,?,?,?)
        """,
        "rows": [
            (2026, "كوفر",            "Cobia",         1432.5, 2900.0),
            (2026, "الروبيان",        "Shrimp",        1047.0, 2600.0),
            (2026, "الكارب",          "Carp",           152.3,  200.0),
            (2026, "البرمندي",        "Barramundi",      38.0,  660.0),
            (2026, "الصفيلح العماني", "Omani Sole",       4.5,  200.0),
            (2026, "البلطي",          "Tilapia",          7.5,   20.0),
            (2026, "المحار الصخري",   "Rock Oyster",      0.0,    0.0),
        ],
    },

    # ---- FactArtisanalTopSpecies ------------------------------------------
    # From embedded chart 3.  SharePct as decimal fraction.
    # QuantityTons estimated as SharePct × total artisanal Q1-2026 (198,471 t).
    "FactArtisanalTopSpecies": {
        "sql": """
            INSERT INTO [FactArtisanalTopSpecies]
                ([FiscalYear],[SpeciesNameAr],[SpeciesNameEn],[SharePct],[QuantityTons])
            VALUES (?,?,?,?,?)
        """,
        "rows": [
            (2026, "عومة",                 "Omua (Grouper)",  0.42749, round(0.42749 * 198471)),
            (2026, "ضلعة",                 "Dhal'a (Trevally)",0.10955, round(0.10955 * 198471)),
            (2026, "جام",                  "Jam (Kingfish)",  0.08662, round(0.08662 * 198471)),
            (2026, "جيذر",                 "Jidhar (Emperor)",0.05330, round(0.05330 * 198471)),
            (2026, "أسماك غير معروفة",    "Unknown Species", 0.03009, round(0.03009 * 198471)),
        ],
    },

    # ---- FactCommercialVessel --------------------------------------------
    # Slides 12/13.  ProductionTons = NULL for vessels in slide 13 with no data.
    # Grand total per slide 13: 26,968 t.
    "FactCommercialVessel": {
        "sql": """
            INSERT INTO [FactCommercialVessel]
                ([FiscalYear],[VesselName],[ProductionTons],[Trips])
            VALUES (?,?,?,?)
        """,
        "rows": [
            # Vessels with production data (slide 12)
            (2026, "الجوهرة",            5017.1, None),
            (2026, "النعمة",             3182.4, None),
            (2026, "النسر",              2987.2, None),
            (2026, "أكجون فيدات",       2840.4, None),
            (2026, "HAWWA",              2613.0, None),
            (2026, "أوشن فرش",          2498.9, None),
            (2026, "NOUR",               2208.0, None),
            (2026, "روسي رسبينا",        1316.2, None),
            (2026, "مارجو حسين",          915.0, None),
            (2026, "ايدن توكر",           656.6, None),
            (2026, "ريسبينا 3",           609.4, None),
            (2026, "LAYLA",               444.0, None),
            (2026, "فخر البحار",          412.0, None),
            (2026, "كازم كوبيا",          404.5, None),
            (2026, "لورنج يوان يو 231",   203.6, None),
            (2026, "لورنج يوان يو 232",   202.7, None),
            (2026, "يوسف الرئيس",         139.0, None),
            (2026, "هيلي 618",            126.0, None),
            (2026, "فايكنج 1",             92.2, None),
            (2026, "Muzdahira",            66.8, None),
            (2026, "zayn",                 53.8, None),
            (2026, "arwa",                 42.3, None),
            (2026, "سينار إبراهيم",         6.5, None),
            # Vessels without production data (slide 12 + slide 13)
            (2026, "البركة",             None,  None),
            (2026, "الزين",              None,  None),
            (2026, "الشرقية 2",          None,  None),
            (2026, "المسرة",             None,  None),
            (2026, "تغريد 2",            None,  None),
            (2026, "تكسوري بيري",        None,  None),
            (2026, "سارة 2",             None,  None),
            (2026, "فيويوان 9992",       None,  None),
            (2026, "فيويوان 9993",       None,  None),
            (2026, "فيويوان 9996",       None,  None),
            (2026, "فيويونج 815",        None,  None),
            (2026, "فيويونج 816",        None,  None),
            (2026, "لورنج يوان يو 237",  None,  None),
            (2026, "هيلي 617",           None,  None),
            (2026, "هيلي 888",           None,  None),
            (2026, "الخير 1",            None,  None),
        ],
    },

    # ---- ReportMeta -------------------------------------------------------
    "ReportMeta": {
        "sql": """
            INSERT INTO [ReportMeta]
                ([ReportTitleAr],[ReportTitleEn],[PeriodLabelAr],[PeriodLabelEn],
                 [FiscalYear],[QuarterNo],[CurrencyAr],[CurrencyEn],[GeneratedOn])
            VALUES (?,?,?,?,?,?,?,?,?)
        """,
        "rows": [
            (
                "البيانات والمؤشرات الإحصائية للقطاع السمكي",
                "Fisheries Sector — Statistical Data & Indicators",
                "الربع الأول (يناير – مارس) 2026",
                "Q1 2026 — January to March",
                2026,
                1,
                "ألف ريال عماني",
                "Thousand OMR",
                datetime.datetime(2026, 4, 1),
            ),
        ],
    },
}


# --------------------------------------------------------------------------- #
# 5. MAIN — create DB, apply DDL, insert seed data                            #
# --------------------------------------------------------------------------- #

def main():
    print("=" * 60)
    print("  Fisheries Dashboard — Database Builder")
    print("=" * 60)

    # Step 1: create blank .accdb file
    print("\n[1] Creating database file …")
    create_accdb(DB_PATH)

    # Step 2: connect and create tables
    print("\n[2] Creating tables …")
    conn = pyodbc.connect(CONN_STR, autocommit=True)
    cursor = conn.cursor()

    for ddl in DDL_STATEMENTS:
        table_name = ddl.strip().split()[2].strip("[]")
        cursor.execute(ddl)
        print(f"    ✓ {table_name}")

    # Step 3: insert seed data
    print("\n[3] Seeding data …")
    conn.autocommit = False  # use explicit transaction for inserts

    for table_name, info in SEED_DATA.items():
        sql  = info["sql"]
        rows = info["rows"]
        for row in rows:
            cursor.execute(sql, row)
        print(f"    ✓ {table_name}: {len(rows)} rows")

    conn.commit()
    cursor.close()
    conn.close()

    print("\n[4] Done!")
    print(f"    Database ready: {DB_PATH}")
    print("    Run the dashboard with:  streamlit run app.py")
    print("=" * 60)


if __name__ == "__main__":
    main()
