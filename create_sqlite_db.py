# =============================================================================
# create_sqlite_db.py
# =============================================================================
# Creates fisheries.db (SQLite) from the same Q1-2026 seed data used by
# build_db.py.  This file is committed to the repo so Streamlit Community
# Cloud (Linux) can read the data without needing the Windows Access ODBC driver.
#
# SQLite is built into Python — no extra packages required.
# Run once locally:   python create_sqlite_db.py
# =============================================================================

import os
import sys
import sqlite3
import datetime

if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8","utf8"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fisheries.db")

# ---------------------------------------------------------------------------
# DDL  (SQLite dialect — TEXT ignores length; REAL instead of DOUBLE)
# ---------------------------------------------------------------------------
DDL = [
    """CREATE TABLE IF NOT EXISTS DimSector (
        SectorID   INTEGER PRIMARY KEY AUTOINCREMENT,
        SectorCode TEXT NOT NULL,
        NameAr     TEXT NOT NULL,
        NameEn     TEXT NOT NULL,
        SortOrder  INTEGER NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS DimGovernorate (
        GovernorateID INTEGER PRIMARY KEY AUTOINCREMENT,
        GovCode       TEXT NOT NULL,
        NameAr        TEXT NOT NULL,
        NameEn        TEXT NOT NULL,
        SortOrder     INTEGER NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS FactSectorSummary (
        SummaryID        INTEGER PRIMARY KEY AUTOINCREMENT,
        SectorID         INTEGER NOT NULL,
        FiscalYear       INTEGER NOT NULL,
        QuantityTons     REAL NOT NULL,
        ValueThousandOMR REAL NOT NULL,
        QtyGrowthPct     REAL,
        ValueGrowthPct   REAL
    )""",
    """CREATE TABLE IF NOT EXISTS FactGovernorateSummary (
        GovSummaryID     INTEGER PRIMARY KEY AUTOINCREMENT,
        GovernorateID    INTEGER NOT NULL,
        FiscalYear       INTEGER NOT NULL,
        QuantityTons     REAL NOT NULL,
        ValueThousandOMR REAL NOT NULL,
        QtyGrowthPct     REAL,
        ValueGrowthPct   REAL
    )""",
    """CREATE TABLE IF NOT EXISTS FactMonthlyProduction (
        MonthlyID        INTEGER PRIMARY KEY AUTOINCREMENT,
        FiscalYear       INTEGER NOT NULL,
        MonthNo          INTEGER NOT NULL,
        SectorID         INTEGER NOT NULL,
        GovernorateID    INTEGER,
        QuantityTons     REAL NOT NULL,
        ValueThousandOMR REAL NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS FactOverallMonthly (
        OverallMonthlyID     INTEGER PRIMARY KEY AUTOINCREMENT,
        FiscalYear           INTEGER NOT NULL,
        MonthNo              INTEGER NOT NULL,
        QuantityThousandTons REAL NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS FactAquacultureSpecies (
        AquaID           INTEGER PRIMARY KEY AUTOINCREMENT,
        FiscalYear       INTEGER NOT NULL,
        SpeciesNameAr    TEXT NOT NULL,
        SpeciesNameEn    TEXT NOT NULL,
        QuantityTons     REAL NOT NULL,
        ValueThousandOMR REAL NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS FactArtisanalTopSpecies (
        TopSpeciesID INTEGER PRIMARY KEY AUTOINCREMENT,
        FiscalYear   INTEGER NOT NULL,
        SpeciesNameAr TEXT NOT NULL,
        SpeciesNameEn TEXT NOT NULL,
        SharePct     REAL NOT NULL,
        QuantityTons REAL
    )""",
    """CREATE TABLE IF NOT EXISTS FactCommercialVessel (
        VesselID       INTEGER PRIMARY KEY AUTOINCREMENT,
        FiscalYear     INTEGER NOT NULL,
        VesselName     TEXT NOT NULL,
        ProductionTons REAL,
        Trips          INTEGER
    )""",
    """CREATE TABLE IF NOT EXISTS ReportMeta (
        MetaID        INTEGER PRIMARY KEY AUTOINCREMENT,
        ReportTitleAr TEXT NOT NULL,
        ReportTitleEn TEXT NOT NULL,
        PeriodLabelAr TEXT NOT NULL,
        PeriodLabelEn TEXT NOT NULL,
        FiscalYear    INTEGER NOT NULL,
        QuarterNo     INTEGER NOT NULL,
        CurrencyAr    TEXT NOT NULL,
        CurrencyEn    TEXT NOT NULL,
        GeneratedOn   TEXT NOT NULL
    )""",
]

# ---------------------------------------------------------------------------
# SEED DATA (identical to build_db.py)
# ---------------------------------------------------------------------------
SEED = {
    "DimSector": {
        "sql": "INSERT INTO DimSector (SectorCode,NameAr,NameEn,SortOrder) VALUES (?,?,?,?)",
        "rows": [
            ("ARTISANAL",   "الصيد الحرفي",      "Artisanal Fishing",  1),
            ("COASTAL",     "الصيد الساحلي",     "Coastal Fishing",    2),
            ("COMMERCIAL",  "الصيد التجاري",     "Commercial Fishing", 3),
            ("AQUACULTURE", "الاستزراع السمكي",  "Aquaculture",        4),
        ],
    },
    "DimGovernorate": {
        "sql": "INSERT INTO DimGovernorate (GovCode,NameAr,NameEn,SortOrder) VALUES (?,?,?,?)",
        "rows": [
            ("MUSANDAM",    "مسندم",                "Musandam",        1),
            ("BATINAH",     "شمال وجنوب الباطنة",  "N & S Batinah",   2),
            ("MUSCAT",      "مسقط",                 "Muscat",          3),
            ("S_SHARQIYAH", "جنوب الشرقية",         "South Sharqiyah", 4),
            ("AL_WUSTA",    "الوسطى",               "Al Wusta",        5),
            ("DHOFAR",      "ظفار",                 "Dhofar",          6),
        ],
    },
    "FactSectorSummary": {
        "sql": "INSERT INTO FactSectorSummary (SectorID,FiscalYear,QuantityTons,ValueThousandOMR,QtyGrowthPct,ValueGrowthPct) VALUES (?,?,?,?,?,?)",
        "rows": [
            (1,2025,189770.0,103935.0,None,None),(1,2026,198471.0,115610.0,0.05,0.11),
            (2,2025, 30835.0, 15298.0,None,None),(2,2026, 31144.0, 15432.0,0.01,0.01),
            (3,2025, 45361.0, 60776.0,None,None),(3,2026, 26968.0, 29595.0,-0.41,-0.51),
            (4,2025,  1912.0,  4200.0,None,None),(4,2026,  2682.0,  6400.0,0.40,0.50),
        ],
    },
    "FactGovernorateSummary": {
        "sql": "INSERT INTO FactGovernorateSummary (GovernorateID,FiscalYear,QuantityTons,ValueThousandOMR,QtyGrowthPct,ValueGrowthPct) VALUES (?,?,?,?,?,?)",
        "rows": [
            (1,2025, 8755.0, 8129.0,None,None),(1,2026, 6557.0, 7663.0,-0.25,-0.06),
            (2,2025,15927.0,18215.0,None,None),(2,2026,18148.0,20910.0, 0.14, 0.15),
            (3,2025, 8038.0, 9906.0,None,None),(3,2026, 9950.0,10599.0, 0.24, 0.07),
            (4,2025,54348.0,26012.0,None,None),(4,2026,57293.0,30502.0, 0.05, 0.17),
            (5,2025,78215.0,30928.0,None,None),(5,2026,82426.0,35113.0, 0.05, 0.14),
            (6,2025,24487.0,10745.0,None,None),(6,2026,24097.0,10822.0,-0.02, 0.01),
        ],
    },
    "FactMonthlyProduction": {
        "sql": "INSERT INTO FactMonthlyProduction (FiscalYear,MonthNo,SectorID,GovernorateID,QuantityTons,ValueThousandOMR) VALUES (?,?,?,?,?,?)",
        "rows": [
            # Artisanal by governorate
            (2026,1,1,1,2663.0,3611.0),(2026,2,1,1,3037.0,3126.0),(2026,3,1,1,857.0,926.0),
            (2026,1,1,2,4757.0,7091.0),(2026,2,1,2,5295.0,6589.0),(2026,3,1,2,8096.0,7230.0),
            (2026,1,1,3,3476.0,4083.0),(2026,2,1,3,2864.0,3100.0),(2026,3,1,3,3611.0,3416.0),
            (2026,1,1,4,22327.0,10896.0),(2026,2,1,4,18415.0,10408.0),(2026,3,1,4,16552.0,9198.0),
            (2026,1,1,5,31150.0,12874.0),(2026,2,1,5,24847.0,11458.0),(2026,3,1,5,26428.0,10782.0),
            (2026,1,1,6,7626.0,2831.0),(2026,2,1,6,6963.0,3550.0),(2026,3,1,6,9507.0,4441.0),
            # Artisanal totals
            (2026,1,1,None,72000.0,41385.0),(2026,2,1,None,61421.0,38231.0),(2026,3,1,None,65050.0,35993.0),
            # Coastal
            (2026,1,2,None,5070.0,2099.0),(2026,2,2,None,7376.0,2474.0),(2026,3,2,None,18697.0,10858.0),
            # Commercial
            (2026,1,3,None,10611.0,7971.0),(2026,2,3,None,10542.0,10759.0),(2026,3,3,None,5815.0,10864.0),
        ],
    },
    "FactOverallMonthly": {
        "sql": "INSERT INTO FactOverallMonthly (FiscalYear,MonthNo,QuantityThousandTons) VALUES (?,?,?)",
        "rows": [
            (2025,1,85.972543540),(2025,2,77.054206360),(2025,3,104.851868520),
            (2026,1,87.681204105),(2026,2,79.339480834),(2026,3,92.243739973),
        ],
    },
    "FactAquacultureSpecies": {
        "sql": "INSERT INTO FactAquacultureSpecies (FiscalYear,SpeciesNameAr,SpeciesNameEn,QuantityTons,ValueThousandOMR) VALUES (?,?,?,?,?)",
        "rows": [
            (2026,"كوفر","Cobia",1432.5,2900.0),
            (2026,"الروبيان","Shrimp",1047.0,2600.0),
            (2026,"الكارب","Carp",152.3,200.0),
            (2026,"البرمندي","Barramundi",38.0,660.0),
            (2026,"الصفيلح العماني","Omani Sole",4.5,200.0),
            (2026,"البلطي","Tilapia",7.5,20.0),
            (2026,"المحار الصخري","Rock Oyster",0.0,0.0),
        ],
    },
    "FactArtisanalTopSpecies": {
        "sql": "INSERT INTO FactArtisanalTopSpecies (FiscalYear,SpeciesNameAr,SpeciesNameEn,SharePct,QuantityTons) VALUES (?,?,?,?,?)",
        "rows": [
            (2026,"عومة","Omua (Grouper)",0.42749,round(0.42749*198471)),
            (2026,"ضلعة","Dhal'a (Trevally)",0.10955,round(0.10955*198471)),
            (2026,"جام","Jam (Kingfish)",0.08662,round(0.08662*198471)),
            (2026,"جيذر","Jidhar (Emperor)",0.05330,round(0.05330*198471)),
            (2026,"أسماك غير معروفة","Unknown Species",0.03009,round(0.03009*198471)),
        ],
    },
    "FactCommercialVessel": {
        "sql": "INSERT INTO FactCommercialVessel (FiscalYear,VesselName,ProductionTons,Trips) VALUES (?,?,?,?)",
        "rows": [
            (2026,"الجوهرة",5017.1,None),(2026,"النعمة",3182.4,None),(2026,"النسر",2987.2,None),
            (2026,"أكجون فيدات",2840.4,None),(2026,"HAWWA",2613.0,None),(2026,"أوشن فرش",2498.9,None),
            (2026,"NOUR",2208.0,None),(2026,"روسي رسبينا",1316.2,None),(2026,"مارجو حسين",915.0,None),
            (2026,"ايدن توكر",656.6,None),(2026,"ريسبينا 3",609.4,None),(2026,"LAYLA",444.0,None),
            (2026,"فخر البحار",412.0,None),(2026,"كازم كوبيا",404.5,None),
            (2026,"لورنج يوان يو 231",203.6,None),(2026,"لورنج يوان يو 232",202.7,None),
            (2026,"يوسف الرئيس",139.0,None),(2026,"هيلي 618",126.0,None),
            (2026,"فايكنج 1",92.2,None),(2026,"Muzdahira",66.8,None),
            (2026,"zayn",53.8,None),(2026,"arwa",42.3,None),(2026,"سينار إبراهيم",6.5,None),
            (2026,"البركة",None,None),(2026,"الزين",None,None),(2026,"الشرقية 2",None,None),
            (2026,"المسرة",None,None),(2026,"تغريد 2",None,None),(2026,"تكسوري بيري",None,None),
            (2026,"سارة 2",None,None),(2026,"فيويوان 9992",None,None),(2026,"فيويوان 9993",None,None),
            (2026,"فيويوان 9996",None,None),(2026,"فيويونج 815",None,None),(2026,"فيويونج 816",None,None),
            (2026,"لورنج يوان يو 237",None,None),(2026,"هيلي 617",None,None),
            (2026,"هيلي 888",None,None),(2026,"الخير 1",None,None),
        ],
    },
    "ReportMeta": {
        "sql": "INSERT INTO ReportMeta (ReportTitleAr,ReportTitleEn,PeriodLabelAr,PeriodLabelEn,FiscalYear,QuarterNo,CurrencyAr,CurrencyEn,GeneratedOn) VALUES (?,?,?,?,?,?,?,?,?)",
        "rows": [(
            "البيانات والمؤشرات الإحصائية للقطاع السمكي",
            "Fisheries Sector — Statistical Data & Indicators",
            "الربع الأول (يناير – مارس) 2026",
            "Q1 2026 — January to March",
            2026, 1,
            "ألف ريال عماني", "Thousand OMR",
            "2026-04-01",
        )],
    },
}


def main():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
        print(f"  Removed existing: {DB_PATH}")

    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()

    print("[1] Creating tables …")
    for ddl in DDL:
        table = ddl.strip().split()[5]  # "CREATE TABLE IF NOT EXISTS <Name>"
        cur.execute(ddl)
        print(f"    ✓ {table}")

    print("[2] Seeding data …")
    for table, info in SEED.items():
        for row in info["rows"]:
            cur.execute(info["sql"], row)
        print(f"    ✓ {table}: {len(info['rows'])} rows")

    conn.commit()
    conn.close()
    size = os.path.getsize(DB_PATH)
    print(f"\n✓ Done — {DB_PATH}  ({size:,} bytes)")


if __name__ == "__main__":
    main()
