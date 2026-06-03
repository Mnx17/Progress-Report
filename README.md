# Fisheries Sector Progress-Report Dashboard
## لوحة بيانات ومؤشرات القطاع السمكي

Bilingual (Arabic / English), read-only, on-premise Streamlit dashboard for the
Q1 2026 Fisheries Sector Statistical Indicators Report — Sultanate of Oman.

---

## Deliverable Files

| File | Purpose |
|------|---------|
| `build_db.py` | Creates and seeds `FisheriesQ1_2026.accdb` with all Q1-2026 data |
| `app.py` | Read-only Streamlit dashboard (bilingual, RTL-aware, Plotly charts) |
| `requirements.txt` | Python dependencies (install on the Windows host) |
| `Dockerfile` | Linux container definition (see Access caveat inside) |
| `.streamlit/config.toml` | Server settings — binds to localhost by default |
| `.streamlit/secrets.toml.example` | Template for configuring the DB path |

---

## Prerequisites (Windows on-prem host)

### 1. Python — 64-bit (recommended)
Download Python 3.10–3.12 (64-bit) from python.org.

### 2. Microsoft Access Database Engine 2016 — **must match Python bitness**

| Python | Driver |
|--------|--------|
| 64-bit | `AccessDatabaseEngine_X64.exe` |
| 32-bit | `AccessDatabaseEngine.exe` |

Download: https://www.microsoft.com/en-us/download/details.aspx?id=54920

> If you already have 32-bit Microsoft Office installed, you must either:
> - Use 32-bit Python + 32-bit driver, **or**
> - Uninstall 32-bit Office and install 64-bit Office/driver.
> There is no mixing of 32-bit Office with a 64-bit ODBC driver.

### 3. Python packages

```powershell
pip install -r requirements.txt
pip install pywin32          # Windows-only; needed by build_db.py
```

---

## Setup & Run (Windows)

### Step 1 — Create the database

```powershell
python build_db.py
```

Expected output:
```
[1] Creating database file …
    Created blank database: …\FisheriesQ1_2026.accdb
[2] Creating tables …
    ✓ DimSector  ✓ DimGovernorate  ✓ FactSectorSummary  …
[3] Seeding data …
    ✓ DimSector: 4 rows  ✓ DimGovernorate: 6 rows  …
[4] Done!
```

### Step 2 — (Optional) Configure the database path

If you move the `.accdb` file, set the path in one of these ways (highest priority first):

**Option A — Streamlit secrets (recommended)**
```toml
# .streamlit/secrets.toml
[database]
path = "C:/path/to/FisheriesQ1_2026.accdb"
```

**Option B — Environment variable**
```powershell
$env:DB_PATH = "C:\path\to\FisheriesQ1_2026.accdb"
```

**Option C — .env file** (create in the same folder as `app.py`)
```
DB_PATH=C:\path\to\FisheriesQ1_2026.accdb
```

If none of the above are set, the app looks for the `.accdb` in its own directory.

### Step 3 — Launch the dashboard

```powershell
streamlit run app.py
```

Open: http://localhost:8501

---

## Security Hardening (on-prem)

The dashboard is strictly **read-only** by design:

| Layer | Control |
|-------|---------|
| SQL | Only `SELECT` statements; no writes anywhere in `app.py` |
| ODBC | `ReadOnly=1` in the connection string |
| pyodbc | `readonly=True` parameter |
| File system | Grant the service account **Read** NTFS permission only on the `.accdb` file |
| Network | `config.toml` binds to `127.0.0.1`; expose via reverse proxy (IIS / Nginx) with Windows Auth or IP allowlist |
| TLS | Terminate SSL at the reverse proxy; do not run Streamlit with a raw self-signed cert |
| VPN | Recommended — restrict access to internal network only |

---

## Database Schema

10 tables — star-style schema derived from the progress report slides.

```
DimSector          — 4 rows: Artisanal / Coastal / Commercial / Aquaculture
DimGovernorate     — 6 rows: Musandam / Batinah / Muscat / S.Sharqiyah / Al Wusta / Dhofar
FactSectorSummary  — sector × year Q1 totals (qty, value, growth)
FactGovernorateSummary — governorate × year Q1 totals (artisanal only)
FactMonthlyProduction  — sector × month × governorate monthly quantities & values
FactOverallMonthly     — overall monthly totals (both years, from embedded chart)
FactAquacultureSpecies — species-level aquaculture qty & value
FactArtisanalTopSpecies — top-5 artisanal species shares
FactCommercialVessel   — 38 vessels with Q1 production
ReportMeta             — bilingual report title, period label, currency
```

---

## Verification Checklist

After running `build_db.py` and `streamlit run app.py`:

- [ ] KPI: Total production ≈ **259.3 k tons** (▼3%)
- [ ] KPI: Total value ≈ **167.0 M OMR** (▼9%)
- [ ] KPI: Artisanal share = **76.6%**
- [ ] KPI: Aquaculture growth = **+40%**
- [ ] Sector donut: Artisanal largest slice (~76.6%), Commercial smallest (~10.4%)
- [ ] Commercial vessels: الجوهرة top vessel at 5,017 t; grand total ~26,968 t
- [ ] Language toggle: switching to **العربية** flips layout to RTL and translates all labels
- [ ] Read-only proof: mark `.accdb` as read-only in Windows Explorer → dashboard still loads

---

## Data Auto-Refresh

`@st.cache_data(ttl=600)` caches every query for **10 minutes**.  
When the `.accdb` file is updated by the data team, the dashboard picks up new
data within 10 minutes without a restart.

To force an immediate refresh: open the Streamlit menu (⋮) → **Clear cache**.

---

## Adding a New Reporting Quarter

1. Insert new rows into the `.accdb` tables (using Access, Excel/ODBC, or a new seed script).
2. The dashboard reads whatever is in the database — no code changes needed for new data.
3. Update `ReportMeta` to reflect the new period label.
