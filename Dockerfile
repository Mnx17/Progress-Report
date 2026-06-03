# =============================================================================
# Dockerfile — Fisheries Dashboard (Streamlit)
# =============================================================================
#
# ⚠️  IMPORTANT — ACCESS DATABASE DRIVER CAVEAT:
#
#   The production database is a Microsoft Access .accdb file.
#   The pyodbc ODBC driver for Access (Microsoft.ACE.OLEDB) is WINDOWS-ONLY.
#
#   This Linux image therefore CANNOT connect to a live .accdb directly.
#
#   Your options:
#
#   OPTION 1 (RECOMMENDED — what you chose):
#     Run the Streamlit app natively on a Windows on-prem host.
#     No Docker required for the data layer.
#     Command:  streamlit run app.py
#
#   OPTION 2 — Windows container:
#     Use a Windows-based base image and install the Access Database Engine:
#       FROM mcr.microsoft.com/windows/servercore:ltsc2022
#     (Much heavier image; requires a Windows container host.)
#
#   OPTION 3 — Linux read-only via mdbtools:
#     Install mdbtools and use the MDB ODBC driver (limited .accdb support):
#       RUN apt-get install -y mdbtools odbc-mdbtools
#     Then adjust CONN_STR to use the mdbtools driver.
#     Note: mdbtools has partial .accdb support and may not read all Access
#     features correctly.  Best for read-only reporting on stable files.
#
#   OPTION 4 — Migrate to SQL Server / PostgreSQL:
#     Mirror the .accdb data to a Linux-friendly database and point app.py
#     at that instead.  Fully Docker/cloud native.
#
# This Dockerfile is provided to satisfy the containerization requirement and
# is ready for Option 3 (mdbtools) or Option 4 if you decide to migrate.
# For Option 1 (Windows native), use the install steps in README.md.
# =============================================================================

FROM python:3.12-slim

# ---- System dependencies ---------------------------------------------------
# libgcc-s1 and unixodbc are required by pyodbc.
# mdbtools-dev + odbc-mdbtools enable the read-only Linux MDB/ACCDB driver
# (Option 3).  Remove these lines if using Option 4 (different DB).
RUN apt-get update && apt-get install -y --no-install-recommends \
        unixodbc \
        unixodbc-dev \
        libgcc-s1 \
        mdbtools \
        odbc-mdbtools \
    && rm -rf /var/lib/apt/lists/*

# ---- Application directory -------------------------------------------------
WORKDIR /app

# ---- Python dependencies ---------------------------------------------------
# requirements.txt intentionally excludes pywin32 (Windows-only).
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# ---- Application source ----------------------------------------------------
COPY app.py            ./
COPY .streamlit/       ./.streamlit/

# ---- Expose Streamlit port -------------------------------------------------
EXPOSE 8501

# ---- Health check ----------------------------------------------------------
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/_stcore/health')"

# ---- Launch command --------------------------------------------------------
# DB_PATH must be injected as an environment variable or mounted via a volume.
# Example:
#   docker run -p 8501:8501 \
#     -e DB_PATH=/data/FisheriesQ1_2026.accdb \
#     -v /path/to/data:/data \
#     fisheries-dash
ENTRYPOINT ["streamlit", "run", "app.py", \
            "--server.address=0.0.0.0", \
            "--server.port=8501", \
            "--server.headless=true", \
            "--browser.gatherUsageStats=false"]
