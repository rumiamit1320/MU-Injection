# MU Injection Manager

Streamlit application for monthly MU injection reporting with Circle/Division views, meter master, monthly readings, Excel import/export, dashboards, and authenticated session control.

## Run locally

1. Install dependencies from requirements.txt.
2. Run: streamlit run app.py
3. Or use Run_MU_Injection_Manager.bat on Windows.

## Authentication

The application initializes an authentication table in SQLite and uses salted PBKDF2-SHA256 password hashes. Configure the first administrator with MU_ADMIN_USERNAME and MU_ADMIN_PASSWORD environment variables before first startup, or use the built-in defaults for local development.

## Included templates

- MU_Injection_Template.xlsx
- MU_Injection_All_Divisions_Template.xlsx
