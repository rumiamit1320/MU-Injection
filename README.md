# MU Injection Manager v5 — Jorhat Circle + Divisions

Single-run Streamlit application for Jorhat Circle monthly MU injection reporting.

## Hierarchy
- Jorhat Circle
- Jorhat-1 — no subdivision
- Jorhat-2 — Titabar, Mariani, Majuli
- Teok — Teok, Kakojan

Circle and division views use the same feeder master and shared monthly meter readings. A meter imported at circle level does not need to be uploaded again in a division when the feeder is mapped to that division.

## Feeder Master
Each feeder stores:
- Feeder name
- Meter number
- MF
- Selection/type
- Voltage
- Division
- Subdivision
- Initial reading

The initial reading is used only when no earlier monthly reading exists. When a historical Excel month is imported, the workbook Last Reading becomes the baseline for that historical month.

## Division calculation
Division workbooks contain separate Import and Export energy columns. The application calculates:
- Difference = Present − Previous
- Import energy = Difference × MF
- Export energy = Difference × MF
- Net Energy Injection = Import − Export
- Net Energy Injection (MU) = Net Energy Injection (MWh) / 1000

Export is therefore subtracted in the Net Energy Injection column. Division-row readings are stored separately when the same meter appears on multiple workbook rows with different row-level readings/baselines.

No manual MU override is supported.

## Subdivision placement
New feeders entered in Feeder Master are assigned to the selected division/subdivision. Generated division workbooks place new feeders under the selected subdivision; Jorhat-1 has no subdivision.

## Import
Import the monthly Excel workbook from the Import Excel page. All available Jorhat division sheets can be processed together. Existing monthly readings are protected unless Overwrite is selected.

## Dashboard
The dashboard provides current totals and a five-month energy trend for the selected circle/division.

## Run
Double-click:
`Run_MU_Injection_Manager.bat`

The launcher creates the virtual environment, installs requirements, initializes the SQLite database and starts Streamlit.


## Login and session control
The application now requires authentication before the main interface is loaded.

- Default local username: `admin`
- Default local password: `admin123`
- Session timeout: 30 minutes of inactivity
- A **Sign out** button is available in the sidebar after login.
- Credentials can be changed for a fresh database through the environment variables `MU_ADMIN_USERNAME` and `MU_ADMIN_PASSWORD` before first startup.
- Passwords are stored as salted PBKDF2-SHA256 hashes in the local SQLite database; the plaintext password is not stored.

For a production deployment, replace the default password and use an appropriate external identity provider or secrets-management system.
