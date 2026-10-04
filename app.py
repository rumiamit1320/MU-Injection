
import calendar
import io
import sqlite3
import re

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    from psycopg2 import errors as pg_errors
except ImportError:
    psycopg2 = None
    RealDictCursor = None
    pg_errors = None
import os
import time
import secrets
import hashlib
from datetime import datetime
from pathlib import Path
from copy import copy
from openpyxl.cell.cell import MergedCell

import openpyxl
import streamlit as st
import pandas as pd
import altair as alt

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "mu_injection.db"
TEMPLATE_PATH = APP_DIR / "MU_Injection_Template.xlsx"

TYPE_LABELS = {
    "A": "Import from GSS",
    "B": "Import from other circle",
    "C": "Export to other circle",
}
LABEL_TO_TYPE = {v: k for k, v in TYPE_LABELS.items()}

SELECTION_OPTIONS = [
    "Import from GSS",
    "Import from other circle",
    "Export to other circle",
    "Import from DSS",
    "Export to DSS",
    "Open Access",
    "None",
]
SELECTION_TO_ENTRY_TYPE = {
    "Import from GSS": "A",
    "Import from other circle": "B",
    "Export to other circle": "C",
    "Import from DSS": "A",
    "Export to DSS": "C",
    "Open Access": "B",
    "None": "A",
}
def selection_label_for_row(row):
    value = row["selection_type"]
    return value if value in SELECTION_OPTIONS else TYPE_LABELS.get(row["entry_type"], "None")

CIRCLE_NAME = "Jorhat Circle"
DIVISIONS = {"Jorhat-1": "Jorhat-ONE", "Jorhat-2": "Jorhat-TWO", "Teok": "Teok^"}
DIVISION_TO_ID = {name: i + 1 for i, name in enumerate(DIVISIONS)}

# ---------------- AUTHENTICATION / SESSION CONTROL ----------------
SESSION_TIMEOUT_SECONDS = 30 * 60
DEFAULT_USERNAME = os.getenv("MU_ADMIN_USERNAME", "admin")
DEFAULT_PASSWORD = os.getenv("MU_ADMIN_PASSWORD", "admin123")

def _hash_password(password, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 240000)
    return salt, digest.hex()

def _verify_password(password, salt, expected_hash):
    _, digest = _hash_password(password, salt)
    return secrets.compare_digest(digest, expected_hash)

def ensure_auth_table(con):
    con.execute("""CREATE TABLE IF NOT EXISTS app_users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE,
        salt TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        last_login TEXT
    )""")
    row = con.execute("SELECT id FROM app_users LIMIT 1").fetchone()
    if row is None:
        salt, password_hash = _hash_password(DEFAULT_PASSWORD)
        con.execute(
            "INSERT INTO app_users(username,salt,password_hash,active) VALUES(?,?,?,1)",
            (DEFAULT_USERNAME, salt, password_hash)
        )
        con.commit()

def logout():
    for key in ("authenticated", "auth_username", "last_activity"):
        st.session_state.pop(key, None)

def render_login(con):
    if st.session_state.get("authenticated"):
        last = st.session_state.get("last_activity", time.time())
        if time.time() - float(last) > SESSION_TIMEOUT_SECONDS:
            logout()
            st.session_state["session_expired"] = True
            st.rerun()
        st.session_state["last_activity"] = time.time()
        return True

    st.markdown("""
    <div class="login-shell">
      <div class="login-brand">⚡</div>
      <div class="login-title">MU Injection Manager</div>
      <div class="login-subtitle">Meter readings • Energy accounting • MU reports</div>
    </div>
    """, unsafe_allow_html=True)

    left, center, right = st.columns([1.0, 1.35, 1.0])
    with center:
        st.markdown('<div class="login-anchor"></div>', unsafe_allow_html=True)
        st.markdown('<div class="login-heading">Welcome back</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-hint">Sign in to continue to your secure workspace.</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-section-label">ACCOUNT</div>', unsafe_allow_html=True)
        username = st.text_input("Username", placeholder="Enter your username", label_visibility="collapsed", key="login_username")
        password = st.text_input("Password", type="password", placeholder="Enter your password", label_visibility="collapsed", key="login_password")
        if st.button("Sign in  →", type="primary", width="stretch", key="login_submit"):
            row = con.execute(
                "SELECT * FROM app_users WHERE username=? AND active=1 LIMIT 1",
                (username.strip(),)
            ).fetchone()
            if row is not None and _verify_password(password, row["salt"], row["password_hash"]):
                con.execute("UPDATE app_users SET last_login=CURRENT_TIMESTAMP WHERE id=?", (row["id"],))
                con.commit()
                st.session_state["authenticated"] = True
                st.session_state["auth_username"] = row["username"]
                st.session_state["last_activity"] = time.time()
                st.session_state.pop("session_expired", None)
                st.rerun()
            else:
                st.error("Invalid username or password.")
        if st.session_state.get("session_expired"):
            st.warning("Your session expired due to inactivity. Please sign in again.")
        st.markdown("""
        <div class="login-security">
          <span class="security-dot">●</span>
          <span>Secure session</span>
          <span class="security-sep">•</span>
          <span>30 min inactivity timeout</span>
        </div>
        """, unsafe_allow_html=True)
    return False

def db_sqlite():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    con.executescript("""
    CREATE TABLE IF NOT EXISTS feeder_master (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        feeder_name TEXT NOT NULL,
        meter_no TEXT NOT NULL,
        mf REAL NOT NULL DEFAULT 1,
        entry_type TEXT NOT NULL CHECK(entry_type IN ('A','B','C')),
        initial_reading_kwh REAL NOT NULL DEFAULT 0,
        division_name TEXT,
        subdivision TEXT,
        voltage_kv REAL,
        energy_direction TEXT NOT NULL DEFAULT 'IMPORT' CHECK(energy_direction IN ('IMPORT','EXPORT')),
        selection_type TEXT,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(meter_no, entry_type)
    );

    CREATE TABLE IF NOT EXISTS monthly_readings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        feeder_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        reading_kwh REAL NOT NULL,
        remarks TEXT,
        last_reading_kwh REAL,
        direct_mu REAL,
        direct_mu_note TEXT,
        report_section TEXT,
        report_order INTEGER,
        report_feeder_name TEXT,
        report_meter_no TEXT,
        report_mf REAL,
        report_sl_no REAL,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(feeder_id, year, month),
        FOREIGN KEY(feeder_id) REFERENCES feeder_master(id) ON DELETE CASCADE
    );

    CREATE INDEX IF NOT EXISTS idx_monthly_feeder_date
      ON monthly_readings(feeder_id, year, month);

    CREATE TABLE IF NOT EXISTS division_master (
        id INTEGER PRIMARY KEY,
        circle_name TEXT NOT NULL,
        division_name TEXT NOT NULL UNIQUE,
        sheet_name TEXT NOT NULL UNIQUE,
        active INTEGER NOT NULL DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS division_row_map (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        division_id INTEGER NOT NULL,
        sheet_row INTEGER NOT NULL,
        sub_division TEXT,
        feeder_id INTEGER NOT NULL,
        flow_direction TEXT NOT NULL CHECK(flow_direction IN ('IMPORT','EXPORT')),
        source_feeder_name TEXT,
        source_meter_no TEXT,
        active INTEGER NOT NULL DEFAULT 1,
        UNIQUE(division_id, sheet_row, flow_direction),
        FOREIGN KEY(division_id) REFERENCES division_master(id) ON DELETE CASCADE,
        FOREIGN KEY(feeder_id) REFERENCES feeder_master(id) ON DELETE CASCADE
    );

    CREATE INDEX IF NOT EXISTS idx_division_map_feeder
      ON division_row_map(feeder_id, division_id);

    CREATE TABLE IF NOT EXISTS division_row_readings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        division_map_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        reading_kwh REAL NOT NULL,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(division_map_id, year, month),
        FOREIGN KEY(division_map_id) REFERENCES division_row_map(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS app_settings (
        key TEXT PRIMARY KEY,
        value TEXT
    );

    CREATE TABLE IF NOT EXISTS meter_change_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        feeder_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        old_feeder_name TEXT,
        old_meter_no TEXT,
        old_mf REAL,
        old_last_reading_kwh REAL,
        old_entry_type TEXT,
        new_meter_no TEXT,
        new_mf REAL,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(feeder_id, year, month),
        FOREIGN KEY(feeder_id) REFERENCES feeder_master(id) ON DELETE CASCADE
    );
    """)
    # Division workbooks can contain the same meter number with different
    # row-level previous readings. Store that historical baseline on the mapping.
    try:
        con.execute("ALTER TABLE division_row_map ADD COLUMN baseline_reading_kwh REAL")
        con.commit()
    except sqlite3.OperationalError:
        pass

    for col, typ in [("division_name", "TEXT"), ("subdivision", "TEXT"), ("voltage_kv", "REAL"), ("energy_direction", "TEXT"), ("selection_type", "TEXT")]:
        try:
            con.execute(f"ALTER TABLE feeder_master ADD COLUMN {col} {typ}")
            con.commit()
        except sqlite3.OperationalError:
            pass

    for col, typ in [("last_reading_kwh", "REAL"), ("direct_mu", "REAL"), ("direct_mu_note", "TEXT"), ("report_section", "TEXT"), ("report_order", "INTEGER"), ("report_feeder_name", "TEXT"), ("report_meter_no", "TEXT"), ("report_mf", "REAL"), ("report_sl_no", "REAL")]:
        try:
            con.execute(f"ALTER TABLE monthly_readings ADD COLUMN {col} {typ}")
            con.commit()
        except sqlite3.OperationalError:
            pass

    # Backward compatibility: older databases get an explicit meter energy direction.
    try:
        con.execute("UPDATE feeder_master SET energy_direction=CASE WHEN entry_type='C' THEN 'EXPORT' ELSE 'IMPORT' END WHERE energy_direction IS NULL OR energy_direction=''")
        con.execute("""UPDATE feeder_master
                       SET selection_type=CASE entry_type
                           WHEN 'A' THEN 'Import from GSS'
                           WHEN 'B' THEN 'Import from other circle'
                           WHEN 'C' THEN 'Export to other circle'
                       END
                       WHERE selection_type IS NULL OR selection_type=''""")
        con.commit()
    except sqlite3.OperationalError:
        pass

    # Backward compatibility: older versions had an energy_override_mu column.
    # Manual MU overrides are no longer supported; clear any legacy values and
    # leave the column unused so existing databases remain compatible.
    try:
        con.execute("UPDATE monthly_readings SET energy_override_mu=NULL")
        con.commit()
    except sqlite3.OperationalError:
        pass
    return con




class PostgresConnection:
    """Small compatibility layer so the existing SQLite-style SQL remains intact."""
    def __init__(self, url):
        if psycopg2 is None:
            raise RuntimeError("PostgreSQL support requires psycopg2-binary. Install the updated requirements.txt.")
        connect_url = url
        if "sslmode=" not in connect_url:
            connect_url += ("&" if "?" in connect_url else "?") + "sslmode=require"
        self._con = psycopg2.connect(connect_url, cursor_factory=RealDictCursor)
        self._initialize_schema()

    def _initialize_schema(self):
        # PostgreSQL must create the same application schema that SQLite creates
        # locally. The application logic remains unchanged; only the database
        # backend differs.
        self.executescript("""
        CREATE TABLE IF NOT EXISTS feeder_master (
            id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            feeder_name TEXT NOT NULL,
            meter_no TEXT NOT NULL,
            mf REAL NOT NULL DEFAULT 1,
            entry_type TEXT NOT NULL CHECK(entry_type IN ('A','B','C')),
            initial_reading_kwh REAL NOT NULL DEFAULT 0,
            division_name TEXT,
            subdivision TEXT,
            voltage_kv REAL,
            energy_direction TEXT NOT NULL DEFAULT 'IMPORT' CHECK(energy_direction IN ('IMPORT','EXPORT')),
            selection_type TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(meter_no, entry_type)
        );
        CREATE TABLE IF NOT EXISTS monthly_readings (
            id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            feeder_id INTEGER NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            reading_kwh REAL NOT NULL,
            remarks TEXT,
            last_reading_kwh REAL,
            direct_mu REAL,
            direct_mu_note TEXT,
            report_section TEXT,
            report_order INTEGER,
            report_feeder_name TEXT,
            report_meter_no TEXT,
            report_mf REAL,
            report_sl_no REAL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(feeder_id, year, month),
            FOREIGN KEY(feeder_id) REFERENCES feeder_master(id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_monthly_feeder_date ON monthly_readings(feeder_id, year, month);
        CREATE TABLE IF NOT EXISTS division_master (
            id INTEGER PRIMARY KEY,
            circle_name TEXT NOT NULL,
            division_name TEXT NOT NULL UNIQUE,
            sheet_name TEXT NOT NULL UNIQUE,
            active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS division_row_map (
            id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            division_id INTEGER NOT NULL,
            sheet_row INTEGER NOT NULL,
            sub_division TEXT,
            feeder_id INTEGER NOT NULL,
            flow_direction TEXT NOT NULL CHECK(flow_direction IN ('IMPORT','EXPORT')),
            source_feeder_name TEXT,
            source_meter_no TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            UNIQUE(division_id, sheet_row, flow_direction),
            FOREIGN KEY(division_id) REFERENCES division_master(id) ON DELETE CASCADE,
            FOREIGN KEY(feeder_id) REFERENCES feeder_master(id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_division_map_feeder ON division_row_map(feeder_id, division_id);
        CREATE TABLE IF NOT EXISTS division_row_readings (
            id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            division_map_id INTEGER NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            reading_kwh REAL NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(division_map_id, year, month),
            FOREIGN KEY(division_map_id) REFERENCES division_row_map(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );
        CREATE TABLE IF NOT EXISTS meter_change_history (
            id INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
            feeder_id INTEGER NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            old_feeder_name TEXT,
            old_meter_no TEXT,
            old_mf REAL,
            old_last_reading_kwh REAL,
            old_entry_type TEXT,
            new_meter_no TEXT,
            new_mf REAL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(feeder_id, year, month),
            FOREIGN KEY(feeder_id) REFERENCES feeder_master(id) ON DELETE CASCADE
        );
        """)
        self.commit()
        # Schema evolution used by the existing SQLite application.
        self.execute("ALTER TABLE division_row_map ADD COLUMN baseline_reading_kwh REAL")
        self.commit()
        for col, typ in [("division_name", "TEXT"), ("subdivision", "TEXT"), ("voltage_kv", "REAL"), ("energy_direction", "TEXT"), ("selection_type", "TEXT")]:
            self.execute(f"ALTER TABLE feeder_master ADD COLUMN {col} {typ}")
            self.commit()
        for col, typ in [("last_reading_kwh", "REAL"), ("direct_mu", "REAL"), ("direct_mu_note", "TEXT"), ("report_section", "TEXT"), ("report_order", "INTEGER"), ("report_feeder_name", "TEXT"), ("report_meter_no", "TEXT"), ("report_mf", "REAL"), ("report_sl_no", "REAL")]:
            self.execute(f"ALTER TABLE monthly_readings ADD COLUMN {col} {typ}")
            self.commit()
        self.execute("""UPDATE feeder_master
                       SET selection_type=CASE entry_type
                           WHEN 'A' THEN 'Import from GSS'
                           WHEN 'B' THEN 'Import from other circle'
                           WHEN 'C' THEN 'Export to other circle'
                       END
                       WHERE selection_type IS NULL OR selection_type=''""")
        self.commit()

    def _sql(self, sql):
        sql = sql.replace("INTEGER PRIMARY KEY AUTOINCREMENT", "INTEGER GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY")
        sql = re.sub(r"ALTER TABLE (\w+) ADD COLUMN ", r"ALTER TABLE \1 ADD COLUMN IF NOT EXISTS ", sql, flags=re.I)
        sql = re.sub(r"^\s*INSERT OR IGNORE INTO ", "INSERT INTO ", sql, flags=re.I)
        sql = sql.replace("last_insert_rowid() AS id", "currval(pg_get_serial_sequence('feeder_master','id')) AS id")
        return sql.replace("?", "%s")

    def execute(self, sql, params=None):
        sql2=self._sql(sql)
        # INSERT OR IGNORE is used only for idempotent mapping/read imports.
        if re.match(r"^\s*INSERT INTO ", sql2, re.I) and re.search(r"\bVALUES\b", sql2, re.I) and "ON CONFLICT" not in sql2.upper():
            original_ignore = bool(re.match(r"^\s*INSERT OR IGNORE INTO ", sql, re.I))
            if original_ignore:
                sql2 = sql2.rstrip().rstrip(";") + " ON CONFLICT DO NOTHING"
        try:
            cur=self._con.cursor()
            cur.execute(sql2, params)
            return cur
        except Exception as exc:
            if pg_errors is not None and isinstance(exc, pg_errors.UniqueViolation):
                self._con.rollback()
                raise sqlite3.IntegrityError(str(exc)) from exc
            if pg_errors is not None and isinstance(exc, (pg_errors.DuplicateColumn, pg_errors.UndefinedColumn, pg_errors.DuplicateTable)):
                self._con.rollback()
                raise sqlite3.OperationalError(str(exc)) from exc
            raise

    def executescript(self, script):
        # The existing schema is deliberately retained; only SQLite-specific
        # identity/ALTER syntax is translated for PostgreSQL.
        for statement in script.split(";"):
            statement=statement.strip()
            if statement:
                self.execute(statement)

    def commit(self):
        self._con.commit()

    def rollback(self):
        self._con.rollback()


def db():
    database_url = os.getenv("DATABASE_URL", "").strip()
    if database_url:
        # Streamlit reruns the script for every widget interaction. Reusing the
        # same PostgreSQL connection for the current browser session avoids a
        # fresh Neon connection + schema initialization on every rerun.
        existing = st.session_state.get("_postgres_db_connection")
        if existing is not None:
            try:
                if getattr(existing._con, "closed", 1) == 0:
                    return existing
            except Exception:
                pass
            st.session_state.pop("_postgres_db_connection", None)

        con = PostgresConnection(database_url)
        st.session_state["_postgres_db_connection"] = con
        return con
    if os.getenv("RENDER") == "true":
        raise RuntimeError(
            "DATABASE_URL is not configured. This Render deployment requires the Neon PostgreSQL connection string."
        )
    return db_sqlite()


def merged_cell_value(ws, row, col):
    value = ws.cell(row, col).value
    if value is not None:
        return value
    coord = ws.cell(row, col).coordinate
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            return ws.cell(rng.min_row, rng.min_col).value
    return None


def preferred_entry_type_for_flow(flow):
    return "C" if flow == "EXPORT" else "A"


def find_or_create_meter_master(con, feeder_name, meter_no, mf, preferred_type="A", year=None, month=None, present=None, division_name=None, subdivision=None, voltage_kv=None, energy_direction=None):
    meter_no = str(meter_no).strip()
    feeder_name = str(feeder_name).strip()
    exact = con.execute(
        "SELECT * FROM feeder_master WHERE meter_no=? AND entry_type=? ORDER BY id LIMIT 1",
        (meter_no, preferred_type)
    ).fetchone()
    if exact is not None:
        con.execute(
            "UPDATE feeder_master SET feeder_name=?, mf=?, division_name=COALESCE(division_name,?), subdivision=COALESCE(subdivision,?), voltage_kv=COALESCE(voltage_kv,?), energy_direction=COALESCE(?,energy_direction), updated_at=CURRENT_TIMESTAMP WHERE id=?",
            (feeder_name, float(mf), division_name, subdivision, voltage_kv, energy_direction or ("EXPORT" if preferred_type=="C" else "IMPORT"), exact["id"])
        )
        return exact["id"], False

    candidates = con.execute(
        "SELECT * FROM feeder_master WHERE meter_no=? ORDER BY id", (meter_no,)
    ).fetchall()
    if candidates:
        priority = {"C": ["C","B","A"], "A": ["B","A","C"], "B": ["B","A","C"]}.get(
            preferred_type, [preferred_type,"B","A","C"]
        )
        ordered = sorted(
            candidates,
            key=lambda x: (priority.index(x["entry_type"]) if x["entry_type"] in priority else 99, x["id"])
        )
        # Division workbooks occasionally show the same meter number twice with
        # different present readings. If a matching monthly reading already
        # exists, bind the division row to that meter instance.
        if year is not None and month is not None and present is not None:
            matching = []
            for candidate in candidates:
                rr = con.execute(
                    "SELECT reading_kwh FROM monthly_readings WHERE feeder_id=? AND year=? AND month=?",
                    (candidate["id"], year, month)
                ).fetchone()
                if rr is not None and abs(float(rr["reading_kwh"]) - float(present)) < 1e-9:
                    matching.append(candidate)
            if matching:
                ordered = sorted(
                    matching,
                    key=lambda x: (priority.index(x["entry_type"]) if x["entry_type"] in priority else 99, x["id"])
                )

        chosen = ordered[0]
        fid = chosen["id"]
        con.execute(
            "UPDATE feeder_master SET feeder_name=?, mf=?, division_name=COALESCE(division_name,?), subdivision=COALESCE(subdivision,?), voltage_kv=COALESCE(voltage_kv,?), energy_direction=COALESCE(?,energy_direction), updated_at=CURRENT_TIMESTAMP WHERE id=?",
            (feeder_name, float(mf), division_name, subdivision, voltage_kv, energy_direction or ("EXPORT" if preferred_type=="C" else "IMPORT"), fid)
        )
        return fid, False

    con.execute(
        "INSERT INTO feeder_master (feeder_name,meter_no,mf,entry_type,initial_reading_kwh,division_name,subdivision,voltage_kv,energy_direction) VALUES(?,?,?,?,0,?,?,?,?)",
        (feeder_name, meter_no, float(mf), preferred_type, division_name, subdivision, voltage_kv, energy_direction or ("EXPORT" if preferred_type=="C" else "IMPORT"))
    )
    return con.execute("SELECT last_insert_rowid() AS id").fetchone()["id"], True


def ensure_division_master(con):
    for name, sheet in DIVISIONS.items():
        con.execute(
            '''INSERT INTO division_master(id,circle_name,division_name,sheet_name)
               VALUES(?,?,?,?)
               ON CONFLICT(id) DO UPDATE SET
                 circle_name=excluded.circle_name,
                 division_name=excluded.division_name,
                 sheet_name=excluded.sheet_name''',
            (DIVISION_TO_ID[name], CIRCLE_NAME, name, sheet)
        )
    con.commit()


def normalize_subdivision(division_name, raw):
    if division_name == "Jorhat-1":
        return None
    text = "" if raw is None else str(raw).strip().upper()
    for key, value in [("TITABAR","Titabar"),("MARIANI","Mariani"),("MAJULI","Majuli"),("KAKOJAN","Kakojan"),("TEOK","Teok"),
                       ("JESD-1","JESD-1"),("JESD-I","JESD-1"),("JESD-2","JESD-2"),("JESD-II","JESD-2"),
                       ("JESD-3","JESD-3"),("JESD-III","JESD-3"),("DERGAON","Dergaon")]:
        if key in text:
            return value
    return None

ALL_SUBDIVISIONS = [
    "Teok", "Kakojan", "Majuli", "Titabar", "Mariani",
    "JESD-1", "JESD-2", "JESD-3", "Dergaon"
]

def allowed_subdivisions(division_name):
    # Keep one consistent master-data vocabulary across the circle.
    # Division-specific workbook sections are still mapped separately during
    # import/report generation; the master dropdown intentionally exposes all
    # available subdivisions so a newly added feeder can be assigned directly.
    return ALL_SUBDIVISIONS

def division_flow_from_entry_type(entry_type):
    return "EXPORT" if entry_type == "C" else "IMPORT"

def division_row_ranges(sheet_name):
    return {
        "Jorhat-ONE": [(4,29)],
        "Jorhat-TWO": [(5,10),(13,21),(24,26)],
        "Teok^": [(5,19),(22,29)],
    }.get(sheet_name, [])


def import_division_workbook(con, uploaded_bytes, year, month, division_name, overwrite=False):
    if division_name not in DIVISIONS:
        raise ValueError("Unknown division.")
    sheet_name = DIVISIONS[division_name]
    wb = openpyxl.load_workbook(io.BytesIO(uploaded_bytes), data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Workbook does not contain the '{sheet_name}' sheet.")
    ws = wb[sheet_name]
    div_id = DIVISION_TO_ID[division_name]

    imported = reused = new_master = mapped = 0
    skipped = []

    try:
        con.execute("BEGIN")
        for start, end in division_row_ranges(sheet_name):
            current_sub = None
            for r in range(start, end + 1):
                a = merged_cell_value(ws, r, 1)
                b = merged_cell_value(ws, r, 2)
                d = merged_cell_value(ws, r, 4)
                e = merged_cell_value(ws, r, 5)
                voltage = merged_cell_value(ws, r, 3)
                f = merged_cell_value(ws, r, 6)
                g = merged_cell_value(ws, r, 7)
                i_cell = ws.cell(r, 9).value
                j_cell = ws.cell(r, 10).value

                if isinstance(a, str) and "SUB-DIVISION" in a.upper():
                    current_sub = a.strip()
                    continue
                if not isinstance(f, (int, float)) and not isinstance(g, (int, float)):
                    continue
                if d is None and b is None:
                    continue

                feeder_name = str(b).strip() if b is not None else f"{division_name} row {r}"
                meter = str(d).strip() if d is not None else f"__{sheet_name}__ROW__{r}"
                mf = _to_float(e) or 1.0
                present = _to_float(g)
                last = _to_float(f)
                if present is None:
                    skipped.append(f"Row {r}: {feeder_name} — present reading is blank")
                    continue

                flows = []
                if i_cell is not None:
                    flows.append("IMPORT")
                if j_cell is not None:
                    flows.append("EXPORT")
                if not flows:
                    continue

                for flow in flows:
                    fid, created = find_or_create_meter_master(
                        con, feeder_name, meter, mf, preferred_entry_type_for_flow(flow),
                        year, month, present, division_name,
                        normalize_subdivision(division_name, current_sub), _to_float(voltage), flow
                    )
                    new_master += int(created)

                    existing = con.execute(
                        "SELECT id FROM monthly_readings WHERE feeder_id=? AND year=? AND month=?",
                        (fid, year, month)
                    ).fetchone()

                    prior = con.execute(
                        '''SELECT 1 FROM monthly_readings
                           WHERE feeder_id=? AND
                           (year < ? OR (year=? AND month < ?)) LIMIT 1''',
                        (fid, year, year, month)
                    ).fetchone()
                    if existing is None and prior is None and last is not None:
                        con.execute(
                            "UPDATE feeder_master SET initial_reading_kwh=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                            (last, fid)
                        )

                    if existing is None:
                        con.execute(
                            "INSERT INTO monthly_readings (feeder_id,year,month,reading_kwh,remarks) VALUES(?,?,?,?,?)",
                            (fid, year, month, present, f"Imported from {division_name}")
                        )
                        imported += 1
                    elif overwrite:
                        # The master monthly reading is shared, but a division
                        # workbook can contain the same meter more than once
                        # with different import/export readings.  The exact
                        # row-level value is stored separately below.
                        con.execute(
                            "UPDATE monthly_readings SET reading_kwh=?, remarks=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                            (present, f"Imported from {division_name}", existing["id"])
                        )
                        reused += 1
                    else:
                        reused += 1

                    con.execute(
                        '''INSERT INTO division_row_map
                           (division_id,sheet_row,sub_division,feeder_id,flow_direction,source_feeder_name,source_meter_no,baseline_reading_kwh)
                           VALUES(?,?,?,?,?,?,?,?)
                           ON CONFLICT(division_id,sheet_row,flow_direction)
                           DO UPDATE SET feeder_id=excluded.feeder_id,
                             sub_division=excluded.sub_division,
                             source_feeder_name=excluded.source_feeder_name,
                             source_meter_no=excluded.source_meter_no,
                             baseline_reading_kwh=excluded.baseline_reading_kwh,
                             active=1''',
                        (div_id, r, current_sub, fid, flow, feeder_name, meter, last)
                    )
                    map_row = con.execute(
                        "SELECT id FROM division_row_map WHERE division_id=? AND sheet_row=? AND flow_direction=?",
                        (div_id, r, flow)
                    ).fetchone()
                    if map_row is not None:
                        con.execute(
                            """INSERT INTO division_row_readings (division_map_id,year,month,reading_kwh)
                               VALUES(?,?,?,?)
                               ON CONFLICT(division_map_id,year,month) DO UPDATE SET
                                 reading_kwh=excluded.reading_kwh, updated_at=CURRENT_TIMESTAMP""",
                            (map_row["id"], year, month, present)
                        )
                    mapped += 1
        con.commit()
    except Exception:
        con.rollback()
        raise

    return {"imported": imported, "reused": reused, "new_master": new_master,
            "mapped": mapped, "skipped": skipped}



def import_all_division_workbook(con, uploaded_bytes, year, month, overwrite=False):
    wb = openpyxl.load_workbook(io.BytesIO(uploaded_bytes), data_only=False)
    results = []
    for division_name, sheet_name in DIVISIONS.items():
        if sheet_name in wb.sheetnames:
            results.append(
                import_division_workbook(
                    con, uploaded_bytes, year, month, division_name, overwrite
                )
            )
    if not results:
        raise ValueError("No Jorhat division sheets were found in the uploaded workbook.")
    return {
        "divisions": len(results),
        "imported": sum(x["imported"] for x in results),
        "reused": sum(x["reused"] for x in results),
        "new_master": sum(x["new_master"] for x in results),
        "mapped": sum(x["mapped"] for x in results),
        "skipped": [item for x in results for item in x["skipped"]],
    }


def bootstrap_division_template(con):
    div_path = APP_DIR / "MU_Injection_All_Divisions_Template.xlsx"
    if not div_path.exists():
        return
    ensure_division_master(con)
    try:
        data = div_path.read_bytes()
        for division_name in DIVISIONS:
            n = con.execute(
                "SELECT COUNT(*) n FROM division_row_map WHERE division_id=?",
                (DIVISION_TO_ID[division_name],)
            ).fetchone()["n"]
            if n == 0:
                import_division_workbook(con, data, 2026, 6, division_name, False)
    except Exception:
        pass


def get_division_id(con, division_name):
    row = con.execute(
        "SELECT id FROM division_master WHERE division_name=?", (division_name,)
    ).fetchone()
    return None if row is None else row["id"]


def get_scope_rows(con, year, month, scope, division_name=None):
    if scope == "Circle":
        return get_reading_rows(con, year, month)
    div_id = get_division_id(con, division_name)
    if div_id is None:
        return []
    return con.execute(
        '''SELECT DISTINCT f.*, r.id AS reading_id, r.reading_kwh, r.remarks, r.direct_mu, r.direct_mu_note
           FROM feeder_master f
           LEFT JOIN division_row_map dm ON dm.feeder_id=f.id AND dm.division_id=? AND dm.active=1
           LEFT JOIN monthly_readings r
             ON r.feeder_id=f.id AND r.year=? AND r.month=?
           WHERE f.active=1 AND (dm.feeder_id IS NOT NULL OR f.division_name=?)
           ORDER BY f.subdivision, f.feeder_name, f.meter_no''',
        (div_id, year, month, division_name)
    ).fetchall()


def division_energy_summary(con, year, month, division_name):
    div_id = get_division_id(con, division_name)
    totals = {"IMPORT": 0.0, "EXPORT": 0.0}
    if div_id is None:
        totals["NET"] = 0.0; return totals
    maps = con.execute(
        """SELECT m.*, f.mf, r.reading_kwh AS master_reading, r.direct_mu AS master_direct_mu, dr.reading_kwh AS row_reading
           FROM division_row_map m JOIN feeder_master f ON f.id=m.feeder_id
           LEFT JOIN monthly_readings r ON r.feeder_id=f.id AND r.year=? AND r.month=?
           LEFT JOIN division_row_readings dr ON dr.division_map_id=m.id AND dr.year=? AND dr.month=?
           WHERE m.division_id=? AND m.active=1""",
        (year, month, year, month, div_id)).fetchall()
    for row in maps:
        if row["master_direct_mu"] is not None:
            energy_mwh=float(row["master_direct_mu"])*1000.0
        else:
            present = row["row_reading"] if row["row_reading"] is not None else row["master_reading"]
            if present is None: continue
            prior = con.execute("SELECT 1 FROM monthly_readings WHERE feeder_id=? AND (year < ? OR (year=? AND month < ?)) LIMIT 1", (row["feeder_id"],year,year,month)).fetchone()
            prev = float(row["baseline_reading_kwh"]) if prior is None and row["baseline_reading_kwh"] is not None else get_previous_reading(con,row["feeder_id"],year,month)
            if prev is None: continue
            energy_mwh=(float(present)-prev)*float(row["mf"])
        totals[row["flow_direction"]]+=energy_mwh
    totals["NET"]=totals["IMPORT"]-totals["EXPORT"]
    return totals


def insert_rows_preserve_merges(ws, idx, amount, copy_from):
    ranges=[]
    for rng in list(ws.merged_cells.ranges):
        ranges.append((rng.min_col,rng.min_row,rng.max_col,rng.max_row))
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))
    ws.insert_rows(idx, amount=amount)
    for rr in range(idx, idx+amount):
        copy_row(ws, copy_from, rr)
    for min_col,min_row,max_col,max_row in ranges:
        if min_row >= idx:
            min_row += amount; max_row += amount
        elif min_row < idx <= max_row:
            max_row += amount
        ws.merge_cells(start_row=min_row,start_column=min_col,end_row=max_row,end_column=max_col)

def division_report(con, year, month, division_name):
    template = APP_DIR / "MU_Injection_All_Divisions_Template.xlsx"
    wb = openpyxl.load_workbook(template)
    sheet_name = DIVISIONS[division_name]
    ws = wb[sheet_name]
    div_id = get_division_id(con, division_name)

    next_month = 1 if month == 12 else month + 1
    next_year = year + 1 if month == 12 else year
    ws["A2"] = (f"For the month of {calendar.month_name[month]}, {year} "
                f"(Billed in {calendar.month_name[next_month]} , {next_year})")

    sections = {
        "Jorhat-ONE": [(None, 4, 29, 31)],
        "Jorhat-TWO": [("Titabar", 5, 10, 11), ("Mariani", 13, 21, 22), ("Majuli", 24, 26, 27)],
        "Teok^": [("Teok", 5, 19, 20), ("Kakojan", 22, 29, 30)],
    }[sheet_name]

    negative_maps = con.execute(
        """SELECT m.*, f.feeder_name, f.meter_no, f.mf, f.voltage_kv, r.reading_kwh
           FROM division_row_map m JOIN feeder_master f ON f.id=m.feeder_id
           LEFT JOIN monthly_readings r ON r.feeder_id=f.id AND r.year=? AND r.month=?
           WHERE m.division_id=? AND m.active=1 AND m.sheet_row < 0
           ORDER BY m.sub_division, f.feeder_name, f.meter_no""",
        (year, month, div_id)).fetchall()

    insertion_by_sub = {}
    for sub, start_row, end_row, total_row in sections:
        insertion_by_sub[sub] = sum(1 for x in negative_maps if (x["sub_division"] or None) == sub)

    for sub, start_row, end_row, total_row in reversed(sections):
        count = insertion_by_sub[sub]
        if count:
            copy_from = start_row
            insert_rows_preserve_merges(ws, total_row, count, copy_from)

    shift = 0
    live_sections = []
    for sub, start_row, end_row, total_row in sections:
        start_row += shift; end_row += shift; total_row += shift
        count = insertion_by_sub[sub]
        end_row += count; total_row += count
        live_sections.append((sub, start_row, end_row, total_row))
        shift += count

    for sub, start_row, end_row, total_row in live_sections:
        for r in range(start_row, end_row + 1):
            for c in range(6, 13):
                ws.cell(r, c).value = None

    maps = con.execute(
        """SELECT m.sheet_row, m.sub_division, m.flow_direction, m.feeder_id,
                  m.baseline_reading_kwh, f.feeder_name, f.meter_no, f.mf,
                  f.voltage_kv, r.reading_kwh AS master_reading, r.direct_mu AS master_direct_mu, dr.reading_kwh AS row_reading
           FROM division_row_map m JOIN feeder_master f ON f.id=m.feeder_id
           LEFT JOIN monthly_readings r ON r.feeder_id=f.id AND r.year=? AND r.month=?
           LEFT JOIN division_row_readings dr ON dr.division_map_id=m.id AND dr.year=? AND dr.month=?
           WHERE m.division_id=? AND m.active=1
           ORDER BY m.sheet_row, m.flow_direction""",
        (year, month, year, month, div_id)).fetchall()

    used_rows = {x["sheet_row"] for x in maps if x["sheet_row"] > 0}
    allocated = []
    for sub, start_row, end_row, total_row in live_sections:
        new_items = [x for x in maps if x["sheet_row"] < 0 and (x["sub_division"] or None) == sub]
        free = [r for r in range(end_row - len(new_items) + 1, end_row + 1) if r not in used_rows]
        for item, rr in zip(new_items, free):
            allocated.append((rr, item)); used_rows.add(rr)

    rows_by_row = {}
    for row in maps:
        if row["sheet_row"] > 0:
            rows_by_row.setdefault(row["sheet_row"], []).append(row)
    for rr, item in allocated:
        rows_by_row.setdefault(rr, []).append(item)

    for r, rowmaps in sorted(rows_by_row.items()):
        base = rowmaps[0]
        present = base["row_reading"] if base["row_reading"] is not None else base["master_reading"]
        direct_mu = base["master_direct_mu"]
        flows = {x["flow_direction"] for x in rowmaps}
        if direct_mu is not None:
            # Meter unavailable/defective: write the supplied MU directly.
            # Import is positive; Export is negative in Net Energy Injection.
            signed_mu = float(direct_mu) * (-1 if "EXPORT" in flows and "IMPORT" not in flows else 1)
            ws.cell(r, 11).value = signed_mu * 1000.0
            ws.cell(r, 12).value = signed_mu
            if not isinstance(ws.cell(r,2), MergedCell): ws.cell(r,2).value = base["feeder_name"]
            if not isinstance(ws.cell(r,4), MergedCell): ws.cell(r,4).value = base["meter_no"]
            if not isinstance(ws.cell(r,5), MergedCell): ws.cell(r,5).value = base["mf"]
            if not isinstance(ws.cell(r,3), MergedCell): ws.cell(r,3).value = base["voltage_kv"]
            continue
        if present is None:
            continue
        prev = get_previous_reading(con, base["feeder_id"], year, month)
        prior = con.execute("SELECT 1 FROM monthly_readings WHERE feeder_id=? AND (year < ? OR (year=? AND month < ?)) LIMIT 1", (base["feeder_id"], year, year, month)).fetchone()
        if prior is None and base["baseline_reading_kwh"] is not None:
            prev = float(base["baseline_reading_kwh"])
        if prev is None:
            continue
        ws.cell(r, 6).value = prev
        ws.cell(r, 7).value = present
        ws.cell(r, 8).value = f"=G{r}-F{r}"
        if not isinstance(ws.cell(r,5), MergedCell): ws.cell(r,5).value = base["mf"]
        if not isinstance(ws.cell(r,4), MergedCell): ws.cell(r,4).value = base["meter_no"]
        if not isinstance(ws.cell(r,3), MergedCell): ws.cell(r,3).value = base["voltage_kv"]
        if not isinstance(ws.cell(r,2), MergedCell): ws.cell(r,2).value = base["feeder_name"]
        if "IMPORT" in flows: ws.cell(r, 9).value = f"=H{r}*E{r}"
        if "EXPORT" in flows: ws.cell(r, 10).value = f"=H{r}*E{r}"
        if flows:
            ws.cell(r, 11).value = f"=I{r}-J{r}"
            ws.cell(r, 12).value = f"=K{r}/1000"

    if sheet_name == "Jorhat-ONE":
        _, a_start, a_end, a_total = live_sections[0]
        ws.cell(a_total, 11).value = "TOTAL:-"; ws.cell(a_total, 12).value = f"=SUM(L{a_start}:L{a_end})"
        summary_row = None
        final_row = None
        for rr in range(1, ws.max_row+1):
            text=ws.cell(rr,2).value
            if isinstance(text,str) and "TOTAL ENERGY INJECTION" in text: summary_row=rr
            if isinstance(text,str) and text.strip().startswith("JORHAT-ONE") : final_row=rr
        if summary_row: ws.cell(summary_row,9).value=f"=L{a_total}"
        if final_row and summary_row: ws.cell(final_row,7).value=f"=I{summary_row}"
    elif sheet_name == "Teok^":
        _, a_start, a_end, a_total = live_sections[0]; _, b_start, b_end, b_total = live_sections[1]
        ws.cell(a_total,11).value="TOTAL:-"; ws.cell(a_total,12).value=f"=SUM(L{a_start}:L{a_end})"
        ws.cell(b_total,11).value="TOTAL:-"; ws.cell(b_total,12).value=f"=SUM(L{b_start}:L{b_end})"
        teok_sum = kakojan_sum = total_power = None
        for r in range(1,ws.max_row+1):
            texts=[ws.cell(r,c).value for c in range(1, min(ws.max_column,12)+1)]
            joined=" | ".join(str(x) for x in texts if x is not None)
            if "ENERGY INJECTED TO TEOK ELEC. SUB-DIV" in joined:
                teok_sum=r; ws.cell(r,9).value=f"=L{a_total}"
            elif "ENERGY INJECTED TO KAKOJAN ELEC. SUB-DIV" in joined:
                kakojan_sum=r; ws.cell(r,9).value=f"=L{b_total}"
            elif "TOTAL POWER" in joined:
                total_power=r
        if total_power and teok_sum and kakojan_sum: ws.cell(total_power,9).value=f"=I{teok_sum}+I{kakojan_sum}"
    elif sheet_name == "Jorhat-TWO":
        totals=[]
        for sub,a_start,a_end,a_total in live_sections:
            ws.cell(a_total,11).value="TOTAL:-"; ws.cell(a_total,12).value=f"=SUM(L{a_start}:L{a_end})"; totals.append(a_total)
        sum_rows=[]; total_div=None
        for r in range(1,ws.max_row+1):
            texts=[ws.cell(r,c).value for c in range(1,min(ws.max_column,12)+1)]
            joined=" | ".join(str(x) for x in texts if x is not None)
            if "ENERGY INJECTED TO TITABAR" in joined: sum_rows.append(r); ws.cell(r,9).value=f"=L{totals[0]}"
            elif "ENERGY INJECTED TO MARIANI" in joined: sum_rows.append(r); ws.cell(r,9).value=f"=L{totals[1]}"
            elif "ENERGY INJECTED TO MAJULI" in joined: sum_rows.append(r); ws.cell(r,9).value=f"=L{totals[2]}"
            elif "TOTAL ENERGY INJECTION" in joined: total_div=r
        if total_div and len(sum_rows)==3: ws.cell(total_div,12).value="="+"+".join(f"I{x}" for x in sum_rows)

    try:
        wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode="auto"
    except Exception: pass
    out=io.BytesIO(); wb.save(out); out.seek(0); return out


def bootstrap_from_template(con):
    """Seed feeder master + June 2026 readings from the supplied workbook once."""
    if not TEMPLATE_PATH.exists():
        return
    existing = con.execute("SELECT COUNT(*) n FROM feeder_master").fetchone()["n"]
    if existing:
        return

    try:
        wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
        ws = wb["MU inj JEC"]
    except Exception:
        return

    ranges = {"A": range(7,31), "B": range(34,43), "C": range(46,51)}
    for sec, rows in ranges.items():
        for r in rows:
            feeder = ws.cell(r,2).value
            meter = ws.cell(r,3).value
            last = ws.cell(r,4).value
            present = ws.cell(r,5).value
            mf = ws.cell(r,7).value
            remarks = ws.cell(r,9).value
            if feeder is None:
                continue

            feeder = str(feeder).strip()
            meter = str(meter).strip() if meter is not None else ""
            if not meter:
                continue
            mf = float(mf) if isinstance(mf,(int,float)) else 1.0
            initial = float(last) if isinstance(last,(int,float)) else 0.0
            present = float(present) if isinstance(present,(int,float)) else None

            try:
                con.execute("""
                    INSERT INTO feeder_master
                    (feeder_name,meter_no,mf,entry_type,initial_reading_kwh,energy_direction)
                    VALUES(?,?,?,?,?,?)
                """,(feeder,meter,mf,sec,initial, "EXPORT" if sec=="C" else "IMPORT"))
            except sqlite3.IntegrityError:
                continue

            fid = con.execute("""
                SELECT id FROM feeder_master
                WHERE meter_no=? AND entry_type=?
            """,(meter,sec)).fetchone()["id"]

            if present is not None:
                # The supplied workbook is June 2026.
                con.execute("""
                    INSERT OR IGNORE INTO monthly_readings
                    (feeder_id,year,month,reading_kwh,remarks)
                    VALUES(?,?,?,?,?)
                """,(fid,2026,6,present,remarks))

    con.commit()

def previous_period(year, month):
    return (year-1,12) if month == 1 else (year,month-1)

def get_previous_reading(con, feeder_id, year, month):
    py, pm = previous_period(year, month)
    row = con.execute("""
        SELECT reading_kwh
        FROM monthly_readings
        WHERE feeder_id=? AND year=? AND month=?
    """, (feeder_id, py, pm)).fetchone()
    if row is not None:
        return float(row["reading_kwh"])

    # If there is no previous monthly record, use the feeder's initial reading
    # only for the first monthly reading.
    row = con.execute("""
        SELECT initial_reading_kwh
        FROM feeder_master WHERE id=?
    """, (feeder_id,)).fetchone()
    return None if row is None else float(row["initial_reading_kwh"])

def save_reading(con, feeder_id, year, month, reading_kwh, remarks=""):
    existing = con.execute(
        "SELECT last_reading_kwh FROM monthly_readings WHERE feeder_id=? AND year=? AND month=?",
        (feeder_id, year, month)
    ).fetchone()
    last_reading = None if existing is None else existing["last_reading_kwh"]
    if last_reading is None:
        last_reading = get_previous_reading(con, feeder_id, year, month)

    con.execute("""
        INSERT INTO monthly_readings
            (feeder_id,year,month,reading_kwh,last_reading_kwh,remarks)
        VALUES(?,?,?,?,?,?)
        ON CONFLICT(feeder_id,year,month)
        DO UPDATE SET
            reading_kwh=excluded.reading_kwh,
            last_reading_kwh=COALESCE(monthly_readings.last_reading_kwh, excluded.last_reading_kwh),
            remarks=excluded.remarks,
            direct_mu=NULL,
            direct_mu_note=NULL,
            updated_at=CURRENT_TIMESTAMP
    """, (feeder_id,year,month,reading_kwh,last_reading,remarks))
    con.execute("""
        UPDATE division_row_readings
        SET reading_kwh=?, updated_at=CURRENT_TIMESTAMP
        WHERE division_map_id IN (SELECT id FROM division_row_map WHERE feeder_id=?)
          AND year=? AND month=?
    """, (reading_kwh, feeder_id, year, month))
    con.commit()

def delete_reading(con, feeder_id, year, month):
    con.execute("""
        DELETE FROM monthly_readings
        WHERE feeder_id=? AND year=? AND month=?
    """, (feeder_id,year,month))
    con.execute("""
        DELETE FROM division_row_readings
        WHERE division_map_id IN (SELECT id FROM division_row_map WHERE feeder_id=?)
          AND year=? AND month=?
    """, (feeder_id,year,month))
    con.commit()

def add_feeder(con, feeder_name, meter_no, mf, entry_type, initial, division_name, subdivision, voltage_kv, energy_direction, selection_type=None):
    flow_dir = energy_direction if energy_direction in ('IMPORT', 'EXPORT') else division_flow_from_entry_type(entry_type)
    con.execute("""INSERT INTO feeder_master
        (feeder_name,meter_no,mf,entry_type,initial_reading_kwh,division_name,subdivision,voltage_kv,energy_direction,selection_type)
        VALUES(?,?,?,?,?,?,?,?,?,?)""",
        (feeder_name.strip(),meter_no.strip(),float(mf),entry_type,float(initial),division_name,subdivision,voltage_kv,flow_dir,selection_type or TYPE_LABELS.get(entry_type,"None")))
    fid=con.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
    if division_name:
        div_id=get_division_id(con, division_name)
        if div_id:
            con.execute("""INSERT OR IGNORE INTO division_row_map
                (division_id,sheet_row,sub_division,feeder_id,flow_direction,source_feeder_name,source_meter_no,baseline_reading_kwh)
                VALUES(?,?,?,?,?,?,?,?)""",
                (div_id,-fid,subdivision,fid,flow_dir,feeder_name,meter_no,initial))
    con.commit()

def get_meter_change(con, feeder_id, year, month):
    return con.execute(
        """SELECT * FROM meter_change_history
           WHERE feeder_id=? AND year=? AND month=? LIMIT 1""",
        (feeder_id, year, month)
    ).fetchone()


def record_meter_change(con, feeder_id, year, month, old_row, new_meter_no, new_mf):
    if old_row["meter_no"] == new_meter_no and abs(float(old_row["mf"]) - float(new_mf)) < 1e-12:
        return
    old_last = get_previous_reading(con, feeder_id, year, month)
    existing = con.execute(
        """SELECT id FROM meter_change_history
           WHERE feeder_id=? AND year=? AND month=? LIMIT 1""",
        (feeder_id, year, month)
    ).fetchone()
    if existing is None:
        con.execute(
            """INSERT INTO meter_change_history
               (feeder_id,year,month,old_feeder_name,old_meter_no,old_mf,
                old_last_reading_kwh,old_entry_type,new_meter_no,new_mf)
               VALUES(?,?,?,?,?,?,?,?,?,?)""",
            (feeder_id,year,month,old_row["feeder_name"],old_row["meter_no"],
             float(old_row["mf"]),old_last,old_row["entry_type"],
             new_meter_no,float(new_mf))
        )
    else:
        # Preserve the first/original old meter for this month. Only the latest
        # new meter/MF is updated if the master is edited again in the same month.
        con.execute(
            """UPDATE meter_change_history
               SET new_meter_no=?, new_mf=?, created_at=CURRENT_TIMESTAMP
               WHERE id=?""",
            (new_meter_no,float(new_mf),existing["id"])
        )


def update_feeder(con, feeder_id, feeder_name, meter_no, mf, entry_type, initial, division_name, subdivision, voltage_kv, energy_direction, selection_type=None, year=None, month=None):
    flow_dir = energy_direction if energy_direction in ('IMPORT', 'EXPORT') else division_flow_from_entry_type(entry_type)
    old_row = con.execute("SELECT * FROM feeder_master WHERE id=?", (feeder_id,)).fetchone()
    if old_row is None:
        raise ValueError("Feeder master record no longer exists.")
    new_meter_no = meter_no.strip()
    new_mf = float(mf)
    if year is not None and month is not None:
        record_meter_change(con, feeder_id, year, month, old_row, new_meter_no, new_mf)
    con.execute("""UPDATE feeder_master
        SET feeder_name=?,meter_no=?,mf=?,entry_type=?,initial_reading_kwh=?,
            division_name=?,subdivision=?,voltage_kv=?,energy_direction=?,selection_type=?,updated_at=CURRENT_TIMESTAMP
        WHERE id=?""",
        (feeder_name.strip(),new_meter_no,new_mf,entry_type,float(initial),division_name,subdivision,voltage_kv,flow_dir,selection_type or TYPE_LABELS.get(entry_type,"None"),feeder_id))
    con.execute("DELETE FROM division_row_map WHERE feeder_id=? AND sheet_row < 0",(feeder_id,))
    if division_name:
        div_id=get_division_id(con, division_name)
        if div_id:
            con.execute("""INSERT OR IGNORE INTO division_row_map
                (division_id,sheet_row,sub_division,feeder_id,flow_direction,source_feeder_name,source_meter_no,baseline_reading_kwh)
                VALUES(?,?,?,?,?,?,?,?)""",
                (div_id,-feeder_id,subdivision,feeder_id,flow_dir,feeder_name,meter_no,initial))
    con.commit()

def delete_feeder(con, feeder_id):
    con.execute("DELETE FROM feeder_master WHERE id=?", (feeder_id,))
    con.commit()

def get_feeders(con, active_only=True):
    sql = "SELECT * FROM feeder_master"
    if active_only:
        sql += " WHERE active=1"
    sql += " ORDER BY entry_type, feeder_name, meter_no"
    return con.execute(sql).fetchall()

def get_reading_rows(con, year, month):
    return con.execute("""
        SELECT f.*, r.id AS reading_id, r.reading_kwh, r.last_reading_kwh, r.remarks, r.direct_mu, r.direct_mu_note
        FROM feeder_master f
        LEFT JOIN monthly_readings r
          ON r.feeder_id=f.id AND r.year=? AND r.month=?
        WHERE f.active=1
        ORDER BY f.entry_type, f.feeder_name, f.meter_no
    """, (year,month)).fetchall()

def get_report_reading_rows(con, year, month):
    """Return the exact monthly rows represented in the generated Circle Excel report.
    When source-report metadata exists for the month, use that set instead of the
    current feeder master so dashboard totals and Excel totals use identical rows.
    """
    rows = con.execute("""
        SELECT f.*, r.id AS reading_id, r.reading_kwh, r.last_reading_kwh,
               r.remarks, r.direct_mu, r.direct_mu_note,
               r.report_section, r.report_order
        FROM monthly_readings r
        JOIN feeder_master f ON f.id=r.feeder_id
        WHERE r.year=? AND r.month=?
          AND r.report_section IN ('A','B','C')
        ORDER BY r.report_section, r.report_order, r.id
    """, (year, month)).fetchall()
    return rows

def _is_mu_template_pending(row):
    # Some legacy/read-only queries only select the feeder master fields.
    # Missing monthly columns must therefore be treated as "not a template
    # placeholder", not as an exception.
    if row["reading_id"] is None:
        return False
    keys = row.keys()
    direct_mu = row["direct_mu"] if "direct_mu" in keys else None
    reading_kwh = row["reading_kwh"] if "reading_kwh" in keys else None
    remarks = row["remarks"] if "remarks" in keys else None
    return (
        direct_mu is None
        and float(reading_kwh or 0) == 0.0
        and str(remarks or "").startswith("Added from MU Template")
    )

def calculate_mu(con, row, year, month):
    if row["reading_id"] is None or _is_mu_template_pending(row):
        return None
    # A direct MU entry is a month-specific fallback for lost/defective meters.
    # It is stored as a positive magnitude; report logic applies Import/Export sign.
    if row["direct_mu"] is not None:
        return float(row["direct_mu"])
    change = get_meter_change(con, row["id"], year, month)
    if change is not None and change["old_meter_no"] != change["new_meter_no"]:
        prev = float(row["initial_reading_kwh"])
    else:
        prev = row["last_reading_kwh"]
        if prev is None:
            prev = get_previous_reading(con, row["id"], year, month)
    if prev is None or row["reading_kwh"] is None:
        return None
    return (float(row["reading_kwh"]) - float(prev)) * float(row["mf"]) / 1000.0


def _to_float(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _find_section_rows(ws, sec):
    """Find data rows for section A/B/C in the supplied monthly workbook."""
    header_tokens = {"A": "A.", "B": "B.", "C": "C."}
    subtotal_tokens = {"A": "SUB-TOTAL (A)", "B": "SUB-TOTAL (B)", "C": "SUB-TOTAL (C)"}
    header = None
    subtotal = None
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 4)]
        text = " ".join(str(v).strip() for v in vals if v is not None)
        if header is None and header_tokens[sec] in text:
            header = r
        if subtotal_tokens[sec] in text:
            subtotal = r
            break
    if header is None or subtotal is None or subtotal <= header + 1:
        return []
    return list(range(header + 1, subtotal))

def import_month_excel(con, uploaded_bytes, year, month, overwrite=False):
    """Import one monthly MU workbook using meter readings/MF plus
    historical Direct MU exceptions.

    The workbook's Last Reading is used as the feeder's initial baseline when
    the database has no monthly reading before the imported month. This makes
    a historical import such as January 2026 self-contained: January Present
    becomes the stored January reading and January Last becomes its baseline.
    When column H contains a non-formula/manual MU value that does not match
    Present-Last times MF, it is preserved as the month-specific Direct MU
    fallback so historical reports remain faithful to the source workbook.
    """
    wb = openpyxl.load_workbook(io.BytesIO(uploaded_bytes), data_only=False)
    wb_values = openpyxl.load_workbook(io.BytesIO(uploaded_bytes), data_only=True)
    ws = wb["MU inj JEC"] if "MU inj JEC" in wb.sheetnames else wb.active
    ws_values = wb_values["MU inj JEC"] if "MU inj JEC" in wb_values.sheetnames else wb_values.active
    imported = []
    skipped = []
    updated_master = 0
    new_master = 0
    overwritten = 0
    direct_mu_imported = 0

    try:
        con.execute("BEGIN")
        for sec in ["A", "B", "C"]:
            for r in _find_section_rows(ws, sec):
                source_feeder_name = ws.cell(r, 2).value
                source_meter_no = ws.cell(r, 3).value
                source_sl_no = _to_float(ws.cell(r, 1).value)
                feeder = source_feeder_name
                meter = source_meter_no
                last = _to_float(ws.cell(r, 4).value)
                present = _to_float(ws.cell(r, 5).value)
                mf = _to_float(ws.cell(r, 7).value)
                energy_mu = _to_float(ws_values.cell(r, 8).value)
                remarks = ws.cell(r, 9).value

                source_feeder_name = "" if source_feeder_name is None else str(source_feeder_name).strip()
                source_meter_no = "" if source_meter_no is None else str(source_meter_no).strip()
                feeder = source_feeder_name
                meter = source_meter_no

                # Some historical circle workbooks contain a manual MU value
                # in column H even when meter readings are unavailable or when
                # Present-Last does not reproduce the stated MU. Preserve that
                # value as the app's explicit Direct MU fallback instead of
                # silently discarding it.
                computed_mu = None
                if last is not None and present is not None:
                    computed_mu = (present - last) * (1.0 if mf is None else mf) / 1000.0
                manual_mu = None
                if energy_mu is not None and (
                    computed_mu is None or abs(float(energy_mu) - float(computed_mu)) > 1e-9
                ):
                    manual_mu = float(energy_mu)

                if not feeder and not meter and present is None and manual_mu is None:
                    skipped.append(f"{sec}: blank source row")
                    continue
                feeder_for_master = feeder or f"__REPORT_ROW__{sec}_{r}"

                if not meter:
                    if manual_mu is None:
                        skipped.append(f"{sec}: {feeder} — meter number is blank")
                        continue
                    # Stable synthetic identifier for a workbook row that has
                    # only a manual MU value and no meter reading.
                    meter = f"__DIRECT_MU__{sec}_{r}"

                if present is None and manual_mu is None:
                    skipped.append(f"{sec}: {feeder} / {meter} — present reading is blank")
                    continue
                mf = 1.0 if mf is None else mf

                feeder_id, created_master = find_or_create_meter_master(
                    con, feeder_for_master, meter, mf, sec
                )
                if created_master:
                    new_master += 1
                    if last is not None:
                        con.execute("""
                            UPDATE feeder_master
                            SET initial_reading_kwh=?, updated_at=CURRENT_TIMESTAMP
                            WHERE id=?
                        """, (last, feeder_id))
                else:
                    updated_master += 1
                    if last is not None:
                        prior = con.execute("""
                            SELECT 1 FROM monthly_readings
                            WHERE feeder_id=? AND (year < ? OR (year=? AND month < ?))
                            LIMIT 1
                        """, (feeder_id, year, year, month)).fetchone()
                        if prior is None:
                            con.execute("""
                                UPDATE feeder_master
                                SET initial_reading_kwh=?, updated_at=CURRENT_TIMESTAMP
                                WHERE id=?
                            """, (last, feeder_id))

                existing = con.execute("""
                    SELECT id FROM monthly_readings
                    WHERE feeder_id=? AND year=? AND month=?
                """, (feeder_id, year, month)).fetchone()
                if existing is not None and not overwrite:
                    skipped.append(f"{sec}: {feeder} / {meter} — monthly reading already exists")
                    continue

                # Normal rows are recalculated from Present-Last and MF.
                # A detected historical exception is stored as Direct MU above.
                con.execute("""
                    INSERT INTO monthly_readings
                    (feeder_id,year,month,reading_kwh,last_reading_kwh,remarks,
                     direct_mu,direct_mu_note,report_section,report_order,
                     report_feeder_name,report_meter_no,report_mf,report_sl_no)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(feeder_id,year,month)
                    DO UPDATE SET
                        reading_kwh=excluded.reading_kwh,
                        last_reading_kwh=excluded.last_reading_kwh,
                        remarks=excluded.remarks,
                        direct_mu=excluded.direct_mu,
                        direct_mu_note=excluded.direct_mu_note,
                        report_section=excluded.report_section,
                        report_order=excluded.report_order,
                        report_feeder_name=excluded.report_feeder_name,
                        report_meter_no=excluded.report_meter_no,
                        report_mf=excluded.report_mf,
                        report_sl_no=excluded.report_sl_no,
                        updated_at=CURRENT_TIMESTAMP
                """, (
                    feeder_id, year, month,
                    0.0 if present is None else present,
                    last,
                    remarks,
                    manual_mu,
                    "Imported Direct MU from workbook column H" if manual_mu is not None else None,
                    sec,
                    r,
                    source_feeder_name,
                    source_meter_no,
                    mf,
                    source_sl_no
                ))
                if manual_mu is not None:
                    direct_mu_imported += 1
                if existing is not None:
                    overwritten += 1
                else:
                    imported.append(feeder)

        con.commit()
    except Exception:
        con.rollback()
        raise

    return {
        "imported": len(imported),
        "new_master": new_master,
        "updated_master": updated_master,
        "overwritten": overwritten,
        "direct_mu_imported": direct_mu_imported,
        "skipped": skipped,
    }


def month_energy_summary(con, year, month):
    # The Circle dashboard must use the same monthly report rows as the
    # generated Excel workbook. This prevents current feeder-master records
    # that were not present in the historical source workbook from changing
    # dashboard totals.
    rows = get_report_reading_rows(con, year, month)
    if not rows:
        # Preserve the existing workflow for months entered manually without
        # imported report-row metadata.
        rows = get_reading_rows(con, year, month)

    totals = {"A": 0.0, "B": 0.0, "C": 0.0}
    for row in rows:
        mu = calculate_mu(con, row, year, month)
        if mu is not None:
            # Historical Excel section is authoritative for the Circle report.
            # Do not use the current feeder-master entry_type because it can be
            # changed later without changing the historical workbook section.
            report_section = row["report_section"] if "report_section" in row.keys() else None
            section = report_section if report_section in ("A", "B", "C") else row["entry_type"]
            totals[section] += mu
    totals["NET"] = totals["A"] + totals["B"] - totals["C"]
    return totals

def copy_style(src, dst):
    if src.has_style:
        dst._style = copy(src._style)
    if src.number_format:
        dst.number_format = src.number_format
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)

def copy_row(ws, src_row, dst_row):
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for c in range(1, ws.max_column+1):
        copy_style(ws.cell(src_row,c), ws.cell(dst_row,c))

def _report_rows_with_meter_continuity(con, year, month, sec):
    rows = list(con.execute("""
        SELECT f.*, r.reading_kwh, r.last_reading_kwh, r.direct_mu, r.remarks,
               r.report_feeder_name, r.report_meter_no, r.report_mf, r.report_sl_no,
               r.report_order
        FROM monthly_readings r
        JOIN feeder_master f ON f.id=r.feeder_id
        WHERE r.year=? AND r.month=? AND r.report_section=?
        ORDER BY r.report_order, r.id
    """, (year, month, sec)).fetchall())

    changes = con.execute(
        """SELECT h.*, f.feeder_name AS current_feeder_name,
                  f.meter_no AS current_meter_no, f.mf AS current_mf
           FROM meter_change_history h
           JOIN feeder_master f ON f.id=h.feeder_id
           WHERE h.year=? AND h.month=? AND h.old_entry_type=?
           ORDER BY h.id""",
        (year, month, sec)
    ).fetchall()

    # Show the previous meter/MF once, only in the month of the change.
    # The continuity row has zero MU so it does not alter report totals.
    for ch in changes:
        current = next((x for x in rows if x["id"] == ch["feeder_id"]), None)
        if current is None:
            continue
        item = dict(current)
        item["id"] = -int(ch["id"])
        item["reading_kwh"] = ch["old_last_reading_kwh"]
        item["last_reading_kwh"] = ch["old_last_reading_kwh"]
        item["direct_mu"] = 0.0
        item["remarks"] = "Previous meter/MF — continuity record"
        item["report_meter_no"] = ch["old_meter_no"]
        item["report_mf"] = ch["old_mf"]
        item["report_sl_no"] = None
        item["report_order"] = (current["report_order"] or 0) - 0.1
        item["report_feeder_name"] = ch["old_feeder_name"] or current["report_feeder_name"]
        rows.append(item)

    rows.sort(key=lambda x: (
        x["report_order"] if x["report_order"] is not None else 10**9,
        x["id"]
    ))
    return rows



def add_feeder_to_mu_template(con, feeder_id, year, month, section):
    """Add or move one feeder in the selected month's MU report section.
    This changes only the month-specific report metadata, not feeder-master
    classification, so historical/current master data remains intact.
    """
    row = con.execute("SELECT * FROM feeder_master WHERE id=? AND active=1 LIMIT 1", (feeder_id,)).fetchone()
    if row is None:
        return False, "Feeder is not present in the active Feeder Master."

    existing = con.execute(
        "SELECT * FROM monthly_readings WHERE feeder_id=? AND year=? AND month=? LIMIT 1",
        (feeder_id, year, month)
    ).fetchone()

    max_row = con.execute(
        "SELECT COALESCE(MAX(report_order),0) AS max_order FROM monthly_readings "
        "WHERE year=? AND month=? AND report_section=?",
        (year, month, section)
    ).fetchone()
    next_order = int(max_row["max_order"] or 0) + 1

    if existing is None:
        con.execute(
            """INSERT INTO monthly_readings
               (feeder_id,year,month,reading_kwh,remarks,report_section,report_order,
                report_feeder_name,report_meter_no,report_mf,report_sl_no)
               VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
            (
                feeder_id, year, month, 0.0,
                "Added from MU Template — pending reading",
                section, next_order,
                row["feeder_name"], row["meter_no"], float(row["mf"]), next_order
            )
        )
    else:
        con.execute(
            """UPDATE monthly_readings
               SET report_section=?,
                   report_order=?,
                   report_feeder_name=?,
                   report_meter_no=?,
                   report_mf=?,
                   report_sl_no=?,
                   updated_at=CURRENT_TIMESTAMP
               WHERE id=?""",
            (
                section, next_order,
                row["feeder_name"], row["meter_no"], float(row["mf"]), next_order,
                existing["id"]
            )
        )
    con.commit()
    return True, row["feeder_name"]


def ensure_mu_template_for_month(con, year, month):
    """Seed a new month from the current MU Template. Once monthly report rows
    exist, those month-specific rows are authoritative until the template is edited.
    """
    existing = con.execute(
        """SELECT 1 FROM monthly_readings
           WHERE year=? AND month=? AND report_section IN ('A','B','C') LIMIT 1""",
        (year, month)
    ).fetchone()
    if existing is not None:
        return

    rows = con.execute(
        """SELECT * FROM feeder_master
           WHERE active=1
             AND selection_type IN ('Import from GSS','Import from other circle','Export to other circle')
             AND meter_no NOT LIKE '__DIRECT_MU__%'
           ORDER BY CASE selection_type
             WHEN 'Import from GSS' THEN 1
             WHEN 'Import from other circle' THEN 2
             WHEN 'Export to other circle' THEN 3
             ELSE 4 END,
             feeder_name, meter_no"""
    ).fetchall()

    counters = {"A": 0, "B": 0, "C": 0}
    for row in rows:
        sec = LABEL_TO_TYPE.get(row["selection_type"])
        if sec not in counters:
            continue
        counters[sec] += 1
        con.execute(
            """INSERT INTO monthly_readings
               (feeder_id,year,month,reading_kwh,remarks,report_section,report_order,
                report_feeder_name,report_meter_no,report_mf,report_sl_no)
               VALUES(?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(feeder_id,year,month) DO NOTHING""",
            (
                row["id"], year, month, 0.0,
                "Added from MU Template — pending reading",
                sec, counters[sec],
                row["feeder_name"], row["meter_no"], float(row["mf"]), counters[sec]
            )
        )
    con.commit()

def build_report(con, year, month):
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["MU inj JEC"]

    # Keep the supplied workbook layout and formatting exactly as the template.
    # Only populate the existing data cells; formulas remain Excel formulas.
    sections = {
        "A": [7,30,32],
        "B": [34,42,44],
        "C": [46,50,51],
    }

    for sec in ["A","B","C"]:
        rows = _report_rows_with_meter_continuity(con, year, month, sec)

        start_row,end_row,subtotal=sections[sec]
        capacity=end_row-start_row+1
        if len(rows)>capacity:
            extra=len(rows)-capacity
            ws.insert_rows(subtotal, amount=extra)
            for rr in range(subtotal,subtotal+extra):
                copy_row(ws,end_row,rr)
            end_row += extra
            subtotal += extra
            sections[sec]=[start_row,end_row,subtotal]
            if sec=="A":
                sections["B"][0]+=extra; sections["B"][1]+=extra; sections["B"][2]+=extra
                sections["C"][0]+=extra; sections["C"][1]+=extra; sections["C"][2]+=extra
            elif sec=="B":
                sections["C"][0]+=extra; sections["C"][1]+=extra; sections["C"][2]+=extra

    for sec in ["A","B","C"]:
        start_row,end_row,subtotal=sections[sec]
        for r in range(start_row,end_row+1):
            for c in range(1,10):
                ws.cell(r,c).value=None

        rows=_report_rows_with_meter_continuity(con,year,month,sec)

        serial=0
        for row in rows:
            if _is_mu_template_pending(row):
                continue
            present=row["reading_kwh"]
            direct_mu=row["direct_mu"]
            source_last=row["last_reading_kwh"]
            change = get_meter_change(con, row["id"], year, month)
            if change is not None and row["meter_no"] == change["new_meter_no"] and change["old_meter_no"] != change["new_meter_no"]:
                source_last = float(row["initial_reading_kwh"])
            if source_last is None and present is not None and direct_mu is None:
                source_last=get_previous_reading(con,row["id"],year,month)

            if present is None and direct_mu is None:
                continue

            serial += 1
            r=start_row+serial-1
            meter_no = row["report_meter_no"]
            if meter_no is None or str(meter_no).startswith("__DIRECT_MU__"):
                meter_no = ""
            report_name = row["report_feeder_name"]
            if report_name is None or str(report_name).startswith("__REPORT_ROW__"):
                report_name = ""
            remarks=row["remarks"]

            ws.cell(r,1).value=row["report_sl_no"]
            ws.cell(r,2).value=report_name
            ws.cell(r,3).value=meter_no

            # Direct-MU-only rows do not invent Last/Present/MF values.
            direct_only = direct_mu is not None and (
                present is None or source_last is None
            )
            if direct_only:
                ws.cell(r,4).value=None
                ws.cell(r,5).value=None
                ws.cell(r,6).value=None
                ws.cell(r,7).value=None
                ws.cell(r,8).value=float(direct_mu)
                ws.cell(r,9).value=remarks
                continue

            if present is None:
                serial -= 1
                for c in range(1,10):
                    ws.cell(r,c).value=None
                continue

            if source_last is None:
                source_last=get_previous_reading(con,row["id"],year,month)
            if source_last is None:
                serial -= 1
                for c in range(1,10):
                    ws.cell(r,c).value=None
                continue

            ws.cell(r,4).value=source_last
            ws.cell(r,5).value=present
            ws.cell(r,6).value=f"=E{r}-D{r}"
            ws.cell(r,7).value=row["report_mf"] if row["report_mf"] is not None else row["mf"]
            ws.cell(r,8).value=(
                float(direct_mu)
                if direct_mu is not None
                else f"=F{r}*G{r}/1000"
            )
            ws.cell(r,9).value=remarks

    a_start,a_end,a_sub=sections["A"]
    b_start,b_end,b_sub=sections["B"]
    c_start,c_end,c_sub=sections["C"]

    ws.cell(a_sub,1).value="SUB-TOTAL (A)"
    ws.cell(a_sub,8).value=f"=SUM(H{a_start}:H{a_end})"
    ws.cell(b_sub,2).value="            SUB-TOTAL (B)"
    ws.cell(b_sub,8).value=f"=SUM(H{b_start}:H{b_end})"
    ws.cell(c_sub,2).value="              SUB-TOTAL (C)"
    ws.cell(c_sub,8).value=f"=SUM(H{c_start}:H{c_end})"

    total_row=c_sub+1
    ws.cell(total_row,1).value="TOTAL ENERGY INJECTED = (A+B)-C"
    ws.cell(total_row,8).value=f"=(H{a_sub}+H{b_sub})-H{c_sub}"

    received_row=total_row+2
    ws.cell(received_row,1).value="TOTAL ENERGY RECEIVED BY THE CIRCLE:-"
    ws.cell(received_row,6).value=f"=H{total_row}"
    ws.cell(received_row,8).value="MU"

    oa_row=received_row+2
    ws.cell(oa_row,1).value="D."
    ws.cell(oa_row,2).value="OPEN ACCESS ENERGY"
    ws.cell(oa_row,8).value="ENERGY (In M.U.)"
    ws.cell(oa_row,9).value="REMARKS"

    try:
        wb.calculation.fullCalcOnLoad=True
        wb.calculation.forceFullCalc=True
        wb.calculation.calcMode="auto"
    except Exception:
        pass

    out=io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def style():
    st.markdown("""
    <style>
    .stApp { background:#0e1117; color:#f8fafc; }
    .block-container { max-width:1500px; padding-top:1rem; }
    .hero {
      background:linear-gradient(135deg,#0f172a,#1d4ed8);
      padding:24px 28px;border-radius:18px;color:white;margin-bottom:18px;
      box-shadow:0 8px 30px rgba(0,0,0,.30);
    }
    .hero h1{margin:0;font-size:30px;color:white}.hero p{margin:5px 0 0;color:#dbeafe}
    .card{background:#161b22;border:1px solid #30363d;border-radius:14px;padding:16px;
          box-shadow:0 3px 12px rgba(0,0,0,.20)}
    .login-shell { text-align:center; margin:7vh auto 30px; }
    .login-brand { width:70px; height:70px; margin:0 auto 17px; border-radius:20px;
      display:flex; align-items:center; justify-content:center; font-size:36px;
      background:linear-gradient(145deg,#ff9d1c 0%,#ff4d4f 100%);
      border:1px solid rgba(255,255,255,.14);
      box-shadow:0 14px 38px rgba(249,115,22,.22), inset 0 1px 0 rgba(255,255,255,.18); }
    .login-title { font-size:36px; font-weight:800; color:#f8fafc; letter-spacing:-1px; }
    .login-subtitle { margin-top:8px; color:#94a3b8; font-size:14px; letter-spacing:.1px; }
    div[data-testid="column"]:has(.login-anchor) {
      background:linear-gradient(180deg,rgba(24,32,47,.98),rgba(14,20,31,.98));
      border:1px solid #293548; border-radius:22px; padding:30px 32px 24px;
      box-shadow:0 24px 70px rgba(0,0,0,.42), 0 0 0 1px rgba(255,255,255,.015) inset;
    }
    .login-anchor { height:0; margin:0; padding:0; }
    .login-heading { font-size:25px; font-weight:750; color:#f8fafc; margin:0 0 5px; letter-spacing:-.35px; }
    .login-hint { color:#8fa0b7; font-size:13px; margin:0 0 22px; }
    .login-section-label { color:#64748b; font-size:10px; font-weight:800; letter-spacing:1.5px; margin:0 0 8px; }
    div[data-testid="column"]:has(.login-anchor) div[data-testid="stTextInput"] input {      background:#0b1220 !important; border:1px solid #2e3b50 !important;
      color:#f8fafc !important; -webkit-text-fill-color:#f8fafc !important;
      border-radius:11px !important; min-height:46px !important; padding:0 14px !important;
      box-sizing:border-box !important; }
    div[data-testid="column"]:has(.login-anchor) div[data-testid="stTextInput"] input::placeholder { color:#64748b !important; opacity:1 !important; }
    div[data-testid="column"]:has(.login-anchor) div[data-testid="stTextInput"] input:focus {
      border-color:#3b82f6 !important; box-shadow:0 0 0 3px rgba(59,130,246,.14) !important; }
    div[data-testid="column"]:has(.login-anchor) div[data-testid="stButton"] > button {
      min-height:46px; height:46px; border-radius:11px; margin-top:7px;
      font-weight:700; font-size:14px; border:1px solid rgba(96,165,250,.35);
      box-shadow:0 8px 22px rgba(37,99,235,.20); }
    .login-security { display:flex; align-items:center; justify-content:center; gap:7px;
      color:#64748b; font-size:10px; margin-top:18px; padding-top:15px;
      border-top:1px solid #253043; }
    .security-dot { color:#22c55e; font-size:8px; }
    .security-sep { color:#334155; }
    .small{color:#94a3b8;font-size:13px}.big{font-size:25px;font-weight:700;color:#f8fafc}
    div[data-testid="stForm"]{border:1px solid #30363d;border-radius:12px;padding:14px;background:#161b22}
    div[data-testid="stMarkdownContainer"] p, div[data-testid="stMarkdownContainer"] span,
    div[data-testid="stMarkdownContainer"] label { color:#f8fafc; }
    div[data-testid="stCaptionContainer"] { color:#94a3b8; }
    div[data-testid="stTextInput"] input, div[data-testid="stNumberInput"] input, div[data-testid="stSelectbox"] input { color:#f8fafc; }
    /* Modern dark sidebar */
    section[data-testid="stSidebar"] { background:linear-gradient(180deg,#0b1220 0%,#111827 55%,#0f172a 100%); border-right:1px solid #263244; }
    section[data-testid="stSidebar"] > div { padding:1.1rem .85rem; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color:#f8fafc; }
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color:#cbd5e1; }
    section[data-testid="stSidebar"] hr { border-color:#263244; margin:.7rem 0; }
    /* Sidebar navigation buttons: real Streamlit buttons, equal sizing and centered text */
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
        width:100%; min-height:42px; height:42px; margin:0;
        border-radius:10px; font-weight:600; text-align:center;
        display:flex; align-items:center; justify-content:center;
        transition:all .15s ease;
    }
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button p {
        width:100%; text-align:center; margin:0;
    }
    section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover {
        border-color:#3b82f6; transform:translateY(-1px);
    }
    /* Keep the two Level controls exactly equal and centered */
    section[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] {
        align-items:center;
    }
    section[data-testid="stSidebar"] [data-testid="stHorizontalBlock"] > div {
        display:flex; align-items:center;
    }
    /* Navigation buttons use a compact, uniform gap */    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:has(> div > div[data-testid="stButton"]) {
        margin-bottom:6px;
    }
    /* Ensure selectbox internal dropdown buttons are not expanded or distorted */
    section[data-testid="stSidebar"] div[data-testid="stSelectbox"] button {
        width: auto !important;
        min-height: unset !important;
        height: auto !important;
        border: none !important;
        background: transparent !important;
        padding: 0 4px !important;
        box-shadow: none !important;
        transform: none !important;
    }
    /* Selectbox container and input text visible and clearly readable */
    div[data-testid="stSelectbox"] input {
        color: #f8fafc !important;
        -webkit-text-fill-color: #f8fafc !important;
        font-size: 14px !important;
        font-weight: 500 !important;
        opacity: 1 !important;
    }
    div[data-testid="stSelectbox"] > div > div {
        background-color: #111827 !important;
        border-color: #334155 !important;
        border-radius: 10px !important;
        color: #f8fafc !important;
    }
    div[data-testid="stSelectbox"] svg {
        fill: #94a3b8 !important;
        color: #94a3b8 !important;
    }
    div[data-testid="stSelectboxVirtualDropdown"] li,
    div[data-testid="stSelectboxVirtualDropdown"] [role="option"] {
        color: #f8fafc !important;
    }
    section[data-testid="stSidebar"] [data-testid="stNumberInput"] > div {
        background:#111827; border-color:#334155; border-radius:10px;
    }
    </style>
    """, unsafe_allow_html=True)

st.set_page_config(page_title="MU Injection Manager", page_icon="⚡", layout="wide")
style()

try:
    con=db()
    # Authentication schema must exist before the login screen queries app_users.
    # This is intentionally kept separate from db() so existing databases are not rebuilt.
    ensure_auth_table(con)
except Exception:
    st.session_state.pop("_postgres_db_connection", None)
    st.error("The MU Injection Manager service is waking up. Please wait a few seconds and retry.", icon=":material/cloud_sync:")
    st.info("No data has been changed. The service is recovering its connection.")
    if st.button("Retry connection", type="primary", width="stretch", key="retry_service_connection"):
        st.rerun()
    st.stop()

if not render_login(con):
    st.stop()

# Perform data/template bootstrapping only once per authenticated
# Streamlit session so page navigation does not repeat the expensive
# first-run database/template initialization.
if not st.session_state.get("_data_bootstrapped", False):
    bootstrap_from_template(con)
    ensure_division_master(con)
    bootstrap_division_template(con)
    st.session_state["_data_bootstrapped"] = True

ensure_mu_template_for_month(con, year, month)

now=datetime.now()
default_year=now.year
default_month=now.month

st.markdown("""
<div class="hero">
<h1>⚡ MU Injection Manager</h1>
<p>Meter master + monthly kWh readings + automatic MU report generation</p>
</div>
""",unsafe_allow_html=True)

with st.sidebar:
    auth_user = st.session_state.get("auth_username", "user")
    st.markdown(f'<div style="color:#94a3b8;font-size:12px;margin-bottom:8px;">Signed in as <b style="color:#f8fafc;">{auth_user}</b></div>', unsafe_allow_html=True)
    if st.button("↪  Sign out", key="logout_button", width="stretch"):
        logout()
        st.rerun()
    st.divider()
    st.markdown("### Period")
    year=st.number_input("Year",2000,2100,default_year,1)
    month_options=list(range(1,13))
    if st.session_state.get("sidebar_month") not in month_options:
        st.session_state["sidebar_month"] = default_month
    month=st.selectbox("Month",month_options,
                       format_func=lambda m:calendar.month_name[m],
                       key="sidebar_month")
    st.divider()
    st.markdown("### View")

    # Keep navigation in session state, but use button callbacks so the
    # selected value is written before Streamlit reruns the script.
    # This prevents page/scope clicks from being applied one rerun late.
    if "scope_level" not in st.session_state:
        st.session_state["scope_level"] = "Circle"
    if "selected_page" not in st.session_state:
        st.session_state["selected_page"] = "Enter Readings"

    def _set_nav_state(value):
        st.session_state["selected_page"] = value

    def _set_scope_state(value):
        st.session_state["scope_level"] = value

    scope=st.session_state["scope_level"]
    c1, c2 = st.columns(2, gap="small")
    with c1:
        st.button("Circle", key="scope_circle", width="stretch",
                  type="primary" if scope=="Circle" else "secondary",
                  on_click=_set_scope_state, args=("Circle",))
    with c2:
        st.button("Division", key="scope_division", width="stretch",
                  type="primary" if scope=="Division" else "secondary",
                  on_click=_set_scope_state, args=("Division",))

    division_name=None
    if scope=="Circle":
        st.caption("Jorhat Circle")
    else:
        division_options=list(DIVISIONS.keys())
        if st.session_state.get("selected_division") not in division_options:
            st.session_state["selected_division"] = division_options[0]
        division_name=st.selectbox("Division",division_options,key="selected_division")

    st.divider()
    st.markdown("### Pages")
    pages=[
        "Enter Readings",
        "Direct MU Entry",
        "All Feeder Readings",
        "Feeder Master",
        "MU Template",
        "Generate Excel",
        "Import Excel",
        "Dashboard"
    ]
    for idx, page_name in enumerate(pages):
        st.button(page_name, key=f"page_nav_{idx}", width="stretch",
                  type="primary" if st.session_state["selected_page"]==page_name else "secondary",
                  on_click=_set_nav_state, args=(page_name,))
    page=st.session_state["selected_page"]

# ---------- ENTER READINGS ----------
if page=="Enter Readings":
    scope_label = CIRCLE_NAME if scope=="Circle" else division_name
    st.subheader(f"Enter Readings — {scope_label} — {calendar.month_name[month]} {year}")
    st.caption("Monthly readings are shared between circle and division views for the same meter master.")
    if scope=="Circle":
            st.subheader(f"Enter Readings — {calendar.month_name[month]} {year}")
            st.caption("Only feeders without a reading for the selected month are shown. Enter the meter reading in kWh and save each row.")

            types=st.tabs(["Import from GSS","Import from other circle","Export to other circle"])

            for tab,sec in zip(types,["A","B","C"]):
                with tab:
                    all_rows=con.execute("""
                        SELECT f.*,r.id reading_id
                        FROM feeder_master f
                        LEFT JOIN monthly_readings r
                          ON r.feeder_id=f.id AND r.year=? AND r.month=?
                        WHERE f.active=1 AND f.entry_type=?
                        ORDER BY f.feeder_name,f.meter_no
                    """,(year,month,sec)).fetchall()

                    pending=[r for r in all_rows if r["reading_id"] is None or _is_mu_template_pending(r)]
                    done=len(all_rows)-len(pending)
                    st.info(f"{len(pending)} pending · {done} already entered")

                    if not pending:
                        st.success("All feeders in this section have a reading for this month.")
                        continue

                    # Exactly 10 feeders per page.
                    total_pages=(len(pending)+9)//10
                    key=f"entry_page_{sec}_{year}_{month}"
                    p=st.number_input("Page",1,total_pages,1,1,key=key)
                    chunk=pending[(p-1)*10:p*10]

                    for row in chunk:
                        c1,c2,c3,c4,c5=st.columns([2.8,1.8,1.2,2.0,1.1])
                        c1.write(f"**{row['feeder_name']}**")
                        c2.write(row["meter_no"])
                        c3.write(f"MF: {row['mf']}")
                        reading=c4.number_input(
                            "Reading (kWh)",min_value=0.0,value=0.0,format="%.3f",
                            key=f"new_{sec}_{year}_{month}_{row['id']}"
                        )
                        if c5.button("Save",key=f"save_{sec}_{year}_{month}_{row['id']}",type="primary"):
                            save_reading(con,row["id"],year,month,reading)
                            st.rerun()

                    st.caption(f"Showing feeders {(p-1)*10+1}–{min(p*10,len(pending))} of {len(pending)} pending.")



    else:
        rows=get_scope_rows(con,year,month,"Division",division_name)
        pending=[r for r in rows if r["reading_id"] is None or _is_mu_template_pending(r)]
        done=len(rows)-len(pending)
        st.info(f"{len(pending)} pending · {done} already entered")
        if not pending:
            st.success("All mapped meters in this division have a reading for this month.")
        else:
            total_pages=(len(pending)+9)//10
            pno=st.number_input(
                "Page",1,total_pages,1,1,
                key=f"division_entry_page_{division_name}_{year}_{month}"
            )
            chunk=pending[(pno-1)*10:pno*10]
            for row in chunk:
                c1,c2,c3,c4,c5=st.columns([2.8,1.8,1.2,2.0,1.1])
                c1.write(f"**{row['feeder_name']}**")
                c2.write(row["meter_no"])
                c3.write(f"MF: {row['mf']}")
                reading=c4.number_input(
                    "Reading (kWh)",min_value=0.0,value=0.0,format="%.3f",
                    key=f"dnew_{division_name}_{year}_{month}_{row['id']}"
                )
                if c5.button("Save",key=f"dsave_{division_name}_{year}_{month}_{row['id']}",type="primary"):
                    save_reading(con,row["id"],year,month,reading)
                    st.rerun()
            st.caption(f"Showing feeders {(pno-1)*10+1}–{min(pno*10,len(pending))} of {len(pending)} pending.")

# ---------- DIRECT MU ENTRY ----------
elif page=="Direct MU Entry":
    scope_label = CIRCLE_NAME if scope=="Circle" else division_name
    st.subheader(f"Direct MU Entry — {scope_label} — {calendar.month_name[month]} {year}")
    st.caption("Search for a specific feeder or meter. Use this only when meter data is lost or the meter is defective.")

    search_text = st.text_input(
        "Search feeder / meter number",
        placeholder="Type feeder name or meter number…",
        key=f"direct_mu_search_{scope}_{division_name}_{year}_{month}"
    ).strip()

    if not search_text:
        st.info("Enter a feeder name or meter number above to make a direct MU entry.")
    else:
        like=f"%{search_text}%"
        if scope=="Circle":
            matches=con.execute("""
                SELECT f.*, r.id reading_id, r.reading_kwh, r.direct_mu, r.direct_mu_note
                FROM feeder_master f LEFT JOIN monthly_readings r
                  ON r.feeder_id=f.id AND r.year=? AND r.month=?
                WHERE f.active=1 AND (f.feeder_name LIKE ? OR f.meter_no LIKE ?)
                ORDER BY f.feeder_name,f.meter_no
            """,(year,month,like,like)).fetchall()
        else:
            div_id = get_division_id(con, division_name)
            if div_id is None:
                matches = []
            else:
                matches = con.execute("""
                    SELECT DISTINCT f.*, r.id reading_id, r.reading_kwh, r.direct_mu, r.direct_mu_note
                    FROM feeder_master f
                    LEFT JOIN division_row_map dm ON dm.feeder_id=f.id AND dm.division_id=? AND dm.active=1
                    LEFT JOIN monthly_readings r
                      ON r.feeder_id=f.id AND r.year=? AND r.month=?
                    WHERE f.active=1 AND (dm.feeder_id IS NOT NULL OR f.division_name=?)
                      AND (f.feeder_name LIKE ? OR f.meter_no LIKE ?)
                    ORDER BY f.feeder_name,f.meter_no
                """,(div_id,year,month,division_name,like,like)).fetchall()

        if not matches:
            st.warning("No feeder or meter found for the current scope.")
        else:
            options=[f"{r['feeder_name']}  |  {r['meter_no']}  |  {r['energy_direction']}" for r in matches]
            selected=st.selectbox("Select feeder",options,key=f"direct_mu_select_{scope}_{division_name}_{year}_{month}")
            row=matches[options.index(selected)]
            st.markdown(f"**{row['feeder_name']}**  ·  Meter: `{row['meter_no']}`  ·  Direction: **{row['energy_direction']}**")
            current=row['direct_mu'] if row['direct_mu'] is not None else 0.0
            mu=st.number_input("Direct MU",min_value=0.0,value=float(current),format="%.6f",key=f"direct_mu_value_{scope}_{division_name}_{year}_{month}_{row['id']}")
            note=st.text_input("Reason / remarks",value=row['direct_mu_note'] or "Meter unavailable/defective",key=f"direct_mu_note_{scope}_{division_name}_{year}_{month}_{row['id']}")
            if st.button("Save Direct MU",type="primary",key=f"direct_mu_save_{scope}_{division_name}_{year}_{month}_{row['id']}"):
                con.execute("""INSERT INTO monthly_readings(feeder_id,year,month,reading_kwh,remarks,direct_mu,direct_mu_note)
                               VALUES(?,?,?,?,?,?,?)
                               ON CONFLICT(feeder_id,year,month) DO UPDATE SET direct_mu=excluded.direct_mu,direct_mu_note=excluded.direct_mu_note,remarks=excluded.remarks,updated_at=CURRENT_TIMESTAMP""",
                            (row['id'],year,month,float(row['reading_kwh'] or 0),"Direct MU fallback",float(mu),note))
                con.commit(); st.success("Direct MU saved."); st.rerun()
            st.info("Enter MU as a positive magnitude. IMPORT contributes positively; EXPORT is subtracted automatically in division net injection.")

# ---------- ALL FEEDER READINGS ----------
elif page=="All Feeder Readings":
    scope_label = CIRCLE_NAME if scope=="Circle" else division_name
    st.subheader(f"All Feeder Readings — {scope_label} — {calendar.month_name[month]} {year}")
    st.caption("This page includes both entered and pending feeders. Use Edit to correct an existing monthly reading.")

    rows=get_scope_rows(con,year,month,scope,division_name)
    if not rows:
        st.info("No feeder master records are mapped to this view yet.")
    else:
        for row in rows:
            prev=get_previous_reading(con,row["id"],year,month)
            entered=row["reading_id"] is not None and not _is_mu_template_pending(row)
            with st.container(border=True):
                c1,c2,c3,c4,c5,c6=st.columns([2.4,1.6,1,1.5,1.5,1])
                c1.write(f"**{row['feeder_name']}**")
                c2.write(row["meter_no"])
                c3.write(TYPE_LABELS[row["entry_type"]])
                c4.write("—" if prev is None else f"Last: {prev:,.3f}")
                if entered:
                    value=c5.number_input("Present kWh",value=float(row["reading_kwh"]),
                                          format="%.3f",key=f"edit_{row['id']}_{year}_{month}")
                    if c6.button("Save",key=f"saveedit_{row['id']}_{year}_{month}"):
                        save_reading(con,row["id"],year,month,value,row["remarks"] or "")
                        st.rerun()
                else:
                    c5.write("**Not entered**")
                    c6.write("Pending")

# ---------- FEEDER MASTER ----------
elif page=="Feeder Master":
    st.subheader("Feeder Master")
    st.caption("Meter information is stored here permanently. Monthly entry pages only ask for the reading.")

    with st.expander("➕ Add New Feeder",expanded=True):
        with st.form("new_feeder"):
            c1,c2,c3=st.columns([2.5,1.8,1])
            feeder=c1.text_input("Feeder / injection point name")
            meter=c2.text_input("Meter number")
            mf=c3.number_input("MF",min_value=0.000001,value=1.0,step=1.0)
            c4,c5,c6=st.columns([1.5,1.5,1.5])
            division=c4.selectbox("Division",list(DIVISIONS.keys()),key="new_division")
            sub_options=[x for x in allowed_subdivisions(division) if x is not None]
            sub_labels=["None"]+sub_options
            subdivision_label=c5.selectbox("Subdivision",sub_labels,key="new_subdivision")
            typ_label=c6.selectbox("Selection",SELECTION_OPTIONS,key="new_selection")
            c7,c8,c9=st.columns([1.3,1.3,2.0])
            energy_direction=c7.selectbox("Energy Direction",["IMPORT","EXPORT"],index=0 if SELECTION_TO_ENTRY_TYPE[typ_label] != "C" else 1,key="new_energy_direction",help="For division reports: IMPORT is added to net injection; EXPORT is subtracted.")
            voltage=c8.number_input("Voltage (kV)",min_value=0.0,value=33.0,step=1.0)
            initial=c9.number_input("Initial reading (kWh)",min_value=0.0,value=0.0,format="%.3f",help="Baseline used only when no earlier monthly reading exists. When a historical Excel month is imported, its Last Reading is used automatically as this baseline.")
            submit=st.form_submit_button("Save new feeder",type="primary",width="stretch")
            if submit:
                if not feeder.strip() or not meter.strip():
                    st.error("Feeder name and meter number are required.")
                else:
                    try:
                        add_feeder(con,feeder,meter,mf,SELECTION_TO_ENTRY_TYPE[typ_label],initial,division,None if subdivision_label=="None" else subdivision_label,voltage,energy_direction,typ_label)
                        st.success("Feeder added.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("This meter number already exists for this selection.")

    st.divider()
    search_text=st.text_input(
        "Search feeder master",
        placeholder="Search by feeder name or meter number",
        key="feeder_master_search",
    )
    rows=get_feeders(con,active_only=True)
    if search_text.strip():
        needle=search_text.strip().casefold()
        rows=[
            row for row in rows
            if needle in str(row["feeder_name"] or "").casefold()
            or needle in str(row["meter_no"] or "").casefold()
        ]
        st.caption(f"{len(rows)} feeder(s) found.")
    else:
        st.caption(f"{len(rows)} active feeder(s).")

    for row in rows:
        with st.expander(f"{row['feeder_name']}  ·  {row['meter_no']}  ·  {row['energy_direction']}  ·  {selection_label_for_row(row)}  ·  {row['division_name'] or 'No division'} / {row['subdivision'] or 'No subdivision'}"):
            with st.form(f"editmaster_{row['id']}"):
                c1,c2,c3=st.columns([2.5,1.8,1])
                feeder=c1.text_input("Feeder name",value=row["feeder_name"],key=f"fn_{row['id']}")
                meter=c2.text_input("Meter number",value=row["meter_no"],key=f"mn_{row['id']}")
                mf=c3.number_input("MF",min_value=0.000001,value=float(row["mf"]),key=f"mf_{row['id']}")
                c4,c5,c6=st.columns([1.5,1.5,1.5])
                division_index=list(DIVISIONS.keys()).index(row["division_name"]) if row["division_name"] in DIVISIONS else 0
                division=c4.selectbox("Division",list(DIVISIONS.keys()),index=division_index,key=f"dv_{row['id']}")
                sub_options=[x for x in allowed_subdivisions(division) if x is not None]
                sub_labels=["None"]+sub_options
                current_sub=row["subdivision"] if row["subdivision"] in sub_options else "None"
                subdivision_label=c5.selectbox("Subdivision",sub_labels,index=sub_labels.index(current_sub),key=f"sd_{row['id']}")
                current_selection=selection_label_for_row(row)
                typ=c6.selectbox("Selection",SELECTION_OPTIONS,
                                 index=SELECTION_OPTIONS.index(current_selection),
                                 key=f"tp_{row['id']}")
                c7,c8=st.columns([1.3,1.3])
                energy_direction=c7.selectbox("Energy Direction",["IMPORT","EXPORT"],index=0 if row["energy_direction"]!="EXPORT" else 1,key=f"ed_{row['id']}",help="For division reports: IMPORT is added to net injection; EXPORT is subtracted.")
                voltage=c8.number_input("Voltage (kV)",min_value=0.0,value=float(row["voltage_kv"] or 33),step=1.0,key=f"vg_{row['id']}")
                c9=st.container()
                initial=c9.number_input("Initial reading (kWh)",min_value=0.0,
                                        value=float(row["initial_reading_kwh"]),
                                        key=f"ir_{row['id']}",format="%.3f",
                                        help="Baseline used only when no earlier monthly reading exists. Historical Excel imports can update this automatically from the workbook's Last Reading.")
                save,delete=st.columns(2)
                if save.form_submit_button("Save changes",type="primary",width="stretch"):
                    try:
                        update_feeder(con,row["id"],feeder,meter,mf,SELECTION_TO_ENTRY_TYPE[typ],initial,division,None if subdivision_label=="None" else subdivision_label,voltage,energy_direction,typ,year,month)
                        st.success("Feeder updated.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("This meter number already exists for that selection.")
                if delete.form_submit_button("Delete feeder",width="stretch"):
                    delete_feeder(con,row["id"])
                    st.success("Feeder deleted.")
                    st.rerun()

# ---------- MU TEMPLATE ----------
elif page=="MU Template":
    st.subheader(f"MU Template — {calendar.month_name[month]} {year}")
    st.caption(
        "Verify the current feeder template by classification. "
        "Feeder name and meter number are shown directly from Feeder Master. "
        "Use Search Feeder Master to add an existing feeder to the selected monthly report section."
    )

    template_sections = [
        ("A", "Import from GSS"),
        ("B", "Import from other circle"),
        ("C", "Export to other circle"),
        ("OA", "Open Access"),
    ]

    tabs = st.tabs([label for _, label in template_sections])
    for tab, (sec, label) in zip(tabs, template_sections):
        with tab:
            st.markdown(f"### {label}")

            if sec == "OA":
                classification_rows = con.execute(
                    """SELECT feeder_name, meter_no, mf, division_name, subdivision
                       FROM feeder_master
                       WHERE active=1 AND selection_type=?
                         AND meter_no NOT LIKE '__DIRECT_MU__%'
                       ORDER BY feeder_name, meter_no""",
                    ("Open Access",)
                ).fetchall()
                st.caption("Current Feeder Master classification — Open Access")
                if classification_rows:
                    st.dataframe(
                        [
                            {
                                "Feeder Name": r["feeder_name"],
                                "Meter No.": r["meter_no"],
                                "MF": r["mf"],
                                "Division": r["division_name"],
                                "Subdivision": r["subdivision"],
                            }
                            for r in classification_rows
                        ],
                        width="stretch",
                        hide_index=True,
                    )
                else:
                    st.info("No feeders are currently configured as Open Access in the template.")

                st.divider()
                st.markdown("**Search Feeder Master**")
                oa_search = st.text_input(
                    "Search feeder or meter number",
                    placeholder="Search by feeder name or meter number…",
                    key=f"mu_template_search_OA_{year}_{month}",
                ).strip()
                if oa_search:
                    needle = oa_search.casefold()
                    matches = [
                        r for r in get_feeders(con, active_only=True)
                        if needle in str(r["feeder_name"] or "").casefold()
                        or needle in str(r["meter_no"] or "").casefold()
                    ]
                    if not matches:
                        st.warning(
                            "No matching feeder was found in Feeder Master. "
                            "Add the feeder in Feeder Master first, then return here."
                        )
                        if st.button(
                            "Open Feeder Master",
                            key=f"mu_template_open_master_OA_{year}_{month}",
                            width="stretch",
                        ):
                            st.session_state["selected_page"] = "Feeder Master"
                            st.rerun()
                    else:
                        st.info(
                            "The feeder exists in Feeder Master. To classify it as Open Access, "
                            "set Selection = Open Access in Feeder Master."
                        )
                        st.dataframe(
                            [
                                {
                                    "Feeder Name": r["feeder_name"],
                                    "Meter No.": r["meter_no"],
                                    "MF": r["mf"],
                                    "Current Selection": selection_label_for_row(r),
                                }
                                for r in matches
                            ],
                            width="stretch",
                            hide_index=True,
                        )
                continue

            # A/B/C show the actual current monthly template. These rows are
            # seeded from the master template for a new month and remain unchanged
            # when the user edits the monthly template.
            classification_rows = con.execute(
                """SELECT r.report_feeder_name AS feeder_name,
                          r.report_meter_no AS meter_no,
                          r.report_mf AS mf,
                          f.division_name, f.subdivision
                   FROM monthly_readings r
                   JOIN feeder_master f ON f.id=r.feeder_id
                   WHERE r.year=? AND r.month=? AND r.report_section=?
                     AND r.report_meter_no NOT LIKE '__DIRECT_MU__%'
                   ORDER BY r.report_order, r.id""",
                (year, month, sec)
            ).fetchall()

            st.caption("Current monthly MU Template")
            if classification_rows:
                st.dataframe(
                    [
                        {
                            "Feeder Name": r["feeder_name"],
                            "Meter No.": r["meter_no"],
                            "MF": r["mf"],
                            "Division": r["division_name"],
                            "Subdivision": r["subdivision"],
                        }
                        for r in classification_rows
                    ],
                    width="stretch",
                    hide_index=True,
                )
            else:
                st.info(f"No feeders are currently configured in the {label} template section.")

            st.divider()
            st.markdown("**Add feeder to this month's report section**")
            st.caption(
                "Search the Feeder Master database. The feeder must already exist in Feeder Master."
            )

            search = st.text_input(
                "Search feeder or meter number",
                placeholder="Search by feeder name or meter number…",
                key=f"mu_template_search_{sec}_{year}_{month}",
            ).strip()

            if search:
                needle = search.casefold()
                matches = [
                    r for r in get_feeders(con, active_only=True)
                    if needle in str(r["feeder_name"] or "").casefold()
                    or needle in str(r["meter_no"] or "").casefold()
                ]

                if not matches:
                    st.warning(
                        "No matching feeder was found in Feeder Master. "
                        "Add the feeder in Feeder Master first, then return here."
                    )
                    if st.button(
                        "Open Feeder Master",
                        key=f"mu_template_open_master_{sec}_{year}_{month}",
                        width="stretch",
                    ):
                        st.session_state["selected_page"] = "Feeder Master"
                        st.rerun()
                else:
                    options = [
                        f"{r['feeder_name']}  |  {r['meter_no']}  |  MF: {r['mf']}"
                        for r in matches
                    ]
                    selected = st.selectbox(
                        "Select feeder",
                        options,
                        key=f"mu_template_select_{sec}_{year}_{month}",
                    )
                    selected_row = matches[options.index(selected)]

                    already = con.execute(
                        """SELECT report_section FROM monthly_readings
                           WHERE feeder_id=? AND year=? AND month=? LIMIT 1""",
                        (selected_row["id"], year, month)
                    ).fetchone()

                    if already is not None and already["report_section"] == sec:
                        st.info("This feeder is already present in this section for the selected month.")
                    elif already is not None and already["report_section"] in ("A", "B", "C"):
                        old_sec = already["report_section"]
                        st.warning(
                            f"This feeder is currently in Section {old_sec} for this month. "
                            "Adding it here will move its month-specific report row to this section."
                        )

                    if st.button(
                        f"Add to {label}",
                        key=f"mu_template_add_{sec}_{year}_{month}_{selected_row['id']}",
                        type="primary",
                        width="stretch",
                    ):
                        ok, message = add_feeder_to_mu_template(
                            con, selected_row["id"], year, month, sec
                        )
                        if ok:
                            st.success(f"{message} added to {label}.")
                            st.rerun()
                        else:
                            st.error(message)

            report_rows = con.execute(
                """SELECT r.report_sl_no, r.report_feeder_name, r.report_meter_no,
                          r.report_mf, r.last_reading_kwh, r.reading_kwh,
                          r.direct_mu, r.remarks, r.report_order
                   FROM monthly_readings r
                   WHERE r.year=? AND r.month=? AND r.report_section=?
                   ORDER BY r.report_order, r.id""",
                (year, month, sec)
            ).fetchall()

            st.divider()
            st.markdown("**Current monthly Excel section preview**")
            if not report_rows:
                st.info("No feeders are currently configured in this monthly template section.")
            else:
                preview = []
                for rr in report_rows:
                    last = rr["last_reading_kwh"]
                    present = rr["reading_kwh"]
                    direct_mu = rr["direct_mu"]
                    mu = direct_mu
                    if mu is None and last is not None and present is not None:
                        mu = (float(present) - float(last)) * float(rr["report_mf"] or 1) / 1000.0
                    preview.append({
                        "SL No.": rr["report_sl_no"],
                        "Feeder Name": rr["report_feeder_name"],
                        "Meter No.": rr["report_meter_no"],
                        "Last Reading": last,
                        "Present Reading": present if present not in (None, 0) else None,
                        "Difference": None if last is None or present in (None, 0) else float(present) - float(last),
                        "MF": rr["report_mf"],
                        "MU": mu,
                        "Remarks": rr["remarks"],
                    })
                st.dataframe(preview, width="stretch", hide_index=True)

# ---------- EXCEL ----------
elif page=="Generate Excel":
    scope_label = CIRCLE_NAME if scope=="Circle" else division_name
    st.subheader(f"Generate Excel — {scope_label} — {calendar.month_name[month]} {year}")
    rows=get_scope_rows(con,year,month,scope,division_name)
    entered=[r for r in rows if r["reading_id"] is not None]
    pending=[r for r in rows if r["reading_id"] is None]

    c1,c2,c3=st.columns(3)
    c1.metric("Total feeders",len(rows))
    c2.metric("Readings entered",len(entered))
    c3.metric("Pending",len(pending))

    if pending:
        st.warning("Some mapped feeders do not have a reading for this month. The generated workbook will show those readings as blank.")
    if st.button("Build exact Excel workbook",type="primary",width="stretch"):
        if scope=="Circle":
            data=build_report(con,year,month)
            filename=f"MU Inj. {calendar.month_name[month]}, {year}.xlsx"
        else:
            data=division_report(con,year,month,division_name)
            filename=f"MU Inj ({division_name}) {calendar.month_name[month]}, {year}.xlsx"
        st.download_button(
            "⬇ Download Excel",data=data,file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch"
        )
        st.success("Report generated from the supplied Excel template.")

# ---------- IMPORT EXCEL ----------
elif page=="Import Excel":
    scope_label = CIRCLE_NAME if scope=="Circle" else division_name
    st.subheader(f"Import Monthly Excel — {scope_label}")
    if scope=="Circle":
        st.caption("Upload the monthly Jorhat Circle workbook. Meter masters and monthly readings are stored in the shared database.")
    else:
        st.caption("Upload the all-division workbook once. All available Jorhat-1, Jorhat-2 and Teok sheets are processed; the selected division only controls the current view.")

    c1,c2=st.columns(2)
    import_year=c1.number_input("Excel month — Year",2000,2100,year,1,key="import_year")
    import_month=c2.selectbox(
        "Excel month — Month",range(1,13),index=month-1,
        format_func=lambda m:calendar.month_name[m],key="import_month"
    )
    uploaded=st.file_uploader(
        "Upload monthly Excel file",type=["xlsx","xlsm"],key="monthly_excel_upload"
    )
    overwrite=st.checkbox(
        "Overwrite readings already stored for this month",value=False,
        help="Leave unchecked to protect existing database readings."
    )

    if scope=="Circle":
        st.info("Importing a circle workbook makes the meter readings immediately available in the mapped divisions. You do not need to upload the same meter data again in a division.")
    else:
        st.info("All available division sheets are mapped to the shared meter masters. If a meter reading was already imported at circle level, the existing monthly reading is retained.")

    if uploaded is not None:
        st.write(f"**Selected:** {uploaded.name}")
        if st.button("Import this month into database",type="primary",width="stretch"):
            try:
                if scope=="Circle":
                    result=import_month_excel(
                        con,uploaded.getvalue(),int(import_year),int(import_month),overwrite
                    )
                    st.success(                        f"Imported {result['imported']} monthly readings. "
                        f"Added {result['new_master']} new feeder records. "
                        f"Updated {result['updated_master']} existing feeder records. "
                        f"Overwritten {result['overwritten']} existing monthly readings. "
                        f"Imported {result.get('direct_mu_imported', 0)} Direct MU exceptions."
                    )
                else:
                    result=import_all_division_workbook(
                        con,uploaded.getvalue(),int(import_year),int(import_month),overwrite
                    )
                    st.success(
                        f"Processed {result['divisions']} division sheets. "
                        f"Mapped {result['mapped']} division rows. "
                        f"Imported {result['imported']} new monthly readings. "
                        f"Reused/protected {result['reused']} existing readings. "
                        f"Added {result['new_master']} new meter masters."
                    )
                if result["skipped"]:
                    st.warning(f"{len(result['skipped'])} rows were skipped.")
                    with st.expander("View skipped rows"):
                        for item in result["skipped"]:
                            st.write(f"• {item}")
            except Exception as exc:
                st.error(f"Import failed: {exc}")

    st.divider()
    st.markdown("**Shared-data workflow**")
    st.markdown(
        "Import a Jorhat Circle workbook once. Then switch **Level → Division** "
        "and select Jorhat-1, Jorhat-2 or Teok. Mapped meters use the same stored "
        "monthly reading, so a second upload is not required."
    )

# ---------- DASHBOARD ----------
else:
    scope_label = CIRCLE_NAME if scope=="Circle" else division_name
    st.subheader(f"Dashboard — {scope_label} — {calendar.month_name[month]} {year}")

    if scope=="Circle":
        # Use the exact monthly rows represented by the generated Circle Excel
        # whenever report metadata is available for this month.
        rows=get_report_reading_rows(con,year,month)
        if not rows:
            rows=get_scope_rows(con,year,month,"Circle",None)
        totals={"A":0.0,"B":0.0,"C":0.0}
        counts={"A":0,"B":0,"C":0}
        for r in rows:
            mu=calculate_mu(con,r,year,month)
            if mu is not None:
                report_section = r["report_section"] if "report_section" in r.keys() else None
                section = report_section if report_section in ("A", "B", "C") else r["entry_type"]
                counts[section]+=1
                totals[section]+=mu
        net=totals["A"]+totals["B"]-totals["C"]

        cols=st.columns(4)
        cols[0].metric("A — GSS",f"{totals['A']:,.6f} MU")
        cols[1].metric("B — Other circle",f"{totals['B']:,.6f} MU")
        cols[2].metric("C — Export",f"{totals['C']:,.6f} MU")
        cols[3].metric("Net injection",f"{net:,.6f} MU")

        st.dataframe({
            "Section":["A — Import from GSS","B — Import from other circle","C — Export to other circle","NET"],
            "Readings entered":[counts["A"],counts["B"],counts["C"],sum(counts.values())],
            "Energy (MU)":[totals["A"],totals["B"],totals["C"],net]
        },width="stretch",hide_index=True)

        st.divider()
        st.subheader("Energy Trend — Last 5 Months")
        trend=[]
        y,m=year,month
        for _ in range(5):
            t=month_energy_summary(con,y,m)
            trend.append({
                "Month":f"{calendar.month_abbr[m]} {y}",
                "Date":datetime(y,m,1),
                "A — GSS":t["A"],
                "B — Other circle":t["B"],
                "C — Export":t["C"],
                "Net injection":t["NET"],
            })
            y,m=previous_period(y,m)
        trend.reverse()
        trend_df=pd.DataFrame(trend).sort_values("Date").reset_index(drop=True)
        chart_long=trend_df.melt(
            id_vars=["Month","Date"],
            value_vars=["A — GSS","B — Other circle","C — Export","Net injection"],
            var_name="Series",
            value_name="MU",
        )
        month_order=trend_df["Month"].tolist()
        chart=alt.Chart(chart_long).mark_line(point=True).encode(
            x=alt.X(
                "Month:N",
                sort=month_order,
                axis=alt.Axis(title=None, labelAngle=0, labelOverlap=False),
            ),
            y=alt.Y("MU:Q", title="MU"),
            color=alt.Color("Series:N", title=None),
            tooltip=[
                alt.Tooltip("Month:N", title="Month"),
                alt.Tooltip("Series:N", title="Series"),
                alt.Tooltip("MU:Q", title="MU", format=".4f"),
            ],
        ).properties(height=360)
        st.altair_chart(chart,width="stretch")
        st.dataframe(trend_df.drop(columns=["Date"]),width="stretch",hide_index=True)

    else:
        totals=division_energy_summary(con,year,month,division_name)
        rows=get_scope_rows(con,year,month,"Division",division_name)
        entered=sum(1 for r in rows if r["reading_id"] is not None)

        cols=st.columns(4)
        cols[0].metric("Import",f"{totals['IMPORT']:,.3f} MWh")
        cols[1].metric("Export",f"{totals['EXPORT']:,.3f} MWh")
        cols[2].metric("Net injection",f"{totals['NET']:,.3f} MWh")
        cols[3].metric("Meters entered",entered)

        st.dataframe({
            "Flow":["Import","Export","NET"],
            "Energy (MWh)":[totals["IMPORT"],totals["EXPORT"],totals["NET"]]
        },width="stretch",hide_index=True)

        st.divider()
        st.subheader("Energy Trend — Last 5 Months")
        trend=[]
        y,m=year,month
        for _ in range(5):
            t=division_energy_summary(con,y,m,division_name)
            trend.append({
                "Month":f"{calendar.month_abbr[m]} {y}",
                "Date":datetime(y,m,1),
                "Import (MWh)":t["IMPORT"],
                "Export (MWh)":t["EXPORT"],
                "Net injection (MWh)":t["NET"],
            })
            y,m=previous_period(y,m)
        trend.reverse()
        trend_df=pd.DataFrame(trend).sort_values("Date").reset_index(drop=True)
        chart_long=trend_df.melt(
            id_vars=["Month","Date"],
            value_vars=["Import (MWh)","Export (MWh)","Net injection (MWh)"],
            var_name="Series",
            value_name="MWh",
        )
        month_order=trend_df["Month"].tolist()
        chart=alt.Chart(chart_long).mark_line(point=True).encode(
            x=alt.X(
                "Month:N",
                sort=month_order,
                axis=alt.Axis(title=None, labelAngle=0, labelOverlap=False),
            ),
            y=alt.Y("MWh:Q", title="MWh"),
            color=alt.Color("Series:N", title=None),
            tooltip=[
                alt.Tooltip("Month:N", title="Month"),
                alt.Tooltip("Series:N", title="Series"),
                alt.Tooltip("MWh:Q", title="MWh", format=".4f"),
            ],
        ).properties(height=360)
        st.altair_chart(chart,width="stretch")
        st.dataframe(trend_df.drop(columns=["Date"]),width="stretch",hide_index=True)