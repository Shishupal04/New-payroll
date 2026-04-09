import sqlite3
import os


def _find_writable_db_path():
    candidates = [
        "/opt/render/project/src/data/payroll.db",
        "/tmp/payroll_data/payroll.db",
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "payroll.db"),
    ]
    for path in candidates:
        try:
            folder = os.path.dirname(path)
            os.makedirs(folder, exist_ok=True)
            conn = sqlite3.connect(path)
            conn.execute("PRAGMA journal_mode=WAL")
            conn.close()
            print(f"[DB] Using path: {path}")
            return path
        except Exception as e:
            print(f"[DB] Skipping {path}: {e}")
            continue
    raise RuntimeError("No writable location found for the database!")


DB_PATH = _find_writable_db_path()


def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    conn   = get_db()
    cursor = conn.cursor()

    cursor.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            role     TEXT NOT NULL,
            email    TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS companies (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            name       TEXT NOT NULL,
            address    TEXT,
            city       TEXT,
            state      TEXT,
            pincode    TEXT,
            phone      TEXT,
            email      TEXT,
            pf_number  TEXT,
            esi_number TEXT,
            pan        TEXT,
            tan        TEXT
        );

        CREATE TABLE IF NOT EXISTS employees (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id      INTEGER DEFAULT 1,
            emp_code        TEXT,
            name            TEXT NOT NULL,
            department      TEXT,
            designation     TEXT,
            location        TEXT,
            doj             TEXT,
            dob             TEXT,
            gender          TEXT DEFAULT 'Male',
            basic           REAL DEFAULT 0,
            hra             REAL DEFAULT 0,
            taxable_allow   REAL DEFAULT 0,
            night_shift_allow REAL DEFAULT 0,
            pf_relief       REAL DEFAULT 0,
            bank_account    TEXT,
            ifsc            TEXT,
            pan             TEXT,
            uan             TEXT,
            pf_number       TEXT,
            esi_number      TEXT,
            eps_number      TEXT,
            status          TEXT DEFAULT 'Active'
        );

        CREATE TABLE IF NOT EXISTS salary_structure (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id          INTEGER NOT NULL UNIQUE,
            basic           REAL DEFAULT 0,
            hra             REAL DEFAULT 0,
            taxable_allow   REAL DEFAULT 0,
            night_shift_allow REAL DEFAULT 0,
            pf_relief       REAL DEFAULT 0,
            effective_date  TEXT,
            FOREIGN KEY (emp_id) REFERENCES employees(id)
        );

        CREATE TABLE IF NOT EXISTS attendance (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id       INTEGER NOT NULL,
            month        TEXT NOT NULL,
            days_in_month INTEGER DEFAULT 26,
            arrear_days  INTEGER DEFAULT 0,
            lopr_days    INTEGER DEFAULT 0,
            lop_days     INTEGER DEFAULT 0,
            present_days INTEGER DEFAULT 0,
            UNIQUE(emp_id, month)
        );

        CREATE TABLE IF NOT EXISTS payroll (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id          INTEGER NOT NULL,
            month           TEXT NOT NULL,
            basic           REAL DEFAULT 0,
            hra             REAL DEFAULT 0,
            taxable_allow   REAL DEFAULT 0,
            night_shift_allow REAL DEFAULT 0,
            ot_allowance    REAL DEFAULT 0,
            pf_relief       REAL DEFAULT 0,
            gross           REAL DEFAULT 0,
            pf              REAL DEFAULT 0,
            esi             REAL DEFAULT 0,
            pt              REAL DEFAULT 0,
            tds             REAL DEFAULT 0,
            lop             REAL DEFAULT 0,
            other_deduction REAL DEFAULT 0,
            net             REAL DEFAULT 0,
            days_in_month   INTEGER DEFAULT 26,
            arrear_days     INTEGER DEFAULT 0,
            lopr_days       INTEGER DEFAULT 0,
            lop_days        INTEGER DEFAULT 0,
            net_days_worked REAL DEFAULT 0,
            UNIQUE(emp_id, month)
        );

        CREATE TABLE IF NOT EXISTS salary_history (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id         INTEGER NOT NULL,
            effective_date TEXT NOT NULL,
            basic          REAL DEFAULT 0,
            hra            REAL DEFAULT 0,
            taxable_allow  REAL DEFAULT 0,
            night_shift_allow REAL DEFAULT 0,
            pf_relief      REAL DEFAULT 0,
            remarks        TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS overtime (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id   INTEGER NOT NULL,
            ot_date  TEXT NOT NULL,
            hours    REAL DEFAULT 0,
            rate     REAL DEFAULT 0,
            amount   REAL DEFAULT 0,
            month    TEXT
        );

        CREATE TABLE IF NOT EXISTS leaves (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_id     INTEGER NOT NULL,
            from_date  TEXT NOT NULL,
            to_date    TEXT NOT NULL,
            leave_type TEXT,
            days       REAL DEFAULT 1,
            reason     TEXT DEFAULT '',
            status     TEXT DEFAULT 'Approved'
        );
    """)

    # --- Migrate: add new columns if they don't exist ---
    _safe_add_column(cursor, "employees", "location", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "employees", "dob", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "employees", "gender", "TEXT DEFAULT 'Male'")
    _safe_add_column(cursor, "employees", "taxable_allow", "REAL DEFAULT 0")
    _safe_add_column(cursor, "employees", "night_shift_allow", "REAL DEFAULT 0")
    _safe_add_column(cursor, "employees", "pf_relief", "REAL DEFAULT 0")
    _safe_add_column(cursor, "employees", "pf_number", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "employees", "esi_number", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "employees", "eps_number", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "employees", "status", "TEXT DEFAULT 'Active'")
    _safe_add_column(cursor, "companies", "city", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "companies", "state", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "companies", "pincode", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "companies", "phone", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "companies", "email", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "companies", "pan", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "companies", "tan", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "users", "email", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "leaves", "to_date", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "leaves", "days", "REAL DEFAULT 1")
    _safe_add_column(cursor, "leaves", "reason", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "leaves", "status", "TEXT DEFAULT 'Approved'")
    _safe_add_column(cursor, "overtime", "rate", "REAL DEFAULT 0")
    _safe_add_column(cursor, "overtime", "amount", "REAL DEFAULT 0")
    _safe_add_column(cursor, "overtime", "month", "TEXT DEFAULT ''")
    _safe_add_column(cursor, "attendance", "days_in_month", "INTEGER DEFAULT 26")
    _safe_add_column(cursor, "attendance", "arrear_days", "INTEGER DEFAULT 0")
    _safe_add_column(cursor, "attendance", "lopr_days", "INTEGER DEFAULT 0")

    # Seed users
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        cursor.executemany(
            "INSERT INTO users (username, password, role) VALUES (?,?,?)",
            [
                ("admin",      "admin123", "admin"),
                ("hr",         "hr123",    "hr"),
                ("accountant", "acc123",   "accountant"),
            ]
        )

    # Seed company
    cursor.execute("SELECT COUNT(*) FROM companies")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            INSERT INTO companies (name, address, city, state, pf_number, esi_number)
            VALUES ('My Company Pvt Ltd', '123 Business Park', 'Hyderabad', 'Telangana', 'PF001', 'ESI001')
        """)

    conn.commit()
    conn.close()
    print(f"[DB] Initialised at {DB_PATH}")


def _safe_add_column(cursor, table, column, definition):
    try:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
    except Exception:
        pass  # Column already exists
