"""SQLite database setup and connection management."""
import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "accounting.db"


def get_connection(db_path: str | Path = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(db_path: str | Path = DB_PATH) -> None:
    with get_connection(db_path) as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS accounts (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                code        TEXT    NOT NULL UNIQUE,
                name        TEXT    NOT NULL,
                type        TEXT    NOT NULL CHECK(type IN ('Asset','Liability','Equity','Revenue','Expense')),
                normal_side TEXT    NOT NULL CHECK(normal_side IN ('Debit','Credit')),
                description TEXT    DEFAULT '',
                is_active   INTEGER NOT NULL DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS journal_entries (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_date  TEXT    NOT NULL,
                reference   TEXT    NOT NULL DEFAULT '',
                description TEXT    NOT NULL DEFAULT '',
                created_at  TEXT    NOT NULL DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS journal_lines (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_id    INTEGER NOT NULL REFERENCES journal_entries(id) ON DELETE CASCADE,
                account_id  INTEGER NOT NULL REFERENCES accounts(id),
                debit       REAL    NOT NULL DEFAULT 0,
                credit      REAL    NOT NULL DEFAULT 0,
                memo        TEXT    NOT NULL DEFAULT '',
                CHECK(debit >= 0 AND credit >= 0),
                CHECK(NOT (debit > 0 AND credit > 0))
            );

            CREATE VIEW IF NOT EXISTS general_ledger AS
                SELECT
                    je.entry_date,
                    je.reference,
                    je.description  AS entry_description,
                    a.code          AS account_code,
                    a.name          AS account_name,
                    a.type          AS account_type,
                    a.normal_side,
                    jl.debit,
                    jl.credit,
                    jl.memo,
                    je.id           AS entry_id,
                    jl.id           AS line_id
                FROM journal_lines jl
                JOIN journal_entries je ON jl.entry_id = je.id
                JOIN accounts a         ON jl.account_id = a.id
                ORDER BY je.entry_date, je.id, jl.id;
        """)
