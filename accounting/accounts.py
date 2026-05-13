"""Chart of Accounts — create, list, and look up accounts."""
from __future__ import annotations
from dataclasses import dataclass
from typing import Optional
from .database import get_connection, DB_PATH

# normal debit/credit side per account type (double-entry convention)
_NORMAL_SIDE = {
    "Asset":     "Debit",
    "Expense":   "Debit",
    "Liability": "Credit",
    "Equity":    "Credit",
    "Revenue":   "Credit",
}


@dataclass
class Account:
    id: int
    code: str
    name: str
    type: str
    normal_side: str
    description: str
    is_active: bool


def create_account(code: str, name: str, account_type: str,
                   description: str = "", db_path=DB_PATH) -> Account:
    if account_type not in _NORMAL_SIDE:
        raise ValueError(f"Unknown account type: {account_type}")
    normal_side = _NORMAL_SIDE[account_type]
    with get_connection(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO accounts (code, name, type, normal_side, description) "
            "VALUES (?, ?, ?, ?, ?)",
            (code, name, account_type, normal_side, description),
        )
        account_id = cur.lastrowid
    # Build from known values — avoids a second round-trip and reads after commit
    return Account(
        id=account_id, code=code, name=name, type=account_type,
        normal_side=normal_side, description=description, is_active=True,
    )


def get_account_by_id(account_id: int, db_path=DB_PATH) -> Account:
    with get_connection(db_path) as conn:
        row = conn.execute("SELECT * FROM accounts WHERE id = ?", (account_id,)).fetchone()
    if not row:
        raise ValueError(f"Account id={account_id} not found")
    return _row_to_account(row)


def get_account_by_code(code: str, db_path=DB_PATH) -> Account:
    with get_connection(db_path) as conn:
        row = conn.execute("SELECT * FROM accounts WHERE code = ?", (code,)).fetchone()
    if not row:
        raise ValueError(f"Account code '{code}' not found")
    return _row_to_account(row)


def list_accounts(account_type: Optional[str] = None, db_path=DB_PATH) -> list[Account]:
    sql = "SELECT * FROM accounts"
    params: tuple = ()
    if account_type:
        sql += " WHERE type = ?"
        params = (account_type,)
    sql += " ORDER BY code"
    with get_connection(db_path) as conn:
        rows = conn.execute(sql, params).fetchall()
    return [_row_to_account(r) for r in rows]


def seed_chart_of_accounts(db_path=DB_PATH) -> None:
    """Populate a standard small-business chart of accounts."""
    accounts = [
        # ── Assets ─────────────────────────────────────────────────────────
        ("1000", "Cash",                        "Asset",     "Primary cash account"),
        ("1010", "Petty Cash",                  "Asset",     "Small cash on hand"),
        ("1100", "Accounts Receivable",         "Asset",     "Amounts owed by customers"),
        ("1200", "Inventory",                   "Asset",     "Goods held for sale"),
        ("1300", "Prepaid Expenses",            "Asset",     "Expenses paid in advance"),
        ("1500", "Equipment",                   "Asset",     "Machinery and equipment"),
        ("1510", "Accumulated Depreciation",    "Asset",     "Contra-asset: equipment"),
        # ── Liabilities ─────────────────────────────────────────────────────
        ("2000", "Accounts Payable",            "Liability", "Amounts owed to vendors"),
        ("2100", "Accrued Liabilities",         "Liability", "Expenses incurred not yet paid"),
        ("2200", "Taxes Payable",               "Liability", "Income and sales tax owed"),
        ("2300", "Notes Payable",               "Liability", "Short-term loans"),
        ("2500", "Long-Term Debt",              "Liability", "Long-term borrowings"),
        # ── Equity ───────────────────────────────────────────────────────────
        ("3000", "Owner's Capital",             "Equity",    "Owner contributions"),
        ("3100", "Retained Earnings",           "Equity",    "Cumulative net income retained"),
        ("3200", "Owner's Draws",               "Equity",    "Owner withdrawals"),
        # ── Revenue ──────────────────────────────────────────────────────────
        ("4000", "Sales Revenue",               "Revenue",   "Revenue from product sales"),
        ("4100", "Service Revenue",             "Revenue",   "Revenue from services"),
        ("4200", "Interest Income",             "Revenue",   "Interest earned"),
        ("4900", "Other Income",                "Revenue",   "Miscellaneous income"),
        # ── Expenses ─────────────────────────────────────────────────────────
        ("5000", "Cost of Goods Sold",          "Expense",   "Direct cost of products sold"),
        ("5100", "Salaries & Wages",            "Expense",   "Employee compensation"),
        ("5200", "Rent Expense",                "Expense",   "Office / warehouse rent"),
        ("5300", "Utilities Expense",           "Expense",   "Electricity, water, internet"),
        ("5400", "Marketing & Advertising",     "Expense",   "Promotional costs"),
        ("5500", "Depreciation Expense",        "Expense",   "Asset depreciation"),
        ("5600", "Interest Expense",            "Expense",   "Interest on loans"),
        ("5700", "Office Supplies",             "Expense",   "Consumable office items"),
        ("5800", "Insurance Expense",           "Expense",   "Business insurance"),
        ("5900", "Other Expenses",              "Expense",   "Miscellaneous expenses"),
    ]
    for code, name, acct_type, desc in accounts:
        try:
            create_account(code, name, acct_type, desc, db_path)
        except Exception:
            pass  # already exists


def _row_to_account(row) -> Account:
    return Account(
        id=row["id"],
        code=row["code"],
        name=row["name"],
        type=row["type"],
        normal_side=row["normal_side"],
        description=row["description"],
        is_active=bool(row["is_active"]),
    )
