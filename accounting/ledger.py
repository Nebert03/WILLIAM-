"""Journal entry creation and general-ledger queries."""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional
from .database import get_connection, DB_PATH
from .accounts import get_account_by_code


@dataclass
class JournalLine:
    account_code: str
    debit: float = 0.0
    credit: float = 0.0
    memo: str = ""


@dataclass
class JournalEntry:
    id: int
    entry_date: str
    reference: str
    description: str
    lines: list[dict] = field(default_factory=list)


def post_journal_entry(
    entry_date: str,
    lines: list[JournalLine],
    description: str = "",
    reference: str = "",
    db_path=DB_PATH,
) -> JournalEntry:
    """Post a balanced journal entry. Raises ValueError if debits ≠ credits."""
    total_debit  = round(sum(l.debit  for l in lines), 2)
    total_credit = round(sum(l.credit for l in lines), 2)
    if total_debit != total_credit:
        raise ValueError(
            f"Journal entry does not balance: debits={total_debit}, credits={total_credit}"
        )
    if total_debit == 0:
        raise ValueError("Journal entry has no amounts")

    with get_connection(db_path) as conn:
        cur = conn.execute(
            "INSERT INTO journal_entries (entry_date, reference, description) VALUES (?,?,?)",
            (entry_date, reference, description),
        )
        entry_id = cur.lastrowid
        for line in lines:
            acct = get_account_by_code(line.account_code, db_path)
            conn.execute(
                "INSERT INTO journal_lines (entry_id, account_id, debit, credit, memo) "
                "VALUES (?,?,?,?,?)",
                (entry_id, acct.id, line.debit, line.credit, line.memo),
            )
    return get_journal_entry(entry_id, db_path)


def get_journal_entry(entry_id: int, db_path=DB_PATH) -> JournalEntry:
    with get_connection(db_path) as conn:
        row = conn.execute(
            "SELECT * FROM journal_entries WHERE id = ?", (entry_id,)
        ).fetchone()
        if not row:
            raise ValueError(f"Journal entry id={entry_id} not found")
        lines = conn.execute(
            """SELECT jl.debit, jl.credit, jl.memo, a.code, a.name
               FROM journal_lines jl JOIN accounts a ON jl.account_id = a.id
               WHERE jl.entry_id = ? ORDER BY jl.id""",
            (entry_id,),
        ).fetchall()
    return JournalEntry(
        id=row["id"],
        entry_date=row["entry_date"],
        reference=row["reference"],
        description=row["description"],
        lines=[dict(r) for r in lines],
    )


def list_journal_entries(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    db_path=DB_PATH,
) -> list[JournalEntry]:
    sql = "SELECT id FROM journal_entries WHERE 1=1"
    params = []
    if start_date:
        sql += " AND entry_date >= ?"
        params.append(start_date)
    if end_date:
        sql += " AND entry_date <= ?"
        params.append(end_date)
    sql += " ORDER BY entry_date, id"
    with get_connection(db_path) as conn:
        ids = [r["id"] for r in conn.execute(sql, params).fetchall()]
    return [get_journal_entry(i, db_path) for i in ids]


def account_balance(
    account_code: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    db_path=DB_PATH,
) -> float:
    """Return the net balance of an account (positive = normal side)."""
    acct = get_account_by_code(account_code, db_path)
    sql = """
        SELECT COALESCE(SUM(jl.debit),0)  AS total_debit,
               COALESCE(SUM(jl.credit),0) AS total_credit
        FROM journal_lines jl
        JOIN journal_entries je ON jl.entry_id = je.id
        WHERE jl.account_id = ?
    """
    params: list = [acct.id]
    if start_date:
        sql += " AND je.entry_date >= ?"
        params.append(start_date)
    if end_date:
        sql += " AND je.entry_date <= ?"
        params.append(end_date)
    with get_connection(db_path) as conn:
        row = conn.execute(sql, params).fetchone()
    debit  = row["total_debit"]
    credit = row["total_credit"]
    if acct.normal_side == "Debit":
        return round(debit - credit, 2)
    return round(credit - debit, 2)


def all_account_balances(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    db_path=DB_PATH,
) -> list[dict]:
    """Return a list of {code, name, type, normal_side, balance} for every account."""
    sql = """
        SELECT a.id, a.code, a.name, a.type, a.normal_side,
               COALESCE(SUM(jl.debit),0)  AS total_debit,
               COALESCE(SUM(jl.credit),0) AS total_credit
        FROM accounts a
        LEFT JOIN journal_lines jl ON jl.account_id = a.id
        LEFT JOIN journal_entries je ON jl.entry_id = je.id
    """
    params = []
    where_clauses = []
    if start_date:
        where_clauses.append("(je.entry_date IS NULL OR je.entry_date >= ?)")
        params.append(start_date)
    if end_date:
        where_clauses.append("(je.entry_date IS NULL OR je.entry_date <= ?)")
        params.append(end_date)
    if where_clauses:
        sql += " WHERE " + " AND ".join(where_clauses)
    sql += " GROUP BY a.id ORDER BY a.code"

    with get_connection(db_path) as conn:
        rows = conn.execute(sql, params).fetchall()

    result = []
    for row in rows:
        dr, cr = row["total_debit"], row["total_credit"]
        if row["normal_side"] == "Debit":
            balance = round(dr - cr, 2)
        else:
            balance = round(cr - dr, 2)
        result.append({
            "code":        row["code"],
            "name":        row["name"],
            "type":        row["type"],
            "normal_side": row["normal_side"],
            "debit":       dr,
            "credit":      cr,
            "balance":     balance,
        })
    return result
