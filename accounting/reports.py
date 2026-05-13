"""Financial report generators: Trial Balance, Income Statement, Balance Sheet."""
from __future__ import annotations
from typing import Optional
from .ledger import all_account_balances


# ── helpers ──────────────────────────────────────────────────────────────────

def _subtotal(rows: list[dict], account_type: str) -> float:
    return round(sum(r["balance"] for r in rows if r["type"] == account_type), 2)


# ── Trial Balance ─────────────────────────────────────────────────────────────

def trial_balance(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    db_path=None,
) -> dict:
    from .database import DB_PATH
    db_path = db_path or DB_PATH
    rows = all_account_balances(start_date, end_date, db_path)

    debit_total  = round(sum(r["debit"]  for r in rows), 2)
    credit_total = round(sum(r["credit"] for r in rows), 2)

    return {
        "title":        "Trial Balance",
        "period":       _period_label(start_date, end_date),
        "rows":         rows,
        "debit_total":  debit_total,
        "credit_total": credit_total,
        "balanced":     debit_total == credit_total,
    }


# ── Income Statement ──────────────────────────────────────────────────────────

def income_statement(
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    db_path=None,
) -> dict:
    from .database import DB_PATH
    db_path = db_path or DB_PATH
    rows = all_account_balances(start_date, end_date, db_path)

    revenue_rows  = [r for r in rows if r["type"] == "Revenue"]
    expense_rows  = [r for r in rows if r["type"] == "Expense"]

    total_revenue  = _subtotal(rows, "Revenue")
    total_expenses = _subtotal(rows, "Expense")
    net_income     = round(total_revenue - total_expenses, 2)

    return {
        "title":          "Income Statement",
        "period":         _period_label(start_date, end_date),
        "revenue_rows":   revenue_rows,
        "expense_rows":   expense_rows,
        "total_revenue":  total_revenue,
        "total_expenses": total_expenses,
        "net_income":     net_income,
    }


# ── Balance Sheet ─────────────────────────────────────────────────────────────

def balance_sheet(
    as_of_date: Optional[str] = None,
    db_path=None,
) -> dict:
    from .database import DB_PATH
    db_path = db_path or DB_PATH
    rows = all_account_balances(end_date=as_of_date, db_path=db_path)

    asset_rows     = [r for r in rows if r["type"] == "Asset"]
    liability_rows = [r for r in rows if r["type"] == "Liability"]
    equity_rows    = [r for r in rows if r["type"] == "Equity"]

    # include net income in equity section
    rev  = _subtotal(rows, "Revenue")
    exp  = _subtotal(rows, "Expense")
    net  = round(rev - exp, 2)

    total_assets      = _subtotal(rows, "Asset")
    total_liabilities = _subtotal(rows, "Liability")
    total_equity      = round(_subtotal(rows, "Equity") + net, 2)
    total_l_e         = round(total_liabilities + total_equity, 2)

    return {
        "title":             "Balance Sheet",
        "as_of":             as_of_date or "All dates",
        "asset_rows":        asset_rows,
        "liability_rows":    liability_rows,
        "equity_rows":       equity_rows,
        "net_income":        net,
        "total_assets":      total_assets,
        "total_liabilities": total_liabilities,
        "total_equity":      total_equity,
        "total_l_e":         total_l_e,
        "balanced":          total_assets == total_l_e,
    }


# ── pretty-print to console ───────────────────────────────────────────────────

def print_trial_balance(tb: dict) -> None:
    print(f"\n{'─'*62}")
    print(f"  {tb['title']}  {tb['period']}")
    print(f"{'─'*62}")
    print(f"  {'Code':<8} {'Account':<28} {'Debit':>10} {'Credit':>10}")
    print(f"  {'─'*8} {'─'*28} {'─'*10} {'─'*10}")
    for r in tb["rows"]:
        dr = f"${r['debit']:>9,.2f}" if r["debit"]  else ""
        cr = f"${r['credit']:>9,.2f}" if r["credit"] else ""
        print(f"  {r['code']:<8} {r['name']:<28} {dr:>10} {cr:>10}")
    print(f"  {'─'*8} {'─'*28} {'─'*10} {'─'*10}")
    print(f"  {'TOTALS':<38} ${tb['debit_total']:>9,.2f} ${tb['credit_total']:>9,.2f}")
    status = "BALANCED" if tb["balanced"] else "*** DOES NOT BALANCE ***"
    print(f"\n  {status}")
    print(f"{'─'*62}\n")


def print_income_statement(is_: dict) -> None:
    print(f"\n{'─'*52}")
    print(f"  {is_['title']}  {is_['period']}")
    print(f"{'─'*52}")
    print("  REVENUE")
    for r in is_["revenue_rows"]:
        print(f"    {r['code']:<8} {r['name']:<24} ${r['balance']:>10,.2f}")
    print(f"  {'Total Revenue':<34} ${is_['total_revenue']:>10,.2f}")
    print()
    print("  EXPENSES")
    for r in is_["expense_rows"]:
        print(f"    {r['code']:<8} {r['name']:<24} ${r['balance']:>10,.2f}")
    print(f"  {'Total Expenses':<34} ${is_['total_expenses']:>10,.2f}")
    print(f"{'─'*52}")
    label = "NET INCOME" if is_["net_income"] >= 0 else "NET LOSS"
    print(f"  {label:<34} ${is_['net_income']:>10,.2f}")
    print(f"{'─'*52}\n")


def print_balance_sheet(bs: dict) -> None:
    print(f"\n{'─'*52}")
    print(f"  {bs['title']}  as of {bs['as_of']}")
    print(f"{'─'*52}")
    print("  ASSETS")
    for r in bs["asset_rows"]:
        print(f"    {r['code']:<8} {r['name']:<24} ${r['balance']:>10,.2f}")
    print(f"  {'Total Assets':<34} ${bs['total_assets']:>10,.2f}")
    print()
    print("  LIABILITIES")
    for r in bs["liability_rows"]:
        print(f"    {r['code']:<8} {r['name']:<24} ${r['balance']:>10,.2f}")
    print(f"  {'Total Liabilities':<34} ${bs['total_liabilities']:>10,.2f}")
    print()
    print("  EQUITY")
    for r in bs["equity_rows"]:
        print(f"    {r['code']:<8} {r['name']:<24} ${r['balance']:>10,.2f}")
    print(f"    {'Net Income (Current Period)':<32} ${bs['net_income']:>10,.2f}")
    print(f"  {'Total Equity':<34} ${bs['total_equity']:>10,.2f}")
    print(f"{'─'*52}")
    print(f"  {'Total Liabilities + Equity':<34} ${bs['total_l_e']:>10,.2f}")
    status = "BALANCED" if bs["balanced"] else "*** DOES NOT BALANCE ***"
    print(f"\n  {status}")
    print(f"{'─'*52}\n")


# ── internal ─────────────────────────────────────────────────────────────────

def _period_label(start: Optional[str], end: Optional[str]) -> str:
    if start and end:
        return f"({start} to {end})"
    if start:
        return f"(from {start})"
    if end:
        return f"(through {end})"
    return "(all dates)"
