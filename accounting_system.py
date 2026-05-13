#!/usr/bin/env python3
"""
Invegrow Accounting System — CLI

Usage:
  python accounting_system.py init
  python accounting_system.py accounts [--type ASSET|LIABILITY|EQUITY|REVENUE|EXPENSE]
  python accounting_system.py post --date DATE --desc TEXT [--ref REF]
                                    --line CODE DEBIT CREDIT [--line ...]
  python accounting_system.py journal [--start DATE] [--end DATE]
  python accounting_system.py balance CODE [--start DATE] [--end DATE]
  python accounting_system.py trial-balance [--start DATE] [--end DATE]
  python accounting_system.py income-statement [--start DATE] [--end DATE]
  python accounting_system.py balance-sheet [--as-of DATE]
  python accounting_system.py export [--start DATE] [--end DATE] [--as-of DATE]
                                      [--output PATH]
"""
import argparse
import sys
from accounting import (
    init_db, seed_chart_of_accounts,
    list_accounts, account_balance,
    JournalLine, post_journal_entry, list_journal_entries,
    trial_balance, income_statement, balance_sheet, export_workbook,
)
from accounting.reports import (
    print_trial_balance, print_income_statement, print_balance_sheet,
)


def cmd_init(args):
    init_db()
    seed_chart_of_accounts()
    print("Database initialised and chart of accounts seeded.")


def cmd_accounts(args):
    accts = list_accounts(args.type)
    print(f"\n  {'Code':<8} {'Name':<32} {'Type':<12} {'Normal Side'}")
    print(f"  {'─'*8} {'─'*32} {'─'*12} {'─'*11}")
    for a in accts:
        print(f"  {a.code:<8} {a.name:<32} {a.type:<12} {a.normal_side}")
    print()


def cmd_post(args):
    lines = []
    for (code, dr, cr) in args.line:
        lines.append(JournalLine(account_code=code, debit=float(dr), credit=float(cr)))
    entry = post_journal_entry(
        entry_date=args.date,
        lines=lines,
        description=args.desc,
        reference=args.ref or "",
    )
    print(f"Posted journal entry #{entry.id}  ({entry.entry_date})  {entry.description}")
    for line in entry.lines:
        dr = f"${line['debit']:,.2f}" if line["debit"] else ""
        cr = f"${line['credit']:,.2f}" if line["credit"] else ""
        print(f"  {line['code']:<8} {line['name']:<28} {dr:>12} {cr:>12}")


def cmd_journal(args):
    entries = list_journal_entries(args.start, args.end)
    if not entries:
        print("No journal entries found.")
        return
    for e in entries:
        print(f"\n  #{e.id:>4}  {e.entry_date}  Ref: {e.reference or '—'}  {e.description}")
        for line in e.lines:
            dr = f"${line['debit']:>10,.2f}" if line["debit"] else " " * 12
            cr = f"${line['credit']:>10,.2f}" if line["credit"] else " " * 12
            print(f"         {line['code']:<8} {line['name']:<28} {dr}  {cr}")
    print()


def cmd_balance(args):
    bal = account_balance(args.code, args.start, args.end)
    print(f"\n  Balance for account {args.code}: ${bal:,.2f}\n")


def cmd_trial_balance(args):
    tb = trial_balance(args.start, args.end)
    print_trial_balance(tb)


def cmd_income_statement(args):
    is_ = income_statement(args.start, args.end)
    print_income_statement(is_)


def cmd_balance_sheet(args):
    bs = balance_sheet(args.as_of)
    print_balance_sheet(bs)


def cmd_export(args):
    tb  = trial_balance(args.start, args.end)
    is_ = income_statement(args.start, args.end)
    bs  = balance_sheet(args.as_of)
    out = args.output or "Invegrow_Accounting_Report.xlsx"
    path = export_workbook(tb, is_, bs, out)
    print(f"Report exported to: {path}")


# ── argument parsing ──────────────────────────────────────────────────────────

def _date_arg(p, *flags, **kw):
    p.add_argument(*flags, metavar="YYYY-MM-DD", **kw)


def build_parser():
    parser = argparse.ArgumentParser(
        prog="accounting_system.py",
        description="Invegrow Double-Entry Accounting System",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # init
    sub.add_parser("init", help="Initialise the database and seed chart of accounts")

    # accounts
    p_accts = sub.add_parser("accounts", help="List chart of accounts")
    p_accts.add_argument("--type", choices=["Asset","Liability","Equity","Revenue","Expense"])

    # post
    p_post = sub.add_parser("post", help="Post a journal entry")
    _date_arg(p_post, "--date", required=True)
    p_post.add_argument("--desc", required=True, metavar="TEXT")
    p_post.add_argument("--ref", metavar="REF")
    p_post.add_argument(
        "--line", nargs=3, metavar=("CODE","DEBIT","CREDIT"),
        action="append", required=True,
    )

    # journal
    p_jnl = sub.add_parser("journal", help="List journal entries")
    _date_arg(p_jnl, "--start"); _date_arg(p_jnl, "--end")

    # balance
    p_bal = sub.add_parser("balance", help="Show balance of one account")
    p_bal.add_argument("code")
    _date_arg(p_bal, "--start"); _date_arg(p_bal, "--end")

    # trial-balance
    p_tb = sub.add_parser("trial-balance", help="Print trial balance")
    _date_arg(p_tb, "--start"); _date_arg(p_tb, "--end")

    # income-statement
    p_is = sub.add_parser("income-statement", help="Print income statement")
    _date_arg(p_is, "--start"); _date_arg(p_is, "--end")

    # balance-sheet
    p_bs = sub.add_parser("balance-sheet", help="Print balance sheet")
    _date_arg(p_bs, "--as-of")

    # export
    p_ex = sub.add_parser("export", help="Export reports to Excel")
    _date_arg(p_ex, "--start"); _date_arg(p_ex, "--end")
    _date_arg(p_ex, "--as-of")
    p_ex.add_argument("--output", metavar="PATH")

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    dispatch = {
        "init":             cmd_init,
        "accounts":         cmd_accounts,
        "post":             cmd_post,
        "journal":          cmd_journal,
        "balance":          cmd_balance,
        "trial-balance":    cmd_trial_balance,
        "income-statement": cmd_income_statement,
        "balance-sheet":    cmd_balance_sheet,
        "export":           cmd_export,
    }
    try:
        dispatch[args.command](args)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
