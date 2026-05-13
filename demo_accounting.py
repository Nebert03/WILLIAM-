#!/usr/bin/env python3
"""
Demo script: populates the accounting system with 6 months of realistic
Invegrow Foods Line transactions and exports a full Excel report.
"""
from pathlib import Path
from accounting import (
    init_db, seed_chart_of_accounts,
    JournalLine, post_journal_entry,
    trial_balance, income_statement, balance_sheet, export_workbook,
)
from accounting.reports import (
    print_trial_balance, print_income_statement, print_balance_sheet,
)

DB = Path("accounting.db")


def main():
    # ── 0. Fresh database ────────────────────────────────────────────────────
    if DB.exists():
        DB.unlink()
    init_db()
    seed_chart_of_accounts()
    print("Database initialised.")

    def je(date, desc, lines, ref=""):
        post_journal_entry(date, [JournalLine(*l) for l in lines], desc, ref)

    # ── 1. Owner invests $150,000 capital ────────────────────────────────────
    je("2026-01-01", "Owner capital contribution",
       [("1000", 150_000, 0),
        ("3000", 0, 150_000)], ref="OC-001")

    # ── 2. Purchase equipment on credit $30,000 ──────────────────────────────
    je("2026-01-05", "Purchase processing equipment",
       [("1500", 30_000, 0),
        ("2300", 0, 30_000)], ref="EQ-001")

    # ── 3. Pay first month rent ──────────────────────────────────────────────
    je("2026-01-06", "January rent payment",
       [("5200", 5_000, 0),
        ("1000", 0, 5_000)], ref="RENT-JAN")

    # ── 4. Purchase inventory on account ─────────────────────────────────────
    je("2026-01-10", "Purchase inventory from supplier",
       [("1200", 40_000, 0),
        ("2000", 0, 40_000)], ref="PO-001")

    # ── 5. First sales – cash + on account ───────────────────────────────────
    je("2026-01-20", "January product sales",
       [("1000", 25_000, 0),
        ("1100", 15_000, 0),
        ("4000", 0, 40_000)], ref="INV-001")

    je("2026-01-20", "Cost of goods sold – January",
       [("5000", 20_000, 0),
        ("1200", 0, 20_000)], ref="COGS-JAN")

    # ── 6. Pay salaries ───────────────────────────────────────────────────────
    je("2026-01-31", "January salaries",
       [("5100", 18_000, 0),
        ("1000", 0, 18_000)], ref="PAY-JAN")

    # ── 7. Collect receivables ────────────────────────────────────────────────
    je("2026-02-05", "Collect receivables from January",
       [("1000", 15_000, 0),
        ("1100", 0, 15_000)], ref="REC-001")

    # ── 8. Pay vendor ─────────────────────────────────────────────────────────
    je("2026-02-10", "Pay supplier invoice PO-001",
       [("2000", 40_000, 0),
        ("1000", 0, 40_000)], ref="PMT-001")

    # ── 9. Feb–June: monthly transactions (loop) ─────────────────────────────
    months = [
        ("2026-02", 52_000, 26_000, 18_500, 5_000),
        ("2026-03", 68_000, 34_000, 19_000, 5_000),
        ("2026-04", 75_000, 37_500, 19_000, 5_000),
        ("2026-05", 80_000, 40_000, 20_000, 5_000),
        ("2026-06", 90_000, 45_000, 21_000, 5_000),
    ]
    for ym, sales, cogs, salaries, rent in months:
        m = ym.replace("-", "")
        je(f"{ym}-15", f"Product sales – {ym}",
           [("1000", sales * 0.6, 0),
            ("1100", sales * 0.4, 0),
            ("4000", 0, sales)], ref=f"INV-{m}")

        je(f"{ym}-15", f"COGS – {ym}",
           [("5000", cogs, 0),
            ("1200", 0, cogs)], ref=f"COGS-{m}")

        je(f"{ym}-25", f"Salaries – {ym}",
           [("5100", salaries, 0),
            ("1000", 0, salaries)], ref=f"PAY-{m}")

        je(f"{ym}-01", f"Rent – {ym}",
           [("5200", rent, 0),
            ("1000", 0, rent)], ref=f"RENT-{m}")

        # replenish inventory monthly
        inv_purchase = cogs * 1.1
        je(f"{ym}-05", f"Inventory purchase – {ym}",
           [("1200", inv_purchase, 0),
            ("2000", 0, inv_purchase)], ref=f"PO-{m}")

        je(f"{ym}-20", f"Pay supplier – {ym}",
           [("2000", inv_purchase, 0),
            ("1000", 0, inv_purchase)], ref=f"PMT-{m}")

    # ── 10. Utilities, marketing, insurance ──────────────────────────────────
    for ym in ["2026-01","2026-02","2026-03","2026-04","2026-05","2026-06"]:
        m = ym.replace("-", "")
        je(f"{ym}-28", f"Utilities – {ym}",
           [("5300", 1_200, 0), ("1000", 0, 1_200)], ref=f"UTIL-{m}")
        je(f"{ym}-28", f"Marketing – {ym}",
           [("5400", 3_500, 0), ("1000", 0, 3_500)], ref=f"MKT-{m}")

    je("2026-01-15", "Annual insurance premium (prepaid)",
       [("1300", 6_000, 0), ("1000", 0, 6_000)], ref="INS-001")
    je("2026-06-30", "Insurance expense recognition H1",
       [("5800", 3_000, 0), ("1300", 0, 3_000)], ref="INS-EXP")

    # ── 11. Depreciation – 6 months ──────────────────────────────────────────
    je("2026-06-30", "Equipment depreciation H1 (5-yr straight-line)",
       [("5500", 3_000, 0),
        ("1510", 0, 3_000)], ref="DEP-H1")

    # ── 12. Note payment ─────────────────────────────────────────────────────
    je("2026-06-30", "Equipment note principal + interest payment",
       [("2300", 6_000, 0),
        ("5600", 750, 0),
        ("1000", 0, 6_750)], ref="LOAN-PMT")

    # ── 13. Print reports ────────────────────────────────────────────────────
    period_start = "2026-01-01"
    period_end   = "2026-06-30"

    tb  = trial_balance(period_start, period_end)
    is_ = income_statement(period_start, period_end)
    bs  = balance_sheet(period_end)

    print_trial_balance(tb)
    print_income_statement(is_)
    print_balance_sheet(bs)

    # ── 14. Export Excel ─────────────────────────────────────────────────────
    out = export_workbook(tb, is_, bs, "Invegrow_Accounting_H1_2026.xlsx")
    print(f"\nExcel report saved to: {out}\n")


if __name__ == "__main__":
    main()
