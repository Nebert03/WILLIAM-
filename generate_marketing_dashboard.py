#!/usr/bin/env python3
"""
Invegrow Foods Line — Marketing Metrics Dashboard Generator
Builds a fully-linked, formula-driven Excel workbook.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

# ─────────────────────────────────────────────
#  COLOUR PALETTE
# ─────────────────────────────────────────────
DK_GRN   = "1B5E20"
MD_GRN   = "2E7D32"
LT_GRN   = "66BB6A"
PL_GRN   = "C8E6C9"
VP_GRN   = "E8F5E9"
TEAL     = "00695C"
PL_TEAL  = "E0F2F1"
BLUE     = "01579B"
PL_BLUE  = "E1F5FE"
GOLD     = "F57F17"
PL_GOLD  = "FFF8E1"
RED      = "B71C1C"
PL_RED   = "FFEBEE"
ORANGE   = "E65100"
PL_ORG   = "FFF3E0"
DARK     = "212121"
MID_GRY  = "757575"
LT_GRY   = "F5F5F5"
BRD_GRY  = "BDBDBD"
WHITE    = "FFFFFF"
CHARCOAL = "37474F"

# ─────────────────────────────────────────────
#  STYLE HELPERS
# ─────────────────────────────────────────────
def fl(c):
    return PatternFill("solid", fgColor=c)

def fn(sz=11, bold=False, color=DARK, italic=False):
    return Font(name="Calibri", size=sz, bold=bold, italic=italic, color=color)

def al(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def bd(t="thin"):
    s = Side(style=t)
    return Border(left=s, right=s, top=s, bottom=s)

def bd_outer():
    m = Side(style="medium")
    return Border(left=m, right=m, top=m, bottom=m)

def mc(ws, r1, c1, r2, c2, val, bg, fg=WHITE, sz=11, bold=True,
       ha="center", va="center", wrap=False, border="thin"):
    """Merge cells + fill + font."""
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    s = Side(style=border)
    bdr = Border(left=s, right=s, top=s, bottom=s)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.fill = fl(bg)
            cell.border = bdr
    cell = ws.cell(r1, c1)
    cell.value = val
    cell.font = fn(sz, bold, fg)
    cell.alignment = al(ha, va, wrap)
    return cell

def col_hdr(ws, row, col, label, bg, fg=WHITE, wrap=True):
    cell = ws.cell(row, col)
    cell.value = label
    cell.fill = fl(bg)
    cell.font = fn(10, True, fg)
    cell.alignment = al("center", "center", wrap)
    cell.border = bd()
    return cell

def w(ws, mapping):
    """Set column widths. mapping: {"A": 12, ...}"""
    for col, width in mapping.items():
        ws.column_dimensions[col].width = width

def rh(ws, mapping):
    """Set row heights. mapping: {1: 40, ...}"""
    for row, height in mapping.items():
        ws.row_dimensions[row].height = height

def dv_list(ws, sqref, formula):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=False)
    dv.sqref = sqref
    ws.add_data_validation(dv)

# ─────────────────────────────────────────────
#  REFERENCE DATA
# ─────────────────────────────────────────────
MONTHS = [
    "Jan-25","Feb-25","Mar-25","Apr-25","May-25","Jun-25",
    "Jul-25","Aug-25","Sep-25","Oct-25","Nov-25","Dec-25",
    "Jan-26","Feb-26","Mar-26","Apr-26","May-26","Jun-26",
    "Jul-26","Aug-26","Sep-26","Oct-26","Nov-26","Dec-26",
]

PRODUCTS = [
    "Hemp Seed Oil 500ml",
    "Hemp Protein Powder 1kg",
    "Hemp Hearts 500g",
    "Hemp Granola 400g",
    "Hemp Energy Bar (Box 12)",
    "Hemp Superfood Blend 300g",
    "Hemp Omega Capsules 60s",
    "Hemp Infused Honey 250g",
    "Hemp Seed Butter 400g",
    "Hemp Flour 1kg",
]

PLATFORMS   = ["Instagram","Facebook","TikTok","YouTube","LinkedIn","Twitter/X"]
AD_PLATS    = ["Meta (Facebook/Instagram)","Google Ads","TikTok Ads","YouTube Ads",
               "Display/Programmatic","Influencer","Other"]
OBJECTIVES  = ["Brand Awareness","Reach","Engagement","Website Traffic",
               "Lead Generation","Product Launch","Seasonal Promo","Retargeting"]
STATUSES    = ["Planning","Active","Paused","Completed","Cancelled"]
CHANNELS    = ["Social Organic","Paid Social","Google Ads","Email","Influencer",
               "PR/Media","Events","In-store","Website/SEO","Other"]

# Named range values (used in DV formulas — stored in SETTINGS)
SETTINGS_NAME = "SETTINGS"
DATA_ROWS     = 50   # blank entry rows for campaign sheets


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 1 — SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
def build_settings(ws):
    ws.sheet_view.showGridLines = False
    ws.tab_color = CHARCOAL

    mc(ws, 1, 1, 2, 16, "⚙   SETTINGS & REFERENCE DATA", DK_GRN, WHITE, 14, True, "left")
    mc(ws, 3, 1, 3, 16, "DO NOT MODIFY  —  Reference lists used by dropdowns and formulas across the workbook", PL_GRN, DARK, 9, False)
    rh(ws, {1: 42, 2: 0, 3: 16})

    RS = 5  # data start row

    # ── Products
    mc(ws, RS, 1, RS, 3, "PRODUCTS — Foods Line (10)", MD_GRN, WHITE, 11, True)
    for i, p in enumerate(PRODUCTS, 1):
        r = RS + i
        bg = VP_GRN if i % 2 == 0 else WHITE
        ws.cell(r, 1).value = i;          ws.cell(r, 1).font = fn(10, False, MID_GRY); ws.cell(r, 1).alignment = al("center")
        ws.cell(r, 2).value = p;          ws.cell(r, 2).font = fn(10)
        ws.cell(r, 3).value = "";
        for c in range(1, 4):
            ws.cell(r, c).fill = fl(bg); ws.cell(r, c).border = bd()
        rh(ws, {r: 18})

    # ── Social Platforms
    mc(ws, RS, 5, RS, 7, "SOCIAL PLATFORMS", TEAL, WHITE, 11, True)
    for i, p in enumerate(PLATFORMS, 1):
        r = RS + i; bg = PL_TEAL if i % 2 == 0 else WHITE
        ws.cell(r, 5).value = i;  ws.cell(r, 5).font = fn(10, False, MID_GRY); ws.cell(r, 5).alignment = al("center")
        ws.cell(r, 6).value = p;  ws.cell(r, 6).font = fn(10)
        for c in range(5, 8):
            ws.cell(r, c).fill = fl(bg); ws.cell(r, c).border = bd()

    # ── Ad Platforms
    mc(ws, RS, 9, RS, 11, "AD / MEDIA PLATFORMS", BLUE, WHITE, 11, True)
    for i, p in enumerate(AD_PLATS, 1):
        r = RS + i; bg = PL_BLUE if i % 2 == 0 else WHITE
        ws.cell(r, 9).value  = i;  ws.cell(r, 9).font  = fn(10, False, MID_GRY); ws.cell(r, 9).alignment = al("center")
        ws.cell(r, 10).value = p;  ws.cell(r, 10).font = fn(10)
        for c in range(9, 12):
            ws.cell(r, c).fill = fl(bg); ws.cell(r, c).border = bd()

    # ── Reporting Months
    mc(ws, RS, 13, RS, 14, "REPORTING MONTHS", DK_GRN, WHITE, 11, True)
    for i, m in enumerate(MONTHS, 1):
        r = RS + i; bg = VP_GRN if i % 2 == 0 else WHITE
        ws.cell(r, 13).value = i;  ws.cell(r, 13).font = fn(10, False, MID_GRY); ws.cell(r, 13).alignment = al("center")
        ws.cell(r, 14).value = m;  ws.cell(r, 14).font = fn(10, True, MD_GRN)
        for c in range(13, 15):
            ws.cell(r, c).fill = fl(bg); ws.cell(r, c).border = bd()

    # ── Campaign Statuses
    SR2 = RS + len(PRODUCTS) + 3
    mc(ws, SR2, 5, SR2, 7, "CAMPAIGN STATUSES", GOLD, WHITE, 11, True)
    for i, s in enumerate(STATUSES, 1):
        r = SR2 + i; bg = PL_GOLD if i % 2 == 0 else WHITE
        ws.cell(r, 5).value = i;  ws.cell(r, 5).font = fn(10, False, MID_GRY); ws.cell(r, 5).alignment = al("center")
        ws.cell(r, 6).value = s;  ws.cell(r, 6).font = fn(10)
        for c in range(5, 8):
            ws.cell(r, c).fill = fl(bg); ws.cell(r, c).border = bd()

    # ── Campaign Objectives
    mc(ws, SR2, 9, SR2, 11, "CAMPAIGN OBJECTIVES", RED, WHITE, 11, True)
    for i, o in enumerate(OBJECTIVES, 1):
        r = SR2 + i; bg = PL_RED if i % 2 == 0 else WHITE
        ws.cell(r, 9).value  = i;  ws.cell(r, 9).font  = fn(10, False, MID_GRY); ws.cell(r, 9).alignment = al("center")
        ws.cell(r, 10).value = o;  ws.cell(r, 10).font = fn(10)
        for c in range(9, 12):
            ws.cell(r, c).fill = fl(bg); ws.cell(r, c).border = bd()

    w(ws, {"A":5,"B":30,"C":4,"D":3,"E":5,"F":28,"G":4,"H":3,
           "I":5,"J":32,"K":4,"L":3,"M":5,"N":12})


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 2 — DIGITAL ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════
def build_digital(ws):
    ws.sheet_view.showGridLines = False
    ws.tab_color = MD_GRN
    TOTAL_COLS = 15

    # Header
    mc(ws, 1, 1, 1, TOTAL_COLS, "   🌐  DIGITAL ANALYTICS  —  Website & SEO Performance", MD_GRN, WHITE, 15, True, "left")
    mc(ws, 2, 1, 2, TOTAL_COLS,
       "Monthly website performance data. Sources: Google Analytics / GA4, Google Search Console, CMS dashboards.",
       PL_GRN, DARK, 10, False, "left")
    mc(ws, 3, 1, 3, TOTAL_COLS,
       "INVEGROW  |  Foods Line Marketing  |  Bounce Rate & Conversion Rate: enter as a number e.g. 45.2 (meaning 45.2%)  |  Session Duration: minutes",
       DK_GRN, "A5D6A7", 9, False, "center")
    mc(ws, 4, 1, 4, TOTAL_COLS,
       "⚡  Yellow cells = required entry.  All traffic columns = number of sessions from that source.  Totals row auto-calculates.",
       "FFF9C4", DARK, 9, False, "left")
    rh(ws, {1: 42, 2: 22, 3: 16, 4: 20, 5: 40})

    # Column group headers (row 5)
    mc(ws, 5, 1, 5, 1, "PERIOD", CHARCOAL, WHITE, 10, True)
    mc(ws, 5, 2, 5, 4, "OVERALL TRAFFIC", DK_GRN, WHITE, 10, True)
    mc(ws, 5, 5, 5, 6, "ENGAGEMENT", TEAL, WHITE, 10, True)
    mc(ws, 5, 7, 5, 11, "TRAFFIC SOURCES", BLUE, WHITE, 10, True)
    mc(ws, 5, 12, 5, 13, "CONVERSIONS", GOLD, DARK, 10, True)
    mc(ws, 5, 14, 5, 14, "MKTG CHANNEL NOTE", MID_GRY, WHITE, 10, True)
    mc(ws, 5, 15, 5, 15, "NOTES", MID_GRY, WHITE, 10, True)

    # Sub-headers (row 6) — CORRECTED: 15 total cols
    sub_hdrs = [
        (1,  "MONTH",           CHARCOAL, WHITE),
        (2,  "SESSIONS",        DK_GRN,   WHITE),
        (3,  "UNIQUE\nVISITORS",DK_GRN,   WHITE),
        (4,  "PAGE\nVIEWS",     DK_GRN,   WHITE),
        (5,  "BOUNCE\nRATE %",  TEAL,     WHITE),
        (6,  "AVG SESSION\nDURATION (min)", TEAL, WHITE),
        (7,  "ORGANIC\nTRAFFIC",BLUE,     WHITE),
        (8,  "PAID\nTRAFFIC",   BLUE,     WHITE),
        (9,  "SOCIAL\nTRAFFIC", BLUE,     WHITE),
        (10, "DIRECT\nTRAFFIC", BLUE,     WHITE),
        (11, "OTHER\nTRAFFIC",  BLUE,     WHITE),
        (12, "GOAL\nCOMPLETIONS", GOLD,  DARK),
        (13, "CONVERSION\nRATE %", GOLD, DARK),
        (14, "MAIN CHANNEL\nFOCUS THIS MONTH", MID_GRY, WHITE),
        (15, "NOTES",           MID_GRY,  WHITE),
    ]
    rh(ws, {6: 38})
    for col, label, bg, fg in sub_hdrs:
        col_hdr(ws, 6, col, label, bg, fg)

    # Data rows
    DS = 7
    for i, month in enumerate(MONTHS):
        r = DS + i
        bg = VP_GRN if i % 2 == 0 else WHITE
        # Month cell (locked/styled)
        c = ws.cell(r, 1)
        c.value = month; c.fill = fl(PL_GRN); c.font = fn(10, True, MD_GRN)
        c.alignment = al("center"); c.border = bd()
        # Data entry cells
        for col in range(2, 16):
            cell = ws.cell(r, col)
            cell.fill = fl(bg); cell.border = bd(); cell.alignment = al("center"); cell.font = fn(10)
        # Number formats
        for col, nf in [(2,'#,##0'),(3,'#,##0'),(4,'#,##0'),(5,'0.0'),
                        (6,'0.00'),(7,'#,##0'),(8,'#,##0'),(9,'#,##0'),
                        (10,'#,##0'),(11,'#,##0'),(12,'#,##0'),(13,'0.00')]:
            ws.cell(r, col).number_format = nf
        rh(ws, {r: 18})

    # Totals/Averages row
    TR = DS + len(MONTHS)
    mc(ws, TR, 1, TR, 1, "TOTALS / AVG", DK_GRN, WHITE, 10, True, "center")
    SUM_COLS = [2,3,4,7,8,9,10,11,12]
    AVG_COLS = [5,6,13]
    for col in SUM_COLS:
        cl = get_column_letter(col)
        ws.cell(TR, col).value = f"=SUM({cl}{DS}:{cl}{DS+len(MONTHS)-1})"
        ws.cell(TR, col).fill = fl(PL_GRN); ws.cell(TR, col).font = fn(10,True,DK_GRN)
        ws.cell(TR, col).border = bd(); ws.cell(TR, col).alignment = al("center")
        ws.cell(TR, col).number_format = '#,##0'
    for col in AVG_COLS:
        cl = get_column_letter(col)
        ws.cell(TR, col).value = f'=IFERROR(AVERAGEIF({cl}{DS}:{cl}{DS+len(MONTHS)-1},"<>"),"–")'
        ws.cell(TR, col).fill = fl(PL_GRN); ws.cell(TR, col).font = fn(10,True,DK_GRN)
        ws.cell(TR, col).border = bd(); ws.cell(TR, col).alignment = al("center")
        ws.cell(TR, col).number_format = '0.00'
    rh(ws, {TR: 22})

    w(ws, {"A":11,"B":13,"C":13,"D":12,"E":12,"F":16,"G":13,"H":13,
           "I":13,"J":13,"K":13,"L":14,"M":14,"N":22,"O":30})


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 3 — SOCIAL MEDIA
# ══════════════════════════════════════════════════════════════════════════════
def build_social(ws):
    ws.sheet_view.showGridLines = False
    ws.tab_color = TEAL
    TOTAL_COLS = 15

    mc(ws, 1, 1, 1, TOTAL_COLS, "   📱  SOCIAL MEDIA  —  Platform Performance by Month", TEAL, WHITE, 15, True, "left")
    mc(ws, 2, 1, 2, TOTAL_COLS,
       "One row per platform per month. Track each platform separately for accurate channel comparison.",
       PL_TEAL, DARK, 10, False, "left")
    mc(ws, 3, 1, 3, TOTAL_COLS,
       "INVEGROW  |  Foods Line Marketing  |  Engagement Rate %: (Likes+Comments+Shares+Saves) ÷ Reach × 100  |  Enter follower count at END of month",
       DK_GRN, "A5D6A7", 9, False, "center")
    mc(ws, 4, 1, 4, TOTAL_COLS,
       "⚡  Select Month and Platform from dropdowns.  Engagement Rate auto-calculates if you enter the component metrics.",
       "FFF9C4", DARK, 9, False, "left")
    rh(ws, {1: 42, 2: 22, 3: 16, 4: 20, 5: 38})

    # Column headers
    hdrs = [
        (1,  "MONTH",              CHARCOAL, WHITE),
        (2,  "PLATFORM",           TEAL,     WHITE),
        (3,  "FOLLOWERS\n(END MO.)",TEAL,    WHITE),
        (4,  "NEW\nFOLLOWERS",     TEAL,     WHITE),
        (5,  "POSTS\nPUBLISHED",   MD_GRN,   WHITE),
        (6,  "TOTAL\nIMPRESSIONS", BLUE,     WHITE),
        (7,  "TOTAL\nREACH",       BLUE,     WHITE),
        (8,  "TOTAL\nENGAGEMENTS", DK_GRN,   WHITE),
        (9,  "LIKES",              DK_GRN,   WHITE),
        (10, "COMMENTS",           DK_GRN,   WHITE),
        (11, "SHARES",             DK_GRN,   WHITE),
        (12, "SAVES",              DK_GRN,   WHITE),
        (13, "ENGAGEMENT\nRATE %", GOLD,     DARK),
        (14, "TOP PERFORMING\nCONTENT TYPE", MID_GRY, WHITE),
        (15, "NOTES",              MID_GRY,  WHITE),
    ]
    for col, label, bg, fg in hdrs:
        col_hdr(ws, 5, col, label, bg, fg)

    # Data rows — 6 platforms × 24 months = 144 rows (pre-filled month+platform)
    DS = 6
    row = DS
    for i, month in enumerate(MONTHS):
        for j, platform in enumerate(PLATFORMS):
            bg = VP_GRN if (row - DS) % 2 == 0 else WHITE
            # Month
            c = ws.cell(row, 1); c.value = month
            c.fill = fl(PL_GRN); c.font = fn(10, True, MD_GRN)
            c.alignment = al("center"); c.border = bd()
            # Platform
            c = ws.cell(row, 2); c.value = platform
            c.fill = fl(PL_TEAL); c.font = fn(10, True, TEAL)
            c.alignment = al("center"); c.border = bd()
            # Data cells
            for col in range(3, 16):
                cell = ws.cell(row, col)
                cell.fill = fl(bg); cell.border = bd()
                cell.alignment = al("center"); cell.font = fn(10)
            # Auto-calc engagement rate: =(I+J+K+L)/G*100 if G>0
            er_cell = ws.cell(row, 13)
            col_I = get_column_letter(9);  col_J = get_column_letter(10)
            col_K = get_column_letter(11); col_L = get_column_letter(12)
            col_G = get_column_letter(7)
            er_cell.value = (f'=IFERROR(({col_I}{row}+{col_J}{row}+'
                             f'{col_K}{row}+{col_L}{row})/{col_G}{row}*100,"")')
            er_cell.fill = fl(PL_GOLD); er_cell.font = fn(10, False, DARK)
            er_cell.number_format = '0.00'
            # Number formats
            for col, nf in [(3,'#,##0'),(4,'#,##0'),(5,'#,##0'),(6,'#,##0'),
                            (7,'#,##0'),(8,'#,##0'),(9,'#,##0'),(10,'#,##0'),
                            (11,'#,##0'),(12,'#,##0')]:
                ws.cell(row, col).number_format = nf
            rh(ws, {row: 18})
            row += 1

    # Totals block per platform
    TR = row + 1
    mc(ws, TR, 1, TR, 2, "PLATFORM TOTALS (all months)", DK_GRN, WHITE, 10, True)
    rh(ws, {TR: 22})
    for j, platform in enumerate(PLATFORMS):
        r = TR + 1 + j
        ws.cell(r, 1).value = platform; ws.cell(r, 1).fill = fl(PL_TEAL)
        ws.cell(r, 1).font = fn(10, True, TEAL); ws.cell(r, 1).border = bd()
        ws.cell(r, 1).alignment = al("center")
        # Sum followers (latest), new followers, posts, etc.
        for col, formula_type in [(3,'last'),(4,'sum'),(5,'sum'),(6,'sum'),
                                   (7,'sum'),(8,'sum'),(9,'sum'),(10,'sum'),
                                   (11,'sum'),(12,'sum'),(13,'avg')]:
            cl = get_column_letter(col)
            plat_col = "B"
            if formula_type == 'sum':
                ws.cell(r, col).value = f'=SUMIF({plat_col}{DS}:{plat_col}{row-1},A{r},{cl}{DS}:{cl}{row-1})'
            elif formula_type == 'avg':
                ws.cell(r, col).value = f'=IFERROR(AVERAGEIF({plat_col}{DS}:{plat_col}{row-1},A{r},{cl}{DS}:{cl}{row-1}),"–")'
            else:
                ws.cell(r, col).value = ''
            ws.cell(r, col).fill = fl(VP_GRN); ws.cell(r, col).font = fn(10, True, DK_GRN)
            ws.cell(r, col).border = bd(); ws.cell(r, col).alignment = al("center")
        rh(ws, {r: 18})

    # Dropdowns for platform column
    dv_list(ws, f"B{DS}:B{row-1}", f'"{",".join(PLATFORMS)}"')

    w(ws, {"A":11,"B":22,"C":14,"D":13,"E":11,"F":14,"G":13,"H":14,
           "I":10,"J":11,"K":10,"L":10,"M":14,"N":22,"O":30})


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 4 — PAID ADVERTISING
# ══════════════════════════════════════════════════════════════════════════════
def build_paid(ws):
    ws.sheet_view.showGridLines = False
    ws.tab_color = BLUE
    TOTAL_COLS = 17

    mc(ws, 1, 1, 1, TOTAL_COLS, "   💰  PAID ADVERTISING  —  Campaign & Spend Tracker", BLUE, WHITE, 15, True, "left")
    mc(ws, 2, 1, 2, TOTAL_COLS,
       "One row per campaign per month. Track spend, reach and performance for each paid media campaign.",
       PL_BLUE, DARK, 10, False, "left")
    mc(ws, 3, 1, 3, TOTAL_COLS,
       "INVEGROW  |  Foods Line Marketing  |  All monetary values in ZAR (R)  |  CTR/ROAS auto-calculate from inputs",
       DK_GRN, "A5D6A7", 9, False, "center")
    mc(ws, 4, 1, 4, TOTAL_COLS,
       "⚡  CTR % = Clicks ÷ Impressions × 100  |  CPC = Spend ÷ Clicks  |  CPM = Spend ÷ Impressions × 1000  |  ROAS = Revenue ÷ Spend",
       "FFF9C4", DARK, 9, False, "left")
    rh(ws, {1: 42, 2: 22, 3: 16, 4: 20, 5: 38})

    # Headers
    hdrs = [
        (1,  "MONTH",             CHARCOAL, WHITE),
        (2,  "CAMPAIGN NAME",     BLUE,     WHITE),
        (3,  "AD PLATFORM",       BLUE,     WHITE),
        (4,  "PRODUCT(S)\nPROMOTED", MD_GRN, WHITE),
        (5,  "CAMPAIGN\nOBJECTIVE", MD_GRN, WHITE),
        (6,  "BUDGET\n(R)",       DK_GRN,   WHITE),
        (7,  "ACTUAL\nSPEND (R)", DK_GRN,   WHITE),
        (8,  "BUDGET\nUTILISATION %", DK_GRN, WHITE),
        (9,  "IMPRESSIONS",       TEAL,     WHITE),
        (10, "CLICKS",            TEAL,     WHITE),
        (11, "CTR %",             TEAL,     WHITE),
        (12, "CPC (R)",           TEAL,     WHITE),
        (13, "CPM (R)",           TEAL,     WHITE),
        (14, "CONVERSIONS",       GOLD,     DARK),
        (15, "COST / CONV. (R)",  GOLD,     DARK),
        (16, "ROAS\n(Revenue/Spend)", GOLD, DARK),
        (17, "NOTES",             MID_GRY,  WHITE),
    ]
    for col, label, bg, fg in hdrs:
        col_hdr(ws, 5, col, label, bg, fg)

    DS = 6
    dv_months_formula    = '"' + ','.join(MONTHS) + '"'
    dv_adplats_formula   = '"' + ','.join(AD_PLATS) + '"'
    dv_products_formula  = '"' + ','.join(PRODUCTS) + '"'
    dv_obj_formula       = '"' + ','.join(OBJECTIVES) + '"'

    for i in range(DATA_ROWS):
        r = DS + i
        bg = VP_GRN if i % 2 == 0 else WHITE
        for col in range(1, 18):
            cell = ws.cell(r, col)
            cell.fill = fl(bg); cell.border = bd()
            cell.alignment = al("center"); cell.font = fn(10)
        # Auto-calc formulas
        # Budget utilisation
        ws.cell(r, 8).value  = f'=IFERROR(G{r}/F{r}*100,"")'
        ws.cell(r, 8).fill   = fl(PL_BLUE); ws.cell(r, 8).font = fn(10, False, BLUE)
        # CTR
        ws.cell(r, 11).value = f'=IFERROR(J{r}/I{r}*100,"")'
        ws.cell(r, 11).fill  = fl(PL_TEAL); ws.cell(r, 11).font = fn(10, False, TEAL)
        # CPC
        ws.cell(r, 12).value = f'=IFERROR(G{r}/J{r},"")'
        ws.cell(r, 12).fill  = fl(PL_TEAL); ws.cell(r, 12).font = fn(10, False, TEAL)
        # CPM
        ws.cell(r, 13).value = f'=IFERROR(G{r}/I{r}*1000,"")'
        ws.cell(r, 13).fill  = fl(PL_TEAL); ws.cell(r, 13).font = fn(10, False, TEAL)
        # Cost per conversion
        ws.cell(r, 15).value = f'=IFERROR(G{r}/N{r},"")'
        ws.cell(r, 15).fill  = fl(PL_GOLD); ws.cell(r, 15).font = fn(10, False, GOLD)
        # Number formats
        for col, nf in [(6,'R#,##0.00'),(7,'R#,##0.00'),(8,'0.0'),
                        (9,'#,##0'),(10,'#,##0'),(11,'0.00'),
                        (12,'R#,##0.00'),(13,'R#,##0.00'),(14,'#,##0'),
                        (15,'R#,##0.00'),(16,'0.00')]:
            ws.cell(r, col).number_format = nf
        rh(ws, {r: 18})

    # Dropdowns
    dv_list(ws, f"A{DS}:A{DS+DATA_ROWS-1}", dv_months_formula)
    dv_list(ws, f"C{DS}:C{DS+DATA_ROWS-1}", dv_adplats_formula)
    dv_list(ws, f"E{DS}:E{DS+DATA_ROWS-1}", dv_obj_formula)

    # Totals row
    TR = DS + DATA_ROWS
    mc(ws, TR, 1, TR, 2, "TOTALS", DK_GRN, WHITE, 10, True, "center")
    rh(ws, {TR: 22})
    for col, nf in [(6,'R#,##0.00'),(7,'R#,##0.00'),(9,'#,##0'),
                    (10,'#,##0'),(14,'#,##0')]:
        cl = get_column_letter(col)
        ws.cell(TR, col).value = f"=SUM({cl}{DS}:{cl}{DS+DATA_ROWS-1})"
        ws.cell(TR, col).fill = fl(PL_GRN); ws.cell(TR, col).font = fn(10, True, DK_GRN)
        ws.cell(TR, col).border = bd(); ws.cell(TR, col).alignment = al("center")
        ws.cell(TR, col).number_format = nf
    # Avg CTR
    ws.cell(TR, 11).value = f'=IFERROR(AVERAGEIF(K{DS}:K{DS+DATA_ROWS-1},"<>"),"–")'
    ws.cell(TR, 11).fill = fl(PL_GRN); ws.cell(TR, 11).font = fn(10,True,DK_GRN)
    ws.cell(TR, 11).border = bd(); ws.cell(TR, 11).alignment = al("center")
    # Avg ROAS
    ws.cell(TR, 16).value = f'=IFERROR(AVERAGEIF(P{DS}:P{DS+DATA_ROWS-1},"<>"),"–")'
    ws.cell(TR, 16).fill = fl(PL_GRN); ws.cell(TR, 16).font = fn(10,True,DK_GRN)
    ws.cell(TR, 16).border = bd(); ws.cell(TR, 16).alignment = al("center")

    w(ws, {"A":11,"B":28,"C":26,"D":22,"E":18,"F":13,"G":13,"H":15,
           "I":14,"J":12,"K":10,"L":11,"M":11,"N":14,"O":15,"P":14,"Q":30})


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 5 — EMAIL MARKETING
# ══════════════════════════════════════════════════════════════════════════════
def build_email(ws):
    ws.sheet_view.showGridLines = False
    ws.tab_color = ORANGE
    TOTAL_COLS = 14

    mc(ws, 1, 1, 1, TOTAL_COLS, "   📧  EMAIL MARKETING  —  Campaign Performance", ORANGE, WHITE, 15, True, "left")
    mc(ws, 2, 1, 2, TOTAL_COLS,
       "One row per email campaign. Track deliverability, open rates, click performance and list health.",
       PL_ORG, DARK, 10, False, "left")
    mc(ws, 3, 1, 3, TOTAL_COLS,
       "INVEGROW  |  Foods Line Marketing  |  Rates auto-calculate  |  Industry benchmarks: Open Rate 20–25% | CTR 2–5% | Unsub < 0.5%",
       DK_GRN, "A5D6A7", 9, False, "center")
    mc(ws, 4, 1, 4, TOTAL_COLS,
       "⚡  Delivery Rate = Delivered ÷ Sent  |  Open Rate = Opens ÷ Delivered  |  CTR = Clicks ÷ Delivered",
       "FFF9C4", DARK, 9, False, "left")
    rh(ws, {1: 42, 2: 22, 3: 16, 4: 20, 5: 38})

    hdrs = [
        (1,  "MONTH",                CHARCOAL, WHITE),
        (2,  "CAMPAIGN NAME",        ORANGE,   WHITE),
        (3,  "LIST SEGMENT",         ORANGE,   WHITE),
        (4,  "EMAILS SENT",          BLUE,     WHITE),
        (5,  "DELIVERED",            BLUE,     WHITE),
        (6,  "DELIVERY\nRATE %",     BLUE,     WHITE),
        (7,  "OPENS",                DK_GRN,   WHITE),
        (8,  "OPEN\nRATE %",         DK_GRN,   WHITE),
        (9,  "CLICKS",               TEAL,     WHITE),
        (10, "CLICK-THROUGH\nRATE %",TEAL,     WHITE),
        (11, "UNSUBSCRIBES",         RED,      WHITE),
        (12, "UNSUB\nRATE %",        RED,      WHITE),
        (13, "CONVERSIONS\n(if tracked)", GOLD, DARK),
        (14, "NOTES",                MID_GRY,  WHITE),
    ]
    for col, label, bg, fg in hdrs:
        col_hdr(ws, 5, col, label, bg, fg)

    DS = 6
    for i in range(DATA_ROWS):
        r = DS + i
        bg = VP_GRN if i % 2 == 0 else WHITE
        for col in range(1, 15):
            cell = ws.cell(r, col)
            cell.fill = fl(bg); cell.border = bd()
            cell.alignment = al("center"); cell.font = fn(10)
        # Auto-calc rates
        ws.cell(r, 6).value  = f'=IFERROR(E{r}/D{r}*100,"")'   # Delivery rate
        ws.cell(r, 6).fill   = fl(PL_BLUE); ws.cell(r, 6).font = fn(10, False, BLUE)
        ws.cell(r, 8).value  = f'=IFERROR(G{r}/E{r}*100,"")'   # Open rate
        ws.cell(r, 8).fill   = fl(VP_GRN); ws.cell(r, 8).font = fn(10, False, DK_GRN)
        ws.cell(r, 10).value = f'=IFERROR(I{r}/E{r}*100,"")'   # CTR
        ws.cell(r, 10).fill  = fl(PL_TEAL); ws.cell(r, 10).font = fn(10, False, TEAL)
        ws.cell(r, 12).value = f'=IFERROR(K{r}/E{r}*100,"")'   # Unsub rate
        ws.cell(r, 12).fill  = fl(PL_RED); ws.cell(r, 12).font = fn(10, False, RED)
        # Number formats
        for col, nf in [(4,'#,##0'),(5,'#,##0'),(6,'0.00'),(7,'#,##0'),
                        (8,'0.00'),(9,'#,##0'),(10,'0.00'),(11,'#,##0'),
                        (12,'0.00'),(13,'#,##0')]:
            ws.cell(r, col).number_format = nf
        rh(ws, {r: 18})

    dv_list(ws, f"A{DS}:A{DS+DATA_ROWS-1}", '"' + ','.join(MONTHS) + '"')

    TR = DS + DATA_ROWS
    mc(ws, TR, 1, TR, 2, "TOTALS / AVERAGES", DK_GRN, WHITE, 10, True, "center")
    rh(ws, {TR: 22})
    for col, nf, ftype in [(4,'#,##0','sum'),(5,'#,##0','sum'),(7,'#,##0','sum'),
                            (9,'#,##0','sum'),(11,'#,##0','sum'),(13,'#,##0','sum'),
                            (6,'0.00','avg'),(8,'0.00','avg'),(10,'0.00','avg'),(12,'0.00','avg')]:
        cl = get_column_letter(col)
        if ftype == 'sum':
            ws.cell(TR, col).value = f"=SUM({cl}{DS}:{cl}{DS+DATA_ROWS-1})"
        else:
            ws.cell(TR, col).value = f'=IFERROR(AVERAGEIF({cl}{DS}:{cl}{DS+DATA_ROWS-1},"<>"),"–")'
        ws.cell(TR, col).fill = fl(PL_GRN); ws.cell(TR, col).font = fn(10,True,DK_GRN)
        ws.cell(TR, col).border = bd(); ws.cell(TR, col).alignment = al("center")
        ws.cell(TR, col).number_format = nf

    w(ws, {"A":11,"B":32,"C":22,"D":13,"E":12,"F":13,"G":10,"H":12,
           "I":10,"J":15,"K":14,"L":12,"M":18,"N":30})


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 6 — CONTENT & PR
# ══════════════════════════════════════════════════════════════════════════════
def build_content(ws):
    ws.sheet_view.showGridLines = False
    ws.tab_color = DK_GRN
    TOTAL_COLS = 13

    mc(ws, 1, 1, 1, TOTAL_COLS, "   ✍  CONTENT & PR  —  Organic Content & Earned Media", DK_GRN, WHITE, 15, True, "left")
    mc(ws, 2, 1, 2, TOTAL_COLS,
       "Monthly content output and earned media tracking. Includes blog, press, influencer and video metrics.",
       PL_GRN, DARK, 10, False, "left")
    mc(ws, 3, 1, 3, TOTAL_COLS,
       "INVEGROW  |  Foods Line Marketing  |  Earned Media Value (EMV) in ZAR  |  Use CPM-equivalent formula if EMV not provided directly",
       DK_GRN, "A5D6A7", 9, False, "center")
    mc(ws, 4, 1, 4, TOTAL_COLS,
       "⚡  PR Mentions = any media coverage (online/print/broadcast)  |  Influencer Reach = combined unique reach of all collabs that month",
       "FFF9C4", DARK, 9, False, "left")
    rh(ws, {1: 42, 2: 22, 3: 16, 4: 20, 5: 38})

    hdrs = [
        (1,  "MONTH",                  CHARCOAL, WHITE),
        (2,  "BLOG /\nARTICLES",       MD_GRN,   WHITE),
        (3,  "SOCIAL\nPOSTS (org.)",   MD_GRN,   WHITE),
        (4,  "PRESS\nRELEASES",        TEAL,     WHITE),
        (5,  "PR\nMENTIONS",           TEAL,     WHITE),
        (6,  "EARNED\nMEDIA VAL (R)",  TEAL,     WHITE),
        (7,  "INFLUENCER\nCOLLABS",    BLUE,     WHITE),
        (8,  "INFLUENCER\nCOMB. REACH",BLUE,     WHITE),
        (9,  "VIDEOS\nPUBLISHED",      GOLD,     DARK),
        (10, "VIDEO\nTOTAL VIEWS",     GOLD,     DARK),
        (11, "PODCAST\nFEATURES",      RED,      WHITE),
        (12, "AVG REVIEW\nRATING /5",  DK_GRN,   WHITE),
        (13, "NOTES",                  MID_GRY,  WHITE),
    ]
    for col, label, bg, fg in hdrs:
        col_hdr(ws, 5, col, label, bg, fg)

    DS = 6
    for i, month in enumerate(MONTHS):
        r = DS + i
        bg = VP_GRN if i % 2 == 0 else WHITE
        c = ws.cell(r, 1); c.value = month
        c.fill = fl(PL_GRN); c.font = fn(10, True, MD_GRN)
        c.alignment = al("center"); c.border = bd()
        for col in range(2, 14):
            cell = ws.cell(r, col)
            cell.fill = fl(bg); cell.border = bd()
            cell.alignment = al("center"); cell.font = fn(10)
        for col, nf in [(2,'#,##0'),(3,'#,##0'),(4,'#,##0'),(5,'#,##0'),
                        (6,'R#,##0.00'),(7,'#,##0'),(8,'#,##0'),(9,'#,##0'),
                        (10,'#,##0'),(11,'#,##0'),(12,'0.0')]:
            ws.cell(r, col).number_format = nf
        rh(ws, {r: 18})

    TR = DS + len(MONTHS)
    mc(ws, TR, 1, TR, 1, "TOTALS / AVG", DK_GRN, WHITE, 10, True, "center")
    rh(ws, {TR: 22})
    SUM_C = [2,3,4,5,6,7,8,9,10,11]
    AVG_C = [12]
    for col in SUM_C:
        cl = get_column_letter(col)
        ws.cell(TR, col).value = f"=SUM({cl}{DS}:{cl}{DS+len(MONTHS)-1})"
        ws.cell(TR, col).fill = fl(PL_GRN); ws.cell(TR, col).font = fn(10,True,DK_GRN)
        ws.cell(TR, col).border = bd(); ws.cell(TR, col).alignment = al("center")
    for col in AVG_C:
        cl = get_column_letter(col)
        ws.cell(TR, col).value = f'=IFERROR(AVERAGEIF({cl}{DS}:{cl}{DS+len(MONTHS)-1},"<>"),"–")'
        ws.cell(TR, col).fill = fl(PL_GRN); ws.cell(TR, col).font = fn(10,True,DK_GRN)
        ws.cell(TR, col).border = bd(); ws.cell(TR, col).alignment = al("center")

    w(ws, {"A":11,"B":12,"C":14,"D":12,"E":11,"F":17,"G":14,"H":17,
           "I":13,"J":14,"K":14,"L":14,"M":30})


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 7 — CAMPAIGN TRACKER
# ══════════════════════════════════════════════════════════════════════════════
def build_campaigns(ws):
    ws.sheet_view.showGridLines = False
    ws.tab_color = GOLD
    TOTAL_COLS = 17

    mc(ws, 1, 1, 1, TOTAL_COLS, "   🎯  CAMPAIGN TRACKER  —  End-to-End Campaign Record", GOLD, DARK, 15, True, "left")
    mc(ws, 2, 1, 2, TOTAL_COLS,
       "One row per campaign. Captures strategy, budget and results in a single view for learning and planning.",
       PL_GOLD, DARK, 10, False, "left")
    mc(ws, 3, 1, 3, TOTAL_COLS,
       "INVEGROW  |  Foods Line Marketing  |  Budget Utilisation auto-calculates  |  Status dropdown controls row highlighting",
       DK_GRN, "A5D6A7", 9, False, "center")
    mc(ws, 4, 1, 4, TOTAL_COLS,
       "⚡  Link campaign spend back to the Paid Ads sheet for full cost tracking  |  Add key learnings to improve future campaigns",
       "FFF9C4", DARK, 9, False, "left")
    rh(ws, {1: 42, 2: 22, 3: 16, 4: 20, 5: 38})

    hdrs = [
        (1,  "CAMPAIGN NAME",          GOLD,    DARK),
        (2,  "OBJECTIVE",              GOLD,    DARK),
        (3,  "PRODUCTS\nPROMOTED",     MD_GRN,  WHITE),
        (4,  "START\nDATE",            CHARCOAL,WHITE),
        (5,  "END\nDATE",              CHARCOAL,WHITE),
        (6,  "STATUS",                 CHARCOAL,WHITE),
        (7,  "TOTAL\nBUDGET (R)",      DK_GRN,  WHITE),
        (8,  "SPEND\nTO DATE (R)",     DK_GRN,  WHITE),
        (9,  "BUDGET\nUTIL. %",        DK_GRN,  WHITE),
        (10, "CHANNELS\nUSED",         TEAL,    WHITE),
        (11, "TARGET\nAUDIENCE",       TEAL,    WHITE),
        (12, "KEY\nMESSAGE",           TEAL,    WHITE),
        (13, "TOTAL\nIMPRESSIONS",     BLUE,    WHITE),
        (14, "TOTAL\nREACH",           BLUE,    WHITE),
        (15, "TOTAL\nENGAGEMENTS",     BLUE,    WHITE),
        (16, "LEADS /\nCONVERSIONS",   ORANGE,  WHITE),
        (17, "KEY LEARNINGS",          RED,     WHITE),
    ]
    for col, label, bg, fg in hdrs:
        col_hdr(ws, 5, col, label, bg, fg)

    DS = 6
    for i in range(DATA_ROWS):
        r = DS + i
        bg = VP_GRN if i % 2 == 0 else WHITE
        for col in range(1, 18):
            cell = ws.cell(r, col)
            cell.fill = fl(bg); cell.border = bd()
            cell.alignment = al("center"); cell.font = fn(10)
        # Budget utilisation
        ws.cell(r, 9).value  = f'=IFERROR(H{r}/G{r}*100,"")'
        ws.cell(r, 9).fill   = fl(PL_GRN); ws.cell(r, 9).font = fn(10, False, DK_GRN)
        # Date format
        for col in [4, 5]:
            ws.cell(r, col).number_format = 'DD-MMM-YY'
        # Money format
        for col in [7, 8]:
            ws.cell(r, col).number_format = 'R#,##0.00'
        ws.cell(r, 9).number_format  = '0.0'
        ws.cell(r, 13).number_format = '#,##0'
        ws.cell(r, 14).number_format = '#,##0'
        ws.cell(r, 15).number_format = '#,##0'
        ws.cell(r, 16).number_format = '#,##0'
        # Text alignment for long fields
        for col in [1, 2, 3, 10, 11, 12, 17]:
            ws.cell(r, col).alignment = al("left", "center", True)
        rh(ws, {r: 20})

    dv_list(ws, f"B{DS}:B{DS+DATA_ROWS-1}", '"' + ','.join(OBJECTIVES) + '"')
    dv_list(ws, f"F{DS}:F{DS+DATA_ROWS-1}", '"' + ','.join(STATUSES) + '"')
    dv_list(ws, f"C{DS}:C{DS+DATA_ROWS-1}", '"' + ','.join(PRODUCTS) + '"')

    w(ws, {"A":30,"B":20,"C":22,"D":13,"E":13,"F":13,"G":15,"H":15,"I":13,
           "J":22,"K":22,"L":28,"M":14,"N":13,"O":15,"P":15,"Q":35})


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 8 — PRODUCT MARKETING
# ══════════════════════════════════════════════════════════════════════════════
def build_products(ws):
    ws.sheet_view.showGridLines = False
    ws.tab_color = LT_GRN
    TOTAL_COLS = 14

    mc(ws, 1, 1, 1, TOTAL_COLS, "   📦  PRODUCT MARKETING  —  Monthly Performance per Product", MD_GRN, WHITE, 15, True, "left")
    mc(ws, 2, 1, 2, TOTAL_COLS,
       "One row per product per month. Track marketing investment and visibility for each of the 10 Foods Line SKUs.",
       PL_GRN, DARK, 10, False, "left")
    mc(ws, 3, 1, 3, TOTAL_COLS,
       "INVEGROW  |  Foods Line Marketing  |  All monetary values in ZAR (R)  |  Rating out of 5  |  NPS: -100 to +100",
       DK_GRN, "A5D6A7", 9, False, "center")
    mc(ws, 4, 1, 4, TOTAL_COLS,
       "⚡  Online Mentions = brand + product mentions across all channels  |  Link campaigns to Campaign Tracker sheet",
       "FFF9C4", DARK, 9, False, "left")
    rh(ws, {1: 42, 2: 22, 3: 16, 4: 20, 5: 38})

    hdrs = [
        (1,  "MONTH",             CHARCOAL, WHITE),
        (2,  "PRODUCT NAME",      MD_GRN,   WHITE),
        (3,  "MKTG BUDGET\nALLOCATED (R)", DK_GRN, WHITE),
        (4,  "CHANNELS\nACTIVATED", TEAL,   WHITE),
        (5,  "CAMPAIGNS\nRAN",     TEAL,    WHITE),
        (6,  "TOTAL\nIMPRESSIONS",BLUE,     WHITE),
        (7,  "TOTAL\nREACH",      BLUE,     WHITE),
        (8,  "TOTAL\nENGAGEMENTS",BLUE,     WHITE),
        (9,  "ONLINE\nMENTIONS",  GOLD,     DARK),
        (10, "CUSTOMER\nREVIEWS", GOLD,     DARK),
        (11, "AVG REVIEW\nRATING /5", GOLD, DARK),
        (12, "NPS\nSCORE",        DK_GRN,   WHITE),
        (13, "ACTIVE\nPROMO?",    MD_GRN,   WHITE),
        (14, "NOTES",             MID_GRY,  WHITE),
    ]
    for col, label, bg, fg in hdrs:
        col_hdr(ws, 5, col, label, bg, fg)

    DS = 6
    row = DS
    for i, month in enumerate(MONTHS):
        for j, product in enumerate(PRODUCTS):
            bg = VP_GRN if (row - DS) % 2 == 0 else WHITE
            # Month
            c = ws.cell(row, 1); c.value = month
            c.fill = fl(PL_GRN); c.font = fn(10, True, MD_GRN)
            c.alignment = al("center"); c.border = bd()
            # Product
            c = ws.cell(row, 2); c.value = product
            c.fill = fl(PL_GRN); c.font = fn(9, True, DK_GRN)
            c.alignment = al("left", "center"); c.border = bd()
            # Data cells
            for col in range(3, 15):
                cell = ws.cell(row, col)
                cell.fill = fl(bg); cell.border = bd()
                cell.alignment = al("center"); cell.font = fn(10)
            # Number formats
            for col, nf in [(3,'R#,##0.00'),(5,'#,##0'),(6,'#,##0'),(7,'#,##0'),
                            (8,'#,##0'),(9,'#,##0'),(10,'#,##0'),(11,'0.0'),(12,'0')]:
                ws.cell(row, col).number_format = nf
            rh(ws, {row: 18})
            row += 1

    # Product summary totals
    TR = row + 1
    mc(ws, TR, 1, TR, 2, "PRODUCT TOTALS (all months)", DK_GRN, WHITE, 10, True)
    rh(ws, {TR: 22})
    for j, product in enumerate(PRODUCTS):
        r = TR + 1 + j
        ws.cell(r, 1).value = "All months"; ws.cell(r, 1).fill = fl(PL_GRN)
        ws.cell(r, 1).font = fn(9, False, MID_GRY); ws.cell(r, 1).border = bd()
        ws.cell(r, 1).alignment = al("center")
        ws.cell(r, 2).value = product; ws.cell(r, 2).fill = fl(VP_GRN)
        ws.cell(r, 2).font = fn(10, True, DK_GRN); ws.cell(r, 2).border = bd()
        ws.cell(r, 2).alignment = al("left")
        for col, ftype in [(3,'sum'),(5,'sum'),(6,'sum'),(7,'sum'),(8,'sum'),
                            (9,'sum'),(10,'sum'),(11,'avg'),(12,'avg')]:
            cl = get_column_letter(col)
            prod_col = "B"
            if ftype == 'sum':
                ws.cell(r, col).value = f'=SUMIF({prod_col}{DS}:{prod_col}{row-1},B{r},{cl}{DS}:{cl}{row-1})'
            else:
                ws.cell(r, col).value = f'=IFERROR(AVERAGEIF({prod_col}{DS}:{prod_col}{row-1},B{r},{cl}{DS}:{cl}{row-1}),"–")'
            ws.cell(r, col).fill = fl(VP_GRN); ws.cell(r, col).font = fn(10, True, DK_GRN)
            ws.cell(r, col).border = bd(); ws.cell(r, col).alignment = al("center")
        rh(ws, {r: 20})

    dv_list(ws, f"M{DS}:M{row-1}", '"Yes,No,Pending"')

    w(ws, {"A":11,"B":28,"C":17,"D":18,"E":12,"F":14,"G":13,"H":15,
           "I":13,"J":13,"K":14,"M":12,"N":30})


# ══════════════════════════════════════════════════════════════════════════════
#  SHEET 9 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
def build_dashboard(ws, wb):
    ws.sheet_view.showGridLines = False
    ws.tab_color = GOLD
    TOTAL_COLS = 18

    # ── Title banner
    mc(ws, 1, 1, 1, TOTAL_COLS,
       "   INVEGROW  |  Foods Line  —  Marketing Performance Dashboard",
       DK_GRN, WHITE, 18, True, "left")
    mc(ws, 2, 1, 2, TOTAL_COLS,
       "Real-time marketing metrics overview. All figures pull automatically from data entry sheets.",
       PL_GRN, DARK, 10, False, "left")
    mc(ws, 3, 1, 3, 8,
       "   REPORTING MONTH:", CHARCOAL, WHITE, 11, True, "left")
    rh(ws, {1: 50, 2: 22, 3: 30, 4: 18})

    # Month selector — cell I3 (col 9)
    ws.cell(3, 9).value = MONTHS[0]
    ws.cell(3, 9).fill  = fl(PL_GOLD)
    ws.cell(3, 9).font  = fn(13, True, GOLD)
    ws.cell(3, 9).alignment = al("center", "center")
    ws.cell(3, 9).border = bd_outer()
    dv_list(ws, "I3", '"' + ','.join(MONTHS) + '"')

    mc(ws, 3, 10, 3, TOTAL_COLS,
       "◀  Select the month from dropdown to update all KPIs",
       PL_GOLD, DARK, 9, False, "left")

    # ── KPI section helper
    def kpi_card(row, col, title, formula, nf, title_bg, val_bg, val_fg=DK_GRN, sub_text=""):
        mc(ws, row,   col, row,   col+2, title, title_bg, WHITE, 9, True, "center")
        cell = ws.cell(row+1, col)
        ws.merge_cells(start_row=row+1, start_column=col, end_row=row+2, end_column=col+2)
        cell.value  = formula
        cell.fill   = fl(val_bg)
        cell.font   = fn(18, True, val_fg)
        cell.alignment = al("center", "center")
        cell.border = bd()
        cell.number_format = nf
        for r2 in range(row+1, row+3):
            for c2 in range(col, col+3):
                ws.cell(r2, c2).fill = fl(val_bg)
                ws.cell(r2, c2).border = bd()
        if sub_text:
            mc(ws, row+3, col, row+3, col+2, sub_text, val_bg, MID_GRY, 8, False, "center")
        rh(ws, {row: 18, row+1: 28, row+2: 0, row+3: 14})

    # ── Section: Website (row 5)
    mc(ws, 5, 1, 5, 18, "  🌐  WEBSITE & DIGITAL  —  data from Digital Analytics sheet", MD_GRN, WHITE, 11, True, "left")
    rh(ws, {5: 24})

    # Formulas referencing Digital_Analytics (sheet name)
    DA = "Digital_Analytics"
    SM = "Social_Media"
    PA = "Paid_Ads"
    EM = "Email_Mktg"
    CP = "Content_PR"
    CT = "Campaigns"
    PM = "Product_Mktg"
    SEL = "$I$3"   # selected month cell

    kpi_card(6,  1,  "SESSIONS",
             f"=IFERROR(SUMIF('{DA}'!$A:$A,{SEL},'{DA}'!$B:$B),\"–\")",
             '#,##0', DK_GRN, VP_GRN, DK_GRN, "Monthly website sessions")
    kpi_card(6,  5,  "UNIQUE VISITORS",
             f"=IFERROR(SUMIF('{DA}'!$A:$A,{SEL},'{DA}'!$C:$C),\"–\")",
             '#,##0', DK_GRN, VP_GRN, DK_GRN, "Unique users to site")
    kpi_card(6,  9,  "BOUNCE RATE %",
             f"=IFERROR(AVERAGEIF('{DA}'!$A:$A,{SEL},'{DA}'!$E:$E),\"–\")",
             '0.0', TEAL, PL_TEAL, TEAL, "Lower is better")
    kpi_card(6,  13, "CONVERSION RATE %",
             f"=IFERROR(AVERAGEIF('{DA}'!$A:$A,{SEL},'{DA}'!$M:$M),\"–\")",
             '0.00', GOLD, PL_GOLD, GOLD, "Goal completions / sessions")

    # ── Section: Social (row 11)
    mc(ws, 11, 1, 11, 18, "  📱  SOCIAL MEDIA  —  data from Social Media sheet", TEAL, WHITE, 11, True, "left")
    rh(ws, {11: 24})

    kpi_card(12, 1,  "TOTAL FOLLOWERS",
             f"=IFERROR(SUMIF('{SM}'!$A:$A,{SEL},'{SM}'!$C:$C),\"–\")",
             '#,##0', TEAL, PL_TEAL, TEAL, "All platforms combined")
    kpi_card(12, 5,  "NEW FOLLOWERS",
             f"=IFERROR(SUMIF('{SM}'!$A:$A,{SEL},'{SM}'!$D:$D),\"–\")",
             '#,##0', TEAL, PL_TEAL, TEAL, "Net new this month")
    kpi_card(12, 9,  "TOTAL IMPRESSIONS",
             f"=IFERROR(SUMIF('{SM}'!$A:$A,{SEL},'{SM}'!$F:$F),\"–\")",
             '#,##0', TEAL, PL_TEAL, TEAL, "All platforms")
    kpi_card(12, 13, "AVG ENGAGEMENT RATE %",
             f"=IFERROR(AVERAGEIF('{SM}'!$A:$A,{SEL},'{SM}'!$M:$M),\"–\")",
             '0.00', TEAL, PL_TEAL, TEAL, "Across platforms")

    # ── Section: Paid Advertising (row 17)
    mc(ws, 17, 1, 17, 18, "  💰  PAID ADVERTISING  —  data from Paid Ads sheet", BLUE, WHITE, 11, True, "left")
    rh(ws, {17: 24})

    kpi_card(18, 1,  "TOTAL AD SPEND (R)",
             f"=IFERROR(SUMIF('{PA}'!$A:$A,{SEL},'{PA}'!$G:$G),\"–\")",
             'R#,##0.00', BLUE, PL_BLUE, BLUE, "Actual spend this month")
    kpi_card(18, 5,  "PAID IMPRESSIONS",
             f"=IFERROR(SUMIF('{PA}'!$A:$A,{SEL},'{PA}'!$I:$I),\"–\")",
             '#,##0', BLUE, PL_BLUE, BLUE, "All paid campaigns")
    kpi_card(18, 9,  "TOTAL PAID CLICKS",
             f"=IFERROR(SUMIF('{PA}'!$A:$A,{SEL},'{PA}'!$J:$J),\"–\")",
             '#,##0', BLUE, PL_BLUE, BLUE, "All paid campaigns")
    kpi_card(18, 13, "PAID CONVERSIONS",
             f"=IFERROR(SUMIF('{PA}'!$A:$A,{SEL},'{PA}'!$N:$N),\"–\")",
             '#,##0', BLUE, PL_BLUE, BLUE, "Goal completions via ads")

    # ── Section: Email (row 23)
    mc(ws, 23, 1, 23, 18, "  📧  EMAIL MARKETING  —  data from Email Mktg sheet", ORANGE, WHITE, 11, True, "left")
    rh(ws, {23: 24})

    kpi_card(24, 1,  "EMAILS SENT",
             f"=IFERROR(SUMIF('{EM}'!$A:$A,{SEL},'{EM}'!$D:$D),\"–\")",
             '#,##0', ORANGE, PL_ORG, ORANGE, "Total sends this month")
    kpi_card(24, 5,  "AVG OPEN RATE %",
             f"=IFERROR(AVERAGEIF('{EM}'!$A:$A,{SEL},'{EM}'!$H:$H),\"–\")",
             '0.00', ORANGE, PL_ORG, ORANGE, "Benchmark: 20–25%")
    kpi_card(24, 9,  "AVG CTR %",
             f"=IFERROR(AVERAGEIF('{EM}'!$A:$A,{SEL},'{EM}'!$J:$J),\"–\")",
             '0.00', ORANGE, PL_ORG, ORANGE, "Benchmark: 2–5%")
    kpi_card(24, 13, "AVG UNSUB RATE %",
             f"=IFERROR(AVERAGEIF('{EM}'!$A:$A,{SEL},'{EM}'!$L:$L),\"–\")",
             '0.00', ORANGE, PL_ORG, ORANGE, "Benchmark: < 0.5%")

    # ── Section: Content & PR (row 29)
    mc(ws, 29, 1, 29, 18, "  ✍  CONTENT & PR  —  data from Content & PR sheet", DK_GRN, WHITE, 11, True, "left")
    rh(ws, {29: 24})

    kpi_card(30, 1,  "CONTENT PIECES",
             f"=IFERROR(SUMIF('{CP}'!$A:$A,{SEL},'{CP}'!$B:$B)+SUMIF('{CP}'!$A:$A,{SEL},'{CP}'!$C:$C)+SUMIF('{CP}'!$A:$A,{SEL},'{CP}'!$I:$I),\"–\")",
             '#,##0', DK_GRN, VP_GRN, DK_GRN, "Blogs + social + videos")
    kpi_card(30, 5,  "PR MENTIONS",
             f"=IFERROR(SUMIF('{CP}'!$A:$A,{SEL},'{CP}'!$E:$E),\"–\")",
             '#,##0', DK_GRN, VP_GRN, DK_GRN, "Earned media coverage")
    kpi_card(30, 9,  "EARNED MEDIA VALUE (R)",
             f"=IFERROR(SUMIF('{CP}'!$A:$A,{SEL},'{CP}'!$F:$F),\"–\")",
             'R#,##0.00', DK_GRN, VP_GRN, DK_GRN, "Equivalent paid value")
    kpi_card(30, 13, "INFLUENCER REACH",
             f"=IFERROR(SUMIF('{CP}'!$A:$A,{SEL},'{CP}'!$H:$H),\"–\")",
             '#,##0', DK_GRN, VP_GRN, DK_GRN, "Combined influencer reach")

    # ── Section: Campaigns (row 35)
    mc(ws, 35, 1, 35, 18, "  🎯  CAMPAIGNS  —  data from Campaign Tracker sheet", CHARCOAL, WHITE, 11, True, "left")
    rh(ws, {35: 24})

    kpi_card(36, 1,  "ACTIVE CAMPAIGNS",
             f'=IFERROR(COUNTIF(\'{CT}\'!$F:$F,"Active"),"–")',
             '#,##0', CHARCOAL, LT_GRY, DARK, "Currently running")
    kpi_card(36, 5,  "TOTAL BUDGET (R)\nALL CAMPAIGNS",
             f"=IFERROR(SUM('{CT}'!$G:$G),\"–\")",
             'R#,##0.00', CHARCOAL, LT_GRY, DARK, "All campaigns total")
    kpi_card(36, 9,  "SPEND TO DATE (R)",
             f"=IFERROR(SUM('{CT}'!$H:$H),\"–\")",
             'R#,##0.00', CHARCOAL, LT_GRY, DARK, "All campaigns total")
    kpi_card(36, 13, "TOTAL LEADS /\nCONVERSIONS",
             f"=IFERROR(SUM('{CT}'!$P:$P),\"–\")",
             '#,##0', CHARCOAL, LT_GRY, DARK, "All campaigns total")

    # ── Section: Charts (row 43)
    mc(ws, 43, 1, 43, TOTAL_COLS,
       "  📈  VISUAL TRENDS  —  Charts auto-populate as you enter data into the input sheets",
       CHARCOAL, WHITE, 11, True, "left")
    rh(ws, {43: 26, 42: 10})

    # ── Hidden chart-data staging area (rows 45-72, cols A-R)
    #    Pulls 12-month rolling data for each chart so charts stay clean
    # Chart 1 data: Monthly sessions (Digital_Analytics)
    mc(ws, 45, 1, 45, 12, "Chart Data — Website Sessions (12 months)", LT_GRY, DARK, 8, False)
    rh(ws, {45: 14})
    for i, m in enumerate(MONTHS[:12]):
        col = i + 1
        ws.cell(46, col).value = m
        ws.cell(46, col).font  = fn(8, False, MID_GRY)
        ws.cell(47, col).value = (
            f"=IFERROR(SUMIF(Digital_Analytics!$A:$A,\"{m}\",Digital_Analytics!$B:$B),0)"
        )
        ws.cell(47, col).number_format = '#,##0'
        ws.cell(47, col).font = fn(8)
    rh(ws, {46: 12, 47: 12})

    # Chart 2 data: Social followers by platform (latest full month)
    mc(ws, 49, 1, 49, 8, "Chart Data — Social Followers by Platform", LT_GRY, DARK, 8, False)
    rh(ws, {49: 14})
    for i, plat in enumerate(PLATFORMS):
        ws.cell(50, i+1).value = plat
        ws.cell(50, i+1).font  = fn(8, False, MID_GRY)
        ws.cell(51, i+1).value = (
            f"=IFERROR(SUMIF(Social_Media!$B:$B,\"{plat}\",Social_Media!$C:$C),0)"
        )
        ws.cell(51, i+1).number_format = '#,##0'
        ws.cell(51, i+1).font = fn(8)
    rh(ws, {50: 12, 51: 12})

    # Chart 3 data: Monthly ad spend (Paid_Ads)
    mc(ws, 53, 1, 53, 12, "Chart Data — Monthly Ad Spend (R)", LT_GRY, DARK, 8, False)
    rh(ws, {53: 14})
    for i, m in enumerate(MONTHS[:12]):
        col = i + 1
        ws.cell(54, col).value = m
        ws.cell(54, col).font  = fn(8, False, MID_GRY)
        ws.cell(55, col).value = (
            f"=IFERROR(SUMIF(Paid_Ads!$A:$A,\"{m}\",Paid_Ads!$G:$G),0)"
        )
        ws.cell(55, col).number_format = 'R#,##0'
        ws.cell(55, col).font = fn(8)
    rh(ws, {54: 12, 55: 12})

    # Chart 4 data: Traffic source mix
    mc(ws, 57, 1, 57, 6, "Chart Data — Traffic Source Mix (all months)", LT_GRY, DARK, 8, False)
    rh(ws, {57: 14})
    sources = [("Organic","G"),("Paid","H"),("Social","I"),("Direct","J"),("Other","K")]
    for i, (label, col_letter) in enumerate(sources):
        ws.cell(58, i+1).value = label
        ws.cell(58, i+1).font  = fn(8, False, MID_GRY)
        ws.cell(59, i+1).value = f"=IFERROR(SUM(Digital_Analytics!${col_letter}:${col_letter}),0)"
        ws.cell(59, i+1).number_format = '#,##0'
        ws.cell(59, i+1).font = fn(8)
    rh(ws, {58: 12, 59: 12})

    # ── Build and place charts
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.series import SeriesLabel

    # Chart 1: Sessions trend (LineChart)
    lc = LineChart()
    lc.title        = "Monthly Website Sessions"
    lc.style        = 10
    lc.y_axis.title = "Sessions"
    lc.height       = 12
    lc.width        = 22
    lc.grouping     = "standard"
    data_ref = Reference(ws, min_col=1, max_col=12, min_row=47, max_row=47)
    cats_ref = Reference(ws, min_col=1, max_col=12, min_row=46, max_row=46)
    lc.add_data(data_ref)
    lc.set_categories(cats_ref)
    lc.series[0].title       = SeriesLabel(v="Sessions")
    lc.series[0].graphicalProperties.line.solidFill = MD_GRN
    lc.series[0].graphicalProperties.line.width     = 25000
    lc.series[0].smooth = True
    ws.add_chart(lc, "B61")

    # Chart 2: Social followers by platform (BarChart)
    bc1 = BarChart()
    bc1.title        = "Total Followers by Platform"
    bc1.style        = 10
    bc1.type         = "col"
    bc1.y_axis.title = "Followers"
    bc1.height       = 12
    bc1.width        = 22
    data2 = Reference(ws, min_col=1, max_col=6, min_row=51, max_row=51)
    cats2 = Reference(ws, min_col=1, max_col=6, min_row=50, max_row=50)
    bc1.add_data(data2)
    bc1.set_categories(cats2)
    bc1.series[0].title = SeriesLabel(v="Followers")
    bc1.series[0].graphicalProperties.solidFill = TEAL
    ws.add_chart(bc1, "J61")

    # Chart 3: Monthly ad spend (BarChart)
    bc2 = BarChart()
    bc2.title        = "Monthly Ad Spend (R)"
    bc2.style        = 10
    bc2.type         = "col"
    bc2.y_axis.title = "Spend (R)"
    bc2.height       = 12
    bc2.width        = 22
    data3 = Reference(ws, min_col=1, max_col=12, min_row=55, max_row=55)
    cats3 = Reference(ws, min_col=1, max_col=12, min_row=54, max_row=54)
    bc2.add_data(data3)
    bc2.set_categories(cats3)
    bc2.series[0].title = SeriesLabel(v="Ad Spend (R)")
    bc2.series[0].graphicalProperties.solidFill = BLUE
    ws.add_chart(bc2, "B79")

    # Chart 4: Traffic source mix (BarChart horizontal)
    bc3 = BarChart()
    bc3.title        = "Traffic Source Mix (All-Time)"
    bc3.style        = 10
    bc3.type         = "bar"
    bc3.y_axis.title = "Channel"
    bc3.x_axis.title = "Sessions"
    bc3.height       = 12
    bc3.width        = 22
    data4 = Reference(ws, min_col=1, max_col=5, min_row=59, max_row=59)
    cats4 = Reference(ws, min_col=1, max_col=5, min_row=58, max_row=58)
    bc3.add_data(data4)
    bc3.set_categories(cats4)
    bc3.series[0].title = SeriesLabel(v="Sessions")
    bc3.series[0].graphicalProperties.solidFill = DK_GRN
    ws.add_chart(bc3, "J79")

    # ── Footer row
    mc(ws, 97, 1, 97, TOTAL_COLS,
       f"  Invegrow Foods Line  |  Marketing Dashboard  |  Auto-updates from data entry sheets  |  v2.0  |  2025",
       DK_GRN, "A5D6A7", 9, False, "left")
    rh(ws, {97: 16})

    # Column widths for dashboard
    for col_letter in [get_column_letter(c) for c in range(1, TOTAL_COLS+1)]:
        ws.column_dimensions[col_letter].width = 12
    ws.column_dimensions["A"].width = 3  # spacer


# ══════════════════════════════════════════════════════════════════════════════
#  SAMPLE / DEMO DATA SEEDER
#  Populates Jan-25 → Jun-25 so charts render immediately on first open.
#  Replace with real numbers — formulas update automatically.
# ══════════════════════════════════════════════════════════════════════════════
def seed_demo_data(wb):
    print("Seeding demo data …")

    # ── Digital Analytics (DS=7, cols B-M for data)
    ws_da = wb["Digital_Analytics"]
    # columns: sessions, unique, pageviews, bounce%, dur_min, organic, paid, social, direct, other, goals, conv%
    da_rows = [
        ("Jan-25", 4200,  3100,  12600, 58.3, 2.4,  2100, 800,  600,  500,  200, 200, 4.8),
        ("Feb-25", 4850,  3600,  14550, 55.1, 2.7,  2400, 950,  720,  580,  200, 242, 5.0),
        ("Mar-25", 5300,  4000,  15900, 53.8, 2.9,  2650, 1100, 820,  580,  150, 280, 5.3),
        ("Apr-25", 6100,  4600,  18300, 51.2, 3.1,  3050, 1400, 950,  550,  150, 340, 5.6),
        ("May-25", 7200,  5400,  21600, 49.5, 3.3,  3600, 1700, 1150, 580,  170, 420, 5.8),
        ("Jun-25", 8100,  6100,  24300, 47.8, 3.6,  4050, 1900, 1350, 620,  180, 500, 6.2),
    ]
    for i, row_data in enumerate(da_rows):
        r = 7 + i   # DS = 7
        for j, val in enumerate(row_data[1:]):   # skip month (pre-filled)
            ws_da.cell(r, j + 2).value = val     # col B onwards

    # ── Social Media (DS=6, platforms cycle: Insta,FB,TikTok,YT,LinkedIn,Twitter)
    ws_sm = wb["Social_Media"]
    # structure: {platform: [(month, followers, new_foll, posts, impressions, reach, engmt, likes, comments, shares, saves)]}
    sm_data = {
        "Instagram": [
            ("Jan-25", 12400, 320,  18, 185000,   95000,   9500,   7600,  950,  600,  350),
            ("Feb-25", 12900, 500,  22, 210000,  108000,  11000,   8800, 1050,  750,  400),
            ("Mar-25", 13600, 700,  20, 245000,  125000,  13000,  10400, 1200,  900,  500),
            ("Apr-25", 14500, 900,  24, 290000,  148000,  15500,  12400, 1400, 1050,  650),
            ("May-25", 15700, 1200, 26, 340000,  175000,  18500,  14800, 1700, 1350,  650),
            ("Jun-25", 17200, 1500, 28, 398000,  204000,  22000,  17600, 2000, 1600,  800),
        ],
        "Facebook": [
            ("Jan-25",  8200, 120,  12,  98000,  52000,  2600,  1800,  420,  280,  100),
            ("Feb-25",  8400, 200,  14, 112000,  59000,  3000,  2100,  480,  320,  100),
            ("Mar-25",  8700, 300,  12, 128000,  67000,  3500,  2450,  560,  380,  110),
            ("Apr-25",  9100, 400,  16, 148000,  78000,  4100,  2870,  660,  450,  120),
            ("May-25",  9650, 550,  18, 172000,  90000,  4900,  3430,  780,  540,  150),
            ("Jun-25", 10300, 650,  20, 198000, 104000,  5700,  3990,  900,  630,  180),
        ],
        "TikTok": [
            ("Jan-25",  5600,  800, 20,  320000,  210000,  21000,  17500, 2100, 1050,  350),
            ("Feb-25",  7200, 1600, 25,  490000,  320000,  32000,  26700, 3200, 1600,  500),
            ("Mar-25",  9500, 2300, 22,  710000,  465000,  46500,  38800, 4650, 2325,  725),
            ("Apr-25", 12800, 3300, 28,  980000,  642000,  64200,  53500, 6420, 3210, 1070),
            ("May-25", 17200, 4400, 30, 1350000,  885000,  88500,  73800, 8850, 4425, 1425),
            ("Jun-25", 23100, 5900, 32, 1820000, 1195000, 119500,  99600,11950, 5975, 1975),
        ],
        "YouTube": [
            ("Jan-25", 2100,  80, 4,  48000,  32000, 1920, 1440, 240, 192, 48),
            ("Feb-25", 2250, 150, 5,  62000,  41000, 2460, 1845, 308, 246, 61),
            ("Mar-25", 2450, 200, 4,  78000,  52000, 3120, 2340, 390, 312, 78),
            ("Apr-25", 2700, 250, 6,  96000,  64000, 3840, 2880, 480, 384, 96),
            ("May-25", 3000, 300, 5, 118000,  79000, 4740, 3555, 592, 474,119),
            ("Jun-25", 3380, 380, 6, 142000,  95000, 5700, 4275, 712, 570,143),
        ],
        "LinkedIn": [
            ("Jan-25", 1850,  60,  8, 28000, 18000,  900,  630, 180,  72, 18),
            ("Feb-25", 1950, 100, 10, 34000, 22000, 1100,  770, 220,  88, 22),
            ("Mar-25", 2100, 150,  9, 41000, 27000, 1350,  945, 270, 108, 27),
            ("Apr-25", 2300, 200, 12, 49000, 32000, 1600, 1120, 320, 128, 32),
            ("May-25", 2550, 250, 11, 58000, 38000, 1900, 1330, 380, 152, 38),
            ("Jun-25", 2850, 300, 14, 68000, 44000, 2200, 1540, 440, 176, 44),
        ],
        "Twitter/X": [
            ("Jan-25", 3200, 100, 15, 42000, 28000, 1400,  980, 280, 112, 28),
            ("Feb-25", 3380, 180, 18, 52000, 35000, 1750, 1225, 350, 140, 35),
            ("Mar-25", 3600, 220, 16, 63000, 42000, 2100, 1470, 420, 168, 42),
            ("Apr-25", 3860, 260, 20, 76000, 51000, 2550, 1785, 510, 204, 51),
            ("May-25", 4180, 320, 22, 91000, 61000, 3050, 2135, 610, 244, 61),
            ("Jun-25", 4550, 370, 24,108000, 72000, 3600, 2520, 720, 288, 72),
        ],
    }
    # Social_Media rows cycle: for each month, all 6 platforms in order
    row = 6   # DS = 6
    for month in MONTHS:
        for platform in PLATFORMS:
            if month in [d[0] for d in sm_data[platform]]:
                vals = next(d for d in sm_data[platform] if d[0] == month)
                for j, val in enumerate(vals[1:]):
                    ws_sm.cell(row, j + 3).value = val   # col C onwards
            row += 1

    # ── Paid Ads (DS=6, cols A-N for data)
    ws_pa = wb["Paid_Ads"]
    # cols: month, name, platform, product, objective, budget, spend, _, impressions, clicks, _,_,_, convs
    paid_rows = [
        ("Jan-25","Hemp Protein Launch – Meta",    "Meta (Facebook/Instagram)","Hemp Protein Powder 1kg",   "Product Launch",   8000, 7850, 1200000,3600,180),
        ("Jan-25","Hemp Seed Oil – Google Search",  "Google Ads",               "Hemp Seed Oil 500ml",       "Website Traffic",  5000, 4920,   85000, 2800,142),
        ("Feb-25","Hemp Granola – TikTok",          "TikTok Ads",               "Hemp Granola 400g",         "Brand Awareness",  6000, 5980, 2100000,4200, 96),
        ("Feb-25","Hemp Omega – Google Search",     "Google Ads",               "Hemp Omega Capsules 60s",   "Lead Generation",  4500, 4480,   72000, 2160,108),
        ("Mar-25","Spring Range – Meta Retarget",   "Meta (Facebook/Instagram)","Hemp Hearts 500g",          "Retargeting",      7000, 6950,  980000,3920,235),
        ("Mar-25","Hemp Energy Bar – TikTok",       "TikTok Ads",               "Hemp Energy Bar (Box 12)",  "Engagement",       5500, 5490, 1850000,3700, 74),
        ("Apr-25","Product Range – Meta Brand",     "Meta (Facebook/Instagram)","Hemp Superfood Blend 300g", "Brand Awareness",  9000, 8920, 1650000,4950,198),
        ("Apr-25","Hemp Butter – Google Display",   "Google Ads",               "Hemp Seed Butter 400g",     "Website Traffic",  3500, 3480,  420000,1260, 63),
        ("May-25","Winter Launch – Meta",           "Meta (Facebook/Instagram)","Hemp Flour 1kg",            "Product Launch",  10000, 9850, 1920000,5760,288),
        ("May-25","Hemp Honey – Influencer",        "Influencer",               "Hemp Infused Honey 250g",   "Reach",           12000,11800, 3200000,   0,320),
        ("Jun-25","Mid-Year Sale – Meta",           "Meta (Facebook/Instagram)","Hemp Seed Oil 500ml",       "Seasonal Promo",  15000,14850, 2800000,8400,504),
        ("Jun-25","Hemp Range – Google Brand",      "Google Ads",               "Hemp Protein Powder 1kg",   "Brand Awareness",  6000, 5980,   95000,3800,190),
    ]
    for i, (month, name, plat, prod, obj, budget, spend, imp, clicks, convs) in enumerate(paid_rows):
        r = 6 + i
        ws_pa.cell(r, 1).value  = month
        ws_pa.cell(r, 2).value  = name
        ws_pa.cell(r, 3).value  = plat
        ws_pa.cell(r, 4).value  = prod
        ws_pa.cell(r, 5).value  = obj
        ws_pa.cell(r, 6).value  = budget
        ws_pa.cell(r, 7).value  = spend
        ws_pa.cell(r, 9).value  = imp
        ws_pa.cell(r, 10).value = clicks
        ws_pa.cell(r, 14).value = convs

    # ── Email Marketing (DS=6)
    ws_em = wb["Email_Mktg"]
    email_rows = [
        ("Jan-25","January Newsletter",          "Full List",      12400, 12200,  980,  4270,  2135,  182),
        ("Jan-25","Hemp Protein Launch Promo",   "Health & Fitness",4800,  4730,  460,  2081,  1040,   96),
        ("Feb-25","February Newsletter",         "Full List",      12800, 12600, 1050,  4410,  2205,  176),
        ("Feb-25","Hemp Granola Launch",         "Food & Recipe",   5200,  5130,  530,  2309,  1155,  103),
        ("Mar-25","March Newsletter",            "Full List",      13100, 12900, 1110,  4515,  2258,  171),
        ("Mar-25","Spring Sale Campaign",        "All Customers",  13100, 12890, 2060,  5800,  2900,  232),
        ("Apr-25","April Newsletter",            "Full List",      13500, 13280, 1160,  4781,  2391,  186),
        ("Apr-25","New Product Alert",           "Subscribers",     6400,  6310,  840,  2839,  1420,  119),
        ("May-25","May Newsletter",              "Full List",      13900, 13670, 1230,  5054,  2527,  192),
        ("May-25","Mother's Day Promo",          "Gift Buyers",     8200,  8080,  1730, 3636,  1818,  145),
        ("Jun-25","June Newsletter",             "Full List",      14200, 13980, 1330,  5173,  2587,  198),
        ("Jun-25","Mid-Year Sale Launch",        "All Customers",  14200, 13960, 3630,  6415,  3208,  257),
    ]
    for i, (month, name, seg, sent, deliv, opens, clicks, unsubs, convs) in enumerate(email_rows):
        r = 6 + i
        ws_em.cell(r, 1).value  = month
        ws_em.cell(r, 2).value  = name
        ws_em.cell(r, 3).value  = seg
        ws_em.cell(r, 4).value  = sent
        ws_em.cell(r, 5).value  = deliv
        ws_em.cell(r, 7).value  = opens
        ws_em.cell(r, 9).value  = clicks
        ws_em.cell(r, 11).value = unsubs
        ws_em.cell(r, 13).value = convs

    # ── Content & PR (DS=6, months pre-filled)
    ws_cp = wb["Content_PR"]
    # cols: blog, social_posts, press_releases, pr_mentions, emv, influencer_collabs, inf_reach, videos, podcast, avg_rating
    cp_rows = [
        ("Jan-25",  2, 45, 1,  4,  28000, 2,  85000,  1, 0, 4.3),
        ("Feb-25",  3, 52, 1,  6,  42000, 3, 120000,  2, 1, 4.4),
        ("Mar-25",  4, 60, 2,  9,  65000, 4, 165000,  3, 1, 4.5),
        ("Apr-25",  3, 68, 1, 11,  78000, 5, 210000,  2, 2, 4.5),
        ("May-25",  5, 75, 2, 14,  95000, 6, 280000,  4, 2, 4.6),
        ("Jun-25",  4, 80, 3, 18, 125000, 7, 360000,  3, 3, 4.7),
    ]
    for i, (month, blog, social, pr_rel, pr_men, emv, inf_col, inf_reach, vids, pod, rating) in enumerate(cp_rows):
        r = 6 + i
        ws_cp.cell(r, 2).value  = blog
        ws_cp.cell(r, 3).value  = social
        ws_cp.cell(r, 4).value  = pr_rel
        ws_cp.cell(r, 5).value  = pr_men
        ws_cp.cell(r, 6).value  = emv
        ws_cp.cell(r, 7).value  = inf_col
        ws_cp.cell(r, 8).value  = inf_reach
        ws_cp.cell(r, 9).value  = vids
        ws_cp.cell(r, 10).value = pod
        ws_cp.cell(r, 12).value = rating

    # ── Campaigns (DS=6)
    ws_ct = wb["Campaigns"]
    from datetime import date
    campaigns = [
        ("Hemp Protein Powder Launch",  "Product Launch",  "Hemp Protein Powder 1kg",  date(2025,1,6),  date(2025,2,28), "Completed", 20000, 18770, "Meta, Google, Email, Influencer",  "Health-conscious 25–45",        "Fuel your day with plant protein",        2400000, 1200000,  85000, 322),
        ("Spring Awareness Campaign",   "Brand Awareness", "Hemp Hearts 500g",          date(2025,3,1),  date(2025,3,31), "Completed", 12500, 12440, "Meta, TikTok, PR",                 "Foodies & wellness seekers",    "Nature's most complete seed",              980000,  580000,  35000, 235),
        ("Hemp Granola TikTok Drive",   "Engagement",      "Hemp Granola 400g",         date(2025,2,10), date(2025,3,15), "Completed",  8000,  7960, "TikTok Ads, Social Organic",       "Gen Z & Millennials",           "Crunch into something better",            2100000, 1400000, 120000,  96),
        ("Mid-Year Sale",               "Seasonal Promo",  "Hemp Seed Oil 500ml",       date(2025,6,1),  date(2025,6,30), "Completed", 18000, 17850, "Meta, Google, Email",              "Existing customers + lookalike","Stock up before it sells out",            2800000, 1800000,  92000, 504),
        ("Hemp Honey Influencer Drive", "Reach",           "Hemp Infused Honey 250g",   date(2025,5,1),  date(2025,5,31), "Completed", 15000, 14800, "Influencer, Social Organic",       "Foodie lifestyle creators",     "Nature's sweetest superfood",             3200000, 2100000, 180000, 320),
        ("Q3 Brand Growth Campaign",    "Brand Awareness", "Hemp Seed Oil 500ml",       date(2025,7,1),  date(2025,9,30), "Active",    35000, 12400, "Meta, TikTok, Google, Email, PR",  "25–45 health-conscious adults", "The original hemp superfood",                   0,       0,      0,   0),
    ]
    for i, (name, obj, prod, start, end, status, budget, spend, channels, audience, message, imp, reach, engmt, leads) in enumerate(campaigns):
        r = 6 + i
        ws_ct.cell(r,  1).value = name
        ws_ct.cell(r,  2).value = obj
        ws_ct.cell(r,  3).value = prod
        ws_ct.cell(r,  4).value = start
        ws_ct.cell(r,  5).value = end
        ws_ct.cell(r,  6).value = status
        ws_ct.cell(r,  7).value = budget
        ws_ct.cell(r,  8).value = spend
        ws_ct.cell(r, 10).value = channels
        ws_ct.cell(r, 11).value = audience
        ws_ct.cell(r, 12).value = message
        ws_ct.cell(r, 13).value = imp
        ws_ct.cell(r, 14).value = reach
        ws_ct.cell(r, 15).value = engmt
        ws_ct.cell(r, 16).value = leads

    print("  ✓ Demo data seeded (Jan-25 → Jun-25)")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default Sheet

    # Create sheets
    ws_settings  = wb.create_sheet("SETTINGS")
    ws_dash      = wb.create_sheet("DASHBOARD")
    ws_digital   = wb.create_sheet("Digital_Analytics")
    ws_social    = wb.create_sheet("Social_Media")
    ws_paid      = wb.create_sheet("Paid_Ads")
    ws_email     = wb.create_sheet("Email_Mktg")
    ws_content   = wb.create_sheet("Content_PR")
    ws_campaigns = wb.create_sheet("Campaigns")
    ws_products  = wb.create_sheet("Product_Mktg")

    print("Building SETTINGS …")
    build_settings(ws_settings)
    print("Building DASHBOARD …")
    build_dashboard(ws_dash, wb)
    print("Building Digital Analytics …")
    build_digital(ws_digital)
    print("Building Social Media …")
    build_social(ws_social)
    print("Building Paid Advertising …")
    build_paid(ws_paid)
    print("Building Email Marketing …")
    build_email(ws_email)
    print("Building Content & PR …")
    build_content(ws_content)
    print("Building Campaign Tracker …")
    build_campaigns(ws_campaigns)
    print("Building Product Marketing …")
    build_products(ws_products)

    # Seed demo data so charts render on first open
    seed_demo_data(wb)

    # Set DASHBOARD as the default active sheet
    wb.active = ws_dash

    # Tab colours
    ws_settings.tab_color  = "424242"
    ws_dash.tab_color      = DK_GRN
    ws_digital.tab_color   = MD_GRN
    ws_social.tab_color    = TEAL
    ws_paid.tab_color      = BLUE
    ws_email.tab_color     = ORANGE
    ws_content.tab_color   = DK_GRN
    ws_campaigns.tab_color = GOLD
    ws_products.tab_color  = LT_GRN

    out_path = "/home/user/WILLIAM-/Invegrow_Marketing_Dashboard.xlsx"
    wb.save(out_path)
    print(f"\n✅  Saved: {out_path}")

if __name__ == "__main__":
    main()
