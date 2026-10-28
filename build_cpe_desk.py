"""
CPE Analyst Desk Builder
Generates:
  - CPE_Desk.xlsx   (9 sheets)
  - CPE_Certificate_Template.docx
  - CPE_Email_Templates.docx
  - CPE_SOP.docx
"""

import random
import math
from datetime import date, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side,
    GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os

random.seed(42)
OUT_DIR = r"c:\Users\Piyush Pathak\Desktop\CPE"
os.makedirs(OUT_DIR, exist_ok=True)

# ─── Palette ────────────────────────────────────────────────────────────────
NAVY   = "1F3864"
BLUE   = "2E75B6"
LBLUE  = "BDD7EE"
GREEN  = "70AD47"
LGREEN = "E2EFDA"
AMBER  = "FFC000"
RED    = "FF0000"
LRED   = "FFE0E0"
GREY   = "F2F2F2"
WHITE  = "FFFFFF"
DKGREY = "595959"

def hdr_fill(hex_col):  return PatternFill("solid", fgColor=hex_col)
def hdr_font(bold=True, color=WHITE, size=11):
    return Font(bold=bold, color=color, size=size)
def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)
def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def left():   return Alignment(horizontal="left",   vertical="center", wrap_text=True)

# ─── Seed Data ───────────────────────────────────────────────────────────────
INSTRUCTORS = [
    "Alice Morrow","Ben Tanaka","Clara Singh","David Okonkwo","Elena Voss",
    "Frank Huang","Grace Patel","Henry Reyes","Irene Schultz","James Osei",
    "Karen Li","Leo Fernandez","Maya Brooks","Nate Johansson","Olivia Chen",
    "Paul Mbeki","Quinn Torres","Rachel Novak","Sam Adeyemi","Tina Williams",
]
REVIEWERS = ["CPE Desk - Priya","CPE Desk - Omar","CPE Desk - Sarah","CPE Desk - Luis"]
DELIVERY  = ["Group","Self-study"]
LEVELS    = ["Basic","Intermediate","Advanced"]
FIELDS    = [
    "Accounting","Auditing","Taxation","Finance","Ethics",
    "Management","Information Technology","Economics","Business Law",
    "Specialized Knowledge",
]
DURATIONS = [20,25,30,45,49,50,60,75,90,100,110,120,140,150,180,200]
COURSE_PREFIXES = [
    "Introduction to","Advanced","Practical","Mastering","Essentials of",
    "Deep Dive into","Workshop:","Seminar:","Foundations of","Current Issues in",
]
COURSE_TOPICS = [
    "GAAP Updates","Revenue Recognition","Lease Accounting","Tax Reform",
    "Cybersecurity Risk","Data Analytics for CPAs","Fraud Prevention",
    "Ethical Decision Making","Excel for Accountants","AI in Finance",
    "Estate Planning","International Tax","QuickBooks Pro","Auditing Standards",
    "Internal Controls","SOX Compliance","IFRS vs GAAP","Financial Modeling",
    "Blockchain in Accounting","Remote Audit Techniques",
]

def rand_course():
    return f"{random.choice(COURSE_PREFIXES)} {random.choice(COURSE_TOPICS)}"

def calc_min_obj(minutes):
    return math.ceil(minutes / 50 * 3)

def calc_credits_half(minutes):
    if minutes < 25: return 0
    return math.floor(minutes / 50 * 2) / 2  # FLOOR(min/50, 0.5)

# Build 100 requests with seeded errors
def build_requests():
    rows = []
    base_date = date(2026, 9, 1)
    used_ids = set()

    error_pool = (
        ["missing_objectives"] * 15 +
        ["no_attendance"]      * 10 +
        ["duration_mismatch"]  * 8  +
        ["missing_fos"]        * 5  +
        ["duplicate"]          * 4  +
        ["clean"]              * 58
    )
    random.shuffle(error_pool)

    dup_rows = []   # will be injected at the end for duplicates

    for i in range(100):
        req_id   = f"CPE-2026-{i+1:03d}"
        title    = rand_course()
        instr    = random.choice(INSTRUCTORS)
        sub_date = base_date + timedelta(days=random.randint(0, 57))
        method   = random.choice(DELIVERY)
        duration = random.choice(DURATIONS)
        fos      = random.choice(FIELDS)
        level    = random.choice(LEVELS)
        prereq   = random.choice(["Y","N"])
        reviewer = random.choice(REVIEWERS)
        sla_due  = sub_date + timedelta(days=5)
        att_checks = max(1, math.ceil(duration / 50 * 3))
        min_obj    = calc_min_obj(duration)
        obj_count  = min_obj  # start correct

        error = error_pool[i] if i < len(error_pool) else "clean"

        if error == "missing_objectives":
            obj_count = max(0, min_obj - random.randint(1, 3))
        elif error == "no_attendance":
            att_checks = 0
        elif error == "duration_mismatch":
            duration = duration + random.choice([-10, 10, 15, -15])
            duration = max(10, duration)
        elif error == "missing_fos":
            fos = ""
        elif error == "duplicate" and dup_rows:
            prev = random.choice(dup_rows)
            title = prev["title"]
            instr = prev["instr"]

        if method == "Self-study":
            att_checks = 0  # attendance not applicable for self-study
            rq = max(1, math.ceil(calc_credits_half(duration) * 3))
        else:
            rq = 0  # review questions for group = N/A

        # status based on error
        if sub_date + timedelta(days=5) < date(2026, 10, 28):
            if error in ("clean",):
                status = "Approved"
            elif error in ("missing_objectives","no_attendance","duration_mismatch"):
                status = "Needs Correction"
            elif error in ("missing_fos",):
                status = "Escalate"
            else:
                status = "Approved"
        else:
            status = "Pending"

        row = {
            "req_id":    req_id,
            "title":     title,
            "instr":     instr,
            "sub_date":  sub_date,
            "method":    method,
            "duration":  duration,
            "fos":       fos,
            "level":     level,
            "obj_count": obj_count,
            "att_checks":att_checks,
            "rq":        rq,
            "prereq":    prereq,
            "sla_due":   sla_due,
            "reviewer":  reviewer,
            "status":    status,
            "error":     error,
        }
        rows.append(row)
        if error != "duplicate":
            dup_rows.append(row)

    return rows

REQUESTS = build_requests()

# ════════════════════════════════════════════════════════════════════════════
# BUILD EXCEL
# ════════════════════════════════════════════════════════════════════════════
wb = Workbook()
wb.remove(wb.active)   # remove default sheet

# ── Helpers ──────────────────────────────────────────────────────────────────
def style_header_row(ws, row_num, fill_hex, font_color=WHITE, height=28):
    ws.row_dimensions[row_num].height = height
    for cell in ws[row_num]:
        cell.fill      = hdr_fill(fill_hex)
        cell.font      = hdr_font(color=font_color)
        cell.alignment = center()
        cell.border    = thin_border()

def auto_width(ws, extra=4):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                v = str(cell.value) if cell.value else ""
                max_len = max(max_len, len(v))
            except: pass
        ws.column_dimensions[col_letter].width = min(max_len + extra, 40)

def write_cell(ws, row, col, val, fill=None, font=None, align=None, border=True, num_fmt=None):
    c = ws.cell(row=row, column=col, value=val)
    if fill:    c.fill      = fill
    if font:    c.font      = font
    if align:   c.alignment = align
    else:       c.alignment = left()
    if border:  c.border    = thin_border()
    if num_fmt: c.number_format = num_fmt
    return c

# ══════════════════════════════════════════════════════════════════════
# SHEET 1: Requests
# ══════════════════════════════════════════════════════════════════════
ws1 = wb.create_sheet("Requests")
ws1.freeze_panes = "A2"

REQ_HEADERS = [
    "Request ID","Course Title","Instructor","Submitted Date",
    "Delivery Method","Duration (min)","Field of Study","Program Level",
    "Learning Objectives","Attendance Checks","Review Questions",
    "Prerequisites","SLA Due Date","Reviewer","Status",
]
for ci, h in enumerate(REQ_HEADERS, 1):
    write_cell(ws1, 1, ci, h,
               fill=hdr_fill(NAVY),
               font=hdr_font(color=WHITE),
               align=center())

# Data rows
for ri, r in enumerate(REQUESTS, 2):
    vals = [
        r["req_id"], r["title"], r["instr"],
        r["sub_date"], r["method"], r["duration"],
        r["fos"], r["level"], r["obj_count"],
        r["att_checks"], r["rq"], r["prereq"],
        r["sla_due"], r["reviewer"], r["status"],
    ]
    for ci, v in enumerate(vals, 1):
        c = write_cell(ws1, ri, ci, v)
        if isinstance(v, date): c.number_format = "DD-MMM-YYYY"

# Status colour fill
status_colors = {"Approved": LGREEN, "Needs Correction": "FFF2CC",
                 "Escalate": "FCE4D6", "Pending": GREY}
for ri in range(2, 102):
    status_val = ws1.cell(row=ri, column=15).value
    color = status_colors.get(str(status_val), WHITE)
    for ci in range(1, 16):
        ws1.cell(row=ri, column=ci).fill = hdr_fill(color)
        ws1.cell(row=ri, column=ci).border = thin_border()

# Freeze and filter
ws1.auto_filter.ref = ws1.dimensions
auto_width(ws1)

# ══════════════════════════════════════════════════════════════════════
# SHEET 2: Rules
# ══════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Rules")

rules_data = [
    # (Label, Parameter Name, Value, Unit, Reference, Note)
    ("NASBA-Style Rules – Editable Parameters","","","","",""),
    ("","","","","",""),
    ("LEARNING OBJECTIVES","","","","",""),
    ("Minimum objectives per 50 min","MIN_OBJ_PER_50",3,"per 50 min","IASA / NASBA","3 objectives per 50-minute credit hour"),
    ("Minimum duration to award any credit","MIN_DURATION_MIN",25,"minutes","NASBA","< 25 min = 0 credits"),
    ("","","","","",""),
    ("ATTENDANCE & ENGAGEMENT (Group)","","","","",""),
    ("Min attendance checks per 50-min hour (Group-live tech)","MIN_ATT_PER_50",3,"per 50 min","LearnSignal / NASBA","Three checks per credit hour"),
    ("","","","","",""),
    ("SELF-STUDY","","","","",""),
    ("Min review questions per credit hour (Self-study)","MIN_RQ_PER_CREDIT",3,"per credit","NASBA SS Standard","Verify against current NASBA self-study standard"),
    ("","","","","",""),
    ("CREDIT ROUNDING","","","","",""),
    ("Half-credit increment (Group programs)","CREDIT_INCREMENT",0.5,"credits","NASBA","FLOOR(min/50, 0.5)"),
    ("Fifth-credit increment (alt rounding)","CREDIT_INCREMENT_5TH",0.2,"credits","NASBA alt","ROUNDDOWN(min/50*5,0)/5"),
    ("","","","","",""),
    ("REQUIRED FIELDS","","","","",""),
    ("Field of Study required","FOS_REQUIRED","Y","—","NASBA","Must not be blank"),
    ("Program Level required","LEVEL_REQUIRED","Y","—","NASBA","Basic / Intermediate / Advanced"),
    ("Delivery Method required","METHOD_REQUIRED","Y","—","NASBA","Group or Self-study"),
    ("Prerequisites documented","PREREQ_REQUIRED","Y","—","NASBA","Y or N must be stated"),
    ("","","","","",""),
    ("SLA","","","","",""),
    ("SLA turnaround days","SLA_DAYS",5,"business days","Firm policy","Submitted date + 5 days"),
    ("Due-soon warning threshold","DUE_SOON_DAYS",1,"days","Firm policy","Amber flag within 1 day"),
]

RULES_HEADERS = ["Label","Parameter","Value","Unit","Source","Note"]
for ci, h in enumerate(RULES_HEADERS, 1):
    write_cell(ws2, 1, ci, h,
               fill=hdr_fill(NAVY), font=hdr_font(), align=center())

for ri, row in enumerate(rules_data, 2):
    lbl = row[0]
    is_section = lbl.isupper() and lbl != ""
    is_title   = ri == 2
    for ci, v in enumerate(row, 1):
        c = ws2.cell(row=ri, column=ci, value=v)
        c.border = thin_border()
        if is_section:
            c.fill = hdr_fill(BLUE)
            c.font = Font(bold=True, color=WHITE, size=10)
            c.alignment = center()
        elif ri == 2 and ci == 1:
            c.fill = hdr_fill(LBLUE)
            c.font = Font(bold=True, color=NAVY, size=12)
            c.alignment = left()
        else:
            c.alignment = left()
            if ci == 2:  # parameter name
                c.font = Font(bold=True, color=NAVY, size=10)
            if ci == 3:  # value – highlight yellow
                c.fill = hdr_fill("FFFF99")
                c.font = Font(bold=True, size=11)
                c.alignment = center()

auto_width(ws2)
ws2.column_dimensions["A"].width = 46
ws2.column_dimensions["F"].width = 42

# ══════════════════════════════════════════════════════════════════════
# SHEET 3: Review
# ══════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Review")
ws3.freeze_panes = "A2"

REV_HEADERS = [
    "Request ID","Course Title","Duration (min)","Delivery Method",
    "Learning Objectives","Min Obj Needed","Obj Check",
    "Attendance Checks","Min Att Needed","Att Check",
    "Field of Study","FOS Check","Program Level","Level Check",
    "Review Questions","RQ Check",
    "Credits (Half)","Credits (Fifth)",
    "SLA Flag","Decision","Reason",
]
for ci, h in enumerate(REV_HEADERS, 1):
    write_cell(ws3, 1, ci, h, fill=hdr_fill(BLUE), font=hdr_font(), align=center())

# Pull data from Requests sheet via formulas
for ri in range(2, 102):
    rn = ri  # same row index in Requests

    # Static refs from Requests
    ws3.cell(ri, 1).value = f"=Requests!A{rn}"
    ws3.cell(ri, 2).value = f"=Requests!B{rn}"
    ws3.cell(ri, 3).value = f"=Requests!F{rn}"
    ws3.cell(ri, 4).value = f"=Requests!E{rn}"
    ws3.cell(ri, 5).value = f"=Requests!I{rn}"   # learning objectives
    # Min obj needed
    ws3.cell(ri, 6).value = f'=ROUNDUP(C{rn}/50*Rules!$C$5,0)'
    # Obj check
    ws3.cell(ri, 7).value = f'=IF(E{rn}>=F{rn},"Pass","Fail")'
    # Attendance checks
    ws3.cell(ri, 8).value = f"=Requests!J{rn}"
    # Min att needed
    ws3.cell(ri, 9).value = f'=IF(D{rn}="Group",ROUNDUP(C{rn}/50*Rules!$C$9,0),0)'
    # Att check
    ws3.cell(ri, 10).value = f'=IF(D{rn}="Self-study","N/A",IF(H{rn}>=I{rn},"Pass","Fail"))'
    # FOS
    ws3.cell(ri, 11).value = f"=Requests!G{rn}"
    # FOS check
    ws3.cell(ri, 12).value = f'=IF(K{rn}<>"","Pass","Fail – Missing FOS")'
    # Level
    ws3.cell(ri, 13).value = f"=Requests!H{rn}"
    # Level check
    ws3.cell(ri, 14).value = f'=IF(OR(M{rn}="Basic",M{rn}="Intermediate",M{rn}="Advanced"),"Pass","Fail")'
    # Review questions
    ws3.cell(ri, 15).value = f"=Requests!K{rn}"
    # RQ check
    ws3.cell(ri, 16).value = (
        f'=IF(D{rn}="Self-study",'
        f'IF(O{rn}>=ROUNDUP(Q{rn}*Rules!$C$12,0),"Pass","Fail"),'
        f'"N/A")'
    )
    # Credits half
    ws3.cell(ri, 17).value = f'=IF(C{rn}<25,0,FLOOR(C{rn}/50,0.5))'
    ws3.cell(ri, 17).number_format = "0.0"
    # Credits fifth
    ws3.cell(ri, 18).value = f'=IF(C{rn}<50,0,ROUNDDOWN(C{rn}/50*5,0)/5)'
    ws3.cell(ri, 18).number_format = "0.0"
    # SLA flag
    ws3.cell(ri, 19).value = (
        f'=IF(AND(Requests!O{rn}<>"Approved",Requests!M{rn}<TODAY()),"Overdue",'
        f'IF(Requests!M{rn}-TODAY()<=Rules!$C$26,"Due Soon","On Track"))'
    )
    # Decision
    ws3.cell(ri, 20).value = (
        f'=IF(AND(G{rn}="Pass",J{rn}<>"Fail",L{rn}="Pass",N{rn}="Pass",P{rn}<>"Fail"),"Approved",'
        f'IF(OR(L{rn}="Fail – Missing FOS",N{rn}="Fail"),"Escalate","Needs Correction"))'
    )
    # Reason
    ws3.cell(ri, 21).value = (
        f'=IFERROR('
        f'IF(G{rn}<>"Pass","Insufficient learning objectives; ","")&'
        f'IF(AND(J{rn}="Fail",D{rn}="Group"),"Missing attendance checks; ","")&'
        f'IF(L{rn}<>"Pass","Missing field of study; ","")&'
        f'IF(N{rn}<>"Pass","Invalid program level; ","")&'
        f'IF(AND(P{rn}="Fail",D{rn}="Self-study"),"Insufficient review questions; ",""),'
        f'"Check data")'
    )

    # Border and alignment
    for ci in range(1, 22):
        ws3.cell(ri, ci).border    = thin_border()
        ws3.cell(ri, ci).alignment = center()

# Dropdown for Decision override
dv = DataValidation(type="list",
                    formula1='"Approved,Needs Correction,Escalate,Pending"',
                    allow_blank=True)
dv.sqref = f"T2:T101"
ws3.add_data_validation(dv)

# Conditional formatting – Decision column (T)
ws3.conditional_formatting.add(
    "T2:T101",
    FormulaRule(formula=['T2="Approved"'],   fill=hdr_fill(LGREEN))
)
ws3.conditional_formatting.add(
    "T2:T101",
    FormulaRule(formula=['T2="Needs Correction"'], fill=hdr_fill("FFF2CC"))
)
ws3.conditional_formatting.add(
    "T2:T101",
    FormulaRule(formula=['T2="Escalate"'],   fill=hdr_fill("FCE4D6"))
)

auto_width(ws3)

# ══════════════════════════════════════════════════════════════════════
# SHEET 4: Credit Calculator
# ══════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("Credit Calculator")

CC_HEADERS = [
    "Request ID","Course Title","Duration (min)",
    "Credits (Half-Credit Rounding)","Credits (Fifth-Credit Rounding)",
    "Min Learning Objectives","Rounding Method","Credit Flag",
]
for ci, h in enumerate(CC_HEADERS, 1):
    write_cell(ws4, 1, ci, h, fill=hdr_fill(NAVY), font=hdr_font(), align=center())

# Explanation block
expl = [
    ("Half-Credit Formula",  "=IF(F<25,0,FLOOR(F/50,0.5))",   "Rounds down to nearest 0.5 credit; awards 0 if < 25 min"),
    ("Fifth-Credit Formula", "=IF(F<50,0,ROUNDDOWN(F/50*5,0)/5)", "Rounds down to nearest 0.2 credit; awards 0 if < 50 min"),
]
ws4.cell(1, 10).value = "Formula Reference"
ws4.cell(1, 10).font  = Font(bold=True, color=NAVY)
ws4.cell(1, 11).value = "Formula"
ws4.cell(1, 11).font  = Font(bold=True, color=NAVY)
ws4.cell(1, 12).value = "Description"
ws4.cell(1, 12).font  = Font(bold=True, color=NAVY)
for ri, (lbl, fml, desc) in enumerate(expl, 2):
    ws4.cell(ri, 10).value = lbl
    ws4.cell(ri, 11).value = fml
    ws4.cell(ri, 12).value = desc
    for ci in [10, 11, 12]:
        ws4.cell(ri, ci).border = thin_border()
        ws4.cell(ri, ci).alignment = left()

for ri in range(2, 102):
    rn = ri
    ws4.cell(ri, 1).value = f"=Requests!A{rn}"
    ws4.cell(ri, 2).value = f"=Requests!B{rn}"
    ws4.cell(ri, 3).value = f"=Requests!F{rn}"
    ws4.cell(ri, 4).value = f'=IF(C{rn}<25,0,FLOOR(C{rn}/50,Rules!$C$15))'
    ws4.cell(ri, 4).number_format = "0.0"
    ws4.cell(ri, 5).value = f'=IF(C{rn}<50,0,ROUNDDOWN(C{rn}/50*5,0)/5)'
    ws4.cell(ri, 5).number_format = "0.0"
    ws4.cell(ri, 6).value = f'=ROUNDUP(C{rn}/50*Rules!$C$5,0)'
    ws4.cell(ri, 7).value = f'=IF(Requests!E{rn}="Group","Half-Credit","Fifth-Credit")'
    ws4.cell(ri, 8).value = (
        f'=IF(C{rn}<25,"Flag: < 25 min – no credit",'
        f'IF(C{rn}<50,"Flag: verify min duration","OK"))'
    )
    for ci in range(1, 9):
        ws4.cell(ri, ci).border    = thin_border()
        ws4.cell(ri, ci).alignment = center()

auto_width(ws4)

# ══════════════════════════════════════════════════════════════════════
# SHEET 5: SLA Monitor
# ══════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet("SLA Monitor")
ws5.freeze_panes = "A2"

SLA_HEADERS = [
    "Request ID","Course Title","Instructor","Submitted Date",
    "SLA Due Date","Days Until Due","Status","SLA Flag","Reviewer",
]
for ci, h in enumerate(SLA_HEADERS, 1):
    write_cell(ws5, 1, ci, h, fill=hdr_fill(NAVY), font=hdr_font(), align=center())

for ri in range(2, 102):
    rn = ri
    ws5.cell(ri, 1).value = f"=Requests!A{rn}"
    ws5.cell(ri, 2).value = f"=Requests!B{rn}"
    ws5.cell(ri, 3).value = f"=Requests!C{rn}"
    ws5.cell(ri, 4).value = f"=Requests!D{rn}"
    ws5.cell(ri, 4).number_format = "DD-MMM-YYYY"
    ws5.cell(ri, 5).value = f"=Requests!M{rn}"
    ws5.cell(ri, 5).number_format = "DD-MMM-YYYY"
    ws5.cell(ri, 6).value = f"=Requests!M{rn}-TODAY()"
    ws5.cell(ri, 6).number_format = "0"
    ws5.cell(ri, 7).value = f"=Requests!O{rn}"
    ws5.cell(ri, 8).value = (
        f'=IF(AND(G{rn}<>"Approved",E{rn}<TODAY()),"Overdue",'
        f'IF(AND(G{rn}<>"Approved",E{rn}-TODAY()<=Rules!$C$26),"Due Soon","On Track"))'
    )
    ws5.cell(ri, 9).value = f"=Requests!N{rn}"
    for ci in range(1, 10):
        ws5.cell(ri, ci).border    = thin_border()
        ws5.cell(ri, ci).alignment = center()

# Conditional formatting for SLA Flag column (H)
ws5.conditional_formatting.add(
    "H2:H101",
    FormulaRule(formula=['H2="Overdue"'],   fill=hdr_fill("FFD9D9"), font=Font(color=RED, bold=True))
)
ws5.conditional_formatting.add(
    "H2:H101",
    FormulaRule(formula=['H2="Due Soon"'],  fill=hdr_fill("FFF2CC"), font=Font(color="FF6600", bold=True))
)
ws5.conditional_formatting.add(
    "H2:H101",
    FormulaRule(formula=['H2="On Track"'],  fill=hdr_fill(LGREEN),   font=Font(color="375623", bold=True))
)

# Also colour Days Until Due column (F)
ws5.conditional_formatting.add(
    "F2:F101",
    CellIsRule(operator="lessThan", formula=["0"],    fill=hdr_fill("FFD9D9"))
)
ws5.conditional_formatting.add(
    "F2:F101",
    CellIsRule(operator="between",  formula=["0","1"], fill=hdr_fill("FFF2CC"))
)

auto_width(ws5)

# ══════════════════════════════════════════════════════════════════════
# SHEET 6: Dashboard (summary stats + chart placeholders)
# ══════════════════════════════════════════════════════════════════════
ws6 = wb.create_sheet("Dashboard")

# Helper: write a section title
def section_title(ws, row, col, text, width=4):
    ws.merge_cells(start_row=row, start_column=col,
                   end_row=row,   end_column=col+width-1)
    c = ws.cell(row=row, column=col, value=text)
    c.fill = hdr_fill(NAVY)
    c.font = Font(bold=True, color=WHITE, size=13)
    c.alignment = center()
    ws.row_dimensions[row].height = 26

section_title(ws6, 1, 1, "📊 CPE Analyst Desk – Dashboard", 10)

# ── Summary KPIs ──
kpi_labels = [
    ("Total Submissions", f'=COUNTA(Requests!A2:A101)'),
    ("Approved",          f'=COUNTIF(Requests!O2:O101,"Approved")'),
    ("Needs Correction",  f'=COUNTIF(Requests!O2:O101,"Needs Correction")'),
    ("Escalated",         f'=COUNTIF(Requests!O2:O101,"Escalate")'),
    ("Pending",           f'=COUNTIF(Requests!O2:O101,"Pending")'),
    ("Overdue",           '=COUNTIF(\'SLA Monitor\'!H2:H101,"Overdue")'),
]
ws6.cell(3, 1).value = "KEY PERFORMANCE INDICATORS"
ws6.cell(3, 1).font  = Font(bold=True, color=NAVY, size=11)

kpi_fills = [LBLUE, LGREEN, "FFF2CC", "FCE4D6", GREY, "FFD9D9"]
for ki, (lbl, fml) in enumerate(kpi_labels):
    col = ki + 1
    ws6.cell(4, col).value = lbl
    ws6.cell(4, col).fill  = hdr_fill(BLUE)
    ws6.cell(4, col).font  = Font(bold=True, color=WHITE, size=10)
    ws6.cell(4, col).alignment = center()
    ws6.cell(4, col).border = thin_border()
    ws6.cell(5, col).value = fml
    ws6.cell(5, col).fill  = hdr_fill(kpi_fills[ki])
    ws6.cell(5, col).font  = Font(bold=True, size=20, color=NAVY)
    ws6.cell(5, col).alignment = center()
    ws6.cell(5, col).border = thin_border()
    ws6.row_dimensions[5].height = 44

# ── Status breakdown pivot (manual COUNTIF table) ──
section_title(ws6, 7, 1, "Requests by Status", 4)
status_rows = [
    ("Approved",         f'=COUNTIF(Requests!O2:O101,"Approved")'),
    ("Needs Correction", f'=COUNTIF(Requests!O2:O101,"Needs Correction")'),
    ("Escalate",         f'=COUNTIF(Requests!O2:O101,"Escalate")'),
    ("Pending",          f'=COUNTIF(Requests!O2:O101,"Pending")'),
]
ws6.cell(8, 1).value = "Status"
ws6.cell(8, 2).value = "Count"
ws6.cell(8, 3).value = "% of Total"
for ci in range(1, 4):
    ws6.cell(8, ci).font   = Font(bold=True, color=NAVY)
    ws6.cell(8, ci).border = thin_border()
    ws6.cell(8, ci).alignment = center()
for sri, (lbl, fml) in enumerate(status_rows, 9):
    ws6.cell(sri, 1).value = lbl
    ws6.cell(sri, 2).value = fml
    ws6.cell(sri, 3).value = f'=B{sri}/COUNTA(Requests!A2:A101)'
    ws6.cell(sri, 3).number_format = "0.0%"
    for ci in range(1, 4):
        ws6.cell(sri, ci).border    = thin_border()
        ws6.cell(sri, ci).alignment = center()

# Pie chart – status
pie = PieChart()
pie.title = "Requests by Status"
pie.style = 10
labels = Reference(ws6, min_col=1, min_row=9, max_row=12)
data   = Reference(ws6, min_col=2, min_row=8, max_row=12)
pie.add_data(data, titles_from_data=True)
pie.set_categories(labels)
pie.width  = 14
pie.height = 10
ws6.add_chart(pie, "E7")

# ── Rejection Reasons ──
section_title(ws6, 14, 1, "Top Rejection Reasons", 4)
rejection_rows = [
    ("Missing Learning Objectives", f'=COUNTIF(Review!U2:U101,"*learning objectives*")'),
    ("Missing Attendance Checks",   f'=COUNTIF(Review!U2:U101,"*attendance*")'),
    ("Missing Field of Study",      f'=COUNTIF(Review!U2:U101,"*field of study*")'),
    ("Insufficient Review Qs",      f'=COUNTIF(Review!U2:U101,"*review questions*")'),
    ("Invalid Program Level",       f'=COUNTIF(Review!U2:U101,"*program level*")'),
]
ws6.cell(15, 1).value = "Rejection Reason"
ws6.cell(15, 2).value = "Count"
ws6.cell(15, 3).value = "% of Non-Approved"
for ci in range(1, 4):
    ws6.cell(15, ci).font   = Font(bold=True, color=NAVY)
    ws6.cell(15, ci).border = thin_border()
    ws6.cell(15, ci).alignment = center()
non_appr_fml = '=COUNTIF(Requests!O2:O101,"<>Approved")'
ws6.cell(15, 4).value = non_appr_fml
ws6.cell(15, 4).font  = Font(italic=True, color=DKGREY)
for rri, (lbl, fml) in enumerate(rejection_rows, 16):
    ws6.cell(rri, 1).value = lbl
    ws6.cell(rri, 2).value = fml
    ws6.cell(rri, 3).value = f'=IFERROR(B{rri}/$D$15,0)'
    ws6.cell(rri, 3).number_format = "0.0%"
    for ci in range(1, 4):
        ws6.cell(rri, ci).border    = thin_border()
        ws6.cell(rri, ci).alignment = center()

# Bar chart – rejection reasons
bar = BarChart()
bar.type   = "col"
bar.title  = "Rejection Reasons"
bar.style  = 10
bar.y_axis.title = "Count"
bar_labels = Reference(ws6, min_col=1, min_row=16, max_row=20)
bar_data   = Reference(ws6, min_col=2, min_row=15, max_row=20)
bar.add_data(bar_data, titles_from_data=True)
bar.set_categories(bar_labels)
bar.width  = 18
bar.height = 10
ws6.add_chart(bar, "E14")

# ── Reviewer Workload ──
section_title(ws6, 22, 1, "Reviewer Workload", 4)
ws6.cell(23, 1).value = "Reviewer"
ws6.cell(23, 2).value = "Assigned"
ws6.cell(23, 3).value = "Approved"
ws6.cell(23, 4).value = "Pending/Open"
for ci in range(1, 5):
    ws6.cell(23, ci).font   = Font(bold=True, color=NAVY)
    ws6.cell(23, ci).border = thin_border()
    ws6.cell(23, ci).alignment = center()
for rri, rv in enumerate(REVIEWERS, 24):
    ws6.cell(rri, 1).value = rv
    ws6.cell(rri, 2).value = f'=COUNTIF(Requests!N2:N101,"{rv}")'
    ws6.cell(rri, 3).value = f'=COUNTIFS(Requests!N2:N101,"{rv}",Requests!O2:O101,"Approved")'
    ws6.cell(rri, 4).value = f'=B{rri}-C{rri}'
    for ci in range(1, 5):
        ws6.cell(rri, ci).border    = thin_border()
        ws6.cell(rri, ci).alignment = center()

# ── Average Turnaround note ──
ws6.cell(29, 1).value = "Avg Turnaround (days, est.)"
ws6.cell(29, 2).value = f"=AVERAGEIF(Requests!O2:O101,\"Approved\",Requests!M2:M101)-AVERAGEIF(Requests!O2:O101,\"Approved\",Requests!D2:D101)"
ws6.cell(29, 2).number_format = "0.0"
ws6.cell(29, 1).font  = Font(bold=True)
ws6.cell(29, 2).font  = Font(bold=True, color=NAVY)
ws6.cell(29, 1).border = thin_border()
ws6.cell(29, 2).border = thin_border()

auto_width(ws6)
ws6.column_dimensions["A"].width = 28

# ══════════════════════════════════════════════════════════════════════
# SHEET 7: Certificates (Mail-merge source)
# ══════════════════════════════════════════════════════════════════════
ws7 = wb.create_sheet("Certificates")

CERT_HEADERS = [
    "Certificate No","Learner Name","Course Title","Instructor",
    "Completion Date","Credits Earned","Field of Study","Program Level",
    "Delivery Method","Certificate Issued By","Issue Date",
]
for ci, h in enumerate(CERT_HEADERS, 1):
    write_cell(ws7, 1, ci, h, fill=hdr_fill(NAVY), font=hdr_font(), align=center())

# Pull only Approved rows (simulate with formulas using IFERROR/IF)
cert_note = (
    "This sheet auto-populates from approved rows via the macro below. "
    "For full mail-merge, link this sheet as the Word data source."
)
ws7.cell(2, 1).value = cert_note
ws7.merge_cells("A2:K2")
ws7.cell(2, 1).fill = hdr_fill(LBLUE)
ws7.cell(2, 1).font = Font(italic=True, color=NAVY)
ws7.cell(2, 1).alignment = left()

# Approved data pulled manually
row_out = 3
for i, r in enumerate(REQUESTS):
    if r["status"] == "Approved":
        credits = calc_credits_half(r["duration"])
        cert_no = f"CERT-2026-{row_out-2:03d}"
        learner = f"Participant – {r['instr'].split()[0]}'s Session"
        vals = [
            cert_no, learner, r["title"], r["instr"],
            r["sla_due"].strftime("%d-%b-%Y"), credits,
            r["fos"] or "General", r["level"],
            r["method"], "CPE Desk, Firm Name", date.today().strftime("%d-%b-%Y"),
        ]
        for ci, v in enumerate(vals, 1):
            c = write_cell(ws7, row_out, ci, v)
            c.fill = hdr_fill(LGREEN)
        row_out += 1

auto_width(ws7)

# ══════════════════════════════════════════════════════════════════════
# SHEET 8: Email Templates
# ══════════════════════════════════════════════════════════════════════
ws8 = wb.create_sheet("Email Templates")
ws8.column_dimensions["A"].width = 20
ws8.column_dimensions["B"].width = 90

emails = [
    ("1. MISSING INFORMATION", "Subject", "CPE Request [Request ID] – Additional Information Required"),
    ("", "Body",
     "Dear [Instructor Name],\n\n"
     "Thank you for submitting [Course Title] (Request ID: [Request ID]) on [Submitted Date].\n\n"
     "During our review, we identified the following items that require clarification or completion "
     "before we can process your CPE credit request:\n\n"
     "  • [List missing items, e.g., Field of Study not specified]\n"
     "  • [Learning objectives provided: X; minimum required: Y]\n\n"
     "Please respond with the updated information by [SLA Due Date] to avoid delays.\n\n"
     "Regards,\nCPE Desk | [Reviewer Name]\n[Firm Name]"),

    ("2. APPROVAL NOTICE", "Subject", "CPE Credits Approved – [Course Title] ([Request ID])"),
    ("", "Body",
     "Dear [Instructor Name],\n\n"
     "We are pleased to confirm that your CPE submission has been reviewed and approved.\n\n"
     "  Course Title : [Course Title]\n"
     "  Request ID   : [Request ID]\n"
     "  Credits      : [Credits Earned]\n"
     "  Field of Study: [Field of Study]\n"
     "  Program Level: [Program Level]\n\n"
     "Certificates will be distributed to participants within 5 business days. "
     "Please retain this notice for your records.\n\n"
     "Regards,\nCPE Desk | [Reviewer Name]\n[Firm Name]"),

    ("3. CORRECTION REQUEST", "Subject", "CPE Request [Request ID] – Correction Required"),
    ("", "Body",
     "Dear [Instructor Name],\n\n"
     "We have reviewed your CPE submission for [Course Title] (Request ID: [Request ID]).\n\n"
     "The following corrections are needed before we can approve this request:\n\n"
     "  • [Specific correction, e.g., Increase learning objectives from X to Y]\n"
     "  • [Specific correction, e.g., Add attendance verification records]\n\n"
     "Please resubmit the corrected materials by [SLA Due Date]. "
     "If you need assistance, please contact us at cpe-desk@firm.com.\n\n"
     "Regards,\nCPE Desk | [Reviewer Name]\n[Firm Name]"),

    ("4. DELAY UPDATE", "Subject", "CPE Request [Request ID] – Processing Update"),
    ("", "Body",
     "Dear [Instructor Name],\n\n"
     "We wanted to provide you with an update on your CPE request for [Course Title] "
     "(Request ID: [Request ID]).\n\n"
     "Your submission is currently under review. Due to [reason for delay], "
     "we anticipate completing our review by [new estimated date]. "
     "We apologise for any inconvenience this may cause.\n\n"
     "If you have any urgent concerns, please reply to this email or call [phone number].\n\n"
     "Regards,\nCPE Desk | [Reviewer Name]\n[Firm Name]"),

    ("5. COMPLAINT REPLY", "Subject", "RE: Your Concern Regarding CPE Request [Request ID]"),
    ("", "Body",
     "Dear [Learner/Instructor Name],\n\n"
     "Thank you for bringing your concern to our attention. We take all feedback seriously.\n\n"
     "We have reviewed Request ID [Request ID] and understand your frustration regarding "
     "[brief description of the complaint].\n\n"
     "After investigation, we have determined that [explanation/resolution]. "
     "[If applicable: We have made the necessary adjustments and the corrected record/certificate "
     "will be issued by [date].]\n\n"
     "Please do not hesitate to contact us if you have any further questions.\n\n"
     "Regards,\nCPE Desk | Senior Reviewer\n[Firm Name]"),
]

ws8.cell(1, 1).value = "CPE DESK – EMAIL TEMPLATES"
ws8.merge_cells("A1:B1")
ws8.cell(1, 1).fill = hdr_fill(NAVY)
ws8.cell(1, 1).font = Font(bold=True, color=WHITE, size=14)
ws8.cell(1, 1).alignment = center()

ri = 2
for item in emails:
    section, label, body = item
    if section:
        ws8.cell(ri, 1).value = section
        ws8.merge_cells(f"A{ri}:B{ri}")
        ws8.cell(ri, 1).fill = hdr_fill(BLUE)
        ws8.cell(ri, 1).font = Font(bold=True, color=WHITE, size=11)
        ws8.cell(ri, 1).alignment = left()
        ri += 1
    ws8.cell(ri, 1).value = label
    ws8.cell(ri, 1).font  = Font(bold=True, color=NAVY)
    ws8.cell(ri, 1).fill  = hdr_fill(LBLUE)
    ws8.cell(ri, 1).alignment = left()
    ws8.cell(ri, 2).value = body
    ws8.cell(ri, 2).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws8.cell(ri, 2).border = thin_border()
    ws8.row_dimensions[ri].height = max(60, body.count("\n") * 14 + 14)
    ri += 1

# ══════════════════════════════════════════════════════════════════════
# SHEET 9: Test Cases
# ══════════════════════════════════════════════════════════════════════
ws9 = wb.create_sheet("Test Cases")

TC_HEADERS = [
    "Test #","Duration (min)","Delivery Method",
    "Expected Credits (Half)","Expected Credits (Fifth)",
    "Formula (Half)","Formula (Fifth)",
    "Actual (Half)","Actual (Fifth)","Pass/Fail",
]
for ci, h in enumerate(TC_HEADERS, 1):
    write_cell(ws9, 1, ci, h, fill=hdr_fill(NAVY), font=hdr_font(), align=center())

test_cases = [
    (1, 75,  "Group",      1.5, 1.4, "FLOOR(75/50,0.5)=1.5", "ROUNDDOWN(75/50*5,0)/5=1.4"),
    (2, 140, "Group",      2.5, 2.8, "FLOOR(140/50,0.5)=2.5","ROUNDDOWN(140/50*5,0)/5=2.8"),
    (3, 60,  "Group",      1.0, 1.2, "FLOOR(60/50,0.5)=1.0", "ROUNDDOWN(60/50*5,0)/5=1.2"),
    (4, 49,  "Group",      0.5, 0.0, "FLOOR(49/50,0.5)=0.5", "IF<50 -> 0 (fifth-credit needs >=50)"),
    (5, 20,  "Group",      0.0, 0.0, "IF<25 -> 0; FLAG",      "IF<50 -> 0; FLAG"),
    (6, 100, "Group",      2.0, 2.0, "FLOOR(100/50,0.5)=2.0", "ROUNDDOWN(100/50*5,0)/5=2.0"),
    (7, 25,  "Group",      0.5, 0.0, "FLOOR(25/50,0.5)=0.5",  "IF<50 -> 0"),
    (8, 50,  "Self-study", 1.0, 1.0, "FLOOR(50/50,0.5)=1.0",  "ROUNDDOWN(50/50*5,0)/5=1.0"),
    (9, 150, "Group",      3.0, 3.0, "FLOOR(150/50,0.5)=3.0", "ROUNDDOWN(150/50*5,0)/5=3.0"),
    (10,200, "Group",      4.0, 4.0, "FLOOR(200/50,0.5)=4.0", "ROUNDDOWN(200/50*5,0)/5=4.0"),
]

def half_credit(m):
    if m < 25: return 0.0
    return math.floor(m / 50 * 2) / 2

def fifth_credit(m):
    if m < 50: return 0.0
    return math.floor(m / 50 * 5) / 5

for ri, (tn, dur, meth, exp_h, exp_f, fml_h, fml_f) in enumerate(test_cases, 2):
    actual_h = half_credit(dur)
    actual_f = fifth_credit(dur)
    h_ok = abs(actual_h - exp_h) < 0.001
    f_ok = abs(actual_f - exp_f) < 0.001
    pf = "PASS" if (h_ok and f_ok) else "FAIL"
    vals = [tn, dur, meth, exp_h, exp_f, fml_h, fml_f, actual_h, actual_f, pf]
    for ci, v in enumerate(vals, 1):
        c = write_cell(ws9, ri, ci, v)
        c.alignment = center()
        if ci in (4,5,8,9):
            c.number_format = "0.0"
    pf_cell = ws9.cell(ri, 10)
    if "PASS" in pf:
        pf_cell.fill = hdr_fill(LGREEN)
        pf_cell.font = Font(bold=True, color="375623")
    else:
        pf_cell.fill = hdr_fill("FFD9D9")
        pf_cell.font = Font(bold=True, color=RED)

# Note row
note_ri = len(test_cases) + 3
ws9.cell(note_ri, 1).value = (
    "Notes: Fifth-credit formula requires min 50 min (no credit below 50 min). "
    "Half-credit formula requires min 25 min. "
    "Test case 4 (49 min): half-credit = 0.5 under revised NASBA standards; fifth-credit = 0 (< 50 min threshold)."
)
ws9.merge_cells(f"A{note_ri}:J{note_ri}")
ws9.cell(note_ri, 1).fill = hdr_fill(LBLUE)
ws9.cell(note_ri, 1).font = Font(italic=True, color=NAVY)
ws9.cell(note_ri, 1).alignment = Alignment(wrap_text=True)
ws9.row_dimensions[note_ri].height = 30

auto_width(ws9)

# ══════════════════════════════════════════════════════════════════════
# SHEET 10: SOP
# ══════════════════════════════════════════════════════════════════════
ws10 = wb.create_sheet("SOP")

sop_lines = [
    ("CPE ANALYST DESK – STANDARD OPERATING PROCEDURE", "title"),
    ("Version 1.0 | Effective: November 2026 | Owner: CPE Desk", "subtitle"),
    ("", ""),
    ("1. PURPOSE", "section"),
    ("To ensure all Continuing Professional Education (CPE) credit requests are reviewed, "
     "calculated, recorded, and communicated within established SLAs and in accordance with "
     "NASBA-style standards.", "body"),
    ("", ""),
    ("2. SCOPE", "section"),
    ("All CPE submissions received from instructors, training coordinators, and third-party "
     "providers for webinars, classroom sessions, and e-learning programs.", "body"),
    ("", ""),
    ("3. ROLES & RESPONSIBILITIES", "section"),
    ("CPE Analyst     | Primary reviewer; completes initial review within 5 business days.", "body"),
    ("Senior Reviewer | Handles escalations, complaints, and edge cases.", "body"),
    ("Instructor/Provider | Submits requests; responds to correction notices.", "body"),
    ("", ""),
    ("4. PROCESS STEPS", "section"),
    ("Step 1 – Intake        : Log submission in the Requests sheet with all required fields.", "body"),
    ("Step 2 – Completeness  : Confirm Field of Study, Program Level, Delivery Method, "
     "Prerequisites are populated.", "body"),
    ("Step 3 – Objectives    : Verify learning objectives ≥ ROUNDUP(duration/50×3, 0).", "body"),
    ("Step 4 – Attendance    : For Group programs, confirm attendance checks ≥ ROUNDUP(duration/50×3, 0).", "body"),
    ("Step 5 – Review Qs     : For Self-study, confirm review questions ≥ credits × 3.", "body"),
    ("Step 6 – Credit Calc   : Use Credit Calculator sheet; apply half-credit rounding for Group, "
     "fifth-credit for Self-study if applicable.", "body"),
    ("Step 7 – Decision      : Record Approved / Needs Correction / Escalate in Review sheet.", "body"),
    ("Step 8 – Communication : Send appropriate email template within SLA.", "body"),
    ("Step 9 – Certificates  : Export Certificates sheet; run Word mail merge for Approved records.", "body"),
    ("Step 10 – Dashboard    : Refresh Dashboard pivot weekly; report KPIs to manager.", "body"),
    ("", ""),
    ("5. SLA TARGETS", "section"),
    ("Initial Review       : 5 business days from submission.", "body"),
    ("Correction Response  : 3 business days after instructor reply.", "body"),
    ("Certificate Issuance : 5 business days after approval.", "body"),
    ("Escalation Resolution: 10 business days.", "body"),
    ("", ""),
    ("6. DECISION CRITERIA (NASBA-STYLE)", "section"),
    ("Approved         : All checks pass (objectives, attendance, FOS, level, review Qs).", "body"),
    ("Needs Correction : One or more fixable issues (too few objectives, missing attendance).", "body"),
    ("Escalate         : Judgment-call issues (unclear FOS, duration discrepancy, complaints).", "body"),
    ("", ""),
    ("7. KEY RULES (see Rules sheet for editable parameters)", "section"),
    ("• Min learning objectives : 3 per 50 minutes of content.", "body"),
    ("• Min attendance checks   : 3 per 50-minute credit hour (Group, technology-based).", "body"),
    ("• Min review questions    : 3 per credit hour (Self-study) – verify vs. current NASBA standard.", "body"),
    ("• Credit rounding (Group) : FLOOR(min/50, 0.5) – half-credit increments.", "body"),
    ("• Min duration for credit : 25 minutes (half-credit) or 50 minutes (fifth-credit).", "body"),
    ("", ""),
    ("8. ESCALATION PATH", "section"),
    ("CPE Analyst → Senior Reviewer → CPE Committee (if unresolved within 10 days).", "body"),
    ("", ""),
    ("9. RECORDS RETENTION", "section"),
    ("All approved submissions, certificates, and correspondence must be retained for 7 years "
     "per firm policy.", "body"),
    ("", ""),
    ("10. REVISION HISTORY", "section"),
    ("v1.0  Nov 2026  Initial release.", "body"),
]

for ri, (text, style) in enumerate(sop_lines, 1):
    c = ws10.cell(ri, 1, value=text)
    if style == "title":
        c.font = Font(bold=True, color=WHITE, size=16)
        c.fill = hdr_fill(NAVY)
        c.alignment = center()
        ws10.row_dimensions[ri].height = 32
    elif style == "subtitle":
        c.font = Font(italic=True, color=DKGREY, size=10)
        c.alignment = center()
    elif style == "section":
        c.font = Font(bold=True, color=WHITE, size=11)
        c.fill = hdr_fill(BLUE)
        c.alignment = left()
        ws10.row_dimensions[ri].height = 22
    elif style == "body":
        c.font = Font(size=10)
        c.alignment = Alignment(horizontal="left", wrap_text=True)
        ws10.row_dimensions[ri].height = 18
    ws10.merge_cells(f"A{ri}:H{ri}")

ws10.column_dimensions["A"].width = 110

# ══════════════════════════════════════════════════════════════════════
# Save Excel
# ══════════════════════════════════════════════════════════════════════
xl_path = os.path.join(OUT_DIR, "CPE_Desk.xlsx")
wb.save(xl_path)
print(f"[OK] Excel saved: {xl_path}")

# ════════════════════════════════════════════════════════════════════════════
# BUILD WORD – Certificate Template
# ════════════════════════════════════════════════════════════════════════════
def build_certificate_doc():
    doc = Document()

    # Page setup – landscape
    section = doc.sections[0]
    section.page_width  = Inches(11)
    section.page_height = Inches(8.5)
    section.left_margin  = Inches(1)
    section.right_margin = Inches(1)
    section.top_margin   = Inches(0.75)
    section.bottom_margin= Inches(0.75)

    def add_para(text, bold=False, size=12, color=None, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=6):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after  = Pt(space_after)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = RGBColor(*color)
        return p

    add_para("CERTIFICATE OF COMPLETION", bold=True, size=28,
             color=(31, 56, 100), space_before=10, space_after=4)
    add_para("Continuing Professional Education", size=16,
             color=(46, 117, 182), space_after=20)

    # Horizontal rule (via a table with one row and border on bottom)
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(9)
    p = cell.paragraphs[0]
    p.text = ""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:color"), "2E75B6")
    tcBorders.append(bottom)
    tcPr.append(tcBorders)

    doc.add_paragraph()
    add_para("This is to certify that", size=13, space_after=4)
    add_para("«Learner Name»", bold=True, size=24,
             color=(31, 56, 100), space_before=6, space_after=6)
    add_para("has successfully completed", size=13, space_after=4)
    add_para("«Course Title»", bold=True, size=20,
             color=(46, 117, 182), space_before=4, space_after=6)

    # Details table
    details_table = doc.add_table(rows=4, cols=4)
    details_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    details_data = [
        ("CPE Credits Earned:", "«Credits Earned»",   "Field of Study:",  "«Field of Study»"),
        ("Program Level:",      "«Program Level»",    "Delivery Method:", "«Delivery Method»"),
        ("Instructor:",         "«Instructor»",        "Completion Date:", "«Completion Date»"),
        ("Certificate No:",     "«Certificate No»",   "Issued By:",       "«Certificate Issued By»"),
    ]
    for r_idx, row_data in enumerate(details_data):
        row = details_table.rows[r_idx]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = val
            run = cell.paragraphs[0].runs[0]
            if c_idx % 2 == 0:
                run.bold = True
                run.font.color.rgb = RGBColor(31, 56, 100)
                run.font.size = Pt(10)
            else:
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(46, 117, 182)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    add_para("_" * 35 + "          " + "_" * 35, size=11, space_before=20)
    add_para("CPE Analyst Signature                                              Authorising Officer",
             size=9, color=(89, 89, 89))
    add_para(f"Issue Date: «Issue Date»", size=10, space_before=10)
    add_para("This certificate is issued in accordance with NASBA-style CPE standards.",
             size=8, color=(89, 89, 89), space_before=6)

    cert_path = os.path.join(OUT_DIR, "CPE_Certificate_Template.docx")
    doc.save(cert_path)
    print(f"[OK] Certificate template saved: {cert_path}")

# ════════════════════════════════════════════════════════════════════════════
# BUILD WORD – Email Templates
# ════════════════════════════════════════════════════════════════════════════
def build_email_doc():
    doc = Document()
    section = doc.sections[0]
    section.left_margin  = Inches(1.2)
    section.right_margin = Inches(1.2)

    def h(text, level=1):
        p = doc.add_heading(text, level=level)
        for run in p.runs:
            run.font.color.rgb = RGBColor(31, 56, 100)
        return p

    def body(text):
        p = doc.add_paragraph(text)
        p.paragraph_format.space_after = Pt(6)
        return p

    doc.add_heading("CPE DESK – EMAIL TEMPLATES", 0)

    email_templates = [
        {
            "title": "Template 1 – Missing Information",
            "subject": "CPE Request [Request ID] – Additional Information Required",
            "body": (
                "Dear [Instructor Name],\n\n"
                "Thank you for submitting [Course Title] (Request ID: [Request ID]) on [Submitted Date].\n\n"
                "During our review, we identified the following items that require clarification or completion "
                "before we can process your CPE credit request:\n\n"
                "  •  [List missing items – e.g., Field of Study not specified]\n"
                "  •  [Learning objectives provided: X; minimum required: Y for this duration]\n"
                "  •  [Attendance verification records not attached (Group programs only)]\n\n"
                "Please respond with the corrected information by [SLA Due Date]. Submissions not resolved "
                "within the SLA window may be closed and will require re-submission.\n\n"
                "If you have any questions, please do not hesitate to contact us.\n\n"
                "Kind regards,\n[Reviewer Name]\nCPE Desk | [Firm Name]\nEmail: cpe-desk@firm.com"
            ),
        },
        {
            "title": "Template 2 – Approval Notice",
            "subject": "CPE Credits Approved – [Course Title] ([Request ID])",
            "body": (
                "Dear [Instructor Name],\n\n"
                "We are pleased to confirm that your CPE submission has been reviewed and approved.\n\n"
                "  Course Title   :  [Course Title]\n"
                "  Request ID     :  [Request ID]\n"
                "  Credits Awarded:  [Credits Earned]\n"
                "  Field of Study :  [Field of Study]\n"
                "  Program Level  :  [Program Level]\n"
                "  Delivery Method:  [Delivery Method]\n\n"
                "Certificates of completion will be distributed to all registered participants within "
                "5 business days. A copy will also be sent to [Instructor Name] for your records.\n\n"
                "Thank you for your commitment to professional development.\n\n"
                "Kind regards,\n[Reviewer Name]\nCPE Desk | [Firm Name]"
            ),
        },
        {
            "title": "Template 3 – Correction Request",
            "subject": "CPE Request [Request ID] – Correction Required Before Approval",
            "body": (
                "Dear [Instructor Name],\n\n"
                "We have completed our initial review of [Course Title] (Request ID: [Request ID]) "
                "and found that the following corrections are required before we can issue CPE credits:\n\n"
                "  1.  [Specific issue – e.g., Learning objectives: 2 provided, minimum 4 required "
                "for a 60-minute course.]\n"
                "  2.  [Specific issue – e.g., Attendance checks: 1 recorded, minimum 3 required "
                "per 50-minute credit hour.]\n\n"
                "Please address these points and resubmit the updated documentation by [SLA Due Date]. "
                "No credits will be awarded until all requirements are met.\n\n"
                "If you believe this assessment is incorrect, please reply with supporting evidence "
                "and we will escalate your case for senior review.\n\n"
                "Kind regards,\n[Reviewer Name]\nCPE Desk | [Firm Name]"
            ),
        },
        {
            "title": "Template 4 – Delay Update",
            "subject": "CPE Request [Request ID] – Processing Update",
            "body": (
                "Dear [Instructor Name],\n\n"
                "We are writing to update you on the status of your CPE credit request for "
                "[Course Title] (Request ID: [Request ID]), submitted on [Submitted Date].\n\n"
                "Your request is currently [under additional review / awaiting documentation / "
                "pending committee decision] due to [brief reason]. We appreciate your patience.\n\n"
                "We anticipate completing our review by [new estimated date]. "
                "You will receive a formal decision or correction notice by that date.\n\n"
                "We apologise for any inconvenience this delay may cause. If you have an urgent need, "
                "please contact us at cpe-desk@firm.com or call [phone number].\n\n"
                "Kind regards,\n[Reviewer Name]\nCPE Desk | [Firm Name]"
            ),
        },
        {
            "title": "Template 5 – Complaint Reply",
            "subject": "RE: Your Concern Regarding CPE Request [Request ID]",
            "body": (
                "Dear [Learner / Instructor Name],\n\n"
                "Thank you for contacting the CPE Desk and for bringing your concern to our attention. "
                "We take all feedback seriously and are committed to resolving issues promptly.\n\n"
                "We have thoroughly reviewed Request ID [Request ID] and wish to address your concern "
                "regarding [brief description of the issue raised].\n\n"
                "Our findings:\n"
                "  [Explanation of what was found during the investigation.]\n\n"
                "Resolution:\n"
                "  [Action taken or outcome – e.g., 'We have updated the credit record; a revised "
                "certificate will be issued within 3 business days.' Or 'Our review confirms the "
                "original decision; the rationale is explained above.']\n\n"
                "If you remain dissatisfied with this outcome, you may escalate your concern to the "
                "CPE Committee by emailing cpe-committee@firm.com within 10 business days.\n\n"
                "Kind regards,\nSenior CPE Reviewer\nCPE Desk | [Firm Name]"
            ),
        },
    ]

    for tmpl in email_templates:
        doc.add_page_break()
        h(tmpl["title"], level=1)
        h("Subject Line", level=2)
        p = doc.add_paragraph(tmpl["subject"])
        p.runs[0].italic = True
        p.runs[0].font.color.rgb = RGBColor(46, 117, 182)
        h("Email Body", level=2)
        body(tmpl["body"])

    email_path = os.path.join(OUT_DIR, "CPE_Email_Templates.docx")
    doc.save(email_path)
    print(f"[OK] Email templates saved: {email_path}")

# ════════════════════════════════════════════════════════════════════════════
# BUILD WORD – SOP
# ════════════════════════════════════════════════════════════════════════════
def build_sop_doc():
    doc = Document()
    section = doc.sections[0]
    section.left_margin  = Inches(1.2)
    section.right_margin = Inches(1.2)

    def h(text, level=1):
        p = doc.add_heading(text, level=level)
        for run in p.runs:
            run.font.color.rgb = RGBColor(31, 56, 100)
        return p

    def body(text, bullet=False):
        style = "List Bullet" if bullet else "Normal"
        p = doc.add_paragraph(text, style=style)
        p.paragraph_format.space_after = Pt(4)
        return p

    # Title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run("CPE ANALYST DESK")
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(31, 56, 100)
    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run("Standard Operating Procedure – v1.0 | Effective November 2026")
    sub_run.font.size = Pt(11)
    sub_run.font.color.rgb = RGBColor(89, 89, 89)
    sub_run.italic = True

    doc.add_paragraph()

    h("1. Purpose", 1)
    body("This SOP governs the end-to-end process for receiving, reviewing, approving, recording, "
         "and communicating Continuing Professional Education (CPE) credit requests submitted by "
         "instructors and training providers. All procedures align with NASBA-style guidelines.")

    h("2. Scope", 1)
    body("All CPE submissions received monthly from up to fifty instructors and teams, covering "
         "webinars, classroom sessions, and e-learning programs (Group and Self-study delivery methods).")

    h("3. Roles & Responsibilities", 1)
    roles = [
        ("CPE Analyst",       "Primary point of contact; owns intake, review, credit calculation, and communications."),
        ("Senior Reviewer",   "Reviews escalations, complaints, and edge cases within 10 business days."),
        ("CPE Committee",     "Final arbitration on disputes not resolved by Senior Reviewer."),
        ("Instructor/Provider","Submits request form; responds to correction or information requests."),
    ]
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    hdr_cells = tbl.rows[0].cells
    hdr_cells[0].text = "Role"
    hdr_cells[1].text = "Responsibility"
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].bold = True
    for role, resp in roles:
        row = tbl.add_row()
        row.cells[0].text = role
        row.cells[1].text = resp

    doc.add_paragraph()
    h("4. Process Overview", 1)

    steps = [
        ("Intake & Logging",     "Log all fields in the Requests sheet: ID, title, instructor, dates, method, duration, FOS, level, objectives, attendance, review Qs, prerequisites."),
        ("Completeness Check",   "Verify required fields (FOS, Program Level, Delivery Method, Prerequisites) are not blank."),
        ("Learning Objectives",  "Check Obj Count ≥ ROUNDUP(Duration/50 × 3, 0). Flag Needs Correction if below minimum."),
        ("Attendance Checks",    "For Group programs: verify checks ≥ ROUNDUP(Duration/50 × 3, 0). Self-study: N/A."),
        ("Review Questions",     "For Self-study: verify review questions ≥ Credits × 3. Verify current NASBA standard before citing publicly."),
        ("Credit Calculation",   "Use Credit Calculator sheet. Group → FLOOR(min/50, 0.5). Self-study → ROUNDDOWN(min/50×5,0)/5."),
        ("Decision & Recording", "Enter Approved / Needs Correction / Escalate in Review sheet with reason codes."),
        ("Communication",        "Send appropriate email template within SLA. Attach correction details or approval notice."),
        ("Certificate Issuance", "For Approved records: export Certificates sheet; run Word mail merge; distribute within 5 days."),
        ("Dashboard Refresh",    "Update Dashboard weekly; report KPIs monthly to manager (status breakdown, rejection reasons, reviewer workload)."),
    ]
    for i, (step, desc) in enumerate(steps, 1):
        p = doc.add_paragraph(style="List Number")
        run_s = p.add_run(f"{step}: ")
        run_s.bold = True
        run_s.font.color.rgb = RGBColor(46, 117, 182)
        run_d = p.add_run(desc)
        run_d.font.size = Pt(10)

    h("5. SLA Targets", 1)
    slas = [
        ("Initial Review",      "5 business days from submission date"),
        ("Correction Response", "3 business days after instructor reply received"),
        ("Certificate Issuance","5 business days after approval decision"),
        ("Escalation Resolution","10 business days"),
        ("Complaint Reply",     "3 business days for acknowledgement; 10 for resolution"),
    ]
    tbl2 = doc.add_table(rows=1, cols=2)
    tbl2.style = "Table Grid"
    h2 = tbl2.rows[0].cells
    h2[0].text = "Activity"
    h2[1].text = "Target Turnaround"
    for cell in h2: cell.paragraphs[0].runs[0].bold = True
    for act, tgt in slas:
        row = tbl2.add_row()
        row.cells[0].text = act
        row.cells[1].text = tgt

    doc.add_paragraph()
    h("6. Decision Logic (NASBA-Style)", 1)
    body("Approved          – All checks pass: objectives, attendance/review Qs, FOS, program level, delivery method.")
    body("Needs Correction  – One or more fixable deficiencies (e.g., too few objectives, missing attendance records). Instructor must resubmit.")
    body("Escalate          – Judgment-call issues (e.g., ambiguous FOS, duration discrepancy, policy interpretation, complaints). Refer to Senior Reviewer.")

    h("7. Key Rules (NASBA-Style – see Rules Sheet)", 1)
    rules = [
        "Learning Objectives: minimum 3 per 50 minutes of content (IASA / NASBA).",
        "Attendance Checks (Group, technology-based): minimum 3 per 50-minute credit hour.",
        "Review Questions (Self-study): minimum 3 per credit hour (verify vs. current NASBA self-study standard).",
        "Credit Rounding (Group): FLOOR(min ÷ 50, 0.5) — half-credit increments.",
        "Credit Rounding (Self-study / alt): ROUNDDOWN(min ÷ 50 × 5, 0) ÷ 5 — fifth-credit increments.",
        "Minimum duration: 25 min for any credit (half-credit); 50 min for fifth-credit rounding.",
    ]
    for r in rules: body(r, bullet=True)

    h("8. Records Retention", 1)
    body("All approved submissions, certificates, review records, and correspondence must be retained "
         "for a minimum of 7 years per firm records policy and applicable regulatory requirements.")

    h("9. Escalation Path", 1)
    body("CPE Analyst  →  Senior Reviewer (within 2 business days of identification)  →  "
         "CPE Committee (if unresolved within 10 business days).")

    h("10. Revision History", 1)
    tbl3 = doc.add_table(rows=2, cols=3)
    tbl3.style = "Table Grid"
    hdr = tbl3.rows[0].cells
    for cell, t in zip(hdr, ["Version","Date","Description"]):
        cell.text = t
        cell.paragraphs[0].runs[0].bold = True
    row = tbl3.rows[1].cells
    row[0].text = "1.0"
    row[1].text = "November 2026"
    row[2].text = "Initial release"

    sop_path = os.path.join(OUT_DIR, "CPE_SOP.docx")
    doc.save(sop_path)
    print(f"[OK] SOP saved: {sop_path}")

# ─── Run all builders ────────────────────────────────────────────────────────
build_certificate_doc()
build_email_doc()
build_sop_doc()

print("\nAll CPE Desk files created successfully!")
print(f"   Location: {OUT_DIR}")
print("   CPE_Desk.xlsx")
print("   CPE_Certificate_Template.docx")
print("   CPE_Email_Templates.docx")
print("   CPE_SOP.docx")
