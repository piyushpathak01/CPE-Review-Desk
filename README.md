# CPE Analyst Desk

> **Continuing Professional Education (CPE) tracking system — NASBA-style**  
> Built for analysts who review, approve, record, and communicate CPE credit requests from up to 50 instructors per month.

---

## Project Structure

```
CPE/
│
├── build_cpe_desk.py            ← Source script — run this to regenerate all files
│
├── CPE_Desk.xlsx                ← Main workbook (10 sheets)
├── CPE_Certificate_Template.docx← Word mail-merge certificate
├── CPE_Email_Templates.docx     ← 5 ready-to-use email templates
├── CPE_SOP.docx                 ← Standard Operating Procedure
│
└── README.md                    ← This file
```

---

## Quick Start

### Prerequisites

```bash
pip install openpyxl python-docx
```

### Generate / Regenerate All Files

```bash
cd "C:\Users\Piyush Pathak\Desktop\CPE"
python build_cpe_desk.py
```

Rebuilds all 4 output files from scratch in under 10 seconds.

---

## Excel Workbook — CPE_Desk.xlsx

| # | Sheet | Purpose |
|---|---|---|
| 1 | **Requests** | 100 mock CPE submissions with seeded errors |
| 2 | **Rules** | Editable NASBA-style thresholds (all formulas reference these cells) |
| 3 | **Review** | Automated Pass/Fail checks + decision + reason per request |
| 4 | **Credit Calculator** | Half-credit & fifth-credit rounding formulas |
| 5 | **SLA Monitor** | Red/Amber/Green flags for overdue and at-risk reviews |
| 6 | **Dashboard** | KPIs, pie chart, bar chart, reviewer workload |
| 7 | **Certificates** | Mail-merge source — 58 approved rows |
| 8 | **Email Templates** | All 5 template bodies inside Excel for quick reference |
| 9 | **Test Cases** | 10 formula test cases with expected vs. actual results |
| 10 | **SOP** | Full process note embedded in the workbook |

---

## Rules Sheet — Key Editable Parameters

All thresholds live in **column C of the Rules sheet**.  
Change a value there → every formula across every sheet updates instantly. No VBA, no macros.

| Cell | Parameter | Default | Source |
|---|---|---|---|
| C5 | Min learning objectives per 50 min | **3** | IASA / NASBA |
| C6 | Min duration for any credit | **25 min** | NASBA |
| C9 | Min attendance checks per 50-min hour | **3** | LearnSignal / NASBA |
| C12 | Min review questions per credit (Self-study) | **3** | NASBA SS Standard* |
| C15 | Half-credit increment | **0.5** | NASBA |
| C16 | Fifth-credit increment | **0.2** | NASBA alt |
| C25 | SLA turnaround | **5 days** | Firm policy |
| C26 | Due-soon warning threshold | **1 day** | Firm policy |

> ⚠️ *Self-study review questions: verify the exact minimum against the current NASBA self-study standard before citing publicly.*

---

## Key Formulas

```excel
-- Credits, half-credit rounding (Group programs)
=IF(F2<25, 0, FLOOR(F2/50, Rules!$C$15))

-- Credits, fifth-credit rounding (Self-study / alternate)
=IF(F2<50, 0, ROUNDDOWN(F2/50*5, 0)/5)

-- Minimum learning objectives needed
=ROUNDUP(F2/50 * Rules!$C$5, 0)

-- Objectives check
=IF(I2 >= ROUNDUP(F2/50*3, 0), "Pass", "Fail")

-- SLA flag
=IF(AND(O2<>"Approved", M2<TODAY()), "Overdue",
   IF(M2-TODAY() <= Rules!$C$26, "Due Soon", "On Track"))

-- Final decision
=IF(AND(G2="Pass", J2<>"Fail", L2="Pass", N2="Pass", P2<>"Fail"),
   "Approved",
   IF(OR(L2="Fail – Missing FOS", N2="Fail"),
      "Escalate", "Needs Correction"))
```

---

## Seeded Errors in Mock Data

The 100 rows in the Requests sheet contain deliberate errors to make the dashboard meaningful:

| Error Type | Count | Triggers |
|---|---|---|
| Missing / insufficient learning objectives | 15 | → Needs Correction |
| No attendance checks (Group programs) | 10 | → Needs Correction |
| Duration mismatches | 8 | → Needs Correction |
| Missing Field of Study | 5 | → Escalate |
| Duplicate submissions | 4 | → Review |

**Resulting status split:** 58 Approved · 27 Needs Correction · 11 Pending · 4 Escalate

---

## Test Cases — Credit Formula Proof

| Duration | Half-Credit (expected) | Fifth-Credit (expected) | Notes |
|---|---|---|---|
| 75 min | **1.5** | 1.4 | — |
| 140 min | **2.5** | 2.8 | — |
| 60 min | **1.0** | 1.0 | — |
| 49 min | **0.5** | 0.0 | Fifth-credit needs ≥ 50 min |
| 20 min | **0.0** | 0.0 | Below both thresholds — flagged |
| 25 min | **0.5** | 0.0 | Minimum for half-credit |
| 50 min | **1.0** | 1.0 | Minimum for fifth-credit |
| 100 min | **2.0** | 2.0 | — |
| 150 min | **3.0** | 3.0 | — |
| 200 min | **4.0** | 4.0 | — |

All test cases show ✔ PASS in the Test Cases sheet.

---

## Word Mail Merge — Certificates

1. Open **`CPE_Certificate_Template.docx`**
2. **Mailings → Select Recipients → Use Existing List**
3. Browse to `CPE_Desk.xlsx` → select sheet **`Certificates`**
4. All `«merge fields»` are pre-placed in the template
5. **Finish & Merge → Print Documents** (or Edit Individual Documents to preview)

Fields used: `«Certificate No»`, `«Learner Name»`, `«Course Title»`, `«Instructor»`,  
`«Completion Date»`, `«Credits Earned»`, `«Field of Study»`, `«Program Level»`,  
`«Delivery Method»`, `«Certificate Issued By»`, `«Issue Date»`

---

## Email Templates

Five templates in **`CPE_Email_Templates.docx`** — copy, fill `[placeholders]`, send:

| # | Template | Trigger |
|---|---|---|
| 1 | Missing Information | Required field or doc not submitted |
| 2 | Approval Notice | Request approved, credits awarded |
| 3 | Correction Request | Fixable issue found — instructor must resubmit |
| 4 | Delay Update | Review will exceed SLA |
| 5 | Complaint Reply | Learner or instructor raises a formal concern |

---

## SLA Targets

| Activity | Target |
|---|---|
| Initial review | 5 business days from submission |
| Correction response | 3 business days after instructor reply |
| Certificate issuance | 5 business days after approval |
| Escalation resolution | 10 business days |
| Complaint acknowledgement | 3 business days |

---

## Decision Logic

```
ALL checks pass?
  └── YES → Approved
  └── NO  → Is the issue a judgment call (missing FOS, invalid level)?
              └── YES → Escalate  (→ Senior Reviewer)
              └── NO  → Needs Correction  (→ email instructor)
```

---

## Reviewer Workload Tracking

The Dashboard sheet tracks per-reviewer:
- **Assigned** — total requests allocated
- **Approved** — completed successfully
- **Open** — pending or in correction cycle

Four reviewers are pre-configured: Priya, Omar, Sarah, Luis.

---

## References

| Rule | Source |
|---|---|
| 3 objectives per 50 min | IASA NASBA Requirements (2025) |
| 3 attendance checks per 50-min hour | LearnSignal / NASBA registry guidance |
| Self-study review questions | NASBA self-study standard (verify current version) |
| Half / fifth-credit rounding | NASBA CPE Standards |

---

## Regenerating the Files

If you need to update seed data, add instructors, or change error distributions:

1. Open [`build_cpe_desk.py`](build_cpe_desk.py)
2. Edit the constants at the top (`INSTRUCTORS`, `DURATIONS`, `error_pool` counts, etc.)
3. Run `python build_cpe_desk.py`
4. All 4 output files are overwritten with fresh data

---

*CPE Analyst Desk v1.0 — Built October 2026*

