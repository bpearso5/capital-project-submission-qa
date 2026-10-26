# Capital Project Submission QA & Weekly Executive Report

A rebuilt case study of work I did on a multi-region healthcare capital portfolio: finding out why submissions kept getting rejected, fixing it at the template level with a standardized pre-submission checklist, and reporting budget position to leadership every week.

> **All data here is synthetic.** Projects, people, amounts, and dates are invented. No client data, files, or screenshots are included. The rejection categories are modeled on the kinds of issues I worked through in practice.

## The problem

On a capital portfolio, every invoice, PO, and commitment change has to clear review before it routes for approval. When submissions come back rejected, the cost isn't just rework. Payment slows down, resubmissions pile up, and the budget position leadership sees falls behind.

As a contract Project Analyst, I administered spend against approved budgets: POs, commitments, invoice coding, approval routing, and remaining balances by project. I acted as the quality gate before anything routed forward. Rejections kept recurring for the same few reasons, so I stopped fixing them one at a time, found the pattern, and built a checklist that prevented them before submission.

This repo rebuilds that work at small scale so it can be shown.

## What's in here

| Path | What it is |
|---|---|
| `workbook/Capital_Project_Submission_QA_and_Weekly_Report.xlsx` | The full work sample: five tabs, every summary built from live formulas |
| `data/submission_log.csv` | 96 synthetic submissions over 12 weeks; checklist introduced at week 7 |
| `data/projects.csv` | 12 synthetic capital projects with budget, commitment, and closeout data |
| `docs/pre_submission_checklist.md` | The checklist, each item mapped to the rejection it prevents |
| `scripts/build_workbook.py` | Generates the data and builds the workbook (seeded, so output is identical every run) |

## Workbook tabs

| Tab | What it shows |
|---|---|
| `Weekly_Exec_Report` | One-page leadership summary: budget position, submission quality, projects needing attention, decisions needed. Change the report date and every figure recalculates |
| `Submission_Log` | Every submission with its outcome, rejection reason, and days to approval |
| `Rejection_Pareto` | Rejections by reason, sorted, with cumulative %, before vs. after the checklist |
| `PreSubmit_Checklist` | Fillable checklist with a "ready to route?" check |
| `Budget_Status` | Approved, committed, invoiced, uncommitted, and open commitment per project, with automatic flags |

## Rejection reasons

| Code | Reason | Preventable at submission? |
|---|---|---|
| R1 | Cost code missing or not on an approved budget line | Yes |
| R2 | Amount exceeds remaining commitment balance | Partly. Balances change upstream |
| R3 | PO number missing or doesn't match the commitment | Yes |
| R4 | Required backup missing (signed pay app, lien waiver) | Yes |
| R5 | Routed to the wrong approver | Partly. Needs a routing matrix |
| R6 | Retainage or math error | Yes |

## Results (synthetic data)

| Measure | Before checklist (wk 1–6) | After checklist (wk 7–12) |
|---|---|---|
| Submissions | 48 | 48 |
| Rejected on first pass | 22 | 5 |
| First-pass approval rate | 54.2% | 89.6% |
| Avg days to final approval | 6.9 | 4.4 |

The checklist eliminates the preventable reasons. R2 and R5 persist because they originate upstream, which is why the executive report asks leadership for a quarterly budget-line refresh and a published approver routing matrix. A checklist fixes the submitter's part. The rest is a process decision.

These figures illustrate the method. They are not measured outcomes from the engagement.

## Run it yourself

```bash
pip install openpyxl
python3 scripts/build_workbook.py   # rebuilds data/ and workbook/
```

Open the workbook in Excel. Formulas calculate on open.

## What this demonstrates

| Skill | Where it shows |
|---|---|
| Root-cause analysis | `Rejection_Pareto`: sorted reasons and cumulative % pointing to the fix |
| Process standardization | `PreSubmit_Checklist`: each control tied to the failure it prevents |
| Budget administration | `Budget_Status`: commitment, invoicing, and remaining balance by project |
| Executive reporting | `Weekly_Exec_Report`: one page, decision-ready, recalculates by date |
| Excel | COUNTIFS, AVERAGEIFS, INDEX/MATCH, data validation, conditional formatting |

## How this was built

The problem, the approach, and the workflow come from my own experience on a healthcare capital portfolio. I used Claude (Anthropic's AI assistant) to help generate the synthetic data, build the workbook, and structure the documentation. I reviewed and tested everything here, and I can walk through any formula and the reasoning behind each check.

## About me

I'm Brystal Pearson, CAPM. I work at the intersection of data and delivery: implementation support, data governance, and the reporting that keeps leadership informed. I'm completing a BBA in Management Information Systems at the University of Memphis.

[LinkedIn](https://www.linkedin.com/in/brystal-pearson-capm) · [More projects](https://github.com/bpearso5)
