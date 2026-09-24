# Risk report

Generate a landscape A4 PDF from an Excel risk register. Each unique Risk ID starts
on a new page; descriptions and actions flow across as many pages as necessary.
Action column headings repeat, and each page identifies its risk.

## Setup

Python 3.10 or later. Only two direct dependencies: openpyxl and ReportLab.
Excel, Word, pandas, and an internet connection at runtime are not required.

```sh
cd risk-report
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

On Windows, activate with `.venv\Scripts\activate` instead.

## Run

```sh
python risk_report.py "/path/to/register.xlsx" \
  --project "My Project" --output report.pdf
```

Optional arguments:

- `--sheet Export`: worksheet name; otherwise uses the active worksheet.
- `--header-row 10`: explicitly select the header row; by default it is detected
  by finding the `Risk ID` heading, even with blank leading columns.
- `--report-date "24/09/2026"`: printed date; defaults to today.
- `--cost actual`: display actual action spend instead of planned spend.

The output directory must already exist. A successful run replaces an existing
output PDF. The source workbook is never modified. An empty export produces a
clear error and no report.

## Multiple actions in one row

Action cells may contain numbered entries on separate lines, such as:

```text
01 - Review the schedule
02 - Confirm resources
```

Only exactly two digits followed by a spaced hyphen (`01 - `) start an action.
Nonbreaking spaces and Unicode line separators are supported. Bare numbers, long
reference numbers, dates and numbered prose such as `01.` are preserved as text.
Each action marker becomes a separate report action row. Values in
all action columns are matched by number, so a missing `02` owner stays blank.
Unnumbered continuation lines remain part of the preceding action. Numbering
must start at the beginning of a line; inline numbers in prose are not split.
Unnumbered populated fields alongside multiple numbered actions are rejected
because their action association is ambiguous. Without an Action ID, the prefix
is displayed in the report ID column. Ordinary one-action-per-row exports remain
supported.

## Data and layout

Uses the exact column headings from the supplied export. Risk ID is required on
every populated row; blank rows are ignored. IDs should be stored as text if leading
zeros matter. Non-adjacent rows with the same Risk ID are grouped, retaining the
first occurrence order of risks and actions.

Shared risk fields use the first nonempty value. Conflicting nonempty values stop
the run with the source row and field, so disagreements can be corrected. Identical
actions with the same Risk ID and Action ID are displayed once; conflicting
duplicates stop the run. Actions without IDs are retained individually. Risks with
no actions display “No actions recorded”. Fields beginning `Action ` are treated
as action data; all other named fields are treated as shared risk data.

The report shows Title, Risk Status, Owner, Type, Description, Cause and Impact.
It uses Initial, Current and Target assessments rather than assuming what the
reference image's “Pre” and “Post” mean. Probability, schedule, cost, technical/
engineering performance and score are displayed for each assessment. The export's
`Targetr Qual Tech/Eng Performance` spelling is supported.

Actions show ID, title, owner, status, forecast start/finish, last comment and
planned (or actual) spend. Missing values appear as a dash. Dates stored as Excel
dates become DD/MM/YYYY; other values use their underlying value, not Excel's
custom display formatting. Formulas require cached results saved by Excel;
openpyxl does not calculate formulas. Built-in Helvetica supports Western European
text; other scripts require adding an appropriate embedded font.

The layout follows the reference structure, without fixed text boxes or fixed row
heights. Very long action cells can split across pages. Page numbers are global.
Additional export columns are checked for consistency but are not printed; adjust
`build_report` to change the selected fields and layout.

## Tests

```sh
python -m unittest -v
```

Tests cover grouping, duplicate actions, inconsistent data, empty input, and risks
with descriptions or individual action comments longer than a page.

Implementation references: [ReportLab tables](https://docs.reportlab.com/reportlab/userguide/ch7_tables/)
and [openpyxl workbook loading](https://openpyxl.readthedocs.io/en/3.1/tutorial.html).
