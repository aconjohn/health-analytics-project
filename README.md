# Health Analytics Project

Parsing and analysing 21 months of personal weight and nutrition data, extracted from an Excel workbook where each sheet is one week.

**Status, 10 August 2026.** The environment and repo are set up and the workbook opens from Python. The notebook in `notebooks/` confirms the sheet layout on one sheet. The parser is in development, targeted at 24 August.

## Version status

* **M0**: environment, repo, README with the scope, workbook readable from Python. Complete, 10 August 2026.
* **M1**: Era 3 parser, validated against each sheet's own totals, charts, finalised README. ~24 August 2026.
* **M2**: food reference table, weight series reconciled against a second calorie source, data-confidence metric, synthetic sample, additional charts. ~14 September 2026.

## Why?

I built the spreadsheet over two years in Excel, and I upgraded the template twice. There are 94 sheets, with no record of when each change occurred. 

Before parsing, I wrote a ten-line loop to confirm the layout of each sheet instead of working from my memory, so as to limit mistakes. The layout was confirmed as a 2x2x2x1 grid, with four header rows, two day-blocks in each of the first three, and one in the last. Header rows are on row 3. Day blocks start at column 2 and sit 9 columns apart. The parser dynamically detects header rows. 

The other issues:
* Inconsistent sheet name formatting (`14Oct24`, `02June25`, `30JUN25`), and five are typos. An override table plus a snap-to-Monday resolves both
* Three template eras with different available fields (the earliest sheets have no quantity column at all)
* Roughly 77 weeks logged across a 91-week span, around 85% coverage, with 14 weeks missing. These are left as gaps rather than filled.

## In scope

* My own data
* One user
* Local application
* Analysis and visualisation

## Out of scope

* Multiple users, accounts, authentication
* Mobile app
* Cloud hosting or deployment
* Commerciality
* Health advice or recommendations framed as guidance. This tracks only my own metrics against my own targets.

## On data availability

For data privacy, the raw workbook is deliberately excluded from this repo. A synthetic sample in the same format ships with M2, so the code can be run without it. Until then this repo is readable but not runnable end to end.

## How to run

Python 3.13.

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows, use source .venv/bin/activate elsewhere
pip install -r requirements.txt
```

A single workaround is needed for the workbook to open. openpyxl validates font records strictly and rejects the file with `ValueError: Max value is 14`. The limit has to be raised before the workbook is loaded:

```python
from openpyxl.styles.fonts import Font
Font.family.max = 100          # must run before load_workbook is called

from openpyxl import load_workbook
wb = load_workbook("data/raw/<workbook>.xlsx", read_only=True, data_only=True)
```

`data_only=True` returns computed values instead of formula strings, which is what makes each sheet's own `Total` rows usable as checksums for validating the parser.

## Layout

```
notebooks/          exploration
data/raw/           the workbook, gitignored
data/processed/     parser output          
src/                parser and reusable code  
reports/            validation output and charts 
```