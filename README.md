# Health Analytics Project

Parsing and analysing 21 months of personal weight and nutrition data, extracted from an Excel workbook where each sheet is one week.

**Status, 07 September 2026.** The environment and repo are set up and the workbook opens from Python. 

## Version status

* **M0**: environment, repo, README with the scope, workbook readable from Python. Complete, 10 August 2026.
* **M1**: Era 3 parser, validated against each sheet's own totals, finalised README. Complete, 7 September 2026.
* **M2**: food reference table, weight series reconciled against a second calorie source, data-confidence metric, synthetic sample, charts.

## Why?

I built the spreadsheet over two years in Excel, and I upgraded the template twice. There are 94 sheets, 81 real weeks, with no record of when each change occurred. 

Before parsing, I wrote a ten-line loop to confirm the layout of each sheet instead of working from my memory, so as to limit mistakes. The layout was confirmed as a 2x2x2x1 grid, with four header rows, two day-blocks in each of the first three, and one in the last. If grouped by header labels, there are three layouts. In actuality, there are four. The parser is keyed on labels and header row numbers to solve that issue.

There is an unclean transition in the data, with the header sitting on row 2 from 14Oct24 to 10Feb25, and then row 2 and row 3 are interwoven until 13May25, then just row 3. This meant layout cannot be decided through date range, even within just one era. It had to be detected per sheet.

For everything positional, it was discovered that hardcoding would not be sufficient. Day-blocks start at column 1 and 6 on the earliest layout, then 2 and 10, then 2 and 11. The height of the blocks vary. A sheet has a "Total" which is unrelated, outside of the block areas, which would have caused a bug in that sheet. This meant pairing headers and totals based on columns was the only solution.

Fields are mapped by names, as the quantity column was inserted instead of appended in the newest template, so each field sits one to the right. The parser had to key by name instead of offset.


The other issues:
* Inconsistent sheet name formatting (`14Oct24`, `02June25`, `30JUN25`). Three sheets are a day later than they should be. The override table plus a snap-to-Monday resolves this.
* Four template eras with different available fields. Earliest sheets have calories and protein only. Full macro analysis can only cover 45 of the 81 weeks, with quantity covering 37.
* 81 real weeks logged across a 91-week span, around 85% coverage, with 14 weeks missing. These are left as gaps rather than filled.

## Validation

The parser was checked against 1130 independent comparisons. For each of the 567 day blocks, the parsed rows' calories and protein were summed. These were then compared to the calculated "Total" row from the original workbook. There are 929 exact matches, with 200 disagreements. All 200 of these have identical cause. The parser skips rows which lack a food name in order to filter out any blank rows. The SUM function from Excel, unlike the parser, is not concerned with the names of food items, and as such calculation continues on. A block in 20OCT25 contains a Total with a non numeric cell. This is reported by the parser.

Interpretation was possible here, as if there was an issue with the parser, there would be Totals which are greater as well as lesser, but here there are only lesser. Keeping unlabelled rows is currently an open decision, but it is likely this will be done. Full comparison is available in reports/validation.csv.

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

## Known Limitations
* Some cells, a small number, have notes instead of numbers, like one cell which contains "^", meaning "same as above". The parser and Excel both ignore these.
* Seven day-blocks per sheet is not a given. There is one sheet without a Saturday and SUnday. The parser records weeks which are short.
* There are some weeks with clearly projected calories, and detecting real vs projected days is currently not built.
* Day names are not reliable, as copy-pasting has resulted in mis labelling. Weekdays are instead assigned by reading order and asserted against the Monday of that sheet.

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