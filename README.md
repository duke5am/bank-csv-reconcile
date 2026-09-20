# bank-csv-reconcile

[![PyPI](https://img.shields.io/pypi/v/bank-csv-reconcile)](https://pypi.org/project/bank-csv-reconcile/)

Turn a bank's CSV or Excel export into **one clean table** — and prove nothing was
lost while doing it. No dependencies for CSV.

## Install it (two minutes, no developer tools)

You do not need to download anything from this page. One command fetches the tool
and puts it on your computer.

**1. Check you have Python.** Open a terminal (Windows: press the Windows key and
type `cmd`; macOS: press `Cmd`+`Space` and type `Terminal`) and type:

```bash
python3 --version
```

If it answers `Python 3.9` or higher, go to step 2. If it says *command not found*
or opens the Microsoft Store, install Python first from
<https://www.python.org/downloads/> — **on Windows, tick "Add python.exe to PATH"
on the first screen of the installer**, then close the terminal and open a new one.

**2. Install the tool.** Type exactly this:

```bash
python3 -m pip install bank-csv-reconcile
```

On Windows the command is usually:

```bash
py -m pip install bank-csv-reconcile
```

**3. Run it.** Nothing else to set up — no folders to activate, no environment to
configure:

```bash
bank-csv-reconcile statement.csv
```

The clean table is written next to your statement as
`statement-normalized.csv` (change that with `--out clean.csv`). Your original
file is never modified.

```bash
bank-csv-reconcile statement.csv --out clean.csv
bank-csv-reconcile statement.csv --date-order dmy     # when dates are ambiguous
bank-csv-reconcile january.xlsx --json-report january-report.json
```

### If `pip` is not found

The message is usually *"No module named pip"* or *"pip: command not found"*.
Python is there, its installer is not. Try these in order:

```bash
python3 -m ensurepip --upgrade          # Windows: py -m ensurepip --upgrade
```

If that does not help, install pip the way your system expects, then run step 2
again:

```bash
sudo apt install python3-pip            # Debian, Ubuntu, Mint, Raspberry Pi OS
```

On macOS, `python3 -m ensurepip --upgrade` is enough; if `python3` itself is
missing, install Python from <https://www.python.org/downloads/>.

### If pip refuses to install it

On newer Debian/Ubuntu systems pip may stop with *"externally-managed-environment"*.
That is your system protecting its own Python, not a fault in the tool. Either of
these two ways around it is fine — pick one and change nothing else:

```bash
sudo apt install pipx && pipx install bank-csv-reconcile
```

or install it into its own private folder (copy both lines as they are):

```bash
python3 -m venv ~/.bank-csv-reconcile
~/.bank-csv-reconcile/bin/pip install bank-csv-reconcile
```

With the second one the command to run is
`~/.bank-csv-reconcile/bin/bank-csv-reconcile statement.csv`.

### If `bank-csv-reconcile` is not found after installing

The install worked, but the folder Python puts commands in is not on your `PATH`.
You can always run the tool this way instead, from any folder:

```bash
python3 -m bank_csv_reconcile.cli statement.csv
```

To fix it properly on Windows, re-run the Python installer, choose *Modify*, and
tick *"Add Python to environment variables"*; on macOS and Linux, add the folder
that `python3 -m site --user-base` prints, followed by `/bin`, to your `PATH`.

### Excel `.xlsx` files (optional)

CSV files need nothing extra. To read or write `.xlsx` files, add the optional
Excel support once:

```bash
python3 -m pip install "bank-csv-reconcile[excel]"
```

If you forget, the tool says so plainly and tells you this exact command — it
never produces a wrong answer because of it.

### Check it worked

```bash
bank-csv-reconcile --version
```

## Or run it straight from a clone (nothing installed)

If you would rather not install anything, clone the repository and run
`normalize.py`; it is the same program:

```bash
git clone https://github.com/duke5am/bank-csv-reconcile.git
cd bank-csv-reconcile
python3 normalize.py statement.csv
python3 normalize.py statement.csv --out clean.csv
python3 normalize.py statement.csv --date-order dmy      # when dates are ambiguous
```

There are four sample statements in `samples/` you can try it on right away, e.g.
`python3 normalize.py samples/sample3-european-format.csv`.

## The twenty minutes, every month

Download the statement. Scroll past the bank's heading block. Work out which
column is the date, which is the description, and whether money is split into
money-in and money-out. Turn `(45.90)` into `-45.90`. Delete the Closing Balance
and Total Debits rows. Re-type the dates so they sort. Stitch back the description
that wrapped onto a second line. Then do it again next month, for a bank that does
all of it differently.

## The part that makes it trustworthy

It **reconciles the running balance** and tells you whether the numbers add up:

```
 RECONCILIATION - proof nothing was lost
  Balance column: yes - "Balance"
  Checks run:     12 balance checks - 12 passed, 0 did not add up
  Totals:         4,850.00 in, 3,472.20 out, net 1,377.80 over 12 transactions
  RESULT: PASS
```

And when a balance does not add up, it names the row and shows the arithmetic:

```
  ROWS THAT DO NOT ADD UP
    line 17: FASTER PAYMENT OUT - A NOTHER WAGES  amount -540.00
        balance above (line 16): 3,514.95   +   amount: -540.00   =   2,974.95
        but the file says the balance here is 2,474.95  ->  difference -500.00
```

That is the difference between "probably fine" and "provably complete". A
transformation that silently drops a row or misreads a sign is otherwise
invisible — and it is the mistake that quietly corrupts a set of books.

## What it handles

- **debit/credit split columns**, or **one signed amount column**
- `1,234.56` · `(123.45)` as negative · `-123.45` · trailing `CR`/`DR` · currency symbols
- **European format** `1.234,56`, semicolon files, and German column headings
- junk heading blocks above the real header, and footer/summary rows
- descriptions that wrapped onto a second line
- Excel `.xlsx` (needs `openpyxl`; CSV needs nothing)
- **newest-first** exports, by running the balance check bottom-up

## It refuses to guess your dates

If day/month ordering is genuinely ambiguous — `03/04/2025` is both 3 April and
4 March — it **stops and asks** rather than picking one. Reading it the wrong way
moves a transaction into a different month, which is a real bookkeeping error.

```
  STOPPED BEFORE GUESSING: every date in this file could be read two ways
  Reading '03/04/2025' as 3 April instead of 4 March moves that transaction
  into a different month, which is a real bookkeeping error.
  --date-order dmy   (day first, e.g. UK, Ireland, Australia, most of Europe)
  --date-order mdy   (month first, e.g. US bank exports)
```

Where the order *can* be proven, it is inferred and says so: `day first -
'14/03/2025' cannot be a month`.

## Exit codes

`0` clean · `1` a balance did not add up, or the file could not be read · `2` a
decision is needed first (`--date-order`).

## What it is not

- Not accounting or tax advice, and not bookkeeping software. It cleans a file and
  proves the file is complete; what the figures *mean* is your judgement.
- It handles the layouts it documents. Banks change formats, and an unrecognised
  one stops with a clear message rather than producing a plausible-looking wrong
  answer.
- No PDF statements, no `.xls`, and nothing posts into Xero, QuickBooks or Sage.
- The balance check needs a **running** balance column. An "available balance"
  column moves on card authorisations and will not reconcile row by row.

## The full pack

The paid pack adds eight sample statements covering more bank layouts, the
for-bookkeepers guide (including how to open a terminal on Windows and Mac), the
column-mapping and reconciliation guides, and 86 tests.

<!-- RELATED:START -->

## Related tools

- **[cur-athena-lint](https://github.com/duke5am/cur-athena-lint)** — Lint AWS Cost and Usage Report Athena SQL for partition pruning and column mistakes, with the schema reference and a FinOps playbook.
  *(if you were searching for "aws cur athena query")*
- **[ga4-bigquery-lint](https://github.com/duke5am/ga4-bigquery-lint)** — Lint GA4 BigQuery SQL for the session, event_params and column mistakes that quietly give you wrong numbers, using sqlglot's real BigQuery grammar.
  *(if you were searching for "ga4 bigquery queries")*

All 28 tools in this set, grouped by what they check: **[dev-tools-index](https://duke5am.github.io/dev-tools-index/)**

If you arrived here searching for one of these, this is the tool: **bank statement csv to excel** · **reconcile bank statement** · **convert bank export to one table** · **ofx qif csv normalise**

<!-- RELATED:END -->

→ **[Bank Statement Normalizer](https://duke5am.gumroad.com/l/31-bank-statement-normalizer)** — $29 on Gumroad <!-- GUMROAD-LINK -->
