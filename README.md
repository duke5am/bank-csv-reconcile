# statement-normalizer

Turn a bank's CSV or Excel export into **one clean table** — and prove nothing was
lost while doing it. No dependencies for CSV.

```bash
python3 normalize.py statement.csv
python3 normalize.py statement.csv --out clean.csv
python3 normalize.py statement.csv --date-order dmy      # when dates are ambiguous
```

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

→ More developer tooling like this: **[duke5am.gumroad.com](https://duke5am.gumroad.com)** <!-- GUMROAD-LINK -->
