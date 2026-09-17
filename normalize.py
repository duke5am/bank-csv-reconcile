#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Power Query Bank Statement Normalizer
=====================================

Turn a messy bank statement export (CSV or Excel .xlsx) into ONE clean table:

    date, description, amount, balance, source_file, source_row

...and then PROVE nothing was lost, by checking that the running balance column
still adds up row by row.

This file is a single self-contained program. It needs Python 3.8 or newer.
It needs no third-party libraries at all for CSV files. For .xlsx files it
needs `openpyxl` (see README.md / docs/FOR-BOOKKEEPERS.md).

Run `python3 normalize.py --help` for all options.

Exit codes
----------
0   success
1   something went wrong (file missing, unreadable, no header found, ...)
2   the file was read, but a decision is needed before it can be trusted
    (currently: ambiguous day/month order in the dates -> use --date-order)
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
import unicodedata
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

VERSION = "1.0.0"

# ---------------------------------------------------------------------------
# Output schema (the clean table every statement is squeezed into)
# ---------------------------------------------------------------------------
OUT_COLUMNS = ["date", "description", "amount", "balance", "source_file", "source_row"]

# How much a running-balance difference may be before we call it a problem.
# One penny: statements round to whole pennies, so anything at or below that is
# rounding, and anything above it is a real discrepancy worth your time.
BALANCE_TOLERANCE = Decimal("0.01")

# Internal field names we try to map columns onto.
FIELD_DATE = "date"
FIELD_DESCRIPTION = "description"
FIELD_AMOUNT = "amount"
FIELD_DEBIT = "debit"
FIELD_CREDIT = "credit"
FIELD_BALANCE = "balance"

# Columns that are not part of the output schema but are worth gluing onto the
# description when the real description column was taken by something else.
DESCRIPTION_EXTRA_HINTS = [
    "memo", "reference", "payee", "particulars", "type", "transactiontype",
    "chequenumber", "checknumber", "cheque", "description2", "details2", "narrative2",
    "verwendungszweck", "additionalinfo", "additionalinformation",
]

# ===========================================================================
# Section 1 - header recognition
# ===========================================================================
#
# Banks use dozens of names for the same four things. We normalise every header
# cell to lowercase letters+digits only ("Transaction Date" -> "transactiondate")
# and look it up in this table.
#
# Entries of 3 characters or fewer are only used for an EXACT match. They are
# too dangerous for the fuzzy "does the header contain this word" pass: "dr"
# appears inside "address", "in" appears inside almost everything.
#
HEADER_ALIASES = {
    FIELD_DATE: [
        "date", "transactiondate", "transdate", "txndate", "posteddate", "postdate",
        "postingdate", "bookingdate", "booking", "valuedate", "valuedt", "effectivedate",
        "entrydate", "processdate", "operationdate", "completeddate", "datumposted",
        "buchungstag", "datum", "dateposted", "operationdate", "transactionposteddate",
    ],
    FIELD_DESCRIPTION: [
        "description", "transactiondescription", "details", "detail", "memo", "narrative",
        "particulars", "payee", "name", "reference", "transaction", "transactiondetails",
        "transactionmemo", "notes", "note", "description1", "merchant", "merchantname",
        "verwendungszweck", "beschreibung", "libelle", "concepto", "omschrijving",
        "transactionnarrative", "info",
    ],
    FIELD_AMOUNT: [
        "amount", "transactionamount", "netamount", "grossamount", "value", "betrag",
        "importe", "bedrag", "montant", "signedamount", "amountinaccountcurrency",
        "transactionvalue",
    ],
    FIELD_DEBIT: [
        "debit", "debits", "debitamount", "debitamt", "withdrawal", "withdrawals",
        "withdrawalamount", "moneyout", "paidout", "paymentsout", "outflow", "soll",
        "debe", "dare", "af", "dr", "out",
    ],
    FIELD_CREDIT: [
        "credit", "credits", "creditamount", "creditamt", "deposit", "deposits",
        "depositamount", "moneyin", "paidin", "paymentsin", "inflow", "haben",
        "haber", "avere", "bij", "cr", "in",
    ],
    FIELD_BALANCE: [
        "balance", "runningbalance", "ledgerbalance", "closingbalance", "balanceamount",
        "newbalance", "balanceafter", "availablebalance", "saldo", "kontostand", "solde",
        "saldodisponible", "balancecarriedforward", "endbalance",
    ],
}

# Which order to test fields in when a header could mean two things.
FIELD_PRIORITY = [FIELD_DATE, FIELD_DEBIT, FIELD_CREDIT, FIELD_BALANCE, FIELD_AMOUNT,
                  FIELD_DESCRIPTION]

_NON_ALNUM = re.compile(r"[^a-z0-9]+")


def norm_header(cell):
    """'Transaction  Date' -> 'transactiondate'. Safe for None/numbers."""
    if cell is None:
        return ""
    s = str(cell).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return _NON_ALNUM.sub("", s)


def _alias_lookup():
    exact, fuzzy = {}, {}
    for field, aliases in HEADER_ALIASES.items():
        for a in aliases:
            exact.setdefault(a, field)
            if len(a) >= 4:
                fuzzy.setdefault(a, field)
    return exact, fuzzy


ALIAS_EXACT, ALIAS_FUZZY = _alias_lookup()


def map_columns(header_cells):
    """Map a header row onto our fields.

    Returns (mapping, extras, ignored) where
      mapping: {field: header cell text}          - the columns we will use
      extras : [header cell text]                 - description-ish leftovers
      ignored: [header cell text]                 - read but not used
    Two passes: exact name matches first (so "Credit Amount" is a credit column,
    not an amount column), then fuzzy substring matches.
    """
    cells = list(header_cells)
    mapping, used, extras, ignored = {}, set(), [], []
    normed = [norm_header(c) for c in cells]

    # Pass 1: exact alias match.
    for idx, n in enumerate(normed):
        if not n or idx in used:
            continue
        f = ALIAS_EXACT.get(n)
        if f and f not in mapping:
            mapping[f] = cells[idx]
            used.add(idx)

    # Pass 2: fuzzy match ("Transaction Amount" -> amount), longest alias wins.
    for field in FIELD_PRIORITY:
        if field in mapping:
            continue
        best, best_len = None, 0
        for idx, n in enumerate(normed):
            if not n or idx in used:
                continue
            for alias, afield in ALIAS_FUZZY.items():
                if afield != field:
                    continue
                if alias in n and len(alias) > best_len:
                    best, best_len = idx, len(alias)
        if best is not None:
            mapping[field] = cells[best]
            used.add(best)

    # Leftover text-ish columns.
    for idx, c in enumerate(cells):
        if idx in used:
            continue
        n = normed[idx]
        if not n:
            continue
        if any(h in n for h in DESCRIPTION_EXTRA_HINTS) and FIELD_DESCRIPTION in mapping:
            extras.append(c)
        else:
            ignored.append(c)
    return mapping, extras, ignored


# ===========================================================================
# Section 2 - numbers (this is where naive tools quietly produce wrong money)
# ===========================================================================

CURRENCY_SYMBOLS = "$€£¥₹₽₪₺₴₦₱₲₵₡﷼"
# Banks write the credit/debit marker attached ("1,234.56CR"), spaced
# ("1,234.56 CR") or in front ("DR 1,234.56"). A digit before the marker is
# not a word boundary, so \b cannot be used here.
_CRDR_TRAIL = re.compile(r"(?<![A-Za-z])(CR|DR)\s*$", re.IGNORECASE)
_CRDR_LEAD = re.compile(r"^\s*(CR|DR)(?![A-Za-z])", re.IGNORECASE)

# Only these two shapes are proof of a number style, because they use both
# separators: 1.234,56 is European, 1,234.56 is UK/US. A bare '1.234' is not
# proof of anything, and the code says so rather than assuming.
EU_STYLE_PAT = re.compile(r"^\d{1,3}(?:\.\d{3})+,\d{1,2}$")
US_STYLE_PAT = re.compile(r"^\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?$")


class AmountParse(object):
    """Result of reading one amount cell."""

    __slots__ = ("value", "notes")

    def __init__(self, value, notes=None):
        self.value = value          # Decimal or None if we could not read it
        self.notes = notes or []    # plain-English notes about what we did


def _strip_amount(raw):
    """Remove currency symbols, spaces, CR/DR markers. Returns (core, negative, notes)."""
    notes = []
    s = str(raw)
    s = s.replace("\u00a0", " ").replace("\u202f", " ")
    for bad, good in (("\u2212", "-"), ("\u2013", "-"), ("\u2014", "-"), ("\u2015", "-")):
        s = s.replace(bad, good)
    s = s.strip()
    negative = False

    m = _CRDR_TRAIL.search(s)
    if m:
        if m.group(1).upper() == "DR":
            negative = True
            notes.append("read a trailing 'DR' as money out")
        else:
            notes.append("read a trailing 'CR' as money in")
        s = s[:m.start()]
    else:
        m = _CRDR_LEAD.match(s)
        if m:
            if m.group(1).upper() == "DR":
                negative = True
                notes.append("read a leading 'DR' as money out")
            else:
                notes.append("read a leading 'CR' as money in")
            s = s[m.end():]

    # Parentheses round a number mean negative in every bank export.
    if s.startswith("(") and s.endswith(")"):
        negative = True
        notes.append("read (brackets) as negative")
        s = s[1:-1]

    # Drop currency symbols, letters and spaces. '-' '+' '.' ',' '(' ')' survive.
    s = "".join(ch for ch in s if ch.isdigit() or ch in ".,+-()")
    s = s.replace("(", "").replace(")", "")
    if s.startswith("-"):
        negative = True
        s = s[1:]
    s = s.replace("-", "").replace("+", "")
    if not s:
        return None, negative, notes
    return s, negative, notes


def detect_number_style(cells):
    """Look at every amount-ish cell in the file and decide EU vs US thousands.

    Returns (style, evidence) where style is 'eu', 'us' or 'plain'.
    European style is '1.234,56'; US/UK style is '1,234.56'. If a file shows
    only one of those shapes we follow it; if it shows only bare numbers like
    '1234.56' we call it 'plain' and say so out loud, because '1.234' on its own
    is genuinely ambiguous.
    """
    eu = us = plain_dot = plain_comma = ambiguous_dot = 0
    for raw in cells:
        core, _neg, _n = _strip_amount(raw)
        if not core:
            continue
        if EU_STYLE_PAT.match(core):
            eu += 1
        elif US_STYLE_PAT.match(core):
            us += 1
        elif "." in core and "," not in core:
            plain_dot += 1
            if re.match(r"^\d{1,3}\.\d{3}$", core):
                ambiguous_dot += 1
        elif "," in core and "." not in core:
            if re.search(r",\d{1,2}$", core):
                plain_comma += 1
            else:
                plain_dot += 1
    evidence = {"european_examples": eu, "us_examples": us,
                "bare_dot_decimals": plain_dot, "bare_comma_decimals": plain_comma,
                "ambiguous_dot_thousands": ambiguous_dot}
    if eu and not us:
        return "eu", evidence
    if us and not eu:
        return "us", evidence
    if eu and us:
        return ("eu" if eu >= us else "us"), evidence
    if plain_comma and not plain_dot:
        return "eu", evidence
    return "plain", evidence


def parse_amount(raw, style="plain"):
    """Read one amount cell into a signed Decimal.

    Handles: 1,234.56 | 1234.56 | (123.45) | -123.45 | 123.45CR | 123.45 DR
             $1,234.56 | 1.234,56 | -1.234,56 | 1 234,56 | 1234 | 0.00
    `style` ('eu' or 'us'/'plain') only settles genuinely ambiguous cells such
    as '1.234' (European thousands, or one-point-two-three-four?).
    """
    if raw is None:
        return AmountParse(None)
    if isinstance(raw, (int, float, Decimal)) and not isinstance(raw, bool):
        try:
            return AmountParse(Decimal(str(raw)).quantize(Decimal("0.01"), ROUND_HALF_UP))
        except InvalidOperation:
            return AmountParse(None)
    text = str(raw).strip()
    if text == "" or text in ("-", "--", "N/A", "n/a", "NA", "None", "nan"):
        return AmountParse(None)
    core, negative, notes = _strip_amount(text)
    if not core:
        return AmountParse(None, notes + ["no digits found"])

    has_dot, has_com = "." in core, "," in core
    if has_dot and has_com:
        dec = "." if core.rfind(".") > core.rfind(",") else ","
        thou = "," if dec == "." else "."
        core = core.replace(thou, "").replace(dec, ".")
        if core.count(".") > 1:
            return AmountParse(None, notes + ["could not read this number"])
    elif has_com:
        last = core.split(",")[-1]
        if len(last) == 2:
            # 412,80 -> European decimal comma
            core = core.replace(".", "").replace(",", ".")
            notes.append("read the comma as a decimal point (European style)")
        elif style == "eu" and len(last) == 3:
            core = core.replace(".", "").replace(",", "")
        else:
            core = core.replace(",", "")
    elif has_dot:
        parts = core.split(".")
        if len(parts) > 2:
            if not re.match(r"^\d{1,3}(?:\.\d{3})+$", core):
                return AmountParse(None, notes + ["could not read this number"])
            # 1.234.567 -> European thousands
            core = core.replace(".", "")
            notes.append("read the dots as thousands separators (European style)")
        elif len(parts[-1]) == 3:
            if style == "eu":
                core = core.replace(".", "")
                notes.append("read the dot as a thousands separator (European style)")
            else:
                notes.append("ambiguous_thousands: could be 1234 or 1.234 -- "
                             "read as a decimal point")
    try:
        value = Decimal(core)
    except InvalidOperation:
        return AmountParse(None, notes + ["could not read this number"])
    if negative:
        value = -value
    return AmountParse(value, notes)


def money(d):
    """Decimal -> '1,234.56' for human reports. Negative as '-1,234.56'."""
    if d is None:
        return ""
    q = Decimal(d).quantize(Decimal("0.01"), ROUND_HALF_UP)
    sign = "-" if q < 0 else ""
    return "{}{:,}".format(sign, abs(q))


def plural(n, singular, plural_form=None):
    """'1 row' / '2 rows'. Small thing; looks careless when it is wrong."""
    if n == 1:
        return "%d %s" % (n, singular)
    return "%d %s" % (n, plural_form or (singular + "s"))


def money_plain(d):
    """Decimal -> '1234.56' for the output file (no separators, no symbols)."""
    if d is None:
        return ""
    q = Decimal(d).quantize(Decimal("0.01"), ROUND_HALF_UP)
    return "{:.2f}".format(q)


# ===========================================================================
# Section 3 - dates
# ===========================================================================

# (format, order, label). order is 'dmy', 'mdy' or 'none' (unambiguous).
DATE_FORMATS = [
    ("%Y-%m-%d %H:%M:%S", "none", "ISO date and time"),
    ("%Y-%m-%d %H:%M", "none", "ISO date and time"),
    ("%Y-%m-%dT%H:%M:%S", "none", "ISO date and time"),
    ("%Y-%m-%d", "none", "ISO (2025-03-31)"),
    ("%Y/%m/%d", "none", "year first (2025/03/31)"),
    ("%d/%m/%Y", "dmy", "day/month/year (31/03/2025)"),
    ("%m/%d/%Y", "mdy", "month/day/year (03/31/2025)"),
    ("%d-%m-%Y", "dmy", "day-month-year (31-03-2025)"),
    ("%m-%d-%Y", "mdy", "month-day-year (03-31-2025)"),
    ("%d.%m.%Y", "dmy", "day.month.year (31.03.2025)"),
    ("%m.%d.%Y", "mdy", "month.day.year (03.31.2025)"),
    ("%d/%m/%Y %H:%M", "dmy", "day/month/year with time"),
    ("%m/%d/%Y %H:%M", "mdy", "month/day/year with time"),
    ("%d/%m/%y", "dmy", "day/month/2-digit year"),
    ("%m/%d/%y", "mdy", "month/day/2-digit year"),
    ("%d-%m-%y", "dmy", "day-month-2-digit year"),
    ("%m-%d-%y", "mdy", "month-day-2-digit year"),
    ("%d.%m.%y", "dmy", "day.month.2-digit year"),
    ("%m.%d.%y", "mdy", "month.day.2-digit year"),
    ("%d %b %Y", "none", "textual (31 Mar 2025)"),
    ("%d %B %Y", "none", "textual (31 March 2025)"),
    ("%d-%b-%Y", "none", "textual (31-Mar-2025)"),
    ("%d %b, %Y", "none", "textual (31 Mar, 2025)"),
    ("%b %d, %Y", "none", "textual (Mar 31, 2025)"),
    ("%b %d %Y", "none", "textual (Mar 31 2025)"),
    ("%B %d, %Y", "none", "textual (March 31, 2025)"),
    ("%b %d, %y", "none", "textual (Mar 31, 25)"),
    ("%Y%m%d", "none", "compact (20250331)"),
]

EXCEL_EPOCH = dt.date(1899, 12, 30)


class NormalizerError(Exception):
    """Anything we can explain to the user in plain English."""


class AmbiguousDateError(NormalizerError):
    def __init__(self, message, examples=None):
        NormalizerError.__init__(self, message)
        self.examples = examples or []


def _try_formats(raw):
    """All (order, date, label) readings that work for one cell."""
    out = []
    text = raw.strip()
    if not text:
        return out
    for fmt, order, label in DATE_FORMATS:
        try:
            out.append((order, dt.datetime.strptime(text, fmt).date(), label, fmt))
        except ValueError:
            continue
    if not out and re.match(r"^\d{5}$", text):
        n = int(text)
        if 20000 <= n <= 80000:   # Excel/exports that lost their date formatting
            out.append(("none", EXCEL_EPOCH + dt.timedelta(days=n),
                        "Excel serial number (%s)" % text, "excel-serial"))
    return out


def classify_dates(raws, date_order=None):
    """Work out the day/month order for a whole column of dates.

    Returns (order, label, evidence, per_row) where per_row is a list of
    (iso_string_or_None, note) aligned with `raws`.

    Raises AmbiguousDateError instead of guessing when a cell could honestly be
    read either way (03/04/2025 is 3 April in the UK and 4 March in the US, and
    getting it wrong moves money into the wrong VAT/tax period).
    """
    parsed = []
    for raw in raws:
        text = "" if raw is None else str(raw).strip()
        if text == "":
            parsed.append(("blank", [], text))
            continue
        cands = _try_formats(text)
        if not cands:
            parsed.append(("unreadable", [], text))
            continue
        orders = set(c[0] for c in cands)
        parsed.append(("ok", cands, text))

    dmy_only, mdy_only, both, unreadable = [], [], [], []
    for kind, cands, text in parsed:
        if kind == "unreadable":
            unreadable.append(text)
            continue
        if kind != "ok":
            continue
        orders = set(c[0] for c in cands)
        if "dmy" in orders and "mdy" in orders:
            both.append(text)
        elif orders == {"dmy"}:
            dmy_only.append(text)
        elif orders == {"mdy"}:
            mdy_only.append(text)

    # An explicit instruction always wins; we only check that it can be applied.
    if date_order in ("dmy", "mdy"):
        broken = []
        for kind, cands, text in parsed:
            if kind != "ok":
                continue
            if date_order not in set(c[0] for c in cands):
                broken.append(text)
        if broken:
            raise NormalizerError(
                "You asked for --date-order {} but these dates cannot be read that "
                "way: {}. Check the flag, or send me the file.".format(
                    date_order, ", ".join(repr(b) for b in broken[:5])))
        order, source = date_order, "you asked for it (--date-order %s)" % date_order
        evidence = ("every date in this file could honestly be read either way, so I "
                    "used the order you gave me")
    elif dmy_only and mdy_only:
        raise NormalizerError(
            "This file mixes both date orders: {dmy} can only be day-first, but "
            "{mdy} can only be month-first. The rows may come from two different "
            "exports. Split the file, or fix the dates, then run it again.".format(
                dmy=", ".join(repr(x) for x in dmy_only[:3]),
                mdy=", ".join(repr(x) for x in mdy_only[:3])))
    elif date_order == "none":
        order, source, evidence = "none", "not needed", "no ambiguous dates"
    elif dmy_only:
        order, source = "dmy", "worked out from the dates themselves"
        evidence = ("%s cannot be a month, so the first number is the day "
                    "(example: %s)" % (repr(dmy_only[0]), dmy_only[0]))
    elif mdy_only:
        order, source = "mdy", "worked out from the dates themselves"
        evidence = ("%s cannot be a month, so the second number is the day "
                    "(example: %s)" % (repr(mdy_only[0]), mdy_only[0]))
    elif both:
        seen, examples = set(), []
        for e in both:                      # same date can appear on many rows
            if e not in seen:
                seen.add(e)
                examples.append(e)
        examples = examples[:6]
        raise AmbiguousDateError(
            "STOPPED BEFORE GUESSING: every date in this file could be read two "
            "ways (day/month/year or month/day/year). Example(s): {}. Reading "
            "'03/04/2025' as 3 April instead of 4 March moves that transaction "
            "into a different month, which is a real bookkeeping error. Tell me "
            "which one it is:\n"
            "    --date-order dmy     (day first, e.g. UK, Ireland, Australia, "
            "most of Europe)\n"
            "    --date-order mdy     (month first, e.g. US bank exports)".format(
                ", ".join(repr(e) for e in examples)),
            examples=examples)
    else:
        order, source, evidence = "none", "not needed", "no ambiguous dates"

    per_row = []
    for kind, cands, text in parsed:
        if kind == "blank":
            per_row.append((None, ""))
            continue
        if kind == "unreadable":
            per_row.append((None, "could not read the date %r" % text))
            continue
        pick = None
        for ordr, d, label, fmt in cands:
            if order in ("dmy", "mdy", "none") and ordr == order:
                pick = (d, label)
                break
        if pick is None:                      # order == 'none': take any reading
            pick = (cands[0][1], cands[0][2])
        note = ""
        if pick[1].startswith("Excel serial"):
            note = "date came through as a number, read as an Excel date"
        per_row.append((pick[0].isoformat(), note))
    return order, source, evidence, per_row, unreadable


# ===========================================================================
# Section 4 - reading files
# ===========================================================================

SUMMARY_KEYWORDS = [
    "closing balance", "opening balance", "ending balance", "beginning balance",
    "balance brought forward", "balance carried forward", "brought forward",
    "carried forward", "balance b/f", "balance c/f", "b/f", "c/f",
    "statement summary", "statement total", "summary", "totals", "total",
    "subtotal", "sub-total", "net total", "page total", "end of statement",
    "continued", "previous balance", "new balance", "account balance",
    "abschlusssaldo", "anfangssaldo", "kontostand", "uebertrag", "übertrag",
    "summe", "saldoauszug", "solde", "saldo final", "saldo inicial",
]
OPENING_KEYWORDS = ["opening balance", "beginning balance", "balance brought forward",
                    "balance b/f", "brought forward", "previous balance",
                    "anfangssaldo", "saldo inicial", "opening"]


def find_summary_keyword(text):
    low = " " + re.sub(r"\s+", " ", str(text)).strip().lower() + " "
    for kw in SUMMARY_KEYWORDS:
        if kw == "total":
            if re.search(r"\btotals?\b", low):
                return kw
        elif kw in ("b/f", "c/f"):
            if re.search(r"\bb/f\b", low) or re.search(r"\bc/f\b", low):
                return kw
        elif kw in low:
            return kw
    return None


def clean_text(v):
    if v is None:
        return ""
    s = str(v).replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def sniff_delimiter(text):
    """Guess the separator: comma, semicolon, tab or pipe.

    Scored by "how many lines agree on a field count", and ties are broken in
    favour of the separator that produces MORE columns. That matters for
    European files: in '1.234,56'-style amounts every comma is a decimal comma,
    so a naive comma-split also "works" but produces nonsense.
    """
    sample = "\n".join(text.splitlines()[:60])
    best, best_score = ",", (-1, -1)
    for d in [",", ";", "\t", "|"]:
        try:
            rows = list(csv.reader(sample.splitlines(), delimiter=d))
        except csv.Error:
            continue
        counts = [len(r) for r in rows if len(r) > 1]
        if not counts:
            continue
        modal = max(set(counts), key=counts.count)
        if modal < 2:
            continue
        agree = sum(1 for c in counts if c == modal)
        score = (agree, modal)
        if score > best_score:
            best, best_score = d, score
    return best, (best_score[0] if best_score[0] > 0 else -1)


class RawTable(object):
    def __init__(self, rows, kind, detail, sheets=None, sheet_used=None):
        self.rows = rows            # list of list of str
        self.kind = kind            # 'CSV' / 'Excel'
        self.detail = detail        # e.g. 'comma-separated, utf-8-sig'
        self.sheets = sheets or []
        self.sheet_used = sheet_used


def read_csv_file(path, encoding=None):
    last_err = None
    for enc in ([encoding] if encoding else ["utf-8-sig", "cp1252", "latin-1"]):
        try:
            with open(path, "r", encoding=enc, newline="") as fh:
                text = fh.read()
            break
        except (UnicodeDecodeError, LookupError) as exc:
            last_err = exc
    else:
        raise NormalizerError("I could not read the text in this file (tried "
                              "utf-8, cp1252 and latin-1). Last error: %s" % last_err)
    if text.strip() == "":
        raise NormalizerError("This file is empty.")
    delim, score = sniff_delimiter(text)
    if score <= 0:
        delim = ","
    rows = [list(r) for r in csv.reader(text.splitlines(), delimiter=delim)]
    names = {",": "comma", ";": "semicolon", "\t": "tab", "|": "pipe"}
    detail = "%s-separated, %s text" % (names.get(delim, repr(delim)), enc)
    if score <= 0:
        detail += " (only one column found - is this really a statement?)"
    return RawTable(rows, "CSV", detail)


def _cell_to_text(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if (v.hour, v.minute, v.second) == (0, 0, 0) \
            else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        s = "{:.10f}".format(v).rstrip("0").rstrip(".")
        return s if s else "0"
    return str(v)


def read_xlsx_file(path, sheet=None):
    try:
        import openpyxl
    except ImportError:
        raise NormalizerError(
            "Reading .xlsx files needs the 'openpyxl' add-on, which is not "
            "installed for this Python.\n"
            "  Fix it with:   pip install openpyxl\n"
            "  Or, in Excel:  File > Save As > CSV UTF-8, then run me on that.")
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        raise NormalizerError(
            "I could not open this .xlsx file (%s). If it was downloaded from "
            "your bank, try opening it in Excel and saving it again, or save it "
            "as CSV UTF-8." % exc)
    names = list(wb.sheetnames)
    ws = None
    if sheet is None:
        ws = wb[names[0]]
    elif isinstance(sheet, int) or str(sheet).isdigit():
        idx = int(sheet)
        if idx < 1 or idx > len(names):
            raise NormalizerError("There is no sheet number %s. This workbook has: %s"
                                  % (sheet, ", ".join(names)))
        ws = wb[names[idx - 1]]
    else:
        if sheet not in names:
            raise NormalizerError("There is no sheet called %r. This workbook has: %s"
                                  % (sheet, ", ".join(names)))
        ws = wb[sheet]
    rows = []
    for row in ws.iter_rows(values_only=True):
        cells = [_cell_to_text(v) for v in (row or ())]
        while cells and cells[-1].strip() == "":
            cells.pop()
        rows.append(cells)
    wb.close()
    if not rows:
        raise NormalizerError("The sheet %r is empty." % ws.title)
    return RawTable(rows, "Excel (.xlsx)",
                    "workbook sheet %r" % ws.title, sheets=names, sheet_used=ws.title)


def read_table(path, sheet=None, encoding=None):
    if not os.path.exists(path):
        raise NormalizerError("I cannot find a file called %r. Check the spelling, "
                              "or drag the file onto this window to get its path." % path)
    if os.path.isdir(path):
        raise NormalizerError("%r is a folder, not a file." % path)
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx_file(path, sheet=sheet)
    if ext == ".xls":
        raise NormalizerError(
            "Old-style .xls files are not supported (they are a completely "
            "different format). Open the file in Excel and use File > Save As > "
            "either .xlsx or CSV UTF-8, then run me on that.")
    return read_csv_file(path, encoding=encoding)


def detect_header_row(rows, max_scan=60):
    """Find the real header row under the bank's junk preamble.

    Scores each of the first `max_scan` rows: it must look like it names a date
    column AND some kind of amount column. Returns (index, mapping, extras,
    ignored, score) or (None, ...) if nothing looks like a header.
    """
    best = (None, None, None, [], -1)
    for i, row in enumerate(rows[:max_scan]):
        cells = [clean_text(c) for c in row]
        nonempty = [c for c in cells if c]
        if len(nonempty) < 2:
            continue
        mapping, extras, ignored = map_columns(cells)
        if FIELD_DATE not in mapping:
            continue
        if not any(f in mapping for f in (FIELD_AMOUNT, FIELD_DEBIT, FIELD_CREDIT)):
            continue
        score = len(mapping) * 10
        if FIELD_DESCRIPTION in mapping:
            score += 3
        if FIELD_BALANCE in mapping:
            score += 2
        # A real header row has far more text than numbers in it.
        numeric_cells = sum(1 for c in nonempty
                            if parse_amount(c).value is not None and not any(
                                ch.isalpha() for ch in c))
        score -= numeric_cells * 4
        if score > best[4]:
            best = (i, mapping, extras, ignored, score)
    return best


# ===========================================================================
# Section 5 - the actual normalisation
# ===========================================================================

class Transaction(object):
    __slots__ = ("date", "description", "amount", "balance", "source_row",
                 "notes", "sort_index")

    def __init__(self, date, description, amount, balance, source_row, notes):
        self.date = date
        self.description = description
        self.amount = amount
        self.balance = balance
        self.source_row = source_row
        self.notes = notes
        self.sort_index = source_row


class Report(object):
    def __init__(self):
        self.data = {
            "tool": "Power Query Bank Statement Normalizer",
            "version": VERSION,
            "source_file": None,
            "file_format": None,
            "sheet": None,
            "sheets_in_workbook": [],
            "header_row": None,
            "header_row_forced_by_user": False,
            "preamble_lines_skipped": 0,
            "column_mapping": {},
            "description_extras": [],
            "ignored_columns": [],
            "number_style": None,
            "number_style_evidence": {},
            "date_order": None,
            "date_order_source": None,
            "date_order_evidence": None,
            "row_order": None,
            "rows_read": 0,
            "transactions": 0,
            "rows_dropped": [],
            "rows_merged": [],
            "blank_lines": 0,
            "date_problems": [],
            "amount_notes": [],
            "warnings": [],
            "reconciliation": {},
            "balance_anchor": {"source_row": None, "balance": None},
            "output_file": None,
            "sort": "source",
            "status": "ok",
        }

    # -- convenience wrappers -------------------------------------------------
    def warn(self, msg):
        if msg not in self.data["warnings"]:
            self.data["warnings"].append(msg)

    def drop(self, row, reason):
        self.data["rows_dropped"].append({"source_row": row, "reason": reason})

    def __getitem__(self, k):
        return self.data[k]

    def __setitem__(self, k, v):
        self.data[k] = v


def build_amount_for_row(cells, mapping, header_index, row_index, style, report,
                         notes_bucket):
    """Return (amount_decimal_or_None, note) for one row.

    Debit/credit columns win when present: they are explicit about direction.
    If both are blank for a row, fall back to a signed amount column.
    """
    def cell(field):
        hdr = mapping.get(field)
        if hdr is None:
            return None
        try:
            idx = header_index.index(hdr)
        except ValueError:
            return None
        return cells[idx] if idx < len(cells) else None

    debit_raw, credit_raw = cell(FIELD_DEBIT), cell(FIELD_CREDIT)
    amount_raw = cell(FIELD_AMOUNT)
    note = ""

    if debit_raw is not None or credit_raw is not None:
        d = parse_amount(debit_raw, style)
        c = parse_amount(credit_raw, style)
        notes_bucket.extend(d.notes)
        notes_bucket.extend(c.notes)
        if d.value is not None and d.value != 0:
            return -abs(d.value), "debit column"
        if c.value is not None and c.value != 0:
            return abs(c.value), "credit column"
        if d.value == 0 or c.value == 0:
            # Both blank/zero: fall through to an amount column if there is one.
            if amount_raw is None or str(amount_raw).strip() == "":
                return Decimal("0.00"), "zero-value row in debit/credit columns"
    if amount_raw is not None:
        a = parse_amount(amount_raw, style)
        notes_bucket.extend(a.notes)
        if a.value is not None:
            return a.value, "amount column"
    return None, note


def find_index(header_cells, header_text):
    for i, c in enumerate(header_cells):
        if c == header_text:
            return i
    return None


def find_header_row_for(rows, mapping_override, max_scan=60):
    """When --map names the columns, find the row that holds them.

    Without this, an unusual layout would fail header detection before --map
    ever got a chance to help - which would make the flag useless exactly when
    it is needed.
    """
    wanted = [t for t in (mapping_override or {}).values() if t]
    named = set()
    for t in wanted:
        s = str(t).strip()
        if s and not s.isdigit():
            named.add(norm_header(s))
    if not named:
        return None
    best, best_hits = None, 0
    for i, row in enumerate(rows[:max_scan]):
        cells = set(norm_header(c) for c in row if clean_text(c))
        if not cells:
            continue
        hits = len(named & cells)
        if hits > best_hits:
            best, best_hits = i, hits
    if best is None:
        return None
    if best_hits < max(2, len(named) - 1):
        return None
    return best


def normalize_table(raw, report, date_order=None, row_order="auto",
                    sort="source", mapping_override=None, opening_balance=None,
                    header_row=None):
    """Turn a RawTable into (transactions, report). Raises NormalizerError."""
    rows = raw.rows
    report["file_format"] = raw.kind + " - " + raw.detail
    report["sheet"] = raw.sheet_used
    report["sheets_in_workbook"] = raw.sheets

    idx = None
    if header_row is not None:
        pos = int(header_row) - 1
        if pos < 0 or pos >= len(rows):
            raise NormalizerError(
                "You said the header row is line %s, but this file only has %d "
                "line(s)." % (header_row, len(rows)))
        idx = pos
        report["header_row_forced_by_user"] = True
    if idx is None and mapping_override:
        idx = find_header_row_for(rows, mapping_override)
    if idx is not None:
        mapping, extras, ignored = map_columns([clean_text(c) for c in rows[idx]])
    else:
        idx, mapping, extras, ignored, score = detect_header_row(rows)
    if idx is None:
        raise NormalizerError(
            "I read the file but could not find the header row - the line that "
            "names the columns (Date, Description, Amount, ...). I looked at the "
            "first 60 lines.\n"
            "  * If the column names are unusual, tell me where they are:\n"
            "        --map date=\"Book Date\",description=\"Narrative\",amount=\"Value\"\n"
            "  * If you also need to say which line they are on:\n"
            "        --header-row 3     (line numbers start at 1)\n"
            "  * If the file has no column names at all, add a first row with the "
            "names and try again.")
    header_cells = [clean_text(c) for c in rows[idx]]

    # An explicit --map overrides detection, so unrecognised layouts still work.
    if mapping_override:
        mapping = dict(mapping)
        for field, target in mapping_override.items():
            if target is None:
                mapping.pop(field, None)
                continue
            pos = find_index(header_cells, target)
            if pos is None and str(target).strip().isdigit():
                pos = int(str(target).strip()) - 1
            if pos is None or pos < 0 or pos >= len(header_cells):
                raise NormalizerError(
                    "In --map you said %s=%r, but there is no column called %r in "
                    "the header row: %s\n"
                    "Column numbers (like --map amount=4) also work."
                    % (field, target, target, " | ".join(h for h in header_cells if h)))
            mapping[field] = header_cells[pos]

    # Recompute which columns end up unused, so the report is truthful even
    # when --map has just changed the answer.
    used_headers = set(mapping.values())
    extras, ignored = [], []
    for c in header_cells:
        if not c or c in used_headers:
            continue
        if FIELD_DESCRIPTION in mapping and any(
                h in norm_header(c) for h in DESCRIPTION_EXTRA_HINTS):
            if c not in extras:
                extras.append(c)
        elif c not in ignored:
            ignored.append(c)

    report["header_row"] = idx + 1
    report["preamble_lines_skipped"] = idx
    report["column_mapping"] = {k: v for k, v in sorted(mapping.items())}
    report["description_extras"] = extras
    report["ignored_columns"] = ignored

    if FIELD_DESCRIPTION not in mapping:
        report.warn("I could not find a description/memo column, so the "
                    "description column in the output will be empty. Use "
                    "--map description=\"Your column name\" if there is one.")
    has_amount_src = any(f in mapping for f in (FIELD_AMOUNT, FIELD_DEBIT, FIELD_CREDIT))
    if not has_amount_src:
        raise NormalizerError(
            "The header row I found (%s) has no amount column I recognise. Use "
            "--map amount=\"...\" (or --map debit=\"...\" --map credit=\"...\") to "
            "tell me which column holds the money."
            % " | ".join(h for h in header_cells if h))

    body = rows[idx + 1:]
    body_start_row = idx + 2            # 1-based line number of the first body row

    # -- number style, decided from the whole file ---------------------------
    amount_cells = []
    for r in body:
        for field in (FIELD_AMOUNT, FIELD_DEBIT, FIELD_CREDIT):
            h = mapping.get(field)
            if h is None:
                continue
            p = find_index(header_cells, h)
            if p is not None and p < len(r):
                amount_cells.append(r[p])
    style, evidence = detect_number_style(amount_cells)
    report["number_style"] = style
    report["number_style_evidence"] = evidence
    if evidence.get("european_examples") and evidence.get("us_examples"):
        report.warn("This file shows BOTH number styles (%d amounts like 1.234,56 and "
                    "%d like 1,234.56). I used the %s one throughout. Check any amount "
                    "that looks the wrong size." % (evidence["european_examples"],
                                                    evidence["us_examples"], style))
    if style == "plain" and "mixed" not in str(style) and evidence.get("ambiguous_dot_thousands"):
        report.warn("Some amounts are written without thousands separators and look "
                    "like they could be European (1.234 = one thousand two hundred "
                    "and thirty-four). I read them as decimal points. If that is "
                    "wrong, check the amounts below.")

    # -- collect date strings so the whole column can be judged at once ------
    date_hdr = mapping.get(FIELD_DATE)
    date_pos = find_index(header_cells, date_hdr) if date_hdr else None
    date_raws = []
    for r in body:
        date_raws.append(r[date_pos] if date_pos is not None and date_pos < len(r) else "")

    order, order_source, order_evidence, per_row_date, unreadable = classify_dates(
        date_raws, date_order=date_order)
    report["date_order"] = order
    report["date_order_source"] = order_source
    report["date_order_evidence"] = order_evidence
    if unreadable:
        for u in unreadable[:10]:
            if u not in report["date_problems"]:
                report["date_problems"].append(u)

    # -- walk the body -------------------------------------------------------
    transactions = []
    notes_bucket = []
    anchor_row = None
    anchor_value = None
    last_txn = None

    for i, r in enumerate(body):
        line_no = body_start_row + i
        cells = list(r)
        while cells and str(cells[-1]).strip() == "":
            cells.pop()
        texts = [clean_text(c) for c in cells]
        nonempty = [t for t in texts if t]

        if not nonempty:
            report["blank_lines"] += 1
            continue

        report["rows_read"] += 1
        iso_date, date_note = per_row_date[i] if i < len(per_row_date) else (None, "")
        amount, amt_note = build_amount_for_row(cells, mapping, header_cells, line_no,
                                                style, report, notes_bucket)
        bal_hdr = mapping.get(FIELD_BALANCE)
        balance = None
        if bal_hdr is not None:
            bp = find_index(header_cells, bal_hdr)
            if bp is not None and bp < len(cells):
                balance = parse_amount(cells[bp], style).value

        joined = " ".join(nonempty)
        keyword = find_summary_keyword(joined)
        desc_parts = []
        if FIELD_DESCRIPTION in mapping:
            dp = find_index(header_cells, mapping[FIELD_DESCRIPTION])
            if dp is not None and dp < len(cells):
                desc_parts.append(texts[dp])
        for extra in extras:
            ep = find_index(header_cells, extra)
            if ep is not None and ep < len(cells) and texts[ep]:
                desc_parts.append(texts[ep])
        description = " | ".join([p for p in desc_parts if p])

        # 1. Opening-balance line: not a transaction, but a gift for the check.
        if keyword and keyword in OPENING_KEYWORDS and balance is not None and amount is None:
            anchor_row, anchor_value = line_no, balance
            report.drop(line_no, "statement opening-balance line (\"%s\") - kept as "
                                 "the starting point for the balance check" % joined[:60])
            continue

        # 2. A real transaction: it has a date and an amount.
        if iso_date and amount is not None:
            if date_note and date_note not in report["date_problems"]:
                report["date_problems"].append(date_note)
            txn = Transaction(iso_date, description, amount, balance, line_no, [])
            transactions.append(txn)
            last_txn = txn
            continue

        # 3. Summary / total / closing-balance / footer rows.
        if keyword:
            report.drop(line_no, "statement summary/footer line (matched \"%s\"): %r"
                        % (keyword, joined[:60]))
            continue

        # 4. Date but no amount: dangerous, say so loudly.
        if iso_date and amount is None:
            report.drop(line_no, "has a date (%s) but no amount I could read - "
                                 "%r. CHECK THIS LINE BY HAND: it may be a real "
                                 "transaction" % (iso_date, joined[:70]))
            continue

        # 5. An amount with no date at all.
        if amount is not None and not iso_date:
            report.drop(line_no, "has an amount but no date: %r - treated as a "
                                 "footer total, not a transaction" % joined[:70])
            continue

        # 6. A continuation of the transaction above it (multi-line description):
        #    the description column has text and every other column is empty.
        others = [t for j, t in enumerate(texts)
                  if t and j != find_index(header_cells, mapping.get(FIELD_DESCRIPTION))]
        if (description and last_txn is not None and balance is None and amount is None
                and not others):
            last_txn.description = (last_txn.description + " | " + description).strip(" |")
            report["rows_merged"].append({"source_row": line_no,
                                          "into_row": last_txn.source_row})
            continue

        report.drop(line_no, "I could not tell what this line is: %r" % joined[:70])

    report["transactions"] = len(transactions)
    report["amount_notes"] = sorted(set(notes_bucket))
    if any(n.startswith("ambiguous_thousands") for n in report["amount_notes"]):
        ambiguous = [n for n in report["amount_notes"] if n.startswith("ambiguous_thousands")]
        report.warn("Some amounts could be read two ways (for example '1.234' can mean "
                    "one thousand two hundred and thirty-four, or one point two three "
                    "four). I read them as decimal points. %s" % "; ".join(ambiguous[:3]))

    # -- chronology: many banks export newest-first ---------------------------
    dates = [t.date for t in transactions if t.date]
    detected = "unknown"
    if len(dates) >= 2:
        asc = all(dates[i] <= dates[i + 1] for i in range(len(dates) - 1))
        desc = all(dates[i] >= dates[i + 1] for i in range(len(dates) - 1))
        detected = "ascending" if asc else ("descending" if desc else "mixed")
    chosen = detected
    if row_order in ("asc", "desc"):
        chosen = "ascending" if row_order == "asc" else "descending"
    elif detected == "unknown":
        chosen = "ascending"
    report["row_order"] = chosen
    if detected == "mixed":
        report.warn("The dates in this file are not in order (they jump around). "
                    "The balance check follows the order of the rows in the file, "
                    "which is usually what the bank intended.")
    elif detected == "descending":
        report.warn("This file is newest-first (the dates go backwards down the "
                    "page). The balance check reads it from the bottom up so the "
                    "arithmetic still works.")

    # -- reconciliation ------------------------------------------------------
    report["balance_anchor"] = {
        "source_row": anchor_row,
        "balance": money_plain(anchor_value) if anchor_value is not None else None,
    }
    report["reconciliation"] = reconcile(transactions, anchor_row, anchor_value,
                                         opening_balance, report)

    # -- sorting -------------------------------------------------------------
    report["sort"] = sort
    if sort == "asc":
        transactions.sort(key=lambda t: (t.date or "", t.sort_index))
    return transactions, report


# ===========================================================================
# Section 6 - the reconciliation check (the reason people trust this)
# ===========================================================================

def _by_source(txns):
    return sorted(txns, key=lambda t: t.sort_index)


def reconcile(transactions, anchor_row, anchor_value, opening_balance, report):
    out = {
        "status": "not_possible",
        "balance_column": bool(report["column_mapping"].get(FIELD_BALANCE)),
        "anchor": None,
        "checks": 0,
        "passed": 0,
        "breaks": [],
        "unchecked_gaps": [],
        "totals": {},
        "sign_looks_inverted": False,
        "message": "",
    }
    t_in = sum((t.amount for t in transactions if t.amount > 0), Decimal("0"))
    t_out = sum((t.amount for t in transactions if t.amount < 0), Decimal("0"))
    out["totals"] = {
        "money_in": money_plain(t_in),
        "money_out": money_plain(t_out),
        "net_change": money_plain(t_in + t_out),
        "transactions": len(transactions),
        "with_balance": sum(1 for t in transactions if t.balance is not None),
    }

    if not out["balance_column"]:
        out["message"] = (
            "This statement has no balance column, so I cannot prove the running "
            "balance adds up. The rows themselves are complete: %s, "
            "%s in, %s out, net %s. To get the balance check as well, export the "
            "statement with a running/closing balance column if your bank offers "
            "one." % (plural(len(transactions), "transaction"), money(t_in),
                      money(abs(t_out)), money(t_in + t_out)))
        return out
    with_balance = len([t for t in transactions if t.balance is not None])
    if with_balance < 2 and not (with_balance and (opening_balance is not None
                                                  or anchor_value is not None)):
        out["message"] = ("The balance column is there but (almost) empty, so there "
                          "is nothing to check. Check the column mapping?")
        return out

    txns = _by_source(transactions)
    if report["row_order"] == "descending":
        txns = list(reversed(txns))

    anchor_txt = None
    if opening_balance is not None:
        anchor_value, anchor_txt = Decimal(str(opening_balance)), "the opening balance you gave me"
    elif anchor_value is not None:
        anchor_txt = "the opening-balance line at row %s" % anchor_row
    out["anchor"] = anchor_txt

    breaks, checks, passed, gaps = [], 0, 0, []
    prev = None
    if anchor_value is not None:
        prev = ("anchor", anchor_row, anchor_value)
    for t in txns:
        if t.balance is None:
            if prev is not None:
                gaps.append({"after_row": prev[1], "at_row": t.source_row,
                             "reason": "this row has no balance to check against"})
                prev = None
            continue
        if prev is None:
            prev = ("row", t.source_row, t.balance)
            continue
        expected = prev[2] + t.amount
        diff = t.balance - expected
        checks += 1
        if abs(diff) <= BALANCE_TOLERANCE:
            passed += 1
        else:
            breaks.append({
                "source_row": t.source_row,
                "previous_row": prev[1] if prev[0] == "row" else prev[1],
                "previous_kind": prev[0],
                "previous_balance": money_plain(prev[2]),
                "amount": money_plain(t.amount),
                "description": t.description[:70],
                "expected_balance": money_plain(expected),
                "actual_balance": money_plain(t.balance),
                "difference": money_plain(diff),
                "difference_value": diff,
            })
        prev = ("row", t.source_row, t.balance)

    out["checks"], out["passed"], out["breaks"], out["unchecked_gaps"] = \
        checks, passed, breaks, gaps

    # If the arithmetic works with the signs flipped, that is a big clue.
    if breaks and checks:
        flipped = 0
        prev = None
        if anchor_value is not None:
            prev = ("anchor", anchor_row, anchor_value)
        for t in txns:
            if t.balance is None:
                prev = None
                continue
            if prev is None:
                prev = ("row", t.source_row, t.balance)
                continue
            if abs(t.balance - (prev[2] - t.amount)) <= BALANCE_TOLERANCE:
                flipped += 1
            prev = ("row", t.source_row, t.balance)
        if flipped >= max(2, checks // 2):
            out["sign_looks_inverted"] = True

    if not breaks:
        first, last = txns[0], txns[-1]
        start_txt = (money(anchor_value) if anchor_value is not None
                     else "the first balance in the file (%s)" % money(first.balance))
        out["status"] = "ok"
        where = ("follows on from the row below it. (This statement is newest-first, "
                 "so I run the check from the bottom up.)"
                 if report["row_order"] == "descending"
                 else "follows on from the row above it.")
        out["message"] = (
            "PASS - every one of the %s I could check %s%s Start: %s, end: %s. "
            "Nothing is missing and no row was read with the wrong sign."
            % (plural(passed, "row"), where,
               (" Starting point: %s." % anchor_txt) if anchor_txt else "",
               start_txt, money(last.balance)))
    else:
        out["status"] = "breaks_found"
        out["message"] = (
            "%d of the %s did NOT add up. The output is still written, but do not "
            "trust the amounts on the rows listed below until you have looked at "
            "them." % (len(breaks), plural(checks, "balance check")))
    return out


# ===========================================================================
# Section 7 - output
# ===========================================================================

def write_csv(path, transactions, source_name):
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(OUT_COLUMNS)
        for t in transactions:
            w.writerow([t.date or "", t.description, money_plain(t.amount),
                        money_plain(t.balance) if t.balance is not None else "",
                        source_name, t.source_row])


def write_xlsx(path, transactions, source_name):
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment
    except ImportError:
        raise NormalizerError("Writing .xlsx needs 'openpyxl' (pip install openpyxl). "
                             "Use --out something.csv instead.")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Normalised statement"
    ws.append(OUT_COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True)
    for t in transactions:
        ws.append([t.date or "", t.description,
                   float(t.amount) if t.amount is not None else None,
                   float(t.balance) if t.balance is not None else None,
                   source_name, t.source_row])
    for col, width in zip("ABCDEF", (12, 60, 14, 14, 28, 12)):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=4):
        for c in row:
            c.number_format = "#,##0.00"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:F%d" % max(1, ws.max_row)
    wb.save(path)


def write_json_report(path, report):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(report.data, fh, indent=2, ensure_ascii=False, default=str)


# ===========================================================================
# Section 8 - the friendly screen report
# ===========================================================================

RULE = "=" * 78


def _kv(label, value, width=15):
    return "  {:<{w}} {}".format(label + ":", value, w=width)


def print_report(report, transactions, verbose=False):
    d = report.data
    print(RULE)
    print(" Power Query Bank Statement Normalizer {}".format(VERSION))
    print(RULE)
    print(_kv("File", "{}".format(d["source_file"])))
    print(_kv("Read as", "{}".format(d["file_format"])))
    if d["sheet"]:
        extra = ""
        if len(d["sheets_in_workbook"]) > 1:
            extra = "  (this workbook also has: {})".format(
                ", ".join(s for s in d["sheets_in_workbook"] if s != d["sheet"]))
        print(_kv("Sheet", "{}{}".format(d["sheet"], extra)))
    print()
    print(" WHAT I FOUND")
    print(_kv("Header row", "line {}".format(d["header_row"])
              + ("  (you told me with --header-row)"
                 if d["header_row_forced_by_user"]
                 else ("  (skipped {} line(s) of bank preamble above it)".format(
                     d["preamble_lines_skipped"]) if d["preamble_lines_skipped"] else ""))))
    print("  Columns I am using:")
    labels = {
        FIELD_DATE: "date",
        FIELD_DESCRIPTION: "description",
        FIELD_DEBIT: "money out (debit)",
        FIELD_CREDIT: "money in (credit)",
        FIELD_AMOUNT: "amount (signed)",
        FIELD_BALANCE: "balance",
    }
    for field in (FIELD_DATE, FIELD_DESCRIPTION, FIELD_DEBIT, FIELD_CREDIT,
                  FIELD_AMOUNT, FIELD_BALANCE):
        if field in d["column_mapping"]:
            print("      {:<22} -> \"{}\"".format(labels[field], d["column_mapping"][field]))
    if d["description_extras"]:
        print("      {:<22} -> {}".format("glued onto description",
                                          ", ".join('"%s"' % e for e in d["description_extras"])))
    if d["ignored_columns"]:
        print(_kv("Not used", ", ".join('"%s"' % c for c in d["ignored_columns"])))
    style_txt = {
        "eu": "European (1.234,56 - the dot is thousands, the comma is pennies)",
        "us": "UK/US (1,234.56 - the comma is thousands, the dot is pennies)",
        "plain": "plain numbers (1234.56, no thousands separators)",
        "mixed": "mixed - see the numbers below",
    }.get(d["number_style"], str(d["number_style"]))
    print(_kv("Numbers", style_txt))
    ev = d["number_style_evidence"]
    if ev.get("european_examples") or ev.get("us_examples"):
        print("      evidence: {} amount(s) written like 1.234,56 and {} written like "
              "1,234.56".format(ev.get("european_examples", 0), ev.get("us_examples", 0)))
    order_txt = {"dmy": "day first (31/03/2025 = 31 March)",
                 "mdy": "month first (03/31/2025 = 31 March)",
                 "none": "not needed - the dates spell out their month, or are ISO"}[d["date_order"]]
    if d["date_order"] == "none":
        print(_kv("Dates", order_txt))
    else:
        print(_kv("Dates", "{} - {}".format(order_txt, d["date_order_source"])))
    if d["date_order_evidence"] and d["date_order"] != "none":
        print("      {}".format(d["date_order_evidence"]))
    print(_kv("Row order", d["row_order"] + (" (read bottom-up for the balance check)"
                                             if d["row_order"] == "descending" else "")))
    print(_kv("Sorting", "output kept in the file's own order" if d["sort"] == "source"
              else "output sorted oldest-first"))
    print()
    print(" WHAT I READ")
    print(_kv("Transaction lines", len(transactions)))
    print(_kv("Lines dropped", len(d["rows_dropped"])))
    for item in d["rows_dropped"]:
        print("      line {:<5} {}".format(item["source_row"], item["reason"]))
    if d["rows_merged"]:
        print(_kv("Lines joined up", len(d["rows_merged"])))
        for m in d["rows_merged"]:
            print("      line {:<5} extra text for the transaction on line {} "
                  "(multi-line description)".format(m["source_row"], m["into_row"]))
    if d["blank_lines"]:
        print(_kv("Blank lines", d["blank_lines"]))
    if d["date_problems"] and verbose:
        print(_kv("Date oddities", "; ".join(d["date_problems"][:6])))
    if d["amount_notes"] and verbose:
        print(_kv("Number notes", "; ".join(d["amount_notes"][:8])))
    print()
    r = d["reconciliation"]
    print(" RECONCILIATION - proof nothing was lost")
    print(_kv("Balance column", "yes - \"%s\"" % d["column_mapping"].get(FIELD_BALANCE)
              if r["balance_column"] else "NO - this statement has no balance column"))
    if r["balance_column"]:
        print(_kv("Starting point", r["anchor"] or
                  "no opening balance available, so the check starts at the second row"))
        print(_kv("Checks run", "{} - {} passed, {} did not add up".format(
            plural(r["checks"], "balance check"), r["passed"], len(r["breaks"]))))
    t = r["totals"]
    print(_kv("Totals", "{} in, {} out, net {} over {} transactions".format(
        money(Decimal(t["money_in"])), money(abs(Decimal(t["money_out"]))),
        money(Decimal(t["net_change"])), t["transactions"])))
    if r["breaks"]:
        print()
        print("  ROWS THAT DO NOT ADD UP")
        for b in r["breaks"][:25]:
            print("    line {}: {}  amount {}".format(
                b["source_row"], b["description"] or "(no description)", b["amount"]))
            print("        balance above (line {}): {}   +   amount: {}   =   {}".format(
                b["previous_row"], money(Decimal(b["previous_balance"])), b["amount"],
                money(Decimal(b["expected_balance"]))))
            print("        but the file says the balance here is {}  ->  difference {}"
                  .format(money(Decimal(b["actual_balance"])),
                          money(Decimal(b["difference"]))))
        if len(r["breaks"]) > 25:
            print("    ... and {} more.".format(len(r["breaks"]) - 25))
        if r["sign_looks_inverted"]:
            print()
            print("  LIKELY CAUSE: the signs in this file look back-to-front. The "
                  "arithmetic")
            print("  works if money IN is treated as money OUT. Re-run with "
                  "--invert-amounts.")
    if r["unchecked_gaps"]:
        print("  Rows I could not check (no balance on them): {}".format(
            ", ".join(str(g["at_row"]) for g in r["unchecked_gaps"][:10])))
    print()
    verdict = {"ok": "PASS", "breaks_found": "PROBLEM FOUND",
               "not_possible": "COULD NOT CHECK"}[r["status"]]
    print("  RESULT: " + verdict)
    for line in _wrap(r["message"], 74):
        print("    " + line)
    print()
    if d["warnings"]:
        print(" THINGS TO LOOK AT")
        for w in d["warnings"]:
            for i, line in enumerate(_wrap(w, 72)):
                print(("  * " if i == 0 else "    ") + line)
        print()
    if d["output_file"]:
        print(_kv("Clean file written", d["output_file"]))
    print(RULE)


def _wrap(text, width):
    words, lines, cur = str(text).split(), [], ""
    for w in words:
        if len(cur) + len(w) + 1 > width:
            lines.append(cur)
            cur = w
        else:
            cur = (cur + " " + w).strip()
    if cur:
        lines.append(cur)
    return lines


# ===========================================================================
# Section 9 - command line
# ===========================================================================

def parse_map_option(values):
    """--map date=Date,description=Memo,amount=Value  ->  {'date': 'Date', ...}"""
    mapping = {}
    if not values:
        return mapping
    for chunk in values:
        for part in chunk.split(","):
            part = part.strip()
            if not part:
                continue
            if "=" not in part:
                raise NormalizerError(
                    "--map needs field=column pairs, e.g. --map date=\"Book Date\","
                    "amount=\"Value\". I got %r." % part)
            field, target = part.split("=", 1)
            field = field.strip().lower()
            target = target.strip().strip('"').strip("'")
            if field not in (FIELD_DATE, FIELD_DESCRIPTION, FIELD_AMOUNT, FIELD_DEBIT,
                             FIELD_CREDIT, FIELD_BALANCE):
                raise NormalizerError(
                    "--map can set these fields: date, description, amount, debit, "
                    "credit, balance. You wrote %r." % field)
            mapping[field] = None if target.lower() in ("none", "ignore", "-") else target
    return mapping


def build_parser():
    p = argparse.ArgumentParser(
        prog="normalize.py",
        description="Turn a messy bank statement export into one clean table - and "
                    "check the running balance adds up so you know nothing was lost.",
        epilog="Example:  python3 normalize.py statement.csv --date-order dmy\n"
               "          python3 normalize.py january.xlsx --out january-clean.csv "
               "--json-report january-report.json",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", help="the statement file (.csv, .xlsx or .xlsm)")
    p.add_argument("--out", "-o", default=None,
                   help="where to write the clean table; .csv or .xlsx decides the "
                        "format (default: <name>-normalized.csv next to the input)")
    p.add_argument("--date-order", choices=["dmy", "mdy"], default=None,
                   help="how to read unclear dates: dmy = day first (31/03/2025), "
                        "mdy = month first (03/31/2025). Only needed when a file's "
                        "dates could honestly be either.")
    p.add_argument("--json-report", default=None, metavar="FILE",
                   help="also write the full findings as a JSON file")
    p.add_argument("--map", action="append", default=None, metavar="FIELD=COLUMN",
                   help="tell me which column is which, e.g. "
                        "--map \"date=Book Date,description=Narrative,amount=Value\". "
                        "Column numbers work too: --map amount=4")
    p.add_argument("--header-row", type=int, default=None, metavar="N",
                   help="the line number of the row that names the columns, when "
                        "it cannot be worked out (use with --map). The first line "
                        "of the file is 1.")
    p.add_argument("--sheet", default=None,
                   help="for .xlsx files with several sheets: a sheet name or number")
    p.add_argument("--encoding", default=None,
                   help="force a text encoding for CSV files (utf-8, cp1252, latin-1)")
    p.add_argument("--row-order", choices=["auto", "asc", "desc"], default="auto",
                   help="whether the file runs oldest-first (asc) or newest-first "
                        "(desc). Default auto-detects from the dates.")
    p.add_argument("--sort", choices=["source", "asc"], default="source",
                   help="source (default) keeps the statement's own line order; asc "
                        "sorts the clean table oldest-first")
    p.add_argument("--opening-balance", default=None, metavar="AMOUNT",
                   help="the balance before the first row, so the first row can be "
                        "checked too, e.g. --opening-balance 2500.00")
    p.add_argument("--invert-amounts", action="store_true",
                   help="flip every sign (use when the reconciliation says the file's "
                        "signs look back-to-front)")
    p.add_argument("--dry-run", action="store_true",
                   help="check everything and report, but do not write any file")
    p.add_argument("--verbose", "-v", action="store_true",
                   help="show the extra detail (date oddities, number notes)")
    p.add_argument("--quiet", "-q", action="store_true",
                   help="print only the result line and any problems")
    p.add_argument("--version", action="version",
                   version="Power Query Bank Statement Normalizer " + VERSION)
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    report = Report()
    src = args.input
    report["source_file"] = os.path.basename(src)
    try:
        mapping_override = parse_map_option(args.map)
        opening = None
        if args.opening_balance is not None:
            opening = parse_amount(args.opening_balance, "plain").value
            if opening is None:
                raise NormalizerError("I could not read --opening-balance %r as a "
                                      "number." % args.opening_balance)
        raw = read_table(src, sheet=args.sheet, encoding=args.encoding)
        txns, report = normalize_table(
            raw, report, date_order=args.date_order, row_order=args.row_order,
            sort=args.sort, mapping_override=mapping_override, opening_balance=opening,
            header_row=args.header_row)
        for t in txns:
            if args.invert_amounts:
                t.amount = -t.amount
        if args.invert_amounts:
            anchor = report["balance_anchor"] or {}
            anchor_bal = (Decimal(anchor["balance"])
                          if anchor.get("balance") is not None else None)
            report["reconciliation"] = reconcile(
                txns, anchor.get("source_row"), anchor_bal, opening, report)
            report["reconciliation"]["message"] += (" (signs flipped as requested "
                                                    "with --invert-amounts)")
            report.warn("You asked me to flip the signs (--invert-amounts).")

        out = args.out
        if out is None:
            base = os.path.splitext(src)[0] + "-normalized.csv"
            out = base
        if not args.dry_run:
            ext = os.path.splitext(out)[1].lower()
            if ext in (".xlsx", ".xlsm"):
                write_xlsx(out, txns, report["source_file"])
            elif ext == ".xls":
                raise NormalizerError("Excel .xls cannot be written; use .xlsx or .csv.")
            else:
                write_csv(out, txns, report["source_file"])
            report["output_file"] = os.path.abspath(out)
        if args.json_report:
            rep = report.data
            if args.dry_run:
                rep = dict(rep)
                rep["output_file"] = None
            write_json_report(args.json_report, report)

        if args.quiet:
            print(report["reconciliation"]["message"])
            if report["reconciliation"]["status"] != "ok":
                print_report(report, txns, verbose=args.verbose)
        else:
            print_report(report, txns, verbose=args.verbose)

        if report["reconciliation"]["status"] == "breaks_found":
            return 1
        return 0

    except AmbiguousDateError as exc:
        print(RULE)
        print(" I NEED ONE PIECE OF INFORMATION BEFORE I TOUCH YOUR NUMBERS")
        print(RULE)
        for line in str(exc).split("\n"):
            for i, w in enumerate(_wrap(line, 74)):
                print(("  " if i == 0 else "  ") + w)
        print()
        print("  Example of how to answer:")
        print("      python3 normalize.py \"{}\" --date-order dmy".format(
            os.path.basename(src)))
        print(RULE)
        return 2
    except NormalizerError as exc:
        print(RULE)
        print(" I COULD NOT DO THIS - HERE IS WHY")
        print(RULE)
        for line in str(exc).split("\n"):
            for w in _wrap(line, 74):
                print("  " + w)
        print(RULE)
        return 1
    except KeyboardInterrupt:
        print("\nStopped.")
        return 1
    except Exception as exc:                      # never a raw traceback
        print(RULE)
        print(" SOMETHING UNEXPECTED HAPPENED")
        print(RULE)
        print("  The technical detail is: {}: {}".format(type(exc).__name__, exc))
        print("  Your file is probably fine - it is just a shape I have not seen.")
        print("  Nothing was written to disk unless it said so above.")
        print("  Please send the statement (or the first 10 lines of it) to the "
              "address in the README and I will add the layout.")
        print(RULE)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
