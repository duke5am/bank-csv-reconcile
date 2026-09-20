#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""normalize.py - keep `python3 normalize.py statement.csv` working from a clone.

This wrapper exists so the workflow the README documents from a checkout keeps
working unchanged. The same CLI is installed as the `bank-csv-reconcile` console
script; the implementation lives in `bank_csv_reconcile/cli.py` so that the
installed package and the checkout are the same code, not two versions of it.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from bank_csv_reconcile.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
