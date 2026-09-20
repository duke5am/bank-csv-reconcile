"""Turn a bank statement export into one clean table, and prove nothing was lost.

The whole tool is the command line: :func:`bank_csv_reconcile.cli.main` is what
the ``bank-csv-reconcile`` console script calls, and what the repo-root
``normalize.py`` wrapper calls.

Deliberately no third-party imports here or in ``cli.py`` at module level: CSV in,
CSV out works on a bare Python install. ``openpyxl`` is imported inside the two
functions that actually touch .xlsx (``read_xlsx_file`` and ``write_xlsx``), so it
is an optional extra (``pip install "bank-csv-reconcile[excel]"``), not a
requirement.
"""
# The distribution version (pyproject.toml / PyPI). The banner the CLI prints
# shows the tool's own VERSION string from cli.py, which is a separate thing.
__version__ = "0.1.0"
__all__ = ["main", "__version__"]


def __getattr__(name):
    """Expose ``bank_csv_reconcile.main`` without importing ``cli`` eagerly.

    A plain ``from .cli import main`` here would put ``bank_csv_reconcile.cli`` in
    sys.modules as a side effect of importing the package, so running
    ``python3 -m bank_csv_reconcile.cli`` would then load the whole module twice
    and print a runpy RuntimeWarning about it. Importing it on first attribute
    access keeps ``import bank_csv_reconcile`` cheap and ``-m`` clean.
    """
    if name == "main":
        from .cli import main
        return main
    raise AttributeError("module {!r} has no attribute {!r}".format(__name__, name))

