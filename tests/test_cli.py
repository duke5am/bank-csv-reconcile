# -*- coding: utf-8 -*-
"""End-to-end tests for the packaged ``bank-csv-reconcile`` command line.

These drive the real CLI in a subprocess - the same ``main()`` that the
``bank-csv-reconcile`` console script and the repo-root ``normalize.py`` wrapper
call - so a wrong exit code, a missing output file or a raw Python traceback is
caught here rather than by a user.

Run from the repository root:

    python3 -m unittest discover -s tests -v

Scratch files are created under /root (NOT /tmp, which this environment treats as
volatile) and are removed in tearDown.
"""
import io
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = os.path.join(REPO_ROOT, "samples")
PYPROJECT = os.path.join(REPO_ROOT, "pyproject.toml")
WRAPPER = os.path.join(REPO_ROOT, "normalize.py")

OUT_HEADER = "date,description,amount,balance,source_file,source_row"

# Where scratch files go. /root by default (as this environment requires); fall
# back to the system temp dir only if /root is genuinely not writable.
_PREFERRED_SCRATCH = os.environ.get("BANK_CSV_RECONCILE_SCRATCH") or "/root"
if not (os.path.isdir(_PREFERRED_SCRATCH) and os.access(_PREFERRED_SCRATCH, os.W_OK)):
    _PREFERRED_SCRATCH = None

# A stand-in module that fails the way an uninstalled openpyxl does.
_NO_OPENPYXL = 'raise ImportError("openpyxl is not installed here")\n'


def _sample(name):
    return os.path.join(SAMPLES, name)


class CLITestCase(unittest.TestCase):
    """Base class: one scratch directory per test, always cleaned up."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bank-csv-reconcile-test-",
                                    dir=_PREFERRED_SCRATCH)
        # cwd for every subprocess: deliberately NOT the source tree, so a test
        # that only passes because the checkout happens to be the cwd would fail.
        self.workdir = os.path.join(self.tmp, "work")
        os.makedirs(self.workdir)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    # -- helpers -----------------------------------------------------------
    def scratch(self, name):
        return os.path.join(self.tmp, name)

    @staticmethod
    def squash(text):
        """Collapse whitespace, so an assertion is not defeated by line wrapping.

        The messages are wrapped to the terminal width on purpose; a test that
        asserts on a phrase must not care where the wrap happens to land.
        """
        return " ".join((text or "").split())

    def write_at(self, path, text):
        with io.open(path, "w", encoding="utf-8") as fh:
            fh.write(text)
        return path

    def write_bytes(self, name, data):
        path = self.scratch(name)
        with open(path, "wb") as fh:
            fh.write(data)
        return path

    def write_text(self, name, text):
        return self.write_bytes(name, text.encode("utf-8"))

    def no_openpyxl_dir(self):
        """A PYTHONPATH entry whose openpyxl import fails, like an absent extra."""
        stub_dir = os.path.join(self.tmp, "no-openpyxl")
        os.makedirs(stub_dir)
        self.write_at(os.path.join(stub_dir, "openpyxl.py"), _NO_OPENPYXL)
        return stub_dir

    def _env(self, extra_path=None):
        env = dict(os.environ)
        parts = [p for p in (extra_path, REPO_ROOT) if p]
        old = env.get("PYTHONPATH")
        if old:
            parts.append(old)
        env["PYTHONPATH"] = os.pathsep.join(parts)
        env["PYTHONIOENCODING"] = "utf-8"
        return env

    def run_packaged(self, *args, **kwargs):
        """Run the packaged CLI the way the console script runs it."""
        extra_path = kwargs.pop("extra_path", None)
        assert not kwargs, kwargs
        return subprocess.run(
            [sys.executable, "-m", "bank_csv_reconcile.cli"] + list(args),
            cwd=self.workdir, env=self._env(extra_path),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            encoding="utf-8", errors="replace", timeout=180)

    def run_wrapper(self, *args):
        """Run the documented checkout workflow: python3 normalize.py ..."""
        return subprocess.run(
            [sys.executable, WRAPPER] + list(args),
            cwd=self.workdir, env=self._env(),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            encoding="utf-8", errors="replace", timeout=180)

    def assert_clean(self, proc):
        """No Python traceback anywhere in the output, and something was said."""
        combined = (proc.stdout or "") + (proc.stderr or "")
        self.assertNotIn("Traceback (most recent call last)", combined,
                         "the CLI printed a Python traceback:\n" + combined)
        self.assertTrue(combined.strip(), "the CLI printed nothing at all")
        return combined

    def read_out(self, path):
        """Read a clean table back (write_csv writes a UTF-8 BOM)."""
        with io.open(path, "r", encoding="utf-8-sig", newline="") as fh:
            return fh.read().splitlines()


class PositiveCaseTests(CLITestCase):
    """The four tracked samples, each behaving as the README documents."""

    def run_sample(self, sample, label):
        out = self.scratch(label + "-clean.csv")
        proc = self.run_packaged(_sample(sample), "--out", out)
        text = self.assert_clean(proc)
        self.assertIn("Transaction lines:", text)
        return proc, text, out

    def test_sample1_uk_debit_credit_columns(self):
        proc, text, out = self.run_sample("sample1-uk-debit-credit-columns.csv", "s1")
        self.assertEqual(proc.returncode, 0)
        self.assertIn("RESULT: PASS", text)
        self.assertIn("Transaction lines: 12", text)
        self.assertIn("Checks run:     12 balance checks - 12 passed, 0 did not "
                      "add up", text)
        self.assertIn('money out (debit)', text)
        self.assertIn('-> "Debit"', text)
        self.assertIn('money in (credit)', text)
        self.assertIn('-> "Credit"', text)
        self.assertIn('date                   -> "Date"', text)
        lines = self.read_out(out)
        self.assertEqual(lines[0], OUT_HEADER)
        self.assertEqual(len(lines), 13)  # header + 12 transactions
        self.assertTrue(lines[1].startswith("2025-03-03,CARD PAYMENT"))

    def test_sample2_us_signed_amount(self):
        proc, text, out = self.run_sample("sample2-us-signed-amount.csv", "s2")
        self.assertEqual(proc.returncode, 0)
        self.assertIn("RESULT: PASS", text)
        self.assertIn("Transaction lines: 12", text)
        self.assertIn("Checks run:     12 balance checks - 12 passed, 0 did not "
                      "add up", text)
        # US dates are proven month-first from the data itself, not guessed.
        self.assertIn("Dates:          month first", text)
        self.assertEqual(len(self.read_out(out)), 13)

    def test_sample3_european_format(self):
        proc, text, out = self.run_sample("sample3-european-format.csv", "s3")
        self.assertEqual(proc.returncode, 0)
        self.assertIn("RESULT: PASS", text)
        self.assertIn("Transaction lines: 9", text)
        self.assertIn("Checks run:     9 balance checks - 9 passed, 0 did not "
                      "add up", text)
        self.assertIn("Numbers:        European (1.234,56", text)
        self.assertEqual(len(self.read_out(out)), 10)  # header + 9 transactions

    def test_sample6_corrupted_balance_exits_1_and_still_writes_output(self):
        """The deliberate bad-balance example: exit 1, and the row is named."""
        proc, text, out = self.run_sample("sample6-corrupted-balance.csv", "s6")
        # Exit code 1 is documented as "a balance did not add up".
        self.assertEqual(proc.returncode, 1)
        self.assertIn("Transaction lines: 12", text)
        self.assertIn("Checks run:     12 balance checks - 10 passed, 2 did not "
                      "add up", text)
        self.assertIn("RESULT: PROBLEM FOUND", text)
        self.assertIn("2 of the 12 balance checks did NOT add up", text)
        self.assertIn("ROWS THAT DO NOT ADD UP", text)
        self.assertIn("line 17", text)
        self.assertIn("line 18", text)
        # The output is still written when a balance breaks, as documented.
        self.assertTrue(os.path.isfile(out))
        self.assertEqual(len(self.read_out(out)), 13)

    def test_default_output_path_is_next_to_the_input(self):
        """README: with no --out, the clean table lands beside the statement."""
        src = self.scratch("statement.csv")
        shutil.copyfile(_sample("sample1-uk-debit-credit-columns.csv"), src)
        proc = self.run_packaged(src)
        self.assert_clean(proc)
        self.assertEqual(proc.returncode, 0)
        expected = self.scratch("statement-normalized.csv")
        self.assertTrue(os.path.isfile(expected),
                        "expected the default output at " + expected)
        self.assertEqual(self.read_out(expected)[0], OUT_HEADER)

    def test_json_report_and_quiet_are_writable(self):
        out = self.scratch("s1-clean.csv")
        rep = self.scratch("s1-report.json")
        proc = self.run_packaged(_sample("sample1-uk-debit-credit-columns.csv"),
                                 "--out", out, "--json-report", rep, "--quiet")
        self.assert_clean(proc)
        self.assertEqual(proc.returncode, 0)
        self.assertTrue(os.path.isfile(rep))
        with io.open(rep, "r", encoding="utf-8") as fh:
            self.assertIn("reconciliation", fh.read())

    def test_documented_exit_code_2_when_dates_are_ambiguous(self):
        """README: 2 = a decision is needed first (--date-order)."""
        src = self.write_text(
            "ambig.csv",
            "Date,Description,Amount,Balance\n"
            "01/03/2025,BANK FEE,-5.00,100.00\n"
            "02/03/2025,INTEREST,1.00,101.00\n")
        proc = self.run_packaged(src, "--dry-run")
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 2)
        self.assertIn("STOPPED BEFORE GUESSING", text)
        self.assertIn("--date-order", text)
        # ...and answering the question makes it work.
        proc2 = self.run_packaged(src, "--dry-run", "--date-order", "dmy")
        self.assert_clean(proc2)
        self.assertEqual(proc2.returncode, 0)


class NegativeCaseTests(CLITestCase):
    """Bad input: non-zero exit, a clear message, never a traceback."""

    def test_missing_file(self):
        missing = self.scratch("no-such-statement.csv")
        proc = self.run_packaged(missing)
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("I cannot find a file called", self.squash(text))
        self.assertIn("no-such-statement.csv", text)
        self.assertIn("Check the spelling", self.squash(text))

    def test_malformed_csv(self):
        """An unbalanced quote swallows the header line, so nothing can be read."""
        bad = self.write_bytes(
            "malformed.csv",
            b'Date,Description,"Amount,Balance\n01/03/2025,"unclosed,12.00\n')
        proc = self.run_packaged(bad, "--dry-run")
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("I COULD NOT DO THIS - HERE IS WHY", text)
        self.assertIn("could not find the header row", self.squash(text))
        # ...and it tells the user what to do next, in plain language.
        self.assertIn("--map", text)
        self.assertIn("--header-row", text)

    def test_file_with_unrecognised_columns(self):
        bad = self.write_text("unrecognised.csv", "Foo,Bar\n1,2\n3,4\n")
        proc = self.run_packaged(bad, "--dry-run")
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("could not find the header row", self.squash(text))

    def test_empty_directory(self):
        empty = os.path.join(self.tmp, "empty-dir")
        os.makedirs(empty)
        self.assertEqual(os.listdir(empty), [])
        proc = self.run_packaged(empty)
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("is a folder, not a file", self.squash(text))
        self.assertIn("empty-dir", text)

    def test_empty_file(self):
        empty = self.write_bytes("empty.csv", b"")
        proc = self.run_packaged(empty, "--dry-run")
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("This file is empty.", text)

    def test_bad_map_option(self):
        proc = self.run_packaged(_sample("sample1-uk-debit-credit-columns.csv"),
                                 "--dry-run", "--map", "nosuchfield=Date")
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("--map can set these fields", self.squash(text))

    def test_bad_opening_balance(self):
        proc = self.run_packaged(_sample("sample1-uk-debit-credit-columns.csv"),
                                 "--dry-run", "--opening-balance", "not-a-number")
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("could not read --opening-balance", self.squash(text))


class PackagingTests(CLITestCase):
    """The parts that only break once the wheel is actually installed."""

    def test_console_script_entry_point_target_is_importable(self):
        with io.open(PYPROJECT, "r", encoding="utf-8") as fh:
            pyproject = fh.read()
        self.assertIn('bank-csv-reconcile = "bank_csv_reconcile.cli:main"',
                      pyproject)

        sys.path.insert(0, REPO_ROOT)
        try:
            import bank_csv_reconcile
            from bank_csv_reconcile import cli
        finally:
            sys.path.remove(REPO_ROOT)
        self.assertTrue(callable(cli.main))
        self.assertIs(bank_csv_reconcile.main, cli.main)

        # main() must RETURN an exit code (the console script wraps it in
        # sys.exit) rather than raising SystemExit on the normal path.
        out = self.scratch("entry-clean.csv")
        buf = io.StringIO()
        real_stdout = sys.stdout
        sys.stdout = buf
        try:
            code = cli.main([_sample("sample1-uk-debit-credit-columns.csv"),
                             "--out", out])
        finally:
            sys.stdout = real_stdout
        self.assertIsInstance(code, int)
        self.assertEqual(code, 0)
        self.assertTrue(os.path.isfile(out))

    def test_csv_path_needs_no_openpyxl(self):
        """The default CSV-in/CSV-out path must work with openpyxl absent."""
        stub_dir = self.no_openpyxl_dir()
        driver = self.write_at(
            self.scratch("driver.py"),
            "import sys\n"
            "from bank_csv_reconcile.cli import main\n"
            "sys.exit(main(sys.argv[1:]))\n")
        out = self.scratch("noopenpyxl-clean.csv")
        proc = subprocess.run(
            [sys.executable, driver,
             _sample("sample1-uk-debit-credit-columns.csv"), "--out", out],
            cwd=self.workdir, env=self._env(extra_path=stub_dir),
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            encoding="utf-8", errors="replace", timeout=180)
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 0)
        self.assertIn("RESULT: PASS", text)
        self.assertTrue(os.path.isfile(out))

    def test_xlsx_path_explains_the_missing_extra(self):
        """With openpyxl absent, .xlsx must fail politely, not with a traceback."""
        stub_dir = self.no_openpyxl_dir()
        out = self.scratch("needs-openpyxl.xlsx")
        proc = self.run_packaged(_sample("sample1-uk-debit-credit-columns.csv"),
                                 "--out", out, extra_path=stub_dir)
        text = self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("openpyxl", text)
        self.assertFalse(os.path.isfile(out))


class WrapperTests(CLITestCase):
    """`python3 normalize.py ...` from a clone must keep working unchanged."""

    def test_wrapper_runs_the_same_code(self):
        out_wrapper = self.scratch("wrapper-clean.csv")
        out_packaged = self.scratch("packaged-clean.csv")
        sample = _sample("sample1-uk-debit-credit-columns.csv")

        proc = self.run_wrapper(sample, "--out", out_wrapper)
        self.assert_clean(proc)
        self.assertEqual(proc.returncode, 0)

        proc2 = self.run_packaged(sample, "--out", out_packaged)
        self.assert_clean(proc2)
        self.assertEqual(proc2.returncode, 0)

        # Byte-for-byte the same table: one implementation, not two.
        with open(out_wrapper, "rb") as fh:
            first = fh.read()
        with open(out_packaged, "rb") as fh:
            second = fh.read()
        self.assertEqual(first, second)
        self.assertEqual(self.read_out(out_wrapper)[0], OUT_HEADER)

    def test_wrapper_keeps_exit_codes(self):
        proc = self.run_wrapper(_sample("sample6-corrupted-balance.csv"),
                                "--dry-run")
        self.assert_clean(proc)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("PROBLEM FOUND", proc.stdout)


if __name__ == "__main__":
    unittest.main()
