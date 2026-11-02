"""UAT-05: full source-to-target reconciliation (master command section 8;
evidence assessment section 3, Q27: "Count equality on INSERT rows only" was
the previous behaviour, transform.py reconcile before this change).

The reconciliation is keyed on the Assignment Identifier, reconciles INSERT
rows for mapped fields and UPDATE rows for address / contact / dependent
groups, and counts every non-blank source cell in exactly one category:
written, blank_token, provisional_conversion, gated_rule, held_row,
excluded_colour, pending_target, failed_conversion, replaced_by_rule,
unexplained. reconciliation_ok only when unexplained == 0 and counts balance.

Fixtures are synthetic: tests/fixtures/topia_sample.xlsx and
tests/fixtures/reconciliation_children.xlsx (built by
tests/reconciliation_fixtures.py). The real-extract test only asserts counts
and column letters and is skipped where the git-ignored extract is absent."""

import io
import json
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))
sys.path.insert(0, str(ROOT / "tests"))
sys.path.insert(0, str(ROOT / "local"))

import mapping_config as cfg  # noqa: E402
import reconcile as cli  # noqa: E402  (local/reconcile.py)
import reconciliation_fixtures as rf  # noqa: E402
import transform as tmod  # noqa: E402
from transform import RECON_CATEGORIES, transform  # noqa: E402

TEMPLATE = ROOT / "lambda" / "src" / "templates" / "equus_eafh_template.xlsx"
SAMPLE = ROOT / "tests" / "fixtures" / "topia_sample.xlsx"
SHIPPED_RULES = ROOT / "lambda" / "src" / "business_rules.json"
REAL = ROOT / "local_s3" / "source_s3" / "topia" / "incoming" / "FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx"
SHEET = "Full Historical Past Assignment"
CMD, IDENT = cfg.COL_ASSIGNMENT_COMMAND, cfg.COL_ASSIGNMENT_IDENTIFIER
DEP_FIRST = "ASSIGNMENT_DEPENDENT!FIRST_NAME"


# --------------------------------------------------------------------------- helpers

@pytest.fixture(scope="module")
def children():
    return transform(rf.CHILDREN.read_bytes(), TEMPLATE.read_bytes())


def _out_rows(xlsx: bytes) -> list[dict]:
    ws = openpyxl.load_workbook(io.BytesIO(xlsx))[cfg.TARGET_SHEET]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    return [{h: ws.cell(r, c).value for c, h in enumerate(headers, 1) if ws.cell(r, c).value is not None}
            for r in range(4, ws.max_row + 1)]


def _tamper(monkeypatch, edit):
    """Monkeypatch write_output so `edit(ws, headers)` changes the written workbook."""
    real = tmod.write_output

    def tampered(*args, **kwargs):
        wb = openpyxl.load_workbook(io.BytesIO(real(*args, **kwargs)))
        ws = wb[cfg.TARGET_SHEET]
        edit(ws, {ws.cell(3, c).value: c for c in range(1, ws.max_column + 1)})
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    monkeypatch.setattr(tmod, "write_output", tampered)


def _field(report, label: str) -> dict:
    col = next(c for c in report.columns if c.label == label)
    return next(r for r in report.reconciliation if r["col"] == col.letter)


def _record(report, source_row: int) -> dict:
    return next(r for r in report.reconciliation_records if r["source_row"] == source_row)


def _source_with(edit) -> bytes:
    wb = openpyxl.load_workbook(SAMPLE)
    edit(wb[SHEET])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _assert_balanced(report):
    for row in report.reconciliation:
        assert row["source"] == sum(row[k] for k in RECON_CATEGORIES), row["col"]


# --------------------------------------------------------------------------- AC-1

def test_children_fixture_reconciles_with_zero_unexplained(children):
    out, _, report = children
    s = report.reconciliation_summary
    assert report.reconciliation_ok and s["result"] == "OK" and s["counts_balance"]
    assert s["unexplained_total"] == 0 and s["cells_unexplained"] == 0 and s["record_unexplained"] == 0
    assert s["reasons"] == []
    rows = _out_rows(out)
    updates = [r for r in rows if r[CMD] == "UPDATE"]
    # Expected expansion rows equal the UPDATE rows actually written.
    assert s["expected_expansion_rows"] == s["update_rows"] == len(updates) == rf.EXPECTED["expansion_rows"]
    assert s["insert_rows"] == s["accepted_assignments"] == 2
    assert s["rows_written"] == len(rows) == s["accepted_assignments"] + s["expected_expansion_rows"]
    for kind in ("address", "contact", "dependent"):
        assert s["child_records_expected"][kind] == s["child_records_written"][kind] == rf.EXPECTED[kind]
    _assert_balanced(report)


def test_children_fixture_per_record_child_counts(children):
    _, _, report = children
    for source_row, want in rf.EXPECTED["rows"].items():
        rec = _record(report, source_row)
        assert rec["status"] == "accepted" and rec["unexplained"] == 0
        assert rec["output_rows"] == want["output_rows"] == rec["insert_rows"] + rec["update_rows"]
        assert rec["expected_expansion_rows"] == want["output_rows"] - 1
        for kind in ("address", "contact", "dependent"):
            assert rec[f"{kind}_expected"] == rec[f"{kind}_written"] == want[kind], (source_row, kind)


def test_children_fixture_child_cells_reconciled_field_by_field(children):
    _, _, report = children
    child_rows = [r for r in report.reconciliation if r["group"] in ("address", "contact", "dependent")]
    assert {r["group"] for r in child_rows} == {"address", "contact", "dependent"}
    for r in child_rows:
        assert r["unexplained"] == 0 and r["source"] == r["written"] + r["blank_token"], r["col"]
    # Placeholders in child columns are explained, not unexplained (FS 'N/A', Child 1 'Not Provided').
    assert sum(r["blank_token"] for r in child_rows) == 2
    # Every populated dependent cell of row 5 (3 dependents) and row 6 (1 dependent) was written.
    assert sum(r["written"] for r in child_rows if r["group"] == "dependent") == 10


def test_reconciliation_is_keyed_on_identifier_not_row_position(monkeypatch):
    def reverse_blocks(ws, headers):
        rows = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)] for r in range(4, ws.max_row + 1)]
        first = [i for i, row in enumerate(rows) if row[headers[CMD] - 1] == "INSERT"]
        blocks = [rows[a:b] for a, b in zip(first, first[1:] + [len(rows)])]
        for r, row in enumerate([row for block in reversed(blocks) for row in block], 4):
            for c, value in enumerate(row, 1):
                ws.cell(r, c).value = value
    _tamper(monkeypatch, reverse_blocks)
    out, _, report = transform(rf.CHILDREN.read_bytes(), TEMPLATE.read_bytes())
    assert [r[IDENT] for r in _out_rows(out) if r[CMD] == "INSERT"] == ["90002", "90001"]
    assert report.reconciliation_ok and report.reconciliation_summary["unexplained_total"] == 0


# --------------------------------------------------------------------------- AC-2

def test_dropped_dependent_first_name_on_update_row_is_unexplained(monkeypatch):
    dropped = []

    def drop(ws, headers):
        r = next(r for r in range(4, ws.max_row + 1)
                 if ws.cell(r, headers[CMD]).value == "UPDATE" and ws.cell(r, headers[DEP_FIRST]).value is not None)
        ws.cell(r, headers[DEP_FIRST]).value = None
        dropped.append(r)
    _tamper(monkeypatch, drop)
    _, _, report = transform(rf.CHILDREN.read_bytes(), TEMPLATE.read_bytes())
    assert dropped == [5]  # second output row: UPDATE row of the second dependent
    assert not report.reconciliation_ok
    bad = [r for r in report.reconciliation if r["unexplained"]]
    assert [(r["group"], r["header"], r["unexplained"]) for r in bad] == [("dependent", DEP_FIRST, 1)]
    assert any(DEP_FIRST in reason for reason in report.reconciliation_summary["reasons"])
    assert any(r["category"] == "reconciliation" for r in report.readiness.reasons)


def test_dropped_insert_first_name_still_caught_on_fixture(monkeypatch):
    _tamper(monkeypatch, lambda ws, h: setattr(ws.cell(4, h["EMPLOYEE!FIRST_NAME"]), "value", None))
    _, _, report = transform(rf.CHILDREN.read_bytes(), TEMPLATE.read_bytes())
    assert not report.reconciliation_ok
    assert next(r for r in report.reconciliation if r["header"] == "EMPLOYEE!FIRST_NAME")["unexplained"] == 1


def test_dropped_address_line_and_changed_contact_are_unexplained(monkeypatch):
    def edit(ws, h):
        ws.cell(4, h["ASSIGNMENT_MAILING_ADDRESS!LINE_2"]).value = None
        ws.cell(5, h["ASSIGNMENT_EMPLOYEE_CONTACT!NUMBER_ADDRESS"]).value = "0000000"
    _tamper(monkeypatch, edit)
    _, _, report = transform(rf.CHILDREN.read_bytes(), TEMPLATE.read_bytes())
    assert not report.reconciliation_ok
    bad = {r["group"]: r["unexplained"] for r in report.reconciliation if r["unexplained"]}
    assert bad == {"address": 1, "contact": 1}


def test_lost_update_row_is_caught_at_record_level(monkeypatch):
    def clear_last(ws, headers):
        for c in range(1, ws.max_column + 1):
            ws.cell(6, c).value = None  # third row of the first assignment (UPDATE, third dependent)
    _tamper(monkeypatch, clear_last)
    _, _, report = transform(rf.CHILDREN.read_bytes(), TEMPLATE.read_bytes())
    s = report.reconciliation_summary
    assert not report.reconciliation_ok and not s["counts_balance"]
    assert s["update_rows"] == 1 and s["expected_expansion_rows"] == 2
    rec = _record(report, 5)
    assert rec["dependent_expected"] == 3 and rec["dependent_written"] == 2 and rec["unexplained"] >= 1


def test_duplicated_and_unknown_output_rows_are_caught(monkeypatch):
    def add_rows(ws, headers):
        last = ws.max_row
        for c in range(1, ws.max_column + 1):
            ws.cell(last + 1, c).value = ws.cell(last, c).value  # duplicate of the last row
        for c in range(1, ws.max_column + 1):
            ws.cell(last + 2, c).value = ws.cell(4, c).value  # second INSERT of the first identifier
        ws.cell(last + 3, headers[CMD]).value = "INSERT"
        ws.cell(last + 3, headers[IDENT]).value = "SYN-UNKNOWN"  # identifier of no source row
    _tamper(monkeypatch, add_rows)
    _, _, report = transform(rf.CHILDREN.read_bytes(), TEMPLATE.read_bytes())
    s = report.reconciliation_summary
    assert not report.reconciliation_ok
    assert s["duplicate_insert_blocks"] == 2 and s["unexplained_total"] > 0  # 90002 and 90001 INSERTed twice
    assert any("INSERTed more than once" in r for r in s["reasons"])
    assert any("identifier of no accepted source row" in r for r in s["reasons"])


# --------------------------------------------------------------------------- AC-3

@pytest.mark.parametrize("case", ["duplicate_row", "missing_identifier"])
def test_held_rows_are_an_explained_category(case):
    def edit(ws):
        labels = {" ".join(str(ws.cell(4, c).value).split()).lower(): c
                  for c in range(1, ws.max_column + 1) if ws.cell(4, c).value is not None}
        if case == "duplicate_row":
            for c in range(1, ws.max_column + 1):
                ws.cell(7, c).value = ws.cell(5, c).value
        else:
            ws.cell(6, labels["assignment id"]).value = None
    src = _source_with(edit)
    _, _, report = transform(src, TEMPLATE.read_bytes())
    held_row = 7 if case == "duplicate_row" else 6
    assert list(report.held_rows) == [held_row]
    s = report.reconciliation_summary
    ws = openpyxl.load_workbook(io.BytesIO(src))[SHEET]
    present = sum(1 for c in report.columns if tmod._present(ws.cell(held_row, c.idx).value))
    assert s["cells_held_row"] == present > 0
    assert s["held_rows"] == 1 and s["accepted_assignments"] == s["source_rows"] - 1
    assert report.reconciliation_ok and s["unexplained_total"] == 0
    rec = _record(report, held_row)
    assert rec["status"] == "held" and rec["output_rows"] == 0 and rec["codes"]
    _assert_balanced(report)


@pytest.mark.parametrize("label, letter, q", [
    ("confidentiality type", "AG", "Q6/Q7/Q29"),
    ("expected assignment outcome", "DG", "Q6/Q7/Q29"),
    ("pii nationality 1", "ED", "Q9"),
    ("assignment status", "AB", "Q34"),
    ("primary mobility email", "EJ", "Q41"),
])
def test_pending_target_columns_are_explained(children, label, letter, q):
    _, _, report = children
    row = _field(report, label)
    assert row["col"] == letter
    assert row["source"] == row["pending_target"] == 2 and row["unexplained"] == 0
    assert row["q"] == q and row["owner"]


def test_gated_excluded_blank_and_provisional_categories(children):
    _, _, report = children
    host_job_level = _field(report, "host job level")  # Q32 rule not approved -> column held
    assert host_job_level["gated_rule"] == host_job_level["source"] == 2 and host_job_level["q"] == "Q32"
    cancellation = _field(report, "cancellation reason")  # Q34 rule not approved
    assert cancellation["gated_rule"] == 2 and cancellation["q"] == "Q34"
    red = [r for r in report.reconciliation if r["group"] == "excluded" and r["source"]]
    assert red and all(r["excluded_colour"] == r["source"] and r["q"] == "Q30" for r in red)
    nationality3 = _field(report, "nationality 3")  # N: 'Not Provided...' placeholders
    assert nationality3["col"] == "N" and nationality3["blank_token"] == 2 and nationality3["q"] == "Q43"
    countries = [r for r in report.reconciliation if r["provisional_conversion"]]
    assert {r["header"] for r in countries} >= {"ASSIGNMENT!FROM_COUNTRY", "ASSIGNMENT!TO_COUNTRY"}
    assert all(r["unexplained"] == 0 for r in report.reconciliation)


def test_approved_rule_replacing_a_source_value_is_explained(tmp_path):
    data = json.loads(SHIPPED_RULES.read_text(encoding="utf-8"))
    rule = next(r for r in data["rules"] if r["id"] == cfg.RULE_CANCELLATION_PHASE)
    rule["approved_for_import"] = True
    rule["approval_reference"] = "test copy only - not an approval"
    path = tmp_path / "business_rules.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    _, _, report = transform(rf.CHILDREN.read_bytes(), TEMPLATE.read_bytes(), business_rules=path)
    assert report.reconciliation_ok and report.reconciliation_summary["unexplained_total"] == 0
    assert _field(report, "cancellation reason")["written"] == 2  # consumed by the approved rule
    status_category = _field(report, "status category")  # mapped to PHASE, overwritten by the rule
    assert status_category["unexplained"] == 0
    assert status_category["written"] + status_category["replaced_by_rule"] == status_category["source"]


# --------------------------------------------------------------------------- report and CLI

def test_report_workbook_has_field_and_record_sheets(children):
    _, report_bytes, report = children
    wb = openpyxl.load_workbook(io.BytesIO(report_bytes))
    header = [c.value for c in wb["Reconciliation"][1]]
    # The pre-UAT-05 columns keep their position.
    assert header[:11] == ["Source Col", "Row 4 (Topia)", "Equus Header", "Source Non-blank", "Written",
                           "Blank Token", "Failed Conversion", "Set By Rule", "Unexplained", "Source Date Total",
                           "Written Date Total"]
    assert {"Pending Target", "Held Row", "Excluded Colour", "Gated Rule", "Provisional Conversion"} <= set(header)
    records = wb["Reconciliation Records"]
    assert [c.value for c in records[1]] == tmod.RECORDS_HEADER
    assert records.max_row == 1 + len(report.reconciliation_records) == 3


def test_cli_evidence_has_no_source_values():
    result = cli.run(rf.CHILDREN)
    ev = dict(result["evidence"])
    ev.pop("source_sha256")
    text = json.dumps(ev, default=str)
    for value in ("SynthA", "SynthB", "SynthD", "Testcase", "Synthetic Lane", "Synthville", "Example Park",
                  "Sample Road", "5550001", "5550002"):
        assert value not in text
    assert all("identifier" not in r and "employee" not in r for r in ev["records"])
    assert all("source_date_total" not in f for f in ev["fields"])


def test_cli_main_writes_report_and_returns_status(tmp_path, capsys):
    assert cli.main([str(rf.CHILDREN), "--out-dir", str(tmp_path)]) == 0
    printed = capsys.readouterr().out
    assert "Reconciliation: OK" in printed and "expected_expansion_rows: 2" in printed
    written = sorted(p.suffix for p in tmp_path.iterdir())
    assert written == [".json", ".xlsx"]
    ev = json.loads(next(tmp_path.glob("*.json")).read_text(encoding="utf-8"))
    assert ev["summary"]["result"] == "OK" and ev["summary"]["unexplained_total"] == 0
    assert cli.main([str(tmp_path / "missing.xlsx"), "--out-dir", str(tmp_path)]) == 2


# --------------------------------------------------------------------------- real extract (counts only)

@pytest.mark.skipif(not REAL.exists(), reason="real Topia extract not present (git-ignored, personal data)")
def test_real_extract_known_facts():
    """Facts recorded in docs/uat/RECONCILIATION_SUMMARY.md; counts and
    column letters only, no value is read into the assertion."""
    _, _, report = transform(REAL.read_bytes(), TEMPLATE.read_bytes())
    s = report.reconciliation_summary
    assert (s["source_rows"], s["source_employees"], s["accepted_assignments"], s["held_rows"]) == (2, 2, 2, 0)
    assert s["expected_expansion_rows"] == s["update_rows"] == 0
    assert {k: s["child_records_expected"][k] for k in ("address", "contact", "dependent")} == \
        {"address": 0, "contact": 0, "dependent": 0}
    by_col = {r["col"]: r for r in report.reconciliation}
    for letter in ("AG", "DG", "ED"):
        assert by_col[letter]["pending_target"] == by_col[letter]["source"] > 0
    assert by_col["N"]["blank_token"] >= 1 and by_col["N"]["unexplained"] == 0
    assert report.reconciliation_ok and s["unexplained_total"] == 0
