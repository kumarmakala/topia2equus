"""UAT-03: identity, duplicate, header and row-level error controls
(BR-01, BR-05, BR-17, BR-18; Q1, Q5, Q17, Q18, Q19).

Fixtures: tests/fixtures/topia_sample.xlsx (synthetic) modified by the
scenarios in tests/fixtures/identity_cases.json (synthetic values only).
Jake's example sheet is read only for its command / blank-cell pattern
(columns A, B, C, D, E), never for its values."""

import functools
import hashlib
import io
import json
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))

import handler  # noqa: E402
import mapping_config as cfg  # noqa: E402
import transform as tmod  # noqa: E402
from mapping_config import ConfigurationError  # noqa: E402
from transform import NOT_READY, READY, transform  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"
SOURCE = FIXTURES / "topia_sample.xlsx"
CASES = json.loads((FIXTURES / "identity_cases.json").read_text(encoding="utf-8"))
TEMPLATE = ROOT / "lambda" / "src" / "templates" / "equus_eafh_template.xlsx"
SHIPPED_RULES = ROOT / "lambda" / "src" / "business_rules.json"
JAKE = ROOT / "clarifications" / "jake-1.xlsx"
SHEET = "Full Historical Past Assignment"
IDENT = "Assignment Identifier"


# --------------------------------------------------------------------------- helpers

def _norm(value) -> str:
    return " ".join(str(value).split()).lower() if value is not None else ""


def _col_of(ws, label: str) -> int:
    for c in range(1, ws.max_column + 1):
        if (_norm(ws.cell(4, c).value) or _norm(ws.cell(3, c).value)) == label:
            return c
    raise AssertionError(f"label {label!r} not in fixture")


def _save(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _apply(ws, rows) -> None:
    for spec in rows:
        for c in range(1, ws.max_column + 1):
            ws.cell(spec["to"], c).value = ws.cell(spec["copy_from"], c).value
        for label, value in spec["set"].items():
            ws.cell(spec["to"], _col_of(ws, label)).value = value


@functools.lru_cache(maxsize=None)
def _case_source(name: str) -> bytes:
    wb = openpyxl.load_workbook(SOURCE)
    _apply(wb[SHEET], CASES["cases"][name]["rows"])
    return _save(wb)


@functools.lru_cache(maxsize=None)
def _run(name: str):
    return transform(_case_source(name), TEMPLATE.read_bytes())


def _rows(xlsx: bytes):
    ws = openpyxl.load_workbook(io.BytesIO(xlsx))["Sheet1"]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    return [{h: ws.cell(r, c).value for c, h in enumerate(headers, 1) if ws.cell(r, c).value is not None}
            for r in range(4, ws.max_row + 1)]


def _inserts(rows):
    return [r for r in rows if r.get("Assignment Command") == "INSERT"]


def _sheet_table(report_bytes: bytes, sheet: str, header_row: int = 1) -> list[dict]:
    ws = openpyxl.load_workbook(io.BytesIO(report_bytes))[sheet]
    header = [c.value for c in ws[header_row]]
    return [dict(zip(header, r)) for r in ws.iter_rows(min_row=header_row + 1, values_only=True)
            if any(v is not None for v in r)]


def _exceptions(report, code):
    return [e for e in report.exceptions if e["code"] == code]


def _shipped() -> dict:
    return json.loads(SHIPPED_RULES.read_text(encoding="utf-8"))


def _write_rules(tmp_path, data) -> Path:
    path = tmp_path / "business_rules.json"
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return path


def _unit_report() -> tmod.Report:
    return tmod.Report(rules=cfg.load_business_rules(), reference_data=cfg.load_reference_data())


# --------------------------------------------------------------------------- baseline

def test_clean_fixture_holds_nothing_and_reconciles():
    out, report_bytes, report = transform(SOURCE.read_bytes(), TEMPLATE.read_bytes())
    assert report.held_rows == {}
    assert not any(e["level"] in (cfg.LEVEL_ROW_HELD, cfg.LEVEL_FILE) for e in report.exceptions)
    assert [a["identifier"] for a in report.accepted] == ["90001", "90002"]
    assert report.reconciliation_ok and report.data_rows == report.assignments == 2
    assert {"Exceptions", "Held Rows", "Accepted Identifiers", "Template Check"} <= \
        set(openpyxl.load_workbook(io.BytesIO(report_bytes)).sheetnames)
    assert _sheet_table(report_bytes, "Exceptions")[0].keys() == set(tmod.EXCEPTIONS_HEADER)


# --------------------------------------------------------------------------- AC-1 missing identifier

@pytest.mark.parametrize("case", ["missing_assignment_id", "placeholder_assignment_id"])
def test_missing_assignment_id_is_held_never_falls_back(case):
    out, report_bytes, report = _run(case)
    employee = CASES["cases"][case]["rows"][0]["set"]["employee id"]
    rows = _rows(out)
    # Absent from the import workbook; only the two valid assignments remain.
    assert [r[IDENT] for r in _inserts(rows)] == ["90001", "90002"]
    assert employee not in {r.get("Employee Number") for r in rows}
    # The Employee Number is never used as identifier (BR-01).
    assert str(employee) not in {str(r.get(IDENT)) for r in rows}
    assert str(employee) not in {str(r.get("ASSIGNMENT_CF!CF_TEXT029")) for r in rows}
    assert report.held_rows == {7: [cfg.EXC_MISSING_ASSIGNMENT_ID]}
    # Exceptions sheet: source row, identifier, employee number, code, level, message, owner, Q.
    exc = [e for e in _sheet_table(report_bytes, "Exceptions") if e["Code"] == cfg.EXC_MISSING_ASSIGNMENT_ID]
    assert len(exc) == 1
    e = exc[0]
    assert e["Source Row"] == 7 and e["Assignment Identifier"] is None and e["Employee Number"] == str(employee)
    assert e["Level"] == "ROW_HELD" and e["Q"] == "Q1"
    assert "Alka" in e["Owner"] and "Jake" in e["Owner"]  # D1 owners
    assert "not used as identifier" in e["Message"]
    # Held Rows sheet carries the source data of the held row.
    held = openpyxl.load_workbook(io.BytesIO(report_bytes))["Held Rows"]
    assert held.cell(3, 1).value == 7 and held.cell(3, 2).value == cfg.EXC_MISSING_ASSIGNMENT_ID
    src_ws = openpyxl.load_workbook(io.BytesIO(_case_source(case)))[SHEET]
    emp_col = _col_of(src_ws, "employee id")
    assert held.cell(2, 2 + emp_col).value == "Employee ID"
    assert held.cell(3, 2 + emp_col).value == employee
    # Fully accounted: accepted + held = source rows; reconciliation balances.
    assert report.data_rows == 3 and report.assignments == 2 and report.reconciliation_ok
    assert report.readiness.status == NOT_READY
    assert any(r["category"] == "row_held" and r["item"] == cfg.EXC_MISSING_ASSIGNMENT_ID and r["q"] == "Q1"
               for r in report.readiness.reasons)


def test_no_identifier_fallback_left_in_engine():
    src = (ROOT / "lambda" / "src" / "transform.py").read_text(encoding="utf-8")
    assert "identifier = employee" not in src
    assert not hasattr(tmod.Report(), "identifier_fallbacks")


def test_identifier_text_preserves_leading_zeros_and_drops_dot_zero():
    assert tmod.identifier_text("00123") == "00123"
    assert tmod.identifier_text(" 00123 ") == "00123"
    assert tmod.identifier_text(90011.0) == "90011"
    assert tmod.identifier_text(90011) == "90011"
    assert tmod.identifier_text(None) is None
    assert tmod.identifier_text("  ") is None
    assert tmod.identifier_text("Not Provided") is None


def test_missing_employee_number_is_held():
    out, report_bytes, report = _run("missing_employee_number")
    assert report.held_rows == {7: [cfg.EXC_MISSING_EMPLOYEE_NUMBER]}
    assert "90021" not in {r.get(IDENT) for r in _rows(out)}
    [e] = _exceptions(report, cfg.EXC_MISSING_EMPLOYEE_NUMBER)
    assert e["identifier"] == "90021" and e["employee"] is None and e["level"] == cfg.LEVEL_ROW_HELD
    assert report.reconciliation_ok


# --------------------------------------------------------------------------- AC-2 duplicates

def test_exact_duplicate_row_one_accepted_one_held():
    out, report_bytes, report = _run("exact_duplicate")
    inserts = _inserts(_rows(out))
    assert [r[IDENT] for r in inserts] == ["90001", "90002"]  # one accepted assignment for 90001
    assert report.held_rows == {7: [cfg.EXC_DUPLICATE_ROW]}
    [e] = _exceptions(report, cfg.EXC_DUPLICATE_ROW)
    assert e["source_row"] == 7 and e["identifier"] == "90001" and "source row 5" in e["message"]
    assert not _exceptions(report, cfg.EXC_CONFLICTING_IDENTIFIER)
    assert [r["Source Row"] for r in _sheet_table(report_bytes, "Held Rows")[1:]] == [7]
    assert report.reconciliation_ok and report.readiness.status == NOT_READY


def test_same_identifier_different_data_holds_both_rows():
    out, report_bytes, report = _run("conflicting_identifier")
    assert [r[IDENT] for r in _inserts(_rows(out))] == ["90001"]
    assert report.held_rows == {6: [cfg.EXC_CONFLICTING_IDENTIFIER], 7: [cfg.EXC_CONFLICTING_IDENTIFIER]}
    found = _exceptions(report, cfg.EXC_CONFLICTING_IDENTIFIER)
    assert sorted(e["source_row"] for e in found) == [6, 7]
    assert all(e["identifier"] == "90002" and e["level"] == cfg.LEVEL_ROW_HELD for e in found)
    # Differing columns are named by letter (no source values in the message).
    src_ws = openpyxl.load_workbook(io.BytesIO(_case_source("conflicting_identifier")))[SHEET]
    letter = openpyxl.utils.get_column_letter(_col_of(src_ws, "first name"))
    assert all(letter in e["message"] and "Synthother" not in e["message"] for e in found)
    assert sorted(r["Source Row"] for r in _sheet_table(report_bytes, "Held Rows")[1:]) == [6, 7]
    assert report.data_rows == 3 and report.assignments == 1 and report.reconciliation_ok


def test_identifier_reused_by_other_employee_is_conflict():
    out, _, report = _run("conflicting_identifier_other_employee")
    assert sorted(report.held_rows) == [5, 7]
    assert [r[IDENT] for r in _inserts(_rows(out))] == ["90002"]


# --------------------------------------------------------------------------- AC-3 multiple assignments

def test_one_employee_three_assignments_three_inserts_text_identifiers():
    out, _, report = _run("multiple_assignments")
    assert report.held_rows == {}  # same Employee Number, different identifiers: valid (Q17)
    rows = _rows(out)
    emp_rows = [r for r in _inserts(rows) if r.get("Employee Number") == 10001]
    assert len(emp_rows) == 3
    assert [r["A"] for r in emp_rows] == ["INSERT", "IGNORE", "IGNORE"]
    assert all(r["Assignment Command"] == "INSERT" and r["Assignment Number"] == "NEW" for r in emp_rows)
    idents = [r[IDENT] for r in emp_rows]
    assert idents == ["90001", "00123", "90011"] and len(set(idents)) == 3
    ws = openpyxl.load_workbook(io.BytesIO(out))["Sheet1"]
    ident_col = next(c for c in range(1, ws.max_column + 1) if ws.cell(3, c).value == IDENT)
    cf29_col = next(c for c in range(1, ws.max_column + 1) if ws.cell(3, c).value == "ASSIGNMENT_CF!CF_TEXT029")
    zero_row = next(r for r in range(4, ws.max_row + 1) if ws.cell(r, ident_col).value == "00123")
    for c in (ident_col, cf29_col):
        cell = ws.cell(zero_row, c)
        assert cell.value == "00123" and cell.data_type == "s" and cell.number_format == "@"
    assert not any(str(r.get(IDENT)).endswith(".0") for r in rows)
    assert not _exceptions(report, cfg.EXC_ROW_STRUCTURE_INVALID)


def test_employee_command_insert_goes_to_first_accepted_assignment():
    # Row 5 (first assignment of 10001) is held as a conflict; the employee's
    # next accepted assignment must then carry Employee command INSERT.
    wb = openpyxl.load_workbook(SOURCE)
    ws = wb[SHEET]
    _apply(ws, CASES["cases"]["multiple_assignments"]["rows"])
    _apply(ws, [{"copy_from": 6, "to": 9, "set": {"assignment id": 90001}}])
    out, _, report = transform(_save(wb), TEMPLATE.read_bytes())
    assert 5 in report.held_rows and 9 in report.held_rows
    emp_rows = [r for r in _inserts(_rows(out)) if r.get("Employee Number") == 10001]
    assert [r["A"] for r in emp_rows] == ["INSERT", "IGNORE"]
    assert not _exceptions(report, cfg.EXC_ROW_STRUCTURE_INVALID)


# --------------------------------------------------------------------------- row structure vs Jake's example

def _jake_pattern():
    ws = openpyxl.load_workbook(JAKE, read_only=True)[" Mny dependents Ex"]
    rows = [r[:5] for r in ws.iter_rows(min_row=4, values_only=True) if r and r[1] is not None]
    first_ident = rows[0][4]
    return [(a, b, c is None, d, e == first_ident) for a, b, c, d, e in rows]


def _source_with_dependents() -> bytes:
    spec = CASES["dependents"]
    wb = openpyxl.load_workbook(SOURCE)
    ws = wb[SHEET]
    start = ws.max_column + 1
    for i, name in enumerate(spec["first_names"]):
        ws.cell(2, start + i, "ASSIGNMENT_DEPENDENT!FIRST_NAME")
        ws.cell(4, start + i, f"Child {i + 1} First Name")
        ws.cell(spec["row"], start + i, name)
    return _save(wb)


def test_generated_rows_follow_jakes_example_structure():
    jake = _jake_pattern()
    assert jake[0][:2] == ("INSERT", "INSERT") and jake[0][3] == "NEW"
    assert len(jake) == len(CASES["dependents"]["first_names"])
    out, _, report = transform(_source_with_dependents(), TEMPLATE.read_bytes())
    rows = [r for r in _rows(out) if r.get(IDENT) == "90001"]
    ours = [(r.get("A"), r.get("Assignment Command"), r.get("Employee Number") is None,
             r.get("Assignment Number"), r.get(IDENT) == "90001") for r in rows]
    assert ours == jake  # INSERT/INSERT/NEW, then IGNORE/UPDATE, blank Employee Number, same identifier
    assert [r.get("ASSIGNMENT_DEPENDENT!FIRST_NAME") for r in rows] == CASES["dependents"]["first_names"]
    assert not _exceptions(report, cfg.EXC_ROW_STRUCTURE_INVALID) and not report.file_failed


def test_row_structure_violations_are_file_level():
    report = _unit_report()
    good_insert = {"A": "INSERT", "Assignment Command": "INSERT", "Employee Number": 1,
                   "Assignment Number": "NEW", IDENT: "X1"}
    tmod.validate_row_structure([good_insert, {"A": "IGNORE", "Assignment Command": "UPDATE", IDENT: "X1"}], report)
    assert report.exceptions == []
    bad = [
        {"A": "IGNORE", "Assignment Command": "UPDATE", IDENT: "X1"},                       # UPDATE before INSERT
        dict(good_insert),
        {"A": "IGNORE", "Assignment Command": "UPDATE", IDENT: "X2"},                       # cross-assignment link
        {"A": "IGNORE", "Assignment Command": "UPDATE", "Employee Number": 1, IDENT: "X1"},  # Employee Number set
        dict(good_insert, A="INSERT", **{IDENT: "X3"}),                                     # 2nd assignment INSERT
        dict(good_insert, **{IDENT: None, "A": "IGNORE"}),                                  # no identifier
    ]
    tmod.validate_row_structure(bad, report)
    found = _exceptions(report, cfg.EXC_ROW_STRUCTURE_INVALID)
    assert len(found) >= 5 and all(e["level"] == cfg.LEVEL_FILE for e in found)
    assert report.file_failed


# --------------------------------------------------------------------------- AC-4 header / template

def _template_with(edits: dict) -> bytes:
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Sheet1"]
    for cell, value in edits.items():
        ws[cell] = value
    return _save(wb)


def test_header_mismatch_fails_file_and_lists_columns():
    tpl = openpyxl.load_workbook(TEMPLATE)["Sheet1"]
    f3 = tpl["F3"].value
    out, report_bytes, report = transform(SOURCE.read_bytes(), _template_with({"F3": f3 + "_X", "H1": "changed"}))
    assert out == b""  # whole file withheld
    assert report.file_failed
    [e] = _exceptions(report, cfg.EXC_HEADER_MISMATCH)
    assert e["level"] == cfg.LEVEL_FILE and e["q"] == "Q5" and "Jake" in e["owner"]
    assert "column(s) F, H" in e["message"] and "F3" in e["message"] and "H1" in e["message"]
    assert [d["cell"] for d in report.template_check["diffs"]] == ["H1", "F3"]
    assert report.readiness.status == NOT_READY
    assert any(r["category"] == "file_rejected" and r["item"] == cfg.EXC_HEADER_MISMATCH
               for r in report.readiness.reasons)
    wb = openpyxl.load_workbook(io.BytesIO(report_bytes))
    tc = wb["Template Check"]
    values = {tc.cell(r, 1).value: tc.cell(r, 2).value for r in range(1, tc.max_row + 1)}
    assert values["Result"] == "FAILED" and values["File status"].startswith("FAILED")
    assert values["Template version"] == cfg.TEMPLATE_VERSION
    assert values["Approved template SHA-256"] == cfg.APPROVED_TEMPLATE_SHA256
    assert values["F3"] == f3  # differing cell row: approved value next to it


def test_employee_command_header_in_template_is_mismatch_not_silently_accepted():
    # D2: Jake's example uses 'Employee Command' in A3; until Jake decides the
    # approved template ('A') stands, so a template with the alternative fails.
    out, _, report = transform(SOURCE.read_bytes(), _template_with({"A3": "Employee Command"}))
    assert out == b"" and report.file_failed
    assert any("A3" in e["message"] for e in _exceptions(report, cfg.EXC_HEADER_MISMATCH))


def test_pending_header_decision_reported_header_unchanged():
    out, report_bytes, report = transform(SOURCE.read_bytes(), TEMPLATE.read_bytes())
    assert not report.file_failed and not _exceptions(report, cfg.EXC_HEADER_MISMATCH)
    assert openpyxl.load_workbook(io.BytesIO(out))["Sheet1"]["A3"].value == "A"  # not changed
    [e] = _exceptions(report, cfg.EXC_HEADER_DECISION_PENDING)
    assert e["level"] == cfg.LEVEL_WARNING and e["q"] == "Q5" and "Jake" in e["owner"]
    assert "'A'" in e["message"] and "'Employee Command'" in e["message"] and "D2" in e["message"]
    assert any(r["category"] == "pending_header_decision" for r in report.readiness.reasons)
    rows = list(openpyxl.load_workbook(io.BytesIO(report_bytes))["Template Check"].iter_rows(values_only=True))
    start = next(i for i, r in enumerate(rows) if r[0] == "Pending Header Decision")
    decision = dict(zip(rows[start], rows[start + 1]))
    assert decision["Pending Header Decision"] == "D2" and decision["Cell"] == "A3"
    assert decision["Approved Template"] == "A" and decision["Alternative"] == "Employee Command"
    assert decision["Status"].startswith("PENDING")
    readiness = {r[0]: r[1] for r in openpyxl.load_workbook(io.BytesIO(report_bytes))["Readiness"].iter_rows(
        values_only=True) if r and r[0]}
    assert readiness["TEMPLATE_VERSION"] == cfg.TEMPLATE_VERSION
    assert readiness["Template SHA-256 (supplied)"] == hashlib.sha256(TEMPLATE.read_bytes()).hexdigest()


def test_same_headers_different_template_bytes_is_warning():
    out, _, report = transform(SOURCE.read_bytes(), _template_with({}))  # re-saved: other bytes, same headers
    assert out and not report.file_failed
    [e] = _exceptions(report, cfg.EXC_TEMPLATE_HASH_DIFFERS)
    assert e["level"] == cfg.LEVEL_WARNING


def test_handler_withholds_import_on_file_level_failure(monkeypatch):
    class FakeS3:
        def __init__(self):
            self.puts = {}

        def get_object(self, Bucket, Key):
            body = SOURCE.read_bytes()
            return {"Body": io.BytesIO(body), "ETag": '"%s"' % hashlib.md5(body).hexdigest()}

        def head_object(self, Bucket, Key):
            if (Bucket, Key) not in self.puts:
                err = Exception("404")
                err.response = {"Error": {"Code": "404"}}
                raise err
            return {}

        def put_object(self, Bucket, Key, Body, IfNoneMatch=None, Metadata=None, **kwargs):
            self.puts[(Bucket, Key)] = Metadata

    s3 = FakeS3()
    monkeypatch.setattr(handler, "_s3", s3)
    monkeypatch.setattr(handler, "_load_template", lambda: _template_with({"F3": "BROKEN"}))
    monkeypatch.setenv("DEST_BUCKET", "dst")
    monkeypatch.setenv("DEST_PREFIX", "equus/")
    monkeypatch.setenv("SOURCE_PREFIX", "topia/incoming/")
    event = {"Records": [{"eventName": "ObjectCreated:Put",
                          "s3": {"bucket": {"name": "src"}, "object": {"key": "topia/incoming/SYN.xlsx", "size": 1}}}]}
    with pytest.raises(RuntimeError, match="File-level validation failed"):
        handler.lambda_handler(event, None)
    [(_, key)] = list(s3.puts)
    assert "FILE_REJECTED" in key and s3.puts[("dst", key)]["file-status"] == "FAILED"


# --------------------------------------------------------------------------- AC-5 readiness

def test_readiness_not_ready_whenever_a_row_is_held():
    report = _unit_report()
    report.reconciliation_ok = True
    assert tmod.compute_readiness(report).status == READY
    report.add_exception(cfg.EXC_DUPLICATE_ROW, "synthetic", 7, "X1", "1")
    readiness = tmod.compute_readiness(report)
    assert readiness.status == NOT_READY
    assert [r["category"] for r in readiness.reasons] == ["row_held"]
    assert readiness.reasons[0]["q"] == "Q5" and readiness.reasons[0]["owner"]


@pytest.mark.parametrize("case", sorted(CASES["cases"]))
def test_every_case_accounts_for_every_source_row(case):
    out, report_bytes, report = _run(case)
    assert report.data_rows == report.assignments + len(report.held_rows)
    assert report.reconciliation_ok
    if report.held_rows:
        assert report.readiness.status == NOT_READY
        held_sheet = {r["Source Row"] for r in _sheet_table(report_bytes, "Held Rows")[1:]}
        exc_rows = {e["Source Row"] for e in _sheet_table(report_bytes, "Exceptions")
                    if e["Level"] in ("ROW_HELD", "FILE") and isinstance(e["Source Row"], int)}
        assert held_sheet == exc_rows == set(report.held_rows)
    for e in _sheet_table(report_bytes, "Exceptions"):
        assert e["Code"] in cfg.EXCEPTION_CODES and e["Level"] in cfg.EXCEPTION_LEVELS
        assert e["Owner"] and e["Q"] and e["Message"]
    accepted = {r["Assignment Identifier"] for r in _sheet_table(report_bytes, "Accepted Identifiers")}
    assert accepted == {r[IDENT] for r in _inserts(_rows(out))}


# --------------------------------------------------------------------------- correction batch / reprocessing

def test_corrected_rerun_imports_only_previously_held_identifiers():
    # Run 1: row 7 held for a missing identifier.
    _, report1, r1 = _run("missing_assignment_id")
    accepted = tmod.accepted_identifiers(report1)
    assert accepted == {"90001", "90002"}
    # Correction at source: the held row gets its identifier; full re-run.
    wb = openpyxl.load_workbook(io.BytesIO(_case_source("missing_assignment_id")))
    ws = wb[SHEET]
    ws.cell(7, _col_of(ws, "assignment id")).value = 90003
    out2, report2, r2 = transform(_save(wb), TEMPLATE.read_bytes(), previously_accepted=accepted)
    rows = _rows(out2)
    assert [r[IDENT] for r in _inserts(rows)] == ["90003"]
    assert sorted(r2.held_rows) == [5, 6]
    assert all(codes == [cfg.EXC_ALREADY_ACCEPTED] for codes in r2.held_rows.values())
    assert _inserts(rows)[0]["A"] == "INSERT"
    assert r2.reconciliation_ok


def test_correction_batch_built_from_held_rows_sheet():
    _, report1, _ = _run("missing_assignment_id")
    held = openpyxl.load_workbook(io.BytesIO(report1))["Held Rows"]
    # Correction batch: source header rows 1-4 + the held rows' data (column C.. of Held Rows = source A..).
    wb = openpyxl.load_workbook(SOURCE)
    ws = wb[SHEET]
    for r in range(5, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None
    for i, r in enumerate(range(3, held.max_row + 1)):
        for c in range(3, held.max_column + 1):
            ws.cell(5 + i, c - 2).value = held.cell(r, c).value
    ws.cell(5, _col_of(ws, "assignment id")).value = 90003  # the correction
    out, _, report = transform(_save(wb), TEMPLATE.read_bytes())
    assert report.held_rows == {} and report.data_rows == 1
    [row] = _inserts(_rows(out))
    assert row[IDENT] == "90003" and row["Employee Number"] == 10003 and row["A"] == "INSERT"


# --------------------------------------------------------------------------- configuration

def test_shipped_identity_controls():
    ic = cfg.load_business_rules().identity
    assert ic.assignment_identifier_label == "assignment id" == cfg.ASSIGNMENT_IDENTIFIER_LABEL
    assert ic.template_sha256 == hashlib.sha256(TEMPLATE.read_bytes()).hexdigest()
    assert set(ic.codes) == set(cfg.EXCEPTION_CODES)
    for code in cfg.ROW_CODES:
        assert ic.code(code).level in (cfg.LEVEL_ROW_HELD, cfg.LEVEL_FILE)
    assert [d["cell"] for d in ic.pending_header_decisions] == ["A3"]


def _m_warn_missing_id(d):
    d["identity_controls"]["exception_codes"]["MISSING_ASSIGNMENT_ID"]["level"] = "WARNING"
    return "$.identity_controls.exception_codes.MISSING_ASSIGNMENT_ID.level"


def _m_missing_code(d):
    del d["identity_controls"]["exception_codes"]["DUPLICATE_ROW"]
    return "$.identity_controls.exception_codes"


def _m_unknown_code(d):
    d["identity_controls"]["exception_codes"]["INVENTED"] = {"level": "WARNING", "Q": "Q5", "owner": "x",
                                                             "description": "x"}
    return "$.identity_controls.exception_codes.INVENTED"


def _m_bad_sha(d):
    d["identity_controls"]["template"]["sha256"] = "0" * 64
    return "$.identity_controls.template.sha256"


def _m_decision_outside_header(d):
    d["identity_controls"]["pending_header_decisions"][0]["cell"] = "A9"
    return "$.identity_controls.pending_header_decisions[0].cell"


def _m_bad_status(d):
    d["identity_controls"]["duplicate_compare_statuses"].append("everything")
    return f"$.identity_controls.duplicate_compare_statuses[{len(d['identity_controls']['duplicate_compare_statuses']) - 1}]"


def _m_missing_identity(d):
    del d["identity_controls"]
    return "$.identity_controls"


def _m_template_outside(d):
    d["identity_controls"]["template"]["file"] = "../../tests/fixtures/topia_sample.xlsx"
    return "$.identity_controls.template.file"


MUTATIONS = [_m_warn_missing_id, _m_missing_code, _m_unknown_code, _m_bad_sha, _m_decision_outside_header,
             _m_bad_status, _m_missing_identity, _m_template_outside]


@pytest.mark.parametrize("mutate", MUTATIONS, ids=[m.__name__[3:] for m in MUTATIONS])
def test_invalid_identity_configuration_rejected(tmp_path, mutate):
    data = _shipped()
    expected = mutate(data)
    with pytest.raises(ConfigurationError) as exc:
        transform(b"not a workbook", TEMPLATE.read_bytes(), business_rules=_write_rules(tmp_path, data))
    assert exc.value.path == expected


def test_fixture_cases_are_synthetic():
    text = (FIXTURES / "identity_cases.json").read_text(encoding="utf-8")
    assert "@" not in text  # no e-mail addresses
    for case in CASES["cases"].values():
        for spec in case["rows"]:
            for label, value in spec["set"].items():
                if label == "first name":
                    assert value.startswith("Synth")


# --------------------------------------------------------------------------- interplay with UAT-04 keyed dependents

def test_keyed_dependents_of_a_held_assignment_are_held_with_it(tmp_path):
    """A held assignment row (here MISSING_EMPLOYEE_NUMBER) never receives
    keyed dependents and they are never attached elsewhere: they are held as
    HELD_DEPENDENTS_ASSIGNMENT_HELD (Q18) until the corrected re-run."""
    sys.path.insert(0, str(ROOT / "tests"))
    import family_fixtures as ff
    data = _shipped()
    dep = data["dependent_source"]
    dep["approved_for_import"] = True
    dep["parameters"].update(sheet=ff.KEYED_SHEET, header_row=1, data_start_row=2)
    rules = _write_rules(tmp_path, data)

    wb = openpyxl.load_workbook(io.BytesIO(ff.KEYED.read_bytes()))
    ws = wb[ff.SHEET]
    ident_c, emp_c = _col_of(ws, "assignment id"), _col_of(ws, "employee id")
    held_row = next(r for r in range(5, ws.max_row + 1) if str(ws.cell(r, ident_c).value) == "90002")
    ws.cell(held_row, emp_c).value = None
    out, report_bytes, report = transform(_save(wb), TEMPLATE.read_bytes(), business_rules=rules)

    assert report.held_rows[held_row] == [cfg.EXC_MISSING_EMPLOYEE_NUMBER]
    written = _rows(out)
    assert "90002" not in {str(r.get(IDENT)) for r in written}
    assert "ChildB1" not in {r.get("ASSIGNMENT_DEPENDENT!FIRST_NAME") for r in written}
    held = [e for e in report.dependent_exceptions if e["code"] == "HELD_DEPENDENTS_ASSIGNMENT_HELD"]
    assert len(held) == 1 and (held[0]["rows"], held[0]["source_rows"]) == (2, "row 4, row 6")
    assert f"assignment source row {held_row} held" in held[0]["detail"]
    unmatched = [e for e in report.dependent_exceptions if e["code"] == "HELD_DEPENDENTS_UNMATCHED"]
    assert not any("row 4" in e["source_rows"] for e in unmatched)
    assert report.readiness.status == NOT_READY
