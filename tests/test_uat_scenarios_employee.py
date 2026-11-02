"""UAT-06 employee / assignment scenarios (master command section 7,
clarifications/UAT_MASTER_COMMAND.md:138-141). Test ids EA-xx are listed in
docs/uat/TEST_RESULTS.md with their scenario and Q / BR mapping.

Fixture: tests/fixtures/uat_synthetic_base.xlsx (synthetic, built by
tests/uat_scenario_fixtures.py); each scenario rewrites data rows by Topia
row-4 label in memory. Local transformation only: no Equus behaviour is
observed or claimed here."""

import io
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import mapping_config as cfg  # noqa: E402
import transform as tmod  # noqa: E402
import uat_scenario_fixtures as fx  # noqa: E402
from transform import NOT_READY, transform  # noqa: E402

IDENT = fx.IDENT
TEMPLATE = fx.TEMPLATE.read_bytes()


def _source(rows: dict) -> bytes:
    """Base header rows; data rows = {source row: (person n, overrides)}."""
    wb = fx.load()
    ws = wb[fx.SHEET]
    fx.clear_data(ws)
    for r, (n, overrides) in rows.items():
        fx.set_person(ws, r, n, **overrides)
    return fx.save(wb)


def _run(rows: dict, **kwargs):
    return transform(_source(rows), TEMPLATE, **kwargs)


def _assert_accounted(report):
    assert report.data_rows == report.assignments + len(report.held_rows)
    assert report.reconciliation_ok, report.reconciliation_summary.get("reasons")


# --------------------------------------------------------------------------- EA-01 .. EA-03

def test_EA01_single_employee_one_assignment():
    out, report_bytes, report = _run({5: (1, {})})
    rows = fx.output_rows(out)
    assert len(rows) == 1
    [row] = rows
    assert (row["A"], row["Assignment Command"], row["Assignment Number"]) == ("INSERT", "INSERT", "NEW")
    assert row["Employee Number"] == fx.employee_id(1)
    # Q1/BR-01: Assignment Identifier and CF_TEXT029 both from Topia 'Assignment ID' (col X), as text.
    assert row[IDENT] == str(fx.assignment_id(1)) == row["ASSIGNMENT_CF!CF_TEXT029"]
    assert row["EMPLOYEE!FIRST_NAME"] == "Syntha" and row["EMPLOYEE!LAST_NAME"] == "Testcase"
    assert report.held_rows == {} and report.assignments == 1
    _assert_accounted(report)
    assert [a["identifier"] for a in report.accepted] == [str(fx.assignment_id(1))]


def test_EA02_one_employee_multiple_historical_assignments_never_merged():
    out, _, report = _run({
        5: (1, {}),
        6: (1, {"assignment_id": 80011, "assignment_start_date": datetime(2021, 6, 1),
                "assignment_end_date": datetime(2022, 6, 1)}),
        7: (1, {"assignment_id": "0080012", "assignment_start_date": datetime(2023, 1, 1),
                "assignment_end_date": datetime(2024, 1, 1)}),
    })
    ins = fx.inserts(fx.output_rows(out))
    # Q17 (Jake J18): one INSERT/NEW per assignment; Employee command INSERT then IGNORE.
    assert [r["A"] for r in ins] == ["INSERT", "IGNORE", "IGNORE"]
    assert all(r["Assignment Number"] == "NEW" and r["Employee Number"] == fx.employee_id(1) for r in ins)
    assert [r[IDENT] for r in ins] == ["80001", "80011", "0080012"]  # leading zeros kept
    assert [r["ASSIGNMENT!START_DATE"].date() for r in ins] == [date(2020, 1, 1), date(2021, 6, 1), date(2023, 1, 1)]
    assert report.held_rows == {} and not fx.warnings_for(report, "overlap")
    _assert_accounted(report)


def test_EA03_multiple_employees_each_insert():
    out, _, report = _run({5: (1, {}), 6: (2, {}), 7: (3, {})})
    ins = fx.inserts(fx.output_rows(out))
    assert [(r["Employee Number"], r[IDENT], r["A"]) for r in ins] == [
        (fx.employee_id(n), str(fx.assignment_id(n)), "INSERT") for n in (1, 2, 3)]
    assert report.reconciliation_summary["source_employees"] == 3
    _assert_accounted(report)


# --------------------------------------------------------------------------- EA-04 .. EA-07 identity controls

def test_EA04_duplicate_source_row_held_once_written():
    out, report_bytes, report = _run({5: (1, {}), 6: (2, {}), 7: (1, {})})
    assert [r[IDENT] for r in fx.inserts(fx.output_rows(out))] == ["80001", "80002"]
    assert report.held_rows == {7: [cfg.EXC_DUPLICATE_ROW]}
    [e] = fx.codes(report, cfg.EXC_DUPLICATE_ROW)
    assert e["level"] == cfg.LEVEL_ROW_HELD and e["q"] == "Q5" and "source row 5" in e["message"]
    assert [r["Source Row"] for r in fx.sheet_table(report_bytes, "Held Rows")[1:]] == [7]
    assert report.readiness.status == NOT_READY
    _assert_accounted(report)


def test_EA05_missing_employee_number_held():
    out, _, report = _run({5: (1, {}), 6: (2, {"employee_id": None})})
    assert report.held_rows == {6: [cfg.EXC_MISSING_EMPLOYEE_NUMBER]}
    assert "80002" not in {r.get(IDENT) for r in fx.output_rows(out)}
    [e] = fx.codes(report, cfg.EXC_MISSING_EMPLOYEE_NUMBER)
    assert e["identifier"] == "80002" and e["employee"] is None and e["q"] == "Q1"
    assert any(r["item"] == cfg.EXC_MISSING_EMPLOYEE_NUMBER for r in fx.reasons(report, "row_held"))
    _assert_accounted(report)


def test_EA06_missing_assignment_identifier_held_no_fallback():
    out, _, report = _run({5: (1, {}), 6: (2, {"assignment_id": None})})
    rows = fx.output_rows(out)
    assert report.held_rows == {6: [cfg.EXC_MISSING_ASSIGNMENT_ID]}
    # BR-01: the Employee Number is never used as Assignment Identifier / CF_TEXT029.
    emp = str(fx.employee_id(2))
    assert emp not in {str(r.get(IDENT)) for r in rows}
    assert emp not in {str(r.get("ASSIGNMENT_CF!CF_TEXT029")) for r in rows}
    assert fx.employee_id(2) not in {r.get("Employee Number") for r in rows}
    [e] = fx.codes(report, cfg.EXC_MISSING_ASSIGNMENT_ID)
    assert e["employee"] == emp and "not used as identifier" in e["message"]
    _assert_accounted(report)


def test_EA07a_same_identifier_different_data_holds_all_rows():
    out, _, report = _run({5: (1, {}), 6: (2, {}), 7: (2, {"first_name": "Synthother"})})
    assert [r[IDENT] for r in fx.inserts(fx.output_rows(out))] == ["80001"]
    assert report.held_rows == {6: [cfg.EXC_CONFLICTING_IDENTIFIER], 7: [cfg.EXC_CONFLICTING_IDENTIFIER]}
    assert all("Synthother" not in e["message"] for e in fx.codes(report, cfg.EXC_CONFLICTING_IDENTIFIER))
    _assert_accounted(report)


def test_EA07b_identifier_reused_by_other_employee_is_conflict():
    out, _, report = _run({5: (1, {}), 6: (2, {"assignment_id": fx.assignment_id(1)})})
    assert fx.output_rows(out) == []
    assert sorted(report.held_rows) == [5, 6]
    assert {e["code"] for e in report.exceptions if e["level"] == cfg.LEVEL_ROW_HELD} == {
        cfg.EXC_CONFLICTING_IDENTIFIER}
    _assert_accounted(report)


# --------------------------------------------------------------------------- EA-08 / EA-09 re-runs

def test_EA08_previously_imported_assignment_not_imported_again():
    rows = {5: (1, {}), 6: (2, {})}
    _, report1, _ = _run(rows)
    accepted = tmod.accepted_identifiers(report1)
    assert accepted == {"80001", "80002"}
    out, _, report = _run(rows, previously_accepted=accepted)
    assert fx.inserts(fx.output_rows(out)) == []
    assert report.held_rows == {5: [cfg.EXC_ALREADY_ACCEPTED], 6: [cfg.EXC_ALREADY_ACCEPTED]}
    assert report.readiness.status == NOT_READY
    _assert_accounted(report)


def _parts(xlsx: bytes) -> dict:
    with zipfile.ZipFile(io.BytesIO(xlsx)) as z:
        return {n: z.read(n) for n in z.namelist()}


def test_EA09_repeated_run_same_file_identical_output_modulo_timestamps():
    source = _source({5: (1, {}), 6: (2, {})})
    out1, rep1, r1 = transform(source, TEMPLATE)
    out2, rep2, r2 = transform(source, TEMPLATE)
    for a, b in ((out1, out2), (rep1, rep2)):
        pa, pb = _parts(a), _parts(b)
        assert pa.keys() == pb.keys()
        differing = sorted(n for n in pa if pa[n] != pb[n])
        # docProps/core.xml carries the workbook created/modified timestamps only.
        assert differing in ([], ["docProps/core.xml"]), differing
    assert fx.output_rows(out1) == fx.output_rows(out2)
    assert r1.readiness.reasons == r2.readiness.reasons and r1.exceptions == r2.exceptions


# --------------------------------------------------------------------------- EA-10 dates

def test_EA10a_overlapping_assignment_dates_warned_never_altered():
    out, _, report = _run({
        5: (1, {}),  # 2020-01-01 .. 2021-01-01
        6: (1, {"assignment_id": 80011, "assignment_start_date": datetime(2020, 6, 1),
                "assignment_end_date": datetime(2021, 6, 1)}),
    })
    ins = fx.inserts(fx.output_rows(out))
    assert [(r["ASSIGNMENT!START_DATE"].date(), r["ASSIGNMENT!ACTUAL_END_DATE"].date()) for r in ins] == [
        (date(2020, 1, 1), date(2021, 1, 1)), (date(2020, 6, 1), date(2021, 6, 1))]
    [w] = fx.warnings_for(report, "overlap")
    assert w[0] == "source row 6" and "source row 5" in w[2] and "kept unchanged" in w[2]
    assert report.held_rows == {}  # warning only, never a hold
    assert any(e["code"] == cfg.EXC_VALIDATION_WARNING and "overlap" in e["message"] for e in report.exceptions)
    _assert_accounted(report)


def test_EA10b_missing_and_inverted_dates_warned_nothing_derived():
    out, _, report = _run({
        5: (1, {"assignment_start_date": None, "arrival_date": None}),
        6: (2, {"assignment_end_date": datetime(2019, 1, 1)}),  # end before start
    })
    rows = {r[IDENT]: r for r in fx.inserts(fx.output_rows(out))}
    assert "ASSIGNMENT!START_DATE" not in rows["80001"]  # left blank, nothing derived (Q23)
    assert rows["80001"]["ASSIGNMENT!ACTUAL_END_DATE"].date() == date(2021, 1, 1)
    assert rows["80002"]["ASSIGNMENT!START_DATE"].date() == date(2020, 1, 1)
    assert rows["80002"]["ASSIGNMENT!ACTUAL_END_DATE"].date() == date(2019, 1, 1)  # kept as in source
    assert [w[0] for w in fx.warnings_for(report, "no assignment start date")] == ["source row 5"]
    assert [w[0] for w in fx.warnings_for(report, "end date before start date")] == ["source row 6"]
    assert report.held_rows == {}
    _assert_accounted(report)


# --------------------------------------------------------------------------- EA-11 existing employee, new assignment

def test_EA11a_existing_employee_new_assignment_same_file_ignore_insert():
    out, _, report = _run({5: (1, {}), 6: (1, {"assignment_id": 80021, "assignment_start_date": datetime(2022, 1, 1),
                                               "assignment_end_date": None})})
    rows = fx.output_rows(out)
    first, second = fx.inserts(rows)
    assert (first["A"], first["Assignment Command"]) == ("INSERT", "INSERT")
    # Jake ' Mny dependents Ex' / Q17: the employee exists -> Employee command IGNORE,
    # the new assignment is Assignment Command INSERT with Assignment Number NEW.
    assert (second["A"], second["Assignment Command"], second["Assignment Number"]) == ("IGNORE", "INSERT", "NEW")
    assert second["Employee Number"] == fx.employee_id(1) and second[IDENT] == "80021"
    assert not fx.codes(report, cfg.EXC_ROW_STRUCTURE_INVALID)
    _assert_accounted(report)


def test_EA11b_existing_employee_from_earlier_run_new_assignment_pending_q26():
    """Run 1 imported 80001 for employee 70001. Run 2 carries 80001 again and a
    new assignment 80031 of the same employee: 80001 is held ALREADY_ACCEPTED,
    80031 is an Assignment INSERT/NEW. The engine only knows identifiers from
    earlier runs, so the Employee command of 80031 is INSERT (first accepted
    assignment of the employee in this file). Whether Equus needs IGNORE for an
    employee created by an earlier batch is Q26 (INSERT/UPDATE/IGNORE semantics,
    OPEN, owner Jake): characterised here, finding F-01 in TEST_RESULTS.md."""
    _, report1, _ = _run({5: (1, {})})
    accepted = tmod.accepted_identifiers(report1)
    out, _, report = _run({5: (1, {}), 6: (1, {"assignment_id": 80031, "assignment_start_date": datetime(2022, 1, 1),
                                               "assignment_end_date": None})}, previously_accepted=accepted)
    [row] = fx.inserts(fx.output_rows(out))
    assert (row[IDENT], row["Assignment Command"], row["Assignment Number"]) == ("80031", "INSERT", "NEW")
    assert row["A"] == cfg.EMPLOYEE_ACTION_FIRST  # characterisation, pending Q26
    assert report.held_rows == {5: [cfg.EXC_ALREADY_ACCEPTED]}
    assert report.readiness.status == NOT_READY
    _assert_accounted(report)
