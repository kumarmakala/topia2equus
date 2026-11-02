"""UAT-04: spouse and dependent rules (Jake Q10, Q11, Q18, Q19, Q20) and a
keyed dependent source.

Jake (clarifications/jake-1.xlsx, sheet 'datamapping clarrification shef'):
  J11 (Q10) "Spouse/Parnter records can be imported without names."
  J12 (Q11) "... the Assignment Identifier will link the dependents to the same assignment."
  J19 (Q18) "Each spouse/dependet row should be accompanied by an Assignment Identifier."
  J20 (Q19) "As many dependents as are needed can be added as long as the rows contain the
            Assignment Identifier" (example sheet ' Mny dependents Ex' rows 4-7).
  J21 (Q20) "While not recommended, only one dependent field is needed to create a dependent record."
Every tracker Status is still Open: these answers are evidence, not approval.

Fixtures are synthetic only: tests/fixtures/topia_sample.xlsx and the two
family_dependents_*.xlsx files built by tests/family_fixtures.py."""

import io
import json
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import family_fixtures as ff  # noqa: E402
import mapping_config as cfg  # noqa: E402
from mapping_config import ConfigurationError  # noqa: E402
from transform import NOT_READY, transform  # noqa: E402

TEMPLATE = ROOT / "lambda" / "src" / "templates" / "equus_eafh_template.xlsx"
SHIPPED_RULES = ROOT / "lambda" / "src" / "business_rules.json"
JAKE = ROOT / "clarifications" / "jake-1.xlsx"
JAKE_EXAMPLE = " Mny dependents Ex"

A, CMD, EMP, NUM, IDENT = "A", "Assignment Command", "Employee Number", "Assignment Number", "Assignment Identifier"
SPOUSE = "SPOUSE_DOMESTIC_PARTNER"
DEP = "ASSIGNMENT_DEPENDENT"
NOTE = "DATA_QUALITY: record has no first/last name; importable per Q10/Q20, business review recommended"


# --------------------------------------------------------------------------- helpers

def _rows(xlsx: bytes) -> list[dict]:
    ws = openpyxl.load_workbook(io.BytesIO(xlsx))["Sheet1"]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    return [{h: ws.cell(r, c).value for c, h in enumerate(headers, 1) if ws.cell(r, c).value is not None}
            for r in range(4, ws.max_row + 1)]


def _family(row: dict, prefix: str) -> dict:
    return {k: v for k, v in row.items() if k.split("!")[0] in (prefix, prefix + "_CF")}


def _rules_copy(tmp_path, **params) -> Path:
    """business_rules.json copy with $.dependent_source changed (test only)."""
    data = json.loads(SHIPPED_RULES.read_text(encoding="utf-8"))
    dep = data["dependent_source"]
    if "approved_for_import" in params:
        dep["approved_for_import"] = params.pop("approved_for_import")
    dep["parameters"].update(params)
    path = tmp_path / "business_rules.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    return path


def _keyed_rules(tmp_path, approved=True) -> Path:
    return _rules_copy(tmp_path, sheet=ff.KEYED_SHEET, header_row=1, data_start_row=2, approved_for_import=approved)


def _with_sheet_dependents(base: bytes, children: dict[int, list[list]]) -> bytes:
    """Append repeated Child N column groups (FIRST_NAME, BIRTH_DATE) to the
    assignment sheet; children = {source row: [[first, birth], ...]}."""
    wb = openpyxl.load_workbook(io.BytesIO(base))
    ws = wb[ff.SHEET]
    n = max(len(v) for v in children.values())
    start = ws.max_column + 1
    for g in range(n):
        for i, (key, label) in enumerate([(DEP + "!FIRST_NAME", f"Child {g + 1} First Name"),
                                          (DEP + "!BIRTH_DATE", f"Child {g + 1} Date of Birth")]):
            ws.cell(2, start + 2 * g + i, key)
            ws.cell(4, start + 2 * g + i, label)
    for r, groups in children.items():
        for g, values in enumerate(groups):
            for i, value in enumerate(values):
                ws.cell(r, start + 2 * g + i, value)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _base() -> bytes:
    return ff.BASE.read_bytes()


def _held(report, code):
    return [e for e in report.dependent_exceptions if e["code"] == code]


def _reasons(report, category):
    return [r for r in report.readiness.reasons if r["category"] == category]


# --------------------------------------------------------------------------- AC-1 Q10 / Q20

def test_spouse_without_names_is_written_with_data_quality_note_not_held():
    out, report_bytes, report = transform(_base(), TEMPLATE.read_bytes())
    rows = _rows(out)
    assert len(rows) == 2  # nothing dropped, nothing extra
    for row in rows:
        spouse = _family(row, SPOUSE)
        assert spouse[SPOUSE + "!RELATIONSHIP"]  # written as sent
        assert SPOUSE + "!FIRST_NAME" not in spouse and SPOUSE + "!LAST_NAME" not in spouse  # never invented
    notes = [m for _, f, m in report.warnings if f == SPOUSE]
    assert len(notes) == 2 and all(m.startswith(NOTE) for m in notes)
    assert not any("Equus may reject" in m for _, _, m in report.warnings)
    assert [n["record"] for n in report.data_quality] == [SPOUSE, SPOUSE]
    assert all(n["missing"] == "FIRST_NAME, LAST_NAME" for n in report.data_quality)
    # A note, not a hold: no held family item, no readiness reason for names.
    assert not report.dependent_exceptions and not _reasons(report, "held_dependents")
    assert not any("first/last name" in r["detail"] for r in report.readiness.reasons)
    wb = openpyxl.load_workbook(io.BytesIO(report_bytes))
    family = [list(r) for r in wb["Family"].iter_rows(values_only=True)]
    assert sum(1 for r in family if r and r[0] == cfg.FAMILY_NAME_NOTE) == 2


def test_dependent_with_only_birth_date_is_written_and_empty_group_creates_no_record():
    src = _with_sheet_dependents(_base(), {5: [[None, datetime(2011, 3, 3)], [None, None]], 6: [[None, None]]})
    out, _, report = transform(src, TEMPLATE.read_bytes())
    rows = _rows(out)
    assert len(rows) == 2  # one INSERT row per assignment; no row for the empty groups
    first = next(r for r in rows if r[IDENT] == "90001")
    assert _family(first, DEP) == {DEP + "!BIRTH_DATE": datetime(2011, 3, 3)}
    second = next(r for r in rows if r[IDENT] == "90002")
    assert _family(second, DEP) == {}
    dep_notes = [n for n in report.data_quality if n["record"] == DEP]
    assert len(dep_notes) == 1 and dep_notes[0]["fields"] == "BIRTH_DATE"
    assert not report.dependent_exceptions


def test_relationship_values_checked_against_equus_list_never_converted():
    # No Equus list bundled (Q14 pending): every value is "not in list".
    out, _, report = transform(_base(), TEMPLATE.read_bytes(), allowed_values={})
    values = {v for (_, v) in report.family_values}
    assert values == {"Spouse", "Partner"}
    for (header, value) in report.family_values:
        assert report.family_value_status(header, value).startswith("not in list (no Equus list; Q14")
    reason = _reasons(report, "family_value_not_in_list")
    assert [(r["item"], r["q"]) for r in reason] == [(SPOUSE + "!RELATIONSHIP", "Q14")]
    assert "Partner" not in reason[0]["detail"]  # readiness carries counts, not source values
    # With a list: checked, reported, still sent unchanged.
    out, _, report = transform(_base(), TEMPLATE.read_bytes(),
                               allowed_values={SPOUSE + "!RELATIONSHIP": ["Spouse"]})
    assert report.family_value_status(SPOUSE + "!RELATIONSHIP", "Spouse") == "in list"
    assert report.family_value_status(SPOUSE + "!RELATIONSHIP", "Partner") == "not in list"
    assert sorted(r[SPOUSE + "!RELATIONSHIP"] for r in _rows(out)) == ["Partner", "Spouse"]
    assert any("'Partner'" in m and "not in list" in m for _, f, m in report.warnings if f == SPOUSE + "!RELATIONSHIP")


# --------------------------------------------------------------------------- AC-2 Q19 layout

def _jake_pattern():
    """Structure of Jake's example (' Mny dependents Ex' rows 4-7): commands,
    which control cells are blank, identifier equality, family presence.
    Only the structure is compared; no example value is copied."""
    ws = openpyxl.load_workbook(JAKE, data_only=True)[JAKE_EXAMPLE]
    headers = {ws.cell(3, c).value: c for c in range(1, ws.max_column + 1)}
    rows = []
    for r in range(4, ws.max_row + 1):
        row = {h: ws.cell(r, c).value for h, c in headers.items() if ws.cell(r, c).value is not None}
        row["__A__"] = ws.cell(r, 1).value  # A3 reads 'Employee Command' there, 'A' in the template (D2)
        rows.append(row)
    return _pattern(rows)


def _pattern(rows):
    ids = {str(r.get(IDENT)) for r in rows}
    return {
        "employee_command": [r.get("__A__", r.get(A)) for r in rows],
        "assignment_command": [r.get(CMD) for r in rows],
        "employee_number_present": [EMP in r for r in rows],
        "assignment_number": [r.get(NUM) for r in rows],
        "identifier_on_every_row": all(IDENT in r for r in rows),
        "one_identifier": len(ids) == 1,
        "spouse_present": [bool(_family(r, SPOUSE)) for r in rows],
        "dependent_present": [bool(_family(r, DEP)) for r in rows],
    }


def test_jake_example_has_the_structure_this_test_encodes():
    assert _jake_pattern() == {
        "employee_command": ["INSERT", "IGNORE", "IGNORE", "IGNORE"],
        "assignment_command": ["INSERT", "UPDATE", "UPDATE", "UPDATE"],
        "employee_number_present": [True, False, False, False],
        "assignment_number": ["NEW", None, None, None],
        "identifier_on_every_row": True,
        "one_identifier": True,
        "spouse_present": [True, False, False, False],
        "dependent_present": [True, True, True, True],
    }


def test_four_sheet_dependents_match_jake_example_structure():
    kids = [[f"Child{i}", datetime(2010 + i, 1, 1)] for i in range(1, 5)]
    out, _, report = transform(_with_sheet_dependents(_base(), {5: kids}), TEMPLATE.read_bytes())
    rows = [r for r in _rows(out) if r[IDENT] == "90001"]
    assert _pattern(rows) == _jake_pattern()
    assert [r[DEP + "!FIRST_NAME"] for r in rows] == ["Child1", "Child2", "Child3", "Child4"]
    assert rows[0][SPOUSE + "!RELATIONSHIP"]  # spouse on the INSERT row with the first dependent
    assert rows[0][EMP] == 10001
    assert report.assignments == 2 and report.reconciliation_ok


def test_four_keyed_dependents_match_jake_example_structure(tmp_path):
    out, _, report = transform(ff.KEYED.read_bytes(), TEMPLATE.read_bytes(), business_rules=_keyed_rules(tmp_path))
    rows = [r for r in _rows(out) if r[IDENT] == "90001"]
    assert _pattern(rows) == _jake_pattern()
    # Q20: the fourth dependent has only a birth date and is still written.
    assert [r.get(DEP + "!FIRST_NAME") for r in rows] == ["Child1", "Child2", "Child3", None]
    assert _family(rows[3], DEP) == {DEP + "!BIRTH_DATE": datetime(2015, 5, 5)}


@pytest.mark.parametrize("count", [1, 2, 3, 7])
def test_any_number_of_dependents_gives_insert_plus_updates(count):
    kids = [[f"Child{i}", None] for i in range(1, count + 1)]
    out, _, _ = transform(_with_sheet_dependents(_base(), {5: kids}), TEMPLATE.read_bytes())
    rows = [r for r in _rows(out) if r[IDENT] == "90001"]
    assert [r[CMD] for r in rows] == ["INSERT"] + ["UPDATE"] * (count - 1)
    assert all(EMP not in r and NUM not in r and r[A] == "IGNORE" for r in rows[1:])


# --------------------------------------------------------------------------- AC-3 Q17 / Q18 isolation

def _same_employee_two_assignments() -> bytes:
    wb = openpyxl.load_workbook(ff.BASE)
    ws = wb[ff.SHEET]
    ws.cell(6, ff._labels(ws)["employee id"], 10001)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_two_assignments_of_one_employee_never_share_sheet_dependents():
    src = _with_sheet_dependents(_same_employee_two_assignments(),
                                 {5: [["ChildA1", None], ["ChildA2", None]], 6: [["ChildB1", None]]})
    out, _, _ = transform(src, TEMPLATE.read_bytes())
    rows = _rows(out)
    by_id = {}
    for row in rows:
        by_id.setdefault(row[IDENT], []).append(row)
    assert set(by_id) == {"90001", "90002"}
    assert [r.get(DEP + "!FIRST_NAME") for r in by_id["90001"]] == ["ChildA1", "ChildA2"]
    assert [r.get(DEP + "!FIRST_NAME") for r in by_id["90002"]] == ["ChildB1"]
    # Second assignment of the same employee: employee IGNORE, assignment INSERT.
    assert (by_id["90002"][0][A], by_id["90002"][0][CMD], by_id["90002"][0][EMP]) == ("IGNORE", "INSERT", 10001)
    # Spouse stays on each assignment's own INSERT row.
    assert all(_family(r, SPOUSE) for r in (by_id["90001"][0], by_id["90002"][0]))
    assert not any(_family(r, SPOUSE) for r in by_id["90001"][1:])


def test_keyed_dependents_stay_with_their_own_assignment(tmp_path):
    out, _, report = transform(ff.KEYED.read_bytes(), TEMPLATE.read_bytes(), business_rules=_keyed_rules(tmp_path))
    rows = _rows(out)
    names = {}
    for row in rows:
        if _family(row, DEP):
            names.setdefault(row[IDENT], []).append(row.get(DEP + "!FIRST_NAME"))
    assert names == {"90001": ["Child1", "Child2", "Child3", None], "90002": ["ChildB1"]}
    assert report.dependents_attached == 5
    assert all(r[IDENT] in ("90001", "90002") for r in rows)


def test_keyed_rows_with_missing_conflicting_or_unknown_keys_are_held(tmp_path):
    out, _, report = transform(ff.KEYED.read_bytes(), TEMPLATE.read_bytes(), business_rules=_keyed_rules(tmp_path))
    written = {r.get(DEP + "!FIRST_NAME") for r in _rows(out)}
    for name in ("ChildNoKey", "ChildWrongEmp", "ChildByEmpNo"):
        assert name not in written
    held = {e["code"]: (e["rows"], e["groups"], e["source_rows"]) for e in report.dependent_exceptions}
    assert held == {
        "HELD_DEPENDENTS_NO_ASSIGNMENT_ID": (1, 1, "row 5"),
        "HELD_DEPENDENTS_EMPLOYEE_MISMATCH": (1, 1, "row 6"),
        # Assignment ID equal to an Employee Number never links to that employee.
        "HELD_DEPENDENTS_UNMATCHED": (1, 1, "row 7"),
    }
    assert report.readiness.status == NOT_READY
    assert {r["item"] for r in _reasons(report, "held_dependents")} == set(held)


def test_keyed_dependents_never_join_through_identifier_fallback(tmp_path):
    """An assignment row without its own Assignment ID never receives keyed
    dependents, even when they carry its Employee Number."""
    wb = openpyxl.load_workbook(io.BytesIO(ff.KEYED.read_bytes()))
    ws = wb[ff.SHEET]
    ws.cell(6, ff._labels(ws)["assignment id"]).value = None
    buf = io.BytesIO()
    wb.save(buf)
    out, _, report = transform(buf.getvalue(), TEMPLATE.read_bytes(), business_rules=_keyed_rules(tmp_path))
    assert "ChildB1" not in {r.get(DEP + "!FIRST_NAME") for r in _rows(out)}
    unmatched = _held(report, "HELD_DEPENDENTS_UNMATCHED")[0]
    assert unmatched["rows"] == 3  # 90002 rows 4 and 6 plus row 7


def test_duplicate_assignment_id_holds_keyed_dependents(tmp_path):
    wb = openpyxl.load_workbook(io.BytesIO(ff.KEYED.read_bytes()))
    ws = wb[ff.SHEET]
    ws.cell(6, ff._labels(ws)["assignment id"], 90001)
    buf = io.BytesIO()
    wb.save(buf)
    out, _, report = transform(buf.getvalue(), TEMPLATE.read_bytes(), business_rules=_keyed_rules(tmp_path))
    assert not any(_family(r, DEP) for r in _rows(out))
    amb = _held(report, "HELD_DEPENDENTS_AMBIGUOUS")
    assert len(amb) == 1 and (amb[0]["rows"], amb[0]["groups"], amb[0]["source_rows"]) == (2, 4, "row 2, row 3")


# --------------------------------------------------------------------------- AC-4 keyed vs unkeyed source

def test_unkeyed_dependent_import_sheet_is_held_and_nothing_attached():
    out, report_bytes, report = transform(ff.UNKEYED.read_bytes(), TEMPLATE.read_bytes())
    rows = _rows(out)
    assert len(rows) == 2 and not any(_family(r, DEP) for r in rows)
    held = _held(report, "HELD_DEPENDENTS_UNKEYED")
    assert len(held) == 1
    assert (held[0]["sheet"], held[0]["rows"], held[0]["groups"]) == (ff.UNKEYED_SHEET, 2, 4)
    assert held[0]["q"] == "Q11" and "Alka" in held[0]["owner"] and "D3" in held[0]["owner"]
    assert report.dependents_attached == 0
    assert report.readiness.status == NOT_READY
    reason = _reasons(report, "held_dependents")
    assert [(r["item"], r["q"]) for r in reason] == [("HELD_DEPENDENTS_UNKEYED", "Q11")]
    assert "2 dependent row(s), 4 dependent record(s) held" in reason[0]["detail"]
    assert "Child" not in reason[0]["detail"]  # counts, never source values
    wb = openpyxl.load_workbook(io.BytesIO(report_bytes))
    family = [list(r) for r in wb["Family"].iter_rows(values_only=True)]
    assert any(r[:4] == ["HELD_DEPENDENTS_UNKEYED", ff.UNKEYED_SHEET, 2, 4] for r in family)


def test_unkeyed_sheet_stays_held_even_with_a_keyed_source_configured(tmp_path):
    rules = _rules_copy(tmp_path, sheet="Some Other Sheet", approved_for_import=True)
    out, _, report = transform(ff.UNKEYED.read_bytes(), TEMPLATE.read_bytes(), business_rules=rules)
    assert {e["code"] for e in report.dependent_exceptions} == {"HELD_DEPENDENTS_UNKEYED",
                                                                "HELD_DEPENDENTS_SHEET_MISSING"}
    assert not any(_family(r, DEP) for r in _rows(out))


def test_keyed_sheet_joins_by_assignment_id_when_configured_and_approved(tmp_path):
    out, report_bytes, report = transform(ff.KEYED.read_bytes(), TEMPLATE.read_bytes(),
                                          business_rules=_keyed_rules(tmp_path))
    rows = [r for r in _rows(out) if _family(r, DEP)]
    assert {r[IDENT] for r in rows} == {"90001", "90002"}
    assert report.dependents_attached == 5
    assert not _held(report, "HELD_DEPENDENTS_UNKEYED")
    wb = openpyxl.load_workbook(io.BytesIO(report_bytes))
    family = dict((r[0], r[1]) for r in wb["Family"].iter_rows(values_only=True) if r and r[0])
    assert family["Dependent source sheet"] == ff.KEYED_SHEET
    assert family["Dependents attached from keyed source"] == 5


def test_keyed_sheet_not_configured_or_not_approved_is_held(tmp_path):
    # Shipped configuration: no keyed sheet configured.
    out, _, report = transform(ff.KEYED.read_bytes(), TEMPLATE.read_bytes())
    assert [(e["code"], e["rows"], e["groups"]) for e in report.dependent_exceptions] == \
        [("HELD_DEPENDENTS_NOT_CONFIGURED", 6, 8)]
    assert not any(_family(r, DEP) for r in _rows(out))
    # Configured but not approved_for_import.
    out, _, report = transform(ff.KEYED.read_bytes(), TEMPLATE.read_bytes(),
                               business_rules=_keyed_rules(tmp_path, approved=False))
    assert [(e["code"], e["rows"]) for e in report.dependent_exceptions] == [("HELD_DEPENDENTS_GATED", 6)]
    assert not any(_family(r, DEP) for r in _rows(out))


def test_configured_keyed_sheet_without_key_columns_is_held_unkeyed(tmp_path):
    rules = _rules_copy(tmp_path, sheet=ff.UNKEYED_SHEET, header_row=1, data_start_row=3, approved_for_import=True)
    out, _, report = transform(ff.UNKEYED.read_bytes(), TEMPLATE.read_bytes(), business_rules=rules)
    held = _held(report, "HELD_DEPENDENTS_UNKEYED")
    assert len(held) == 1 and (held[0]["rows"], held[0]["groups"]) == (2, 4)
    assert "employee id, assignment id" in held[0]["detail"]
    assert not any(_family(r, DEP) for r in _rows(out))


# --------------------------------------------------------------------------- configuration

def test_shipped_dependent_source_is_unconfigured_and_unapproved():
    dep = cfg.load_business_rules().dependent_source
    assert dep is not None and dep.sheet is None and not dep.approved_for_import and not dep.enabled
    assert dep.q == "Q11/Q18" and dep.status in cfg.RULE_STATUSES and "Alka" in dep.owner
    assert (dep.employee_label, dep.assignment_label) == ("employee id", "assignment id")
    assert dep.version == cfg.load_business_rules().version
    assert cfg.DEPENDENT_SOURCE == dep


@pytest.mark.parametrize("change, path", [
    ({"approved_for_import": True}, "$.dependent_source.approved_for_import"),
    ({"data_start_row": 1}, "$.dependent_source.parameters.data_start_row"),
    ({"header_row": 0}, "$.dependent_source.parameters.header_row"),
    ({"child_groups": "guess"}, "$.dependent_source.parameters.child_groups"),
    ({"sheet": cfg.SOURCE_SHEET}, "$.dependent_source.parameters.sheet"),
    ({"key_columns": {"employee": "assignment id", "assignment": "assignment id"}},
     "$.dependent_source.parameters.key_columns"),
    ({"key_columns": {"assignment": "assignment id"}}, "$.dependent_source.parameters.key_columns.employee"),
    ({"key_columns": {"employee": "Employee ID", "assignment": "assignment id"}},
     "$.dependent_source.parameters.key_columns.employee"),
    ({"join_on": "employee"}, "$.dependent_source.parameters.join_on"),
])
def test_invalid_dependent_source_rejected_at_load(tmp_path, change, path):
    with pytest.raises(ConfigurationError) as exc:
        cfg.load_business_rules(_rules_copy(tmp_path, **change))
    assert exc.value.path == path
