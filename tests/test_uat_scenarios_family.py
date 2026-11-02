"""UAT-06 spouse / dependent scenarios (master command section 7,
clarifications/UAT_MASTER_COMMAND.md:142-144; BR-10, BR-11, BR-18, BR-19,
BR-20; Jake Q10, Q11, Q18, Q19, Q20). Test ids FM-xx are listed in
docs/uat/TEST_RESULTS.md.

Fixtures (synthetic, tests/uat_scenario_fixtures.py):
- tests/fixtures/uat_family.xlsx: rows 5-10 per FAMILY_ROWS, dependents in
  appended 'Child k' ASSIGNMENT_DEPENDENT columns (Topia 'Dependent Import'
  layout), spouse in the GS-GX spouse columns.
- tests/fixtures/uat_family_linkage.xlsx: a keyed dependent sheet
  (LINKAGE_ROWS) for missing / incorrect linkage (F-10)."""

import functools
import json
import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import mapping_config as cfg  # noqa: E402
import uat_scenario_fixtures as fx  # noqa: E402
from transform import NOT_READY, transform  # noqa: E402

IDENT = fx.IDENT
TEMPLATE = fx.TEMPLATE.read_bytes()
DEP = "ASSIGNMENT_DEPENDENT!"
SPOUSE = "SPOUSE_DOMESTIC_PARTNER!"


@functools.lru_cache(maxsize=None)
def _family():
    return transform(fx.FAMILY.read_bytes(), TEMPLATE)


def _block_of_row(r: int) -> list[dict]:
    out, _, report = _family()
    ident = report.row_identity[r][0]
    return fx.block(fx.output_rows(out), ident)


def _deps(block: list[dict]) -> list[dict]:
    return [{k: v for k, v in row.items() if k.startswith(DEP)} for row in block
            if any(k.startswith(DEP) for k in row)]


def _spouse(block: list[dict]) -> dict:
    return {k: v for k, v in block[0].items() if k.startswith(SPOUSE)}


def test_family_fixture_accounts_for_every_row():
    _, _, report = _family()
    assert report.held_rows == {} and report.reconciliation_ok, report.reconciliation_summary.get("reasons")
    s = report.reconciliation_summary
    written = sum(v[0] for v in fx.FAMILY_EXPECTED.values())
    assert s["child_records_expected"]["dependent"] == s["child_records_written"]["dependent"] == written


@pytest.mark.parametrize("r", sorted(fx.FAMILY_EXPECTED))
def test_FM_rows_match_expected_counts(r):
    n_deps, has_spouse, _ = fx.FAMILY_EXPECTED[r]
    block = _block_of_row(r)
    assert len(_deps(block)) == n_deps
    assert bool(_spouse(block)) is has_spouse
    assert len(block) == max(1, n_deps)  # one INSERT row + one UPDATE row per further dependent (Q19)


def test_FM01_no_spouse_no_dependents_single_insert_row():
    block = _block_of_row(5)
    assert len(block) == 1 and block[0]["Assignment Command"] == "INSERT"
    assert not _deps(block) and not _spouse(block)


def test_FM02_one_dependent_on_insert_row():
    block = _block_of_row(6)
    [dep] = _deps(block)
    assert dep[DEP + "FIRST_NAME"] == "Synthkidb1" and dep[DEP + "RELATIONSHIP"] == "Child"
    assert len(block) == 1  # the only dependent sits on the INSERT row


def test_FM03_two_dependents_insert_then_update_same_identifier():
    block = _block_of_row(7)
    assert [r["Assignment Command"] for r in block] == ["INSERT", "UPDATE"]
    assert [r["A"] for r in block] == ["INSERT", "IGNORE"]
    assert "Employee Number" not in block[1] and "Assignment Number" not in block[1]  # Jake example rows 5-7
    assert {r[IDENT] for r in block} == {"80003"}
    assert [d[DEP + "FIRST_NAME"] for d in _deps(block)] == ["Synthkidc1", "Synthkidc2"]


def test_FM04_four_dependents_no_capacity_limit():
    block = _block_of_row(8)
    assert [r["Assignment Command"] for r in block] == ["INSERT", "UPDATE", "UPDATE", "UPDATE"]
    assert {r[IDENT] for r in block} == {"80004"}
    assert [d.get(DEP + "FIRST_NAME") for d in _deps(block)] == ["Synthkidd1", "Synthkidd2", "Synthkidd3", None]


def test_FM05_spouse_with_name_written_on_insert_row():
    spouse = _spouse(_block_of_row(6))
    assert spouse == {SPOUSE + "FIRST_NAME": "Synthspouse", SPOUSE + "LAST_NAME": "Testcase",
                      SPOUSE + "RELATIONSHIP": "Spouse"}
    _, _, report = _family()
    assert not [d for d in report.data_quality if d["where"] == "source row 6"]


def test_FM06_spouse_without_name_imported_with_data_quality_note_only():
    # Q10 (Jake J11): importable without names; never invented, never held.
    spouse = _spouse(_block_of_row(7))
    assert SPOUSE + "FIRST_NAME" not in spouse and SPOUSE + "LAST_NAME" not in spouse
    assert spouse[SPOUSE + "RELATIONSHIP"] == "Partner"
    assert spouse[SPOUSE + "BIRTH_DATE"].date() == date(1991, 3, 3)
    _, _, report = _family()
    [dq] = [d for d in report.data_quality if d["where"] == "source row 7" and d["record"] == SPOUSE[:-1]]
    assert dq["missing"] == "FIRST_NAME, LAST_NAME"
    # Relationship values are reported against the (pending) Equus list, never converted (Q14).
    assert any(r["item"] == SPOUSE + "RELATIONSHIP" for r in fx.reasons(report, "family_value_not_in_list"))


def test_FM07_dependent_with_one_populated_field_is_created():
    # Q20 (Jake J21): one populated field creates a dependent; no name invented.
    last = _deps(_block_of_row(8))[-1]
    assert last == {DEP + "BIRTH_DATE": last[DEP + "BIRTH_DATE"]} and last[DEP + "BIRTH_DATE"].date() == date(2020, 2, 2)
    _, _, report = _family()
    assert any(d["where"] == "source row 8" and d["fields"] == "BIRTH_DATE" for d in report.data_quality)


def test_FM08_multiple_assignments_each_with_own_dependents():
    a, b = _block_of_row(8), _block_of_row(9)
    assert a[0]["Employee Number"] == b[0]["Employee Number"] == fx.employee_id(4)
    assert (a[0]["A"], b[0]["A"]) == ("INSERT", "IGNORE")
    names_a = {d.get(DEP + "FIRST_NAME") for d in _deps(a)}
    names_b = {d.get(DEP + "FIRST_NAME") for d in _deps(b)}
    assert names_b == {"Synthkide1"} and not names_a & names_b  # Q18: no cross-assignment linkage
    assert {r[IDENT] for r in a} == {"80004"} and {r[IDENT] for r in b} == {"80005"}


def test_FM09_repeated_dependent_record_held_as_duplicate():
    block = _block_of_row(10)
    assert [d[DEP + "FIRST_NAME"] for d in _deps(block)] == ["Synthkidf1", "Synthkidf2"]  # written once
    _, report_bytes, report = _family()
    [e] = [e for e in report.dependent_exceptions if e["code"] == "HELD_DEPENDENTS_DUPLICATE"]
    assert (e["groups"], e["source_rows"], e["q"]) == (1, "source row 10", "Q5/Q19")
    assert any(r["item"] == "HELD_DEPENDENTS_DUPLICATE" for r in fx.reasons(report, "held_dependents"))
    assert report.readiness.status == NOT_READY
    # Reconciliation: the 4 cells of the held group are explained as held, nothing unexplained.
    assert report.reconciliation_summary["cells_held_row"] == 4
    assert report.reconciliation_summary["cells_unexplained"] == 0 and report.reconciliation_ok


# --------------------------------------------------------------------------- FM-10 linkage (keyed dependent source)

def _linkage_rules(tmp_path) -> Path:
    data = json.loads(cfg.BUSINESS_RULES_FILE.read_text(encoding="utf-8"))
    dep = data["dependent_source"]
    # Test-only configuration: the shipped dependent_source stays unconfigured
    # and not approved (Q11, D3); this switches it on for the synthetic sheet.
    dep["approved_for_import"] = True
    dep["parameters"].update(sheet=fx.KEYED_SHEET, header_row=1, data_start_row=2)
    path = tmp_path / "business_rules.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    return path


def test_FM10a_missing_or_incorrect_linkage_is_held_never_guessed(tmp_path):
    out, _, report = transform(fx.LINKAGE.read_bytes(), TEMPLATE, business_rules=_linkage_rules(tmp_path))
    rows = fx.output_rows(out)
    linked = _deps(fx.block(rows, "80001"))
    assert [d[DEP + "FIRST_NAME"] for d in linked] == ["Synthlinka"]  # row 2 only
    written = {d.get(DEP + "FIRST_NAME") for r in rows for d in _deps([r])}
    assert written == {"Synthlinka"}
    held = {e["code"]: e for e in report.dependent_exceptions}
    assert held["HELD_DEPENDENTS_NO_ASSIGNMENT_ID"]["source_rows"] == "row 3"  # never linked by Employee Number
    assert held["HELD_DEPENDENTS_EMPLOYEE_MISMATCH"]["source_rows"] == "row 5"
    assert set(held["HELD_DEPENDENTS_UNMATCHED"]["source_rows"].split(", ")) == {"row 4", "row 6"}
    assert report.readiness.status == NOT_READY and report.reconciliation_ok


def test_FM10b_keyed_sheet_without_approval_attaches_nothing():
    # Shipped configuration: dependent_source not approved (Q11, D3) -> held, nothing attached.
    out, _, report = transform(fx.LINKAGE.read_bytes(), TEMPLATE)
    assert not any(k.startswith(DEP) for r in fx.output_rows(out) for k in r)
    assert [e["code"] for e in report.dependent_exceptions] == ["HELD_DEPENDENTS_NOT_CONFIGURED"]
    assert report.dependent_exceptions[0]["groups"] == len(fx.LINKAGE_ROWS)
