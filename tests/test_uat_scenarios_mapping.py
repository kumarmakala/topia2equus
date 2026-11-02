"""UAT-06 mapping / data-quality scenarios (master command section 7,
clarifications/UAT_MASTER_COMMAND.md:145-148; BR-04, BR-06, BR-08, BR-12,
BR-13, BR-14, BR-16; Q3, Q4, Q5, Q6, Q8, Q12, Q13, Q14, Q22-Q24, Q29, Q41,
Q43). Test ids MD-xx are listed in docs/uat/TEST_RESULTS.md.

Fixture: tests/fixtures/uat_synthetic_base.xlsx (synthetic). Approved
reference lists used here are test-only copies in tmp_path; the shipped lists
stay empty / not approved, nothing here approves a value."""

import json
import shutil
import sys
from datetime import date
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import mapping_config as cfg  # noqa: E402
import uat_scenario_fixtures as fx  # noqa: E402
from transform import NOT_READY, PendingFieldsError, transform  # noqa: E402

IDENT = fx.IDENT
TEMPLATE = fx.TEMPLATE.read_bytes()


def _source(**row5) -> bytes:
    """Base fixture with overrides (label with '_' for spaces) on source row 5."""
    wb = fx.load()
    ws = wb[fx.SHEET]
    if row5:
        fx.write_row(ws, 5, {k.replace("_", " "): v for k, v in row5.items()})
    return fx.save(wb)


def _run(source: bytes | None = None, **kwargs):
    return transform(source if source is not None else fx.BASE.read_bytes(), TEMPLATE, **kwargs)


def _row(out: bytes, ident: str = "80001") -> dict:
    [row] = [r for r in fx.inserts(fx.output_rows(out)) if r[IDENT] == ident]
    return row


def _recon(report, letter: str) -> dict:
    return next(r for r in report.reconciliation if r["col"] == letter)


def _col_letter(label: str) -> str:
    ws = fx.load()[fx.SHEET]
    return openpyxl.utils.get_column_letter(fx.labels(ws)[label])


def _approved_country_state(tmp_path, values: dict) -> Path:
    """Test-only reference data dir: shipped lists + an approved country_state."""
    ref = tmp_path / "reference_data"
    shutil.copytree(cfg.REFERENCE_DATA_DIR, ref)
    data = json.loads((ref / "country_state.json").read_text(encoding="utf-8"))
    data.update(approved_for_import=True, status="partially_clarified", version="test-1", effective_version="test-1",
                values=values)
    (ref / "country_state.json").write_text(json.dumps(data), encoding="utf-8")
    return ref


# --------------------------------------------------------------------------- MD-01 / MD-02 country and state (Q12, Q3)

def test_MD01a_country_names_provisional_unknown_passed_through():
    out, _, report = _run(_source(host_country="Atlantis"))
    row = _row(out)
    assert row["ASSIGNMENT!FROM_COUNTRY"] == "CAN"  # provisional ISO-3 conversion (rule not approved)
    assert row["ASSIGNMENT!TO_COUNTRY"] == "Atlantis"  # unknown: sent unchanged, never guessed
    assert fx.warnings_for(report, "no ISO-3 code for country 'Atlantis'")
    assert any(r["item"] == "ASSIGNMENT!FROM_COUNTRY" for r in fx.reasons(report, "provisional_conversion"))
    assert any(r["item"] == "country_name_to_iso3" and r["q"] == "Q12/Q3" for r in fx.reasons(report, "gated_rule"))
    assert any(r["item"] == "country_state" for r in fx.reasons(report, "reference_list_unavailable"))


def test_MD01b_country_against_approved_list_unknown_reported(tmp_path):
    refs = _approved_country_state(tmp_path, {"CAN": ["Ontario"], "DEU": ["Bavaria"]})
    out, _, report = _run(_source(host_country="Atlantis"), reference_data=refs)
    assert _row(out)["ASSIGNMENT!TO_COUNTRY"] == "Atlantis"
    [r] = [r for r in fx.reasons(report, "reference_value_unknown") if r["item"] == "country_state"]
    assert "ASSIGNMENT!TO_COUNTRY" in r["detail"] and r["q"] == "Q12/Q3"
    assert fx.warnings_for(report, "'Atlantis' (1x) not in approved reference list country_state")


def test_MD02a_state_not_in_approved_list_reported_never_converted(tmp_path):
    refs = _approved_country_state(tmp_path, {"CAN": ["Ontario"], "DEU": ["Bavaria"]})
    out, _, report = _run(_source(host_state="Nowhere"), reference_data=refs)
    row = _row(out)
    assert row["ASSIGNMENT!TO_STATE_PROVINCE"] == "Nowhere" and row["ASSIGNMENT!FROM_STATE_PROVINCE"] == "Ontario"
    [r] = [r for r in fx.reasons(report, "reference_value_unknown") if r["item"] == "country_state (state)"]
    assert r["detail"].startswith("1 distinct state value(s)") and r["q"] == "Q12/Q3"
    [w] = fx.warnings_for(report, "state 'Nowhere'")
    assert w[1] == "ASSIGNMENT!TO_STATE_PROVINCE" and "'DEU'" in w[2]
    assert not fx.warnings_for(report, "state 'Ontario'") and not fx.warnings_for(report, "state 'Bavaria'")
    # Row 6 (person 2) has valid states: no reason other than the one unknown value.
    assert report.readiness.status == NOT_READY


def test_MD02b_state_unvalidated_while_list_not_approved():
    out, _, report = _run()
    assert _row(out)["ASSIGNMENT!TO_STATE_PROVINCE"] == "Bavaria"
    [r] = [r for r in fx.reasons(report, "reference_list_unavailable") if r["item"] == "country_state (state)"]
    lst = cfg.load_reference_data()[cfg.STATE_LIST_ID]
    assert not lst.usable and f"list {lst.state}:" in r["detail"] and "cannot be validated" in r["detail"]
    assert r["q"] == "Q12/Q3" and "Jake" in r["owner"]


# --------------------------------------------------------------------------- MD-03 currency (Q4, BR-04)

CURRENCY_HEADERS = ("ASSIGNMENT!FROM_CURRENCY", "ASSIGNMENT!TO_CURRENCY", "ASSIGNMENT!DEFAULT_CURRENCY",
                    "ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_CURRENCY")


def test_MD03_currency_blank_never_derived_from_country():
    out, _, report = _run(_source(home_annualized_bse_pay_amount=50000))
    row = _row(out)
    assert row["ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_AMOUNT"] == 50000
    assert row["ASSIGNMENT!FROM_COUNTRY"] == "CAN" and row["ASSIGNMENT!TO_COUNTRY"] == "DEU"
    for r in fx.output_rows(out):
        assert not set(CURRENCY_HEADERS) & set(r)  # no home/host/salary/default currency invented


# --------------------------------------------------------------------------- MD-04 dates (Q24)

def test_MD04a_ambiguous_text_date_read_with_configured_order_and_reported():
    out, _, report = _run(_source(arrival_date="05/06/2024"))
    assert _row(out)["ASSIGNMENT!HOST_COUNTRY_ARRIVAL_DATE"].date() == date(2024, 5, 6)  # month/day only
    [w] = fx.warnings_for(report, "ambiguous date text")
    assert w[1] == "ASSIGNMENT!HOST_COUNTRY_ARRIVAL_DATE" and "MM/DD/YYYY" in w[2] and "Q24" in w[2]
    assert report.reconciliation_ok


def test_MD04b_unambiguous_text_date_not_reported():
    out, _, report = _run(_source(arrival_date="06/13/2024"))
    assert _row(out)["ASSIGNMENT!HOST_COUNTRY_ARRIVAL_DATE"].date() == date(2024, 6, 13)
    assert not fx.warnings_for(report, "ambiguous date text")


@pytest.mark.parametrize("text", ["13/06/2024", "31/31/2024", "not a date"])
def test_MD04c_invalid_date_left_blank_warned_and_reconciled(text):
    # Day-first 13/06/2024 is not converted: the configured order is month/day.
    out, _, report = _run(_source(arrival_date=text))
    assert "ASSIGNMENT!HOST_COUNTRY_ARRIVAL_DATE" not in _row(out)
    assert fx.warnings_for(report, f"unparseable date {text!r}; left blank")
    assert _recon(report, _col_letter("arrival date"))["failed_conversion"] == 1
    assert report.reconciliation_ok


# --------------------------------------------------------------------------- MD-05 salary / effective date (Q21, Q22)

def test_MD05_salary_without_effective_date_no_default_exception():
    out, _, report = _run(_source(home_annualized_bse_pay_amount=50000))
    row = _row(out)
    assert "ASSIGNMENT_COMPENSATION_SALARY!EFFECTIVE_DATE" not in row  # no default from the start date
    [r] = [r for r in fx.reasons(report, "gated_rule") if r["item"] == cfg.RULE_DEFAULT_DATES]
    assert r["q"] == "Q22/Q23" and "rule not applied" in r["detail"]
    assert 5 in report.rule_rows[cfg.RULE_DEFAULT_DATES]
    assert report.readiness.status == NOT_READY


# --------------------------------------------------------------------------- MD-06 optional vs mandatory

def test_MD06_blank_optional_fields_accepted_mandatory_identity_held():
    out, _, report = _run(_source(date_of_birth=None, gender=None, marital_status=None, business_title=None))
    row = _row(out)
    assert not {"EMPLOYEE!BIRTH_DATE", "EMPLOYEE!GENDER", "EMPLOYEE!MARITAL_STATUS"} & set(row)
    assert report.held_rows == {}  # optional blanks never hold a row
    _, _, held = _run(_source(employee_id=None, assignment_id=None))
    assert held.held_rows == {5: [cfg.EXC_MISSING_ASSIGNMENT_ID, cfg.EXC_MISSING_EMPLOYEE_NUMBER]}


# --------------------------------------------------------------------------- MD-07 / MD-08 colours and custom fields

def test_MD07_red_yellow_cyan_brown_flags():
    out, _, report = _run(_source(preferred_pronouns="synthetic/value"))
    cols = {c.label: c for c in report.columns}
    # red = not required -> excluded, never written (Q30 disposition pending)
    assert cols["preferred pronouns"].status == "excluded"
    assert _recon(report, cols["preferred pronouns"].letter)["excluded_colour"] == 1  # row 5 only holds data
    assert "synthetic/value" not in {v for r in fx.output_rows(out) for v in r.values()}
    # yellow = pending -> PENDING flag, readiness reason, value written for review only
    assert "PENDING" in cols["person status"].flags
    assert any(r["item"].startswith("source S ") for r in fx.reasons(report, "pending_column"))
    # cyan without Equus column -> MISSING_TARGET; brown = no meaning (Q8), never UNKNOWN_COLOUR
    assert "MISSING_TARGET" in cols["expected assignment outcome"].flags
    assert cols["spouse first name"].colour == cfg.NO_MEANING_COLOUR
    assert not fx.reasons(report, "unknown_colour")
    with pytest.raises(PendingFieldsError):
        _run(strict_pending=True)


def test_MD08_approved_custom_field_targets_written_as_text():
    out, _, _ = _run()
    row = _row(out)
    assert row["EMPLOYEE_CF!CF_TEXT014"] == str(990000000000001)  # Global ID (cyan)
    assert row["EMPLOYEE_CF!CF_TEXT015"] == "SYNPH001"  # Pre-Hire ID (cyan)
    assert row["ASSIGNMENT_CF!CF_TEXT029"] == "80001"  # Topia Assignment ID (Q1, D1 open)


# --------------------------------------------------------------------------- MD-09 e-mail (Q13, Q41)

def test_MD09_email_mapping_no_contact_records():
    out, _, report = _run()
    row = _row(out)
    assert row["EMPLOYEE!USER_ID_EMAIL"] == "synth.a@example.invalid"
    assert row["EMPLOYEE_CF!CF_TEXT008"] == "synth.a.alt@example.invalid"
    contact = [r for r in fx.output_rows(out) if any(k.startswith("ASSIGNMENT_EMPLOYEE_CONTACT!") for k in r)]
    assert contact == []  # Q13: no e-mail contact type rows
    written = {v for r in fx.output_rows(out) for v in r.values() if isinstance(v, str)}
    assert "synth.coord@example.invalid" not in written  # resource e-mails unmapped (Q41)
    cols = {c.label: c for c in report.columns}
    assert _recon(report, cols["primary mobility email"].letter)["q"] == "Q41"


# --------------------------------------------------------------------------- MD-10 placeholders (Q43)

def test_MD10_placeholders_blank_and_glued_placeholder_reported():
    out, _, report = _run(_source(nationality_3="Not ProvidedHome", middle_name="N/A"))
    row = _row(out)
    assert "EMPLOYEE_CF!CF_TEXT013" not in row and "EMPLOYEE!MIDDLE_NAME" not in row
    assert "EMPLOYEE_CF!CF_TEXT005" not in row  # Nationality 2 'Not Provided'
    assert fx.warnings_for(report, "placeholder with extra text 'Not ProvidedHome'")
    assert _recon(report, _col_letter("nationality 3"))["blank_token"] == 2
    assert report.reconciliation_ok


# --------------------------------------------------------------------------- MD-11 confidential (Q6, Q28)

def test_MD11_confidential_attribute_held_never_written():
    out, _, report = _run(_source(confidentiality_type="Synthetic Confidential", is_confidential=True))
    written = {v for r in fx.output_rows(out) for v in r.values()}
    assert "Synthetic Confidential" not in written
    letter = _col_letter("confidentiality type")
    [r] = [r for r in fx.reasons(report, "missing_target") if r["item"].startswith(f"source {letter} ")]
    assert r["q"] == "Q6/Q7/Q29"
    recon = _recon(report, letter)
    assert recon["pending_target"] == 2 and "Q6" in recon["q"] and recon["unexplained"] == 0
    assert report.readiness.status == NOT_READY


# --------------------------------------------------------------------------- MD-12 header mismatch (Q5)

def test_MD12_header_mismatch_fails_whole_file():
    wb = openpyxl.load_workbook(fx.TEMPLATE)
    wb["Sheet1"]["G3"] = "EMPLOYEE!LAST_NAME_X"
    out, report_bytes, report = transform(fx.BASE.read_bytes(), fx.save(wb))
    assert out == b"" and report.file_failed
    [e] = fx.codes(report, cfg.EXC_HEADER_MISMATCH)
    assert e["level"] == cfg.LEVEL_FILE and "G3" in e["message"]
    assert fx.sheet_table(report_bytes, "Exceptions")  # diagnostics kept


# --------------------------------------------------------------------------- MD-13 .. MD-15 values outside lists

def test_MD13_value_outside_equus_list_reported_unchanged():
    out, _, report = _run(allowed_values={"EMPLOYEE!GENDER": ["Female"]})
    assert _row(out, "80002")["EMPLOYEE!GENDER"] == "Male"  # never converted
    assert any(r["item"] == "EMPLOYEE!GENDER" for r in fx.reasons(report, "dropdown_value_not_allowed"))
    assert fx.warnings_for(report, "'Male' (1x) not in Equus list")
    assert not any(r["item"] == "EMPLOYEE!GENDER" for r in fx.reasons(report, "dropdown_list_missing"))


def test_MD14_unknown_reference_values_resource_without_username():
    out, _, report = _run()
    assert _row(out)["ASSIGNMENT!COORDINATOR_RESOURCE_ASSIGNMENT_ID"] == "Synthcoord One"  # not derived
    [r] = fx.reasons(report, "reference_lookup_missing")
    assert r["item"] == "resource_usernames" and r["q"] == "Q2"


def test_MD15_weekly_hours_outside_range_kept_and_warned():
    out, _, report = _run(_source(host_scheduled_weekly_hours=158))
    assert _row(out)["ASSIGNMENT_CF!CF_NUMBER003"] == 158  # Q16: never capped
    assert fx.warnings_for(report, "value 158 outside expected range")
