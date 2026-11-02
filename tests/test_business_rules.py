"""UAT-02: configuration-driven business rules with approval status.

Every rule of assessment D4 is declared in lambda/src/business_rules.json and
is not approved_for_import, so the default output must not contain what the
rule would produce (AC-2); enabling one rule in a test copy of the
configuration restores its previous behaviour and nothing else (AC-3).
Fixture: tests/fixtures/topia_sample.xlsx (synthetic data only)."""

import copy
import functools
import hashlib
import io
import json
import logging
import sys
from datetime import date, datetime
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

SOURCE = ROOT / "tests" / "fixtures" / "topia_sample.xlsx"
TEMPLATE = ROOT / "lambda" / "src" / "templates" / "equus_eafh_template.xlsx"
SHIPPED_RULES = ROOT / "lambda" / "src" / "business_rules.json"
SHIPPED_REFS = ROOT / "lambda" / "src" / "reference_data"
SHEET = "Full Historical Past Assignment"

D4_RULES = {
    cfg.RULE_INTERNATIONAL_RELO: "Q40",
    cfg.RULE_CANCELLATION_PHASE: "Q34",
    cfg.RULE_DEFAULT_DATES: "Q22/Q23",
    cfg.RULE_VALUE_MAPS: "Q35",
    cfg.RULE_HOST_JOB_LEVEL: "Q32",
    cfg.RULE_COUNTRY_ISO3: "Q12/Q3",
}
ADDR_BEGIN = "ASSIGNMENT_MAILING_ADDRESS!BEGIN_DATE"
CONTACT_START = "ASSIGNMENT_EMPLOYEE_CONTACT!START_DATE"
SALARY_EFFECTIVE = "ASSIGNMENT_COMPENSATION_SALARY!EFFECTIVE_DATE"
FAMILY_EFFECTIVE = "ASSIGNMENT_FAMILY_SIZE!EFFECTIVE_DATE"
DATE_HEADERS = {ADDR_BEGIN, CONTACT_START, SALARY_EFFECTIVE, FAMILY_EFFECTIVE}


# --------------------------------------------------------------------------- helpers

def _norm(value) -> str:
    return " ".join(str(value).split()).lower() if value is not None else ""


def _col_of(ws, label: str):
    for c in range(1, ws.max_column + 1):
        if (_norm(ws.cell(4, c).value) or _norm(ws.cell(3, c).value)) == label:
            return c
    return None


@functools.lru_cache(maxsize=None)
def _source() -> bytes:
    """Fixture plus synthetic child records, a cancellation reason and a
    country name, so every D4 rule has a row it would affect."""
    wb = openpyxl.load_workbook(SOURCE)
    ws = wb[SHEET]
    for label, value in [("personal home address with effective date", "12 Big Street\nDown the Road"),
                         ("work host address with effective date", "1 Office Park"),
                         ("mobile phone number", 1234567.0),
                         ("country of birth", "India")]:
        c = _col_of(ws, label)
        assert c is not None, label
        ws.cell(5, c, value)
    cancel = _col_of(ws, cfg.CANCELLATION_REASON_LABEL)
    if cancel is not None:
        ws.cell(5, cancel, "Synthetic cancellation reason")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _rows(xlsx: bytes):
    ws = openpyxl.load_workbook(io.BytesIO(xlsx))["Sheet1"]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    return [{h: ws.cell(r, c).value for c, h in enumerate(headers, 1) if ws.cell(r, c).value is not None}
            for r in range(4, ws.max_row + 1)]


def _inserts(rows):
    return [r for r in rows if r.get("Assignment Command") == "INSERT"]


def _mapped_source_value(report, header, row):
    """Value a mapped source column provides for `header` (None if none)."""
    col = next((c for c in report.columns if c.status == "mapped" and c.target == header), None)
    if col is None:
        return None
    value = openpyxl.load_workbook(io.BytesIO(_source()))[SHEET].cell(row, col.idx).value
    if isinstance(value, str):
        value = value.strip()
    return None if value in (None, "") else value


def _shipped() -> dict:
    return json.loads(SHIPPED_RULES.read_text(encoding="utf-8"))


def _write_rules(tmp_path, data) -> Path:
    path = tmp_path / "business_rules.json"
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return path


def _enabled_copy(tmp_path, rule_id) -> Path:
    data = _shipped()
    rule = next(r for r in data["rules"] if r["id"] == rule_id)
    rule["approved_for_import"] = True
    rule["approval_reference"] = "test copy only - not an approval"
    return _write_rules(tmp_path, data)


def _table(ws, first_header: str) -> dict:
    """Rows of the table in `ws` whose header row starts with `first_header`, keyed by column A."""
    rows = list(ws.iter_rows(values_only=True))
    start = next(i for i, r in enumerate(rows) if r and r[0] == first_header)
    header = rows[start]
    out = {}
    for r in rows[start + 1:]:
        if not r or r[0] is None:
            break
        out[r[0]] = dict(zip(header, r))
    return out


def _gated_reasons(report) -> set:
    return {r["item"] for r in report.readiness.reasons if r["category"] == "gated_rule"}


def _affected(report) -> set:
    return {rule_id for rule_id, rows in report.rule_rows.items() if rows}


def _diff_headers(a_rows, b_rows) -> set:
    assert len(a_rows) == len(b_rows)
    diff = set()
    for a, b in zip(a_rows, b_rows):
        diff |= {h for h in set(a) | set(b) if a.get(h) != b.get(h)}
    return diff


@pytest.fixture(scope="module")
def gated():
    return transform(_source(), TEMPLATE.read_bytes())


def _run_enabled(tmp_path, rule_id):
    return transform(_source(), TEMPLATE.read_bytes(), business_rules=_enabled_copy(tmp_path, rule_id))


# --------------------------------------------------------------------------- shipped configuration

def test_shipped_configuration_declares_every_d4_rule_unapproved():
    rules = cfg.load_business_rules()
    assert {r.id: r.q for r in rules.rules.values()} == D4_RULES
    for rule in rules.rules.values():
        assert rule.approved_for_import is False, rule.id
        assert rule.status in cfg.RULE_STATUSES and rule.owner
    assert rules.version and rules.effective_version and rules.generated_at


def test_shipped_reference_lists_are_empty_placeholders():
    refs = cfg.load_reference_data()
    assert set(refs) == {"resource_usernames", "cost_centres", "country_state"}
    for lst in refs.values():
        assert lst.present and lst.values == {} and lst.approved_for_import is False
        assert lst.version and lst.q and lst.owner
        assert not lst.usable and lst.state == "empty"
    assert cfg.RESOURCE_MAP == {}


def test_absent_reference_list_means_exception_not_guess(tmp_path):
    refs = cfg.load_reference_data(tmp_path / "missing")
    assert all(not lst.present and not lst.usable for lst in refs.values())
    _, _, report = transform(_source(), TEMPLATE.read_bytes(), reference_data=tmp_path / "missing")
    assert report.readiness.status == NOT_READY
    assert any(r["category"] == "reference_list_unavailable" and r["item"] == "country_state"
               and "file absent" in r["detail"] for r in report.readiness.reasons)


# --------------------------------------------------------------------------- AC-1 strict loading

def _m_unknown_status(d):
    d["rules"][0]["status"] = "signed_off"
    return "$.rules[0].status"


def _m_duplicate_id(d):
    d["rules"].append(copy.deepcopy(d["rules"][0]))
    return f"$.rules[{len(d['rules']) - 1}].id"


def _m_missing_owner(d):
    del d["rules"][2]["owner"]
    return "$.rules[2].owner"


def _m_unknown_rule_key(d):
    d["rules"][1]["approved"] = True
    return "$.rules[1].approved"


def _m_unknown_top_key(d):
    d["rules_v2"] = []
    return "$.rules_v2"


def _m_bad_q(d):
    d["rules"][0]["Q"] = "40"
    return "$.rules[0].Q"


def _m_non_bool_approval(d):
    d["rules"][4]["approved_for_import"] = "false"
    return "$.rules[4].approved_for_import"


def _m_unsupported_rule(d):
    d["rules"][0]["id"] = "invented_rule"
    return "$.rules[0].id"


def _m_missing_rule(d):
    d["rules"].pop()
    return "$.rules"


def _m_unknown_parameter(d):
    d["rules"][0]["parameters"]["fallback"] = "No"
    return "$.rules[0].parameters.fallback"


def _m_missing_parameter(d):
    del d["rules"][1]["parameters"]["value"]
    return "$.rules[1].parameters.value"


def _m_bad_header(d):
    d["rules"][0]["parameters"]["target"] = "international relo"
    return "$.rules[0].parameters.target"


def _m_value_map_not_lower_case(d):
    d["rules"][3]["parameters"]["maps"]["EMPLOYEE!TYPE"]["Contractor"] = "Contractor"
    return "$.rules[3].parameters.maps.EMPLOYEE!TYPE.Contractor"


def _m_missing_version(d):
    del d["version"]
    return "$.version"


def _m_bad_generated_at(d):
    d["generated_at"] = "last week"
    return "$.generated_at"


def _m_empty_owner(d):
    d["rules"][5]["owner"] = " "
    return "$.rules[5].owner"


MUTATIONS = [_m_unknown_status, _m_duplicate_id, _m_missing_owner, _m_unknown_rule_key, _m_unknown_top_key,
             _m_bad_q, _m_non_bool_approval, _m_unsupported_rule, _m_missing_rule, _m_unknown_parameter,
             _m_missing_parameter, _m_bad_header, _m_value_map_not_lower_case, _m_missing_version,
             _m_bad_generated_at, _m_empty_owner]


@pytest.mark.parametrize("mutate", MUTATIONS, ids=[m.__name__[3:] for m in MUTATIONS])
def test_malformed_business_rules_fail_fast_naming_path(tmp_path, mutate):
    data = _shipped()
    expected_path = mutate(data)
    path = _write_rules(tmp_path, data)
    # Not even a workbook: configuration is validated before any input is read.
    with pytest.raises(ConfigurationError) as exc:
        transform(b"not a workbook", TEMPLATE.read_bytes(), business_rules=path)
    assert exc.value.path == expected_path
    assert str(path) in str(exc.value) and expected_path in str(exc.value)


def test_duplicate_json_key_invalid_json_and_missing_file_rejected(tmp_path):
    text = SHIPPED_RULES.read_text(encoding="utf-8")
    assert text.count('"version": "') == 1
    dup = tmp_path / "dup.json"
    dup.write_text(text.replace('"version": "', '"version": "9", "version": "', 1), encoding="utf-8")
    with pytest.raises(ConfigurationError, match="duplicate key 'version'"):
        transform(b"x", TEMPLATE.read_bytes(), business_rules=dup)
    broken = tmp_path / "broken.json"
    broken.write_text(text[:-10], encoding="utf-8")
    with pytest.raises(ConfigurationError, match="invalid JSON"):
        transform(b"x", TEMPLATE.read_bytes(), business_rules=broken)
    with pytest.raises(ConfigurationError, match="cannot read"):
        transform(b"x", TEMPLATE.read_bytes(), business_rules=tmp_path / "absent.json")


def _ref_copy(tmp_path) -> Path:
    d = tmp_path / "reference_data"
    d.mkdir()
    for p in SHIPPED_REFS.glob("*.json"):
        (d / p.name).write_text(p.read_text(encoding="utf-8"), encoding="utf-8")
    return d


@pytest.mark.parametrize("key,value,expected_path", [
    ("status", "approved", "$.status"),
    ("approved_for_import", 1, "$.approved_for_import"),
    ("list_id", "cost_centers", "$.list_id"),
    ("values", ["1000"], "$.values"),
    ("source_system", "Topia", "$.source_system"),
])
def test_malformed_reference_list_fails_fast(tmp_path, key, value, expected_path):
    d = _ref_copy(tmp_path)
    path = d / "cost_centres.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    data[key] = value
    path.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(ConfigurationError) as exc:
        transform(b"x", TEMPLATE.read_bytes(), reference_data=d)
    assert exc.value.path == expected_path and "cost_centres.json" in str(exc.value)


def test_unsupported_reference_list_file_rejected(tmp_path):
    d = _ref_copy(tmp_path)
    (d / "invented_list.json").write_text("{}", encoding="utf-8")
    with pytest.raises(ConfigurationError, match="unsupported reference list"):
        transform(b"x", TEMPLATE.read_bytes(), reference_data=d)


def test_valid_configuration_loads_and_version_is_reported(tmp_path):
    data = _shipped()
    data["version"] = "9.9.9-test"
    _, report_bytes, report = transform(_source(), TEMPLATE.read_bytes(), business_rules=_write_rules(tmp_path, data))
    assert report.business_rules_version == "9.9.9-test"
    wb = openpyxl.load_workbook(io.BytesIO(report_bytes))
    assert wb["Rules"]["A2"].value == "Version" and wb["Rules"]["B2"].value == "9.9.9-test"
    assert wb["Readiness"]["A2"].value == "Business rules version" and wb["Readiness"]["B2"].value == "9.9.9-test"


# --------------------------------------------------------------------------- AC-2 gated default

def test_default_output_has_no_unapproved_rule_values(gated):
    out, _, report = gated
    rows = _rows(out)
    assert any("ASSIGNMENT_MAILING_ADDRESS!LOCATION_TYPE" in r for r in rows)  # child rows exist
    for source_row, row in zip((5, 6), _inserts(rows)):
        # Q40 / Q34: nothing derived; only a mapped source column may fill these.
        for header in ("ASSIGNMENT!INTERNATIONAL_RELO", "ASSIGNMENT!PHASE"):
            assert row.get(header) == _mapped_source_value(report, header, source_row), header
        # Q22/Q23: no effective date defaulted from the start date.
        for header in (SALARY_EFFECTIVE, FAMILY_EFFECTIVE):
            if _mapped_source_value(report, header, source_row) is None:
                assert header not in row, header
    for row in rows:
        assert ADDR_BEGIN not in row and CONTACT_START not in row  # Q22/Q23
    assert report.rule_rows[cfg.RULE_INTERNATIONAL_RELO]
    assert report.rule_rows[cfg.RULE_DEFAULT_DATES]


def test_default_readiness_not_ready_and_lists_every_gated_and_pending_item(gated):
    _, report_bytes, report = gated
    readiness = report.readiness
    assert readiness.status == NOT_READY and readiness.reasons
    for reason in readiness.reasons:
        assert reason["q"] and reason["owner"] and reason["detail"], reason
    # Every unapproved rule that affected a row is a reason, with its Q number.
    assert _gated_reasons(report) == _affected(report)
    assert {cfg.RULE_INTERNATIONAL_RELO, cfg.RULE_DEFAULT_DATES, cfg.RULE_HOST_JOB_LEVEL,
            cfg.RULE_COUNTRY_ISO3} <= _gated_reasons(report)
    by_item = {r["item"]: r for r in readiness.reasons if r["category"] == "gated_rule"}
    for rule_id, q in D4_RULES.items():
        if rule_id in by_item:
            assert by_item[rule_id]["q"] == q
    # Every held and every pending (yellow) column with data is listed.
    items = {r["item"] for r in readiness.reasons}
    for c in report.columns:
        if c.has_data and ("HELD" in c.flags or "PENDING" in c.flags):
            assert f"source {c.letter} ({c.topia or c.ui})" in items, c.letter
    assert any("HELD" in c.flags and c.has_data for c in report.columns)
    # Provisional country conversion (Q12/Q3) is listed and counts against readiness.
    assert report.provisional[("EMPLOYEE!BIRTH_COUNTRY", "India", "IND")] >= 1
    assert any(r["category"] == "provisional_conversion" and r["item"] == "EMPLOYEE!BIRTH_COUNTRY"
               for r in readiness.reasons)
    # Empty reference lists: unknown value = exception, never a guess.
    assert any(r["category"] == "reference_list_unavailable" and r["item"] == "country_state"
               for r in readiness.reasons)


def test_rules_and_readiness_sheets(gated):
    _, report_bytes, report = gated
    wb = openpyxl.load_workbook(io.BytesIO(report_bytes))
    assert {"Rules", "Readiness"} <= set(wb.sheetnames)
    table = _table(wb["Rules"], "Rule ID")
    assert set(table) == set(D4_RULES)
    for rule_id, q in D4_RULES.items():
        row = table[rule_id]
        assert row["Q"] == q and row["Approved For Import"] == "no"
        assert row["Rows Affected"] == len(report.rule_rows.get(rule_id, ()))
        assert row["Status"] in cfg.RULE_STATUSES and row["Owner"]
    assert table[cfg.RULE_INTERNATIONAL_RELO]["Applied?"] == "no (gated)"
    assert table[cfg.RULE_COUNTRY_ISO3]["Applied?"] == "provisional (diagnostic only)"
    provisional = _table(wb["Rules"], "Provisional Conversion")
    assert provisional["provisional conversion"]["Converted"] in cfg.ISO3_CODES
    refs = _table(wb["Rules"], "Reference List")
    assert set(refs) == set(cfg.REFERENCE_LISTS) and all(r["State"] == "empty" for r in refs.values())
    ws = wb["Readiness"]
    assert ws["A1"].value == "Status" and ws["B1"].value == NOT_READY
    reasons = list(ws.iter_rows(values_only=True))
    header_at = next(i for i, r in enumerate(reasons) if r and r[0] == "Category")
    assert len([r for r in reasons[header_at + 1:] if r and r[0]]) == len(report.readiness.reasons)
    assert "SIGNED OFF" not in json.dumps([list(r) for r in reasons], default=str).upper()


def test_brown_fill_is_no_meaning_not_unknown(gated):
    # Q8 (Jake): brown has no business meaning.
    _, _, report = gated
    brown = [c for c in report.columns if c.colour == cfg.NO_MEANING_COLOUR]
    assert brown
    assert all("UNKNOWN_COLOUR" not in c.flags for c in brown)
    assert not any(r["category"] == "unknown_colour" and "Spouse" in r["item"] for r in report.readiness.reasons)


# --------------------------------------------------------------------------- AC-3 one rule enabled

def _assert_only_rule_enabled(report, rule_id):
    assert [r.id for r in report.rules.rules.values() if r.approved_for_import] == [rule_id]
    assert rule_id not in _gated_reasons(report)
    assert _gated_reasons(report) == _affected(report) - {rule_id}
    assert report.readiness.status == NOT_READY  # other items stay open


def test_enable_international_relo_restores_derivation(tmp_path, gated):
    out, report_bytes, report = _run_enabled(tmp_path, cfg.RULE_INTERNATIONAL_RELO)
    raj = _inserts(_rows(out))[0]
    assert (raj["ASSIGNMENT!FROM_COUNTRY"], raj["ASSIGNMENT!TO_COUNTRY"]) == ("EGY", "IND")
    assert raj["ASSIGNMENT!INTERNATIONAL_RELO"] == "Yes"  # behaviour before UAT-02 (Q40)
    assert _diff_headers(_rows(gated[0]), _rows(out)) == {"ASSIGNMENT!INTERNATIONAL_RELO"}
    _assert_only_rule_enabled(report, cfg.RULE_INTERNATIONAL_RELO)
    table = _table(openpyxl.load_workbook(io.BytesIO(report_bytes))["Rules"], "Rule ID")
    assert table[cfg.RULE_INTERNATIONAL_RELO]["Applied?"] == "yes"
    assert table[cfg.RULE_INTERNATIONAL_RELO]["Approved For Import"] == "yes"


def test_enable_cancellation_phase_restores_cancelled(tmp_path, gated):
    out, _, report = _run_enabled(tmp_path, cfg.RULE_CANCELLATION_PHASE)
    raj = _inserts(_rows(out))[0]
    assert raj["ASSIGNMENT!PHASE"] == "Cancelled"  # behaviour before UAT-02 (Q34)
    assert _diff_headers(_rows(gated[0]), _rows(out)) <= {"ASSIGNMENT!PHASE"}
    cancel = next((c for c in report.columns if c.label == cfg.CANCELLATION_REASON_LABEL), None)
    if cancel is not None and cancel.status != "excluded":
        assert cancel.status == "rule" and "HELD" not in cancel.flags
        assert report.rule_rows[cfg.RULE_CANCELLATION_PHASE]
        gated_col = next(c for c in gated[2].columns if c.idx == cancel.idx)
        assert gated_col.status == "held" and "HELD" in gated_col.flags
    _assert_only_rule_enabled(report, cfg.RULE_CANCELLATION_PHASE)


def test_enable_default_dates_restores_start_date_defaults(tmp_path, gated):
    out, _, report = _run_enabled(tmp_path, cfg.RULE_DEFAULT_DATES)
    rows = _rows(out)
    assert _diff_headers(_rows(gated[0]), rows) <= DATE_HEADERS
    raj_rows = [r for r in rows if r.get("Assignment Identifier") == _inserts(rows)[0]["Assignment Identifier"]]
    start = raj_rows[0]["ASSIGNMENT!START_DATE"]
    tpl = openpyxl.load_workbook(TEMPLATE)["Sheet1"]
    tpl_headers = {tpl.cell(3, c).value for c in range(1, tpl.max_column + 1)}
    for row in raj_rows:
        if "ASSIGNMENT_MAILING_ADDRESS!LOCATION_TYPE" in row and ADDR_BEGIN in tpl_headers:
            assert row[ADDR_BEGIN] == start  # behaviour before UAT-02 (Q22/Q23)
        if "ASSIGNMENT_EMPLOYEE_CONTACT!NUMBER_ADDRESS" in row and CONTACT_START in tpl_headers:
            assert row[CONTACT_START] == start
    assert report.rule_rows[cfg.RULE_DEFAULT_DATES]
    _assert_only_rule_enabled(report, cfg.RULE_DEFAULT_DATES)


def test_enable_value_maps_restores_translation(tmp_path, gated):
    out, _, report = _run_enabled(tmp_path, cfg.RULE_VALUE_MAPS)
    raj = _inserts(_rows(out))[0]
    assert raj["EMPLOYEE!EMPLOYMENT_STATUS"] == "Active"  # behaviour before UAT-02 (Q35)
    assert raj["EMPLOYEE!TYPE"] == "Employee"
    assert _diff_headers(_rows(gated[0]), _rows(out)) <= {"EMPLOYEE!TYPE", "EMPLOYEE!EMPLOYMENT_STATUS"}
    _assert_only_rule_enabled(report, cfg.RULE_VALUE_MAPS)


def test_enable_host_job_level_restores_salary_grade(tmp_path, gated):
    out, _, report = _run_enabled(tmp_path, cfg.RULE_HOST_JOB_LEVEL)
    nic = _inserts(_rows(out))[1]
    assert nic["ASSIGNMENT!TO_SALARY_GRADE"] == 8  # behaviour before UAT-02 (Q32)
    assert nic["ASSIGNMENT!TO_TITLE"] == "Solution Architect"
    col = next(c for c in report.columns if c.label == "host job level")
    assert col.status == "mapped" and col.target == "ASSIGNMENT!TO_SALARY_GRADE"
    assert _diff_headers(_rows(gated[0]), _rows(out)) <= {"ASSIGNMENT!TO_SALARY_GRADE"}
    _assert_only_rule_enabled(report, cfg.RULE_HOST_JOB_LEVEL)


def test_enable_country_iso3_removes_provisional_marking(tmp_path, gated):
    out, report_bytes, report = _run_enabled(tmp_path, cfg.RULE_COUNTRY_ISO3)
    assert _diff_headers(_rows(gated[0]), _rows(out)) == set()  # same conversion, now approved
    assert _inserts(_rows(out))[0]["EMPLOYEE!BIRTH_COUNTRY"] == "IND"
    assert not report.provisional
    assert not any(r["category"] == "provisional_conversion" for r in report.readiness.reasons)
    _assert_only_rule_enabled(report, cfg.RULE_COUNTRY_ISO3)
    assert not _table(openpyxl.load_workbook(io.BytesIO(report_bytes))["Rules"], "Provisional Conversion")


def _unit_report(rules_path=None) -> tmod.Report:
    report = tmod.Report(rules=cfg.load_business_rules(rules_path), reference_data=cfg.load_reference_data())
    report.current_row = 7
    return report


def test_rule_functions_gated_and_enabled(tmp_path):
    start = date(2020, 1, 1)
    main = {"ASSIGNMENT!START_DATE": start, "ASSIGNMENT!FROM_COUNTRY": "AAA", "ASSIGNMENT!TO_COUNTRY": "BBB",
            "ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_AMOUNT": 100,
            "ASSIGNMENT_FAMILY_SIZE!HOME_FAMILY_SIZE": 3}
    report = _unit_report()
    gated_main = dict(main)
    tmod.apply_assignment_rules(gated_main, "a reason", report)
    addr, contact = {"ASSIGNMENT_MAILING_ADDRESS!LINE_1": "x"}, {"ASSIGNMENT_EMPLOYEE_CONTACT!NUMBER_ADDRESS": "1"}
    tmod.default_child_dates(addr, start, report)
    tmod.default_child_dates(contact, start, report)
    assert gated_main == main and ADDR_BEGIN not in addr and CONTACT_START not in contact
    for rule_id in (cfg.RULE_CANCELLATION_PHASE, cfg.RULE_INTERNATIONAL_RELO, cfg.RULE_DEFAULT_DATES):
        assert report.rule_rows[rule_id] == {7}

    expected = {
        cfg.RULE_CANCELLATION_PHASE: {"ASSIGNMENT!PHASE": "Cancelled"},
        cfg.RULE_INTERNATIONAL_RELO: {"ASSIGNMENT!INTERNATIONAL_RELO": "Yes"},
        cfg.RULE_DEFAULT_DATES: {SALARY_EFFECTIVE: start, FAMILY_EFFECTIVE: start},
    }
    for rule_id, added in expected.items():
        sub = tmp_path / rule_id
        sub.mkdir()
        report = _unit_report(_enabled_copy(sub, rule_id))
        enabled_main = dict(main)
        tmod.apply_assignment_rules(enabled_main, "a reason", report)
        assert enabled_main == {**main, **added}, rule_id
        addr, contact = {"ASSIGNMENT_MAILING_ADDRESS!LINE_1": "x"}, {"ASSIGNMENT_EMPLOYEE_CONTACT!NUMBER_ADDRESS": "1"}
        tmod.default_child_dates(addr, start, report)
        tmod.default_child_dates(contact, start, report)
        if rule_id == cfg.RULE_DEFAULT_DATES:
            assert addr[ADDR_BEGIN] == start and contact[CONTACT_START] == start
        else:
            assert ADDR_BEGIN not in addr and CONTACT_START not in contact
    same = {"ASSIGNMENT!FROM_COUNTRY": "AAA", "ASSIGNMENT!TO_COUNTRY": "AAA"}
    sub = tmp_path / "same"
    sub.mkdir()
    tmod.apply_assignment_rules(same, None, _unit_report(_enabled_copy(sub, cfg.RULE_INTERNATIONAL_RELO)))
    assert same["ASSIGNMENT!INTERNATIONAL_RELO"] == "No"


def test_helper_reads_flag_from_configuration(tmp_path):
    assert cfg.rule_enabled(cfg.RULE_INTERNATIONAL_RELO) is False
    enabled = cfg.load_business_rules(_enabled_copy(tmp_path, cfg.RULE_INTERNATIONAL_RELO))
    assert cfg.rule_enabled(cfg.RULE_INTERNATIONAL_RELO, enabled) is True
    with pytest.raises(ConfigurationError):
        cfg.rule_enabled("invented_rule")


# --------------------------------------------------------------------------- readiness

def test_readiness_ready_only_when_nothing_is_open():
    report = _unit_report()
    report.reconciliation_ok = True
    assert tmod.compute_readiness(report).status == READY
    report.rule_hit(cfg.RULE_INTERNATIONAL_RELO, 5)
    readiness = tmod.compute_readiness(report)
    assert readiness.status == NOT_READY
    assert readiness.reasons == [{"category": "gated_rule", "item": cfg.RULE_INTERNATIONAL_RELO, "q": "Q40",
                                  "owner": "Schneider", "detail": readiness.reasons[0]["detail"]}]
    report = _unit_report()
    assert tmod.compute_readiness(report).status == NOT_READY  # reconciliation not proven


def test_readiness_independent_of_strict_pending(gated):
    # STRICT_PENDING keeps refusing output; without it readiness still says NOT_READY.
    from transform import PendingFieldsError
    with pytest.raises(PendingFieldsError):
        transform(_source(), TEMPLATE.read_bytes(), strict_pending=True)
    assert gated[2].readiness.status == NOT_READY


# --------------------------------------------------------------------------- AC-4 handler

class S3Error(Exception):
    def __init__(self, code):
        super().__init__(code)
        self.response = {"Error": {"Code": code}}


class FakeS3:
    def __init__(self, objects):
        self.objects = objects
        self.puts = {}
        self.meta = {}

    def get_object(self, Bucket, Key):
        body = self.objects[(Bucket, Key)]
        return {"Body": io.BytesIO(body), "ETag": '"%s"' % hashlib.md5(body).hexdigest()}

    def head_object(self, Bucket, Key):
        if (Bucket, Key) not in self.puts:
            raise S3Error("404")
        return {}

    def put_object(self, Bucket, Key, Body, IfNoneMatch=None, Metadata=None, **kwargs):
        if IfNoneMatch == "*" and (Bucket, Key) in self.puts:
            raise S3Error("PreconditionFailed")
        self.puts[(Bucket, Key)] = Body
        self.meta[(Bucket, Key)] = Metadata


EVENT = {"Records": [{"eventName": "ObjectCreated:Put",
                      "s3": {"bucket": {"name": "src"}, "object": {"key": "topia/incoming/SYNTHETIC.xlsx", "size": 1}}}]}


@pytest.fixture
def fake_s3(monkeypatch):
    s3 = FakeS3({("src", "topia/incoming/SYNTHETIC.xlsx"): SOURCE.read_bytes()})
    monkeypatch.setattr(handler, "_s3", s3)
    monkeypatch.setenv("DEST_BUCKET", "dst")
    monkeypatch.setenv("DEST_PREFIX", "equus/")
    monkeypatch.setenv("SOURCE_PREFIX", "topia/incoming/")
    monkeypatch.delenv("TEMPLATE_S3_URI", raising=False)
    monkeypatch.delenv("STRICT_PENDING", raising=False)
    return s3


def _outputs(s3):
    return [k for (_, k) in s3.puts if not k.startswith("equus/reports/")]


def test_handler_names_draft_and_records_readiness(fake_s3, caplog):
    caplog.set_level(logging.INFO)
    summary = handler.lambda_handler(EVENT, None)["results"][0]
    [out_key] = _outputs(fake_s3)
    assert out_key.startswith("equus/SYNTHETIC_EQUUS_DRAFT_")
    meta = fake_s3.meta[("dst", out_key)]
    assert meta["readiness"] == NOT_READY
    assert int(meta["readiness-reasons"]) > 0
    assert meta["business-rules-version"] == _shipped()["version"]
    assert summary["readiness"] == NOT_READY and summary["readiness_reasons"] > 0
    assert '"readiness": "NOT_READY"' in caplog.text
    report_key = next(k for (_, k) in fake_s3.puts if k.startswith("equus/reports/"))
    assert fake_s3.meta[("dst", report_key)]["readiness"] == NOT_READY
    wb = openpyxl.load_workbook(io.BytesIO(fake_s3.puts[("dst", report_key)]))
    assert {"Rules", "Readiness"} <= set(wb.sheetnames) and wb["Readiness"]["B1"].value == NOT_READY
    # A duplicate delivery is still skipped although the name is the draft one.
    again = handler.lambda_handler(EVENT, None)["results"][0]
    assert again["skipped"] == "already processed" and len(fake_s3.puts) == 2


def test_handler_uses_import_name_only_when_ready(fake_s3, monkeypatch):
    real = handler.transform

    def ready_transform(*args, **kwargs):
        out, report_bytes, report = real(*args, **kwargs)
        report.readiness = tmod.Readiness(status=READY, reasons=[])
        return out, report_bytes, report

    monkeypatch.setattr(handler, "transform", ready_transform)
    handler.lambda_handler(EVENT, None)
    [out_key] = _outputs(fake_s3)
    assert out_key.startswith("equus/SYNTHETIC_EQUUS_") and "_DRAFT_" not in out_key
    assert fake_s3.meta[("dst", out_key)]["readiness"] == READY


def test_output_key_changes_when_rule_approval_changes(monkeypatch, tmp_path):
    base = ROOT / "lambda" / "src"
    before = handler._code_version(b"tpl")
    copy_dir = tmp_path / "src"
    (copy_dir / "reference_data").mkdir(parents=True)
    for name in ("transform.py", "mapping_config.py"):
        (copy_dir / name).write_bytes((base / name).read_bytes())
    for p in (base / "reference_data").glob("*.json"):
        (copy_dir / "reference_data" / p.name).write_bytes(p.read_bytes())
    data = _shipped()
    data["rules"][0]["approved_for_import"] = True
    (copy_dir / "business_rules.json").write_text(json.dumps(data), encoding="utf-8")
    monkeypatch.setattr(handler, "__file__", str(copy_dir / "handler.py"))
    assert handler._code_version(b"tpl") != before


def test_no_real_or_dated_values_leak_into_shipped_config():
    # Shipped reference lists carry metadata only (no personal data, no guesses).
    for p in SHIPPED_REFS.glob("*.json"):
        assert json.loads(p.read_text(encoding="utf-8"))["values"] == {}
    assert isinstance(datetime.fromisoformat(_shipped()["generated_at"]), datetime)
