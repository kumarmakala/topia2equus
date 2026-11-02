"""Transform a Topia Full Historical Past Assignment export into the Equus
EAFH import layout. Pure functions: bytes in, bytes out, no AWS calls.

Business rules are read from the versioned configuration
(mapping_config.load_business_rules / load_reference_data). A rule changes
output only while its entry is approved_for_import; otherwise the affected
rows are counted and reported in the Rules and Readiness sheets.

Identity, duplicate, header and row-structure controls (UAT-03: screen_rows,
check_template_headers, validate_row_structure) never change a value: a row
is accepted or HELD (excluded from the import workbook, listed in the
Exceptions and Held Rows sheets), and a FILE-level finding withholds the whole
import workbook. There is no fallback from a missing Assignment Identifier to
the Employee Number (BR-01)."""

from __future__ import annotations

import hashlib
import io
import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from fnmatch import fnmatchcase
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

import mapping_config as cfg
from mapping_config import ConfigurationError  # noqa: F401  (re-exported for callers)

KEY_RE = re.compile(r"^[A-Z0-9_]+![A-Z0-9_]+$")
DATE_FORMAT = "MM/DD/YYYY"
_DATE_PATTERNS = ("%m/%d/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y", "%Y-%m-%d %H:%M:%S", "%m/%d/%y")

ADDR = "ASSIGNMENT_MAILING_ADDRESS!"
CONTACT = "ASSIGNMENT_EMPLOYEE_CONTACT!"

# Import-readiness statuses. Without Equus access nothing better than
# READY_FOR_UAT_EXECUTION can be computed here (never "signed off").
READY = "READY_FOR_UAT_EXECUTION"
NOT_READY = "NOT_READY"


def norm_label(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value)).strip().lower() if value is not None else ""


def norm_key(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    key = value.strip().upper()
    if not KEY_RE.match(key):
        return None
    for pattern, repl in cfg.KEY_ALIASES:
        key = re.sub(pattern, repl, key)
    return key


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == "" or value.strip().lower() in cfg.BLANK_TOKENS
    return False


class PendingFieldsError(ValueError):
    """Raised in strict mode when yellow (pending) columns hold data."""


class TemplateError(ValueError):
    """The Equus template's sheet or header row does not match what the
    engine writes. Q5 (Jake): header issues fail the whole import, so the run
    stops instead of producing a file with a missing or doubled column."""


class SourceError(ValueError):
    """The source workbook does not have the expected structure."""



@dataclass
class SourceColumn:
    idx: int
    key: Any
    ui: Any
    topia: Any
    status: str = "unmapped"  # mapped | address | contact | dependent | rule | held | ignored | excluded | unmapped | duplicate
    target: str | None = None
    note: str = ""
    colour: str = "none"  # red | yellow | cyan | peach | none | none (Q8: no meaning) | unknown:<ARGB>
    flags: list[str] = field(default_factory=list)  # PENDING | MISSING_TARGET | UNKNOWN_COLOUR | HELD
    has_data: bool = False
    rule: str | None = None  # business rule id that claims or holds this column

    @property
    def letter(self) -> str:
        return get_column_letter(self.idx)

    @property
    def label(self) -> str:
        return norm_label(self.topia) or norm_label(self.ui)


@dataclass
class TargetLayout:
    headers: dict[str, int]  # header -> column index
    rules: dict[str, str]  # header -> lower-cased row-1 instruction

    def is_date(self, header: str) -> bool:
        return "date format" in self.rules.get(header, "") or header.endswith("_DATE")

    def is_country(self, header: str) -> bool:
        return header in cfg.COUNTRY_HEADERS

    def is_resource(self, header: str) -> bool:
        return "RESOURCE" in header

    def is_dropdown(self, header: str) -> bool:
        return "dropdown" in self.rules.get(header, "") or "yesno" in self.rules.get(header, "")


@dataclass
class Readiness:
    status: str
    reasons: list[dict] = field(default_factory=list)  # category, item, q, owner, detail


@dataclass
class Report:
    columns: list[SourceColumn] = field(default_factory=list)
    warnings: list[tuple[Any, str, str]] = field(default_factory=list)
    dropdown_values: Counter = field(default_factory=Counter)
    unmapped_resources: Counter = field(default_factory=Counter)
    allowed_values: dict[str, set[str]] = field(default_factory=dict)  # header -> lower-cased values
    coverage: Counter = field(default_factory=Counter)
    assignments: int = 0
    rows_written: int = 0
    data_rows: int = 0
    source_sha256: str = ""
    reconciliation: list[dict] = field(default_factory=list)  # one row per source column (see reconcile)
    reconciliation_ok: bool = False
    reconciliation_records: list[dict] = field(default_factory=list)  # one row per source data row
    reconciliation_summary: dict = field(default_factory=dict)
    keyed_children: dict = field(default_factory=dict)  # source row -> [{"sheet", "row", "record"}] attached
    duplicate_dependent_groups: dict = field(default_factory=dict)  # source row -> {sheet group index} held (UAT-06)
    # Configuration in force and what the rules did.
    rules: Any = None  # mapping_config.BusinessRules
    reference_data: dict = field(default_factory=dict)  # list id -> mapping_config.ReferenceList
    rule_rows: dict = field(default_factory=lambda: defaultdict(set))  # rule id -> source rows affected
    provisional: Counter = field(default_factory=Counter)  # (header, source text, converted) -> n
    reference_values: Counter = field(default_factory=Counter)  # (list id, header, value) -> n
    state_values: Counter = field(default_factory=Counter)  # (state header, country code, state) -> n (Q12)
    config_issues: list[tuple[str, str]] = field(default_factory=list)  # (config path, message)
    current_row: Any = None
    readiness: Readiness | None = None
    # Spouse / dependent records (Q10, Q11, Q18, Q19, Q20).
    data_quality: list[dict] = field(default_factory=list)  # where, record, fields, missing
    family_values: Counter = field(default_factory=Counter)  # (header, value) -> n
    dependent_exceptions: list[dict] = field(default_factory=list)  # code, sheet, rows, groups, q, owner, detail
    dependents_attached: int = 0  # child records attached from the keyed dependent source
    dependent_overlaps: int = 0  # assignments with dependents from the main sheet and the keyed source
    # Identity, duplicate, header and row-level controls (UAT-03).
    exceptions: list[dict] = field(default_factory=list)  # see EXCEPTIONS_HEADER
    held_rows: dict = field(default_factory=dict)  # source row -> [codes]; excluded from the import workbook
    held_row_values: dict = field(default_factory=dict)  # source row -> raw cell values (columns A..)
    source_headers: list = field(default_factory=list)  # [(column letter, label)] for the Held Rows sheet
    row_identity: dict = field(default_factory=dict)  # source row -> (identifier text, employee number)
    accepted: list[dict] = field(default_factory=list)  # {"row", "identifier", "employee"} written to the import
    previously_accepted: set = field(default_factory=set)  # identifiers accepted by an earlier run
    template_check: dict = field(default_factory=dict)

    @property
    def business_rules_version(self) -> str:
        return self.rules.version if self.rules is not None else ""

    @property
    def file_failed(self) -> bool:
        """True when any FILE-level finding withholds the whole import workbook."""
        return any(e["level"] == cfg.LEVEL_FILE for e in self.exceptions)

    def add_exception(self, code: str, message: str, source_row: Any = None, identifier: Any = None,
                      employee: Any = None) -> dict:
        """Record a finding with its configured level, Q and owner. Row-level
        codes also hold `source_row` (kept out of the import workbook)."""
        spec = self.rules.identity.code(code)
        entry = {"source_row": source_row, "identifier": identifier, "employee": employee, "code": code,
                 "level": spec.level, "message": message, "owner": spec.owner, "q": spec.q}
        self.exceptions.append(entry)
        if code in cfg.ROW_CODES and isinstance(source_row, int):
            codes = self.held_rows.setdefault(source_row, [])
            if code not in codes:
                codes.append(code)
        return entry

    def warn(self, row: Any, fld: str, msg: str) -> None:
        self.warnings.append((row, fld, msg))

    def rule_hit(self, rule_id: str, row: Any = None) -> None:
        """Record that `rule_id` applies (or would apply) to a source row."""
        self.rule_rows[rule_id].add(self.current_row if row is None else row)

    def rule_enabled(self, rule_id: str) -> bool:
        return self.rules.rule_enabled(rule_id)

    def lookup_values(self, header: str) -> dict:
        """Approved lookup values for `header` (empty unless a list is usable)."""
        out: dict = {}
        for lst in self.reference_data.values():
            if lst.kind == "lookup" and lst.usable and lst.applies(header):
                out.update(lst.values)
        return out

    def note_reference_value(self, header: str, value: Any) -> None:
        for lst in self.reference_data.values():
            if lst.kind == "allowed_values" and lst.applies(header):
                self.reference_values[(lst.list_id, header, str(value))] += 1

    def allowed(self, header: str, value: Any) -> str:
        if header not in self.allowed_values:
            return "no list"
        return "yes" if str(value).lower() in self.allowed_values[header] else "NO"

    def family_value_status(self, header: str, value: Any) -> str:
        """Q14: a family dropdown value is checked against the Equus list when
        one is bundled; without a list it is 'not in list' (never converted)."""
        if header not in self.allowed_values:
            return f"not in list (no Equus list; {cfg.FAMILY_LIST_OWNER[0]} pending)"
        return "in list" if self.allowed(header, value) == "yes" else "not in list"

    def state_in_list(self, country: Any, state: str) -> bool | None:
        """Q12: True/False against the approved country_state list; None
        while the list is not usable (nothing can be validated)."""
        lst = self.reference_data.get(cfg.STATE_LIST_ID)
        if lst is None or not lst.usable:
            return None
        return state in lst.values.get(str(country), ()) if country else False

    def hold_dependents(self, code: str, sheet: str, rows: list, groups: int, detail: str = "") -> None:
        q, owner, meaning = cfg.DEPENDENT_HOLDS[code]
        entry = {"code": code, "sheet": sheet, "rows": len(rows), "groups": groups, "q": q, "owner": owner,
                 "detail": meaning + (f"; {detail}" if detail else ""),
                 "source_rows": ", ".join(str(r) for r in rows)}
        self.dependent_exceptions.append(entry)
        self.warn(f"sheet {sheet!r}", cfg.FAMILY_PREFIXES[1],
                  f"{code}: {len(rows)} dependent row(s), {groups} dependent record(s) held ({q}, owner {owner}); "
                  f"{entry['detail']}")

    def summarise_values(self) -> None:
        """One warning per distinct unmapped resource / disallowed dropdown value."""
        for (header, value), n in sorted(self.family_values.items(), key=str):
            status = self.family_value_status(header, value)
            if status != "in list":
                self.warn("all rows", header, f"{value!r} ({n}x) {status}; sent unchanged, never converted "
                                              f"({cfg.FAMILY_LIST_OWNER[0]}, owner {cfg.FAMILY_LIST_OWNER[1]})")
        for (header, value), n in sorted(self.unmapped_resources.items(), key=str):
            self.warn("all rows", header, f"{value!r} ({n}x) not in RESOURCE_MAP "
                                          f"(reference_data/{cfg.RESOURCE_LIST_ID}.json); Equus expects a username")
        for (header, value), n in sorted(self.dropdown_values.items(), key=str):
            if self.allowed(header, value) == "NO":
                self.warn("all rows", header, f"{value!r} ({n}x) not in Equus list")
        for (list_id, header, value), n in sorted(self.reference_values.items(), key=str):
            lst = self.reference_data.get(list_id)
            if lst is not None and lst.usable and value not in lst.values:
                self.warn("all rows", header, f"{value!r} ({n}x) not in approved reference list {list_id} ({lst.q})")
        for (header, country, state), n in sorted(self.state_values.items(), key=str):
            if self.state_in_list(country, state) is False:
                self.warn("all rows", header, f"state {state!r} ({n}x) not in approved reference list "
                                              f"{cfg.STATE_LIST_ID} for country {country!r} (Q12); sent unchanged")


# --------------------------------------------------------------------------- layout

def read_target_layout(ws) -> TargetLayout:
    headers, rules = {}, {}
    for c in range(1, ws.max_column + 1):
        raw = ws.cell(cfg.TARGET_HEADER_ROW, c).value
        if raw is None:
            continue
        header = norm_key(raw) or str(raw).strip()
        if header in headers:
            raise TemplateError(f"Duplicate Equus header {header!r} in template")
        headers[header] = c
        rules[header] = norm_label(ws.cell(cfg.TARGET_RULE_ROW, c).value)
    return TargetLayout(headers, rules)


def check_template_structure(tpl_wb) -> TargetLayout:
    """Target sheet present and header row unique; otherwise TemplateError
    (whole-file failure, Q5). A renamed or missing header is not raised here:
    check_template_headers (UAT-03) records it as a FILE-level HEADER_MISMATCH
    against the approved template, which withholds the import workbook."""
    if cfg.TARGET_SHEET not in tpl_wb.sheetnames:
        raise TemplateError(f"Equus template has no sheet {cfg.TARGET_SHEET!r}; got {tpl_wb.sheetnames}")
    return read_target_layout(tpl_wb[cfg.TARGET_SHEET])


def column_colour(ws, c: int) -> str:
    """Business colour of a source column, from the fills of its header rows."""
    found = []
    for r in cfg.SOURCE_COLOUR_ROWS:
        fill = ws.cell(r, c).fill
        if fill is None or fill.fill_type is None or fill.fgColor.type != "rgb":
            continue
        found.append(fill.fgColor.rgb)
    names = {cfg.FILL_COLOURS[f] for f in found if f in cfg.FILL_COLOURS}
    for name in cfg.COLOUR_PRECEDENCE[:3]:  # colours with a business meaning
        if name in names:
            return name
    unknown = [f for f in found if f not in cfg.FILL_COLOURS]
    if unknown:
        return f"unknown:{unknown[0]}"
    for name in cfg.COLOUR_PRECEDENCE[3:]:
        if name in names:
            return name
    return "none"


def check_template(layout: TargetLayout, report: Report) -> None:
    """Warn where the template's row-1 country rules disagree with COUNTRY_HEADERS."""
    for header in layout.headers:
        has_rule = "country code" in layout.rules.get(header, "")
        if header in cfg.COUNTRY_HEADERS and not has_rule:
            report.warn("template", header, "treated as country code, but row 1 has no 'Match Country Code' rule")
        elif has_rule and header not in cfg.COUNTRY_HEADERS:
            report.warn("template", header, "row 1 says 'Match Country Code' but header is not in COUNTRY_HEADERS; "
                                            "rules may be shifted against row 3")


def check_configuration(layout: TargetLayout, report: Report) -> None:
    """Every Equus header the configuration refers to must exist in the template."""
    rules = report.rules
    for path, header in rules.header_refs():
        if header not in layout.headers:
            msg = f"{rules.file_name} {path}: {header} not in Equus template"
            report.config_issues.append((path, msg))
            report.warn("configuration", header, msg)
    checked = [*cfg.STATE_COUNTRY_HEADERS.items(), (cfg.ASSIGNMENT_START_HEADER, cfg.ASSIGNMENT_END_HEADER)]
    for header in sorted({h for pair in checked for h in pair}):
        if header not in layout.headers:
            msg = f"mapping_config (UAT-06 checks): {header} not in Equus template"
            report.config_issues.append(("mapping_config", msg))
            report.warn("configuration", header, msg)
    for lst in report.reference_data.values():
        for pattern in lst.applies_to:
            if not any(fnmatchcase(h, pattern) for h in layout.headers):
                msg = f"reference list {lst.list_id}: applies_to {pattern!r} matches no Equus template header"
                report.config_issues.append((lst.list_id, msg))
                report.warn("configuration", pattern, msg)


def build_column_plan(ws, layout: TargetLayout, rules: cfg.BusinessRules | None = None) -> list[SourceColumn]:
    rules = rules if rules is not None else cfg.BUSINESS_RULES
    cols = []
    for c in range(cfg.SOURCE_FIRST_COL, ws.max_column + 1):
        col = SourceColumn(
            idx=c,
            key=ws.cell(cfg.SOURCE_KEY_ROW, c).value,
            ui=ws.cell(cfg.SOURCE_UI_ROW, c).value,
            topia=ws.cell(cfg.SOURCE_TOPIA_ROW, c).value,
            colour=column_colour(ws, c),
        )
        if col.key is None and col.ui is None and col.topia is None:
            continue
        col.has_data = any(not is_blank(ws.cell(r, c).value)
                           for r in range(cfg.SOURCE_DATA_START_ROW, ws.max_row + 1))
        cols.append(col)

    # Red = not required: excluded before any other rule can claim the column.
    for col in cols:
        if col.colour == "red":
            col.status, col.note = "excluded", "red: not required"

    claimed: dict[str, SourceColumn] = {}

    def claim(col: SourceColumn, target: str, note: str) -> None:
        if target not in layout.headers:
            col.status, col.note = "unmapped", f"override target {target} not in Equus template"
        elif target in claimed:
            col.status, col.target = "duplicate", target
            col.note = f"{target} already filled from source column {claimed[target].letter}"
        else:
            col.status, col.target, col.note = "mapped", target, note
            claimed[target] = col

    def hold(col: SourceColumn, rule: cfg.BusinessRule) -> None:
        col.status, col.rule = "held", rule.id
        col.note = (f"held: rule {rule.id} ({rule.q}, {rule.status}, owner {rule.owner}) "
                    "not approved_for_import; not migrated")
        col.flags.append("HELD")

    cancel = rules.rule(cfg.RULE_CANCELLATION_PHASE)
    overrides = {rules.params(rule_id)["source_label"]: rules.rule(rule_id) for rule_id in cfg.COLUMN_OVERRIDE_RULES}

    # Pass 1: special roles, rule columns and explicit overrides (these win over row 2).
    for col in cols:
        label = col.label
        ui_label = norm_label(col.ui)
        if col.status == "excluded":
            continue
        if label in cfg.ADDRESS_COLUMNS:
            col.status, col.note = "address", f"LOCATION_TYPE = {cfg.ADDRESS_COLUMNS[label]}"
        elif label in cfg.CONTACT_COLUMNS:
            col.status, col.note = "contact", "LOCATION_TYPE/CONTACT_TYPE = %s / %s" % cfg.CONTACT_COLUMNS[label]
        elif label.startswith(cfg.IGNORED_LABEL_PREFIXES):
            col.status, col.note = "ignored", "describes a constant location type"
        elif label == cancel.parameters["source_label"]:
            if cancel.approved_for_import:
                col.status, col.rule = "rule", cancel.id
                col.note = (f"non-empty -> {cancel.parameters['target']} = {cancel.parameters['value']} "
                            f"(rule {cancel.id}, {cancel.q})")
            else:
                hold(col, cancel)
        elif label in overrides or ui_label in overrides:
            rule = overrides[label] if label in overrides else overrides[ui_label]
            if rule.approved_for_import:
                claim(col, rule.parameters["target"], f"override by rule {rule.id} ({rule.q})")
                col.rule = rule.id
            else:
                hold(col, rule)
        elif label in cfg.FIELD_OVERRIDES or ui_label in cfg.FIELD_OVERRIDES:
            target = cfg.FIELD_OVERRIDES.get(label, cfg.FIELD_OVERRIDES.get(ui_label))
            if target is None:
                col.status, col.note = "ignored", "excluded in mapping_config.FIELD_OVERRIDES"
            else:
                claim(col, target, "override (row 2 not usable)")

    # Pass 2: row-2 technical field vs Equus row-3 header.
    for col in cols:
        if col.status != "unmapped":
            continue
        key = norm_key(col.key)
        if key is None:
            col.note = "row 2 has no technical field" if col.key is None else "row 2 is a note, not a field"
        elif key.startswith("ASSIGNMENT_DEPENDENT"):
            col.status, col.target = "dependent", key
            if key not in layout.headers:
                col.status, col.note = "unmapped", f"{key} not in Equus template"
        elif key in layout.headers:
            claim(col, key, "row 2 = row 3" if key == str(col.key).strip().upper() else f"alias of {col.key}")
        else:
            col.note = f"{key} not in Equus template"

    # Governance flags from the colour coding.
    for col in cols:
        if col.colour == "yellow" and col.status != "excluded":
            col.flags.append("PENDING")
        if col.colour == "cyan" and col.status == "unmapped":
            col.flags.append("MISSING_TARGET")
        if col.colour.startswith("unknown:"):
            col.flags.append("UNKNOWN_COLOUR")
    return cols


# --------------------------------------------------------------------------- values

def to_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date() if value.time() == datetime.min.time() else value
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):  # Excel serial
        return openpyxl.utils.datetime.from_excel(value)
    text = str(value).strip()
    for fmt in _DATE_PATTERNS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def is_dirty_placeholder(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    text = value.strip().lower()
    return any(text.startswith(t) and text != t for t in cfg.BLANK_PREFIX_TOKENS)


def convert(header: str, value: Any, layout: TargetLayout, report: Report, row: Any) -> Any:
    value = _convert(header, value, layout, report, row)
    if value is not None:
        report.note_reference_value(header, value)
    return value


def _convert(header: str, value: Any, layout: TargetLayout, report: Report, row: Any) -> Any:
    if is_blank(value):
        return None
    if isinstance(value, str):
        value = value.strip()
    if is_dirty_placeholder(value):
        report.warn(row, header, f"placeholder with extra text {value!r}; left blank")
        return None

    if layout.is_date(header):
        parsed = to_date(value)
        if parsed is None:
            report.warn(row, header, f"unparseable date {value!r}; left blank")
        elif isinstance(value, str):
            m = re.match(cfg.AMBIGUOUS_DATE_RE, value)
            if m and int(m.group(1)) <= 12 and int(m.group(2)) <= 12 and m.group(1) != m.group(2):
                # UAT-06 (Q24 open): read only with the configured order, reported.
                report.warn(row, header, f"ambiguous date text {value!r} read as {cfg.DATE_ORDER} "
                                         f"(configured order) = {parsed.isoformat()}; day and month could be "
                                         "swapped, verify (Q24 pending)")
        return parsed

    if layout.is_country(header):
        text = str(value)
        if text.upper() in cfg.ISO3_CODES:
            return text.upper()
        iso = cfg.COUNTRY_ISO3.get(text.lower())
        if iso is None:
            report.warn(row, header, f"no ISO-3 code for country {text!r}; passed through")
            return text
        # Performed even while the rule is not approved (diagnostics), but
        # then every converted value is a provisional conversion.
        report.rule_hit(cfg.RULE_COUNTRY_ISO3)
        if not report.rule_enabled(cfg.RULE_COUNTRY_ISO3):
            report.provisional[(header, text, iso)] += 1
        return iso

    if isinstance(value, bool):
        value = "Yes" if value else "No"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if "CF_TEXT" in header and not isinstance(value, str):
        value = str(value)

    if header in cfg.NUMERIC_RANGES and isinstance(value, (int, float)):
        low, high = cfg.NUMERIC_RANGES[header]
        if not low <= value <= high:
            report.warn(row, header, f"value {value} outside expected range {low}-{high}; kept, please verify")

    value_maps = report.rules.params(cfg.RULE_VALUE_MAPS)["maps"]
    if header in value_maps and isinstance(value, str):
        mapped = value_maps[header].get(value.lower(), value)
        if mapped != value:
            report.rule_hit(cfg.RULE_VALUE_MAPS)
            if report.rule_enabled(cfg.RULE_VALUE_MAPS):
                value = mapped
    if layout.is_resource(header) and isinstance(value, str):
        lookup = report.lookup_values(header)
        if value in lookup:
            value = lookup[value]
        else:
            report.unmapped_resources[(header, value)] += 1
    if layout.is_dropdown(header) or layout.is_resource(header):
        report.dropdown_values[(header, value)] += 1
    return value


def validate_record(main: dict, report: Report, row_ref: str) -> None:
    """Cross-field checks on one assignment. Warns only; never changes values."""
    for code_h, name_h in cfg.CODE_NAME_PAIRS:
        code, name = main.get(code_h), main.get(name_h)
        if isinstance(name, str) and (name == code or re.match(cfg.COST_CENTER_CODE_RE, name)):
            report.warn(row_ref, name_h, f"name {name!r} looks like a code (code column holds {code!r})")
    validate_family(main, cfg.FAMILY_PREFIXES[0], report, row_ref)


def validate_family(rec: dict, prefix: str, report: Report, row_ref: str) -> None:
    """Q10 / Q20: a spouse/partner or dependent record with at least one
    populated field is written as it is. Missing names are a DATA_QUALITY
    note for business review (never a hold, a drop or an invented name);
    relationship values are counted for the Equus list check (Q14)."""
    values = {k: v for k, v in rec.items() if k.split("!")[0] in (prefix, prefix + "_CF") and v is not None}
    if not values:
        return
    for header in cfg.FAMILY_LIST_HEADERS:
        if header in values:
            report.family_values[(header, str(values[header]))] += 1
    missing = [f for f in cfg.FAMILY_NAME_FIELDS if not values.get(f"{prefix}!{f}")]
    if missing:
        fields = ", ".join(sorted(k.split("!")[1] for k in values))
        report.data_quality.append({"where": row_ref, "record": prefix, "fields": fields,
                                    "missing": ", ".join(missing)})
        report.warn(row_ref, prefix, f"DATA_QUALITY: {cfg.FAMILY_NAME_NOTE} (has {fields}; missing {', '.join(missing)})")


# --------------------------------------------------------------------------- rules

def apply_assignment_rules(main: dict, cancel_value: Any, report: Report) -> None:
    """Assignment-level business rules. Each one changes `main` only while its
    configuration entry is approved_for_import; otherwise the row is only
    counted as affected for the Rules and Readiness sheets."""
    rule_id = cfg.RULE_CANCELLATION_PHASE
    if not is_blank(cancel_value):
        report.rule_hit(rule_id)
        if report.rule_enabled(rule_id):
            p = report.rules.params(rule_id)
            main[p["target"]] = p["value"]

    rule_id = cfg.RULE_INTERNATIONAL_RELO
    p = report.rules.params(rule_id)
    home, host = main.get(p["home_country"]), main.get(p["host_country"])
    if main.get(p["target"]) is None and home and host:
        report.rule_hit(rule_id)
        if report.rule_enabled(rule_id):
            main[p["target"]] = p["value_if_different"] if home != host else p["value_if_same"]

    rule_id = cfg.RULE_DEFAULT_DATES
    p = report.rules.params(rule_id)
    start = main.get(p["start_header"])
    if start:
        for date_header, triggers in p["effective_dates"].items():
            if main.get(date_header) is None and any(main.get(t) is not None for t in triggers):
                report.rule_hit(rule_id)
                if report.rule_enabled(rule_id):
                    main[date_header] = start


def default_child_dates(rec: dict, start: Any, report: Report) -> None:
    """Begin/start date of a child record (address, contact) from the
    assignment start date, only while the date rule is approved."""
    rule_id = cfg.RULE_DEFAULT_DATES
    if not start or not rec:
        return
    tables = {k.split("!")[0] for k in rec}
    for header in report.rules.params(rule_id)["child_dates"]:
        if header.split("!")[0] in tables and rec.get(header) is None:
            report.rule_hit(rule_id)
            if report.rule_enabled(rule_id):
                rec[header] = start


# --------------------------------------------------------------------------- identity controls

def identifier_text(value: Any) -> str | None:
    """Assignment Identifier as text: leading zeros of a text cell are kept,
    a numeric cell is rendered without '.0'. Blank / placeholder = None (never
    replaced by another field, BR-01)."""
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, str):
        return value.strip()
    return str(value)


def _canonical(value: Any) -> Any:
    """Comparable form of a source cell for duplicate detection."""
    if is_blank(value):
        return None
    if isinstance(value, str):
        return ("s", value.strip())
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, (int, float)):
        return ("n", int(value) if float(value).is_integer() else value)
    if isinstance(value, (date, datetime)):
        return ("d", value.isoformat())
    return ("o", str(value))


def _data_rows(src_ws, cols: list[SourceColumn]) -> list[int]:
    return [r for r in range(cfg.SOURCE_DATA_START_ROW, src_ws.max_row + 1)
            if not all(is_blank(src_ws.cell(r, c.idx).value) for c in cols)]


def screen_rows(src_ws, cols: list[SourceColumn], report: Report) -> None:
    """Identity and duplicate controls over every source data row (BR-01,
    BR-05, BR-17). Fills report.row_identity and holds rows through
    report.add_exception; never changes a value.

    * no Assignment Identifier            -> MISSING_ASSIGNMENT_ID (no fallback)
    * no Employee Number                  -> MISSING_EMPLOYEE_NUMBER
    * all compared cells equal an earlier row -> DUPLICATE_ROW (earlier row kept)
    * one identifier, rows with different data -> CONFLICTING_IDENTIFIER (all rows)
    * identifier accepted by an earlier run   -> ALREADY_ACCEPTED
    The same Employee Number with different identifiers is valid (Q17)."""
    ic = report.rules.identity
    ident_col = next((c for c in cols if c.label == ic.assignment_identifier_label), None)
    emp_col = next((c for c in cols if c.status == "mapped" and c.target == cfg.COL_EMPLOYEE_NUMBER), None)
    compare = [c for c in cols if c.status in ic.duplicate_compare_statuses]
    if ident_col is not None and ident_col not in compare:
        compare.append(ident_col)
    if emp_col is not None and emp_col not in compare:
        compare.append(emp_col)

    max_col = src_ws.max_column
    report.source_headers = []
    for c in range(1, max_col + 1):
        label = next((v for v in (src_ws.cell(r, c).value for r in (cfg.SOURCE_TOPIA_ROW, cfg.SOURCE_UI_ROW,
                                                                     cfg.SOURCE_KEY_ROW)) if v is not None), None)
        report.source_headers.append((get_column_letter(c), label))

    rows = _data_rows(src_ws, cols)
    signatures: dict[int, tuple] = {}
    first_seen: dict[tuple, int] = {}
    by_identifier: dict[str, list[int]] = defaultdict(list)
    for r in rows:
        identifier = identifier_text(src_ws.cell(r, ident_col.idx).value) if ident_col else None
        employee = identifier_text(src_ws.cell(r, emp_col.idx).value) if emp_col else None
        report.row_identity[r] = (identifier, employee)
        sig = tuple(_canonical(src_ws.cell(r, c.idx).value) for c in compare)
        signatures[r] = sig
        where = f"source {ident_col.letter}{r}" if ident_col else f"source row {r}"
        if identifier is None:
            report.add_exception(cfg.EXC_MISSING_ASSIGNMENT_ID,
                                 f"{where}: no Assignment Identifier (Topia "
                                 f"'{ic.assignment_identifier_label}'); row held, the Employee Number is not "
                                 "used as identifier (BR-01, D1)", r, identifier, employee)
        else:
            by_identifier[identifier].append(r)
        if employee is None:
            report.add_exception(cfg.EXC_MISSING_EMPLOYEE_NUMBER,
                                 f"source row {r}: no Employee Number (Talent Link ID); row held", r, identifier,
                                 employee)
        if sig in first_seen:
            report.add_exception(cfg.EXC_DUPLICATE_ROW,
                                 f"source row {r}: every compared cell equals source row {first_seen[sig]} "
                                 f"({len(compare)} columns); row {first_seen[sig]} is kept, this row is held",
                                 r, identifier, employee)
        else:
            first_seen[sig] = r

    for identifier, group in by_identifier.items():
        if len({signatures[r] for r in group}) < 2:
            continue
        differing = [c.letter for i, c in enumerate(compare) if len({signatures[r][i] for r in group}) > 1]
        for r in group:
            others = ", ".join(str(o) for o in group if o != r)
            report.add_exception(cfg.EXC_CONFLICTING_IDENTIFIER,
                                 f"source row {r}: Assignment Identifier also on source row(s) {others} with "
                                 f"different data in column(s) {', '.join(differing)}; all rows held",
                                 r, identifier, report.row_identity[r][1])

    for r in rows:
        identifier, employee = report.row_identity[r]
        if identifier is not None and identifier in report.previously_accepted and r not in report.held_rows:
            report.add_exception(cfg.EXC_ALREADY_ACCEPTED,
                                 f"source row {r}: identifier was accepted by an earlier run; not imported again",
                                 r, identifier, employee)

    for r in report.held_rows:
        report.held_row_values[r] = [src_ws.cell(r, c).value for c in range(1, max_col + 1)]


def validate_row_structure(rows: list[dict], report: Report) -> None:
    """Generated rows must follow Jake's example (' Mny dependents Ex', Q17-Q19):
    the first row of an assignment is Employee command INSERT (first assignment
    of the employee in the file) or IGNORE, Assignment Command INSERT,
    Assignment Number NEW; further rows are IGNORE / UPDATE with blank Employee
    and Assignment Number and the same Assignment Identifier. A violation is an
    engine defect and withholds the whole file."""
    problems: list[str] = []
    current = None
    seen_employees: set = set()
    identifiers: set = set()
    for i, row in enumerate(rows):
        out_row = cfg.TARGET_DATA_START_ROW + i
        command = row.get(cfg.COL_ASSIGNMENT_COMMAND)
        ident = row.get(cfg.COL_ASSIGNMENT_IDENTIFIER)
        if ident is None or (isinstance(ident, str) and not ident.strip()):
            problems.append(f"output row {out_row}: no Assignment Identifier")
        if command == "INSERT":
            employee = row.get(cfg.COL_EMPLOYEE_NUMBER)
            expected = cfg.EMPLOYEE_ACTION_REPEAT if employee in seen_employees else cfg.EMPLOYEE_ACTION_FIRST
            if employee is None:
                problems.append(f"output row {out_row}: INSERT row without Employee Number")
            if row.get(cfg.COL_EMPLOYEE_ACTION) != expected:
                problems.append(f"output row {out_row}: Employee command {row.get(cfg.COL_EMPLOYEE_ACTION)!r}, "
                                f"expected {expected!r}")
            if row.get(cfg.COL_ASSIGNMENT_NUMBER) != cfg.NEW_ASSIGNMENT_NUMBER:
                problems.append(f"output row {out_row}: Assignment Number is not {cfg.NEW_ASSIGNMENT_NUMBER!r}")
            if ident in identifiers:
                problems.append(f"output row {out_row}: Assignment Identifier INSERTed twice")
            identifiers.add(ident)
            seen_employees.add(employee)
            current = ident
        elif command == "UPDATE":
            if current is None:
                problems.append(f"output row {out_row}: UPDATE row before any INSERT row")
            elif ident != current:
                problems.append(f"output row {out_row}: UPDATE row carries a different Assignment Identifier "
                                "than its INSERT row (cross-assignment linkage, Q18)")
            if row.get(cfg.COL_EMPLOYEE_ACTION) != cfg.EMPLOYEE_ACTION_REPEAT:
                problems.append(f"output row {out_row}: UPDATE row Employee command is not "
                                f"{cfg.EMPLOYEE_ACTION_REPEAT!r}")
            for header in (cfg.COL_EMPLOYEE_NUMBER, cfg.COL_ASSIGNMENT_NUMBER):
                if row.get(header) is not None:
                    problems.append(f"output row {out_row}: UPDATE row has {header} (must be blank)")
        else:
            problems.append(f"output row {out_row}: Assignment Command {command!r} is neither INSERT nor UPDATE")
    for problem in problems:
        report.add_exception(cfg.EXC_ROW_STRUCTURE_INVALID, problem, "output")


def _header_value(value: Any) -> Any:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    return value


def check_template_headers(output: bytes, template_bytes: bytes, report: Report) -> None:
    """Compare header rows of the output workbook with the approved (bundled)
    template version; record template version and SHA-256 (BR-05, Q5: header
    issues fail the whole import). Pending header decisions (D2) are reported,
    the header itself is never changed here."""
    ic = report.rules.identity
    approved_bytes = ic.template_file.read_bytes()
    approved_ws = openpyxl.load_workbook(io.BytesIO(approved_bytes))[cfg.TARGET_SHEET]
    out_ws = openpyxl.load_workbook(io.BytesIO(output))[cfg.TARGET_SHEET]
    supplied_sha = hashlib.sha256(template_bytes).hexdigest()
    diffs = []
    for r in ic.header_rows:
        for c in range(1, max(approved_ws.max_column, out_ws.max_column) + 1):
            a, o = _header_value(approved_ws.cell(r, c).value), _header_value(out_ws.cell(r, c).value)
            if a != o:
                diffs.append({"cell": f"{get_column_letter(c)}{r}", "column": get_column_letter(c),
                              "approved": a, "found": o})
    report.template_check = {
        "template_version": ic.template_version, "approved_sha256": ic.template_sha256,
        "supplied_sha256": supplied_sha, "header_rows": ic.header_rows, "diffs": diffs,
        "decisions": [], "result": "FAILED" if diffs else "OK",
    }
    if diffs:
        columns = sorted({d["column"] for d in diffs}, key=lambda s: (len(s), s))
        detail = "; ".join(f"{d['cell']} approved {d['approved']!r}, found {d['found']!r}" for d in diffs[:20])
        more = f" (+{len(diffs) - 20} more)" if len(diffs) > 20 else ""
        report.add_exception(cfg.EXC_HEADER_MISMATCH,
                             f"{len(diffs)} header cell(s) differ from approved template {ic.template_version} in "
                             f"column(s) {', '.join(columns)}: {detail}{more}; import workbook withheld",
                             "template")
    elif supplied_sha != ic.template_sha256:
        report.add_exception(cfg.EXC_TEMPLATE_HASH_DIFFERS,
                             f"supplied template SHA-256 {supplied_sha} differs from approved "
                             f"{ic.template_version} ({ic.template_sha256}); header rows are identical",
                             "template")
    for d in ic.pending_header_decisions:
        approved_value = _header_value(approved_ws[d["cell"]].value)
        output_value = _header_value(out_ws[d["cell"]].value)
        report.template_check["decisions"].append({**d, "template_value": approved_value,
                                                   "output_value": output_value})
        if approved_value != d["approved_value"]:
            msg = (f"pending header decision {d['id']}: {d['cell']} of the approved template is "
                   f"{approved_value!r}, configuration expects {d['approved_value']!r}")
            report.config_issues.append((f"identity_controls.pending_header_decisions.{d['id']}", msg))
            report.warn("configuration", d["cell"], msg)
        report.add_exception(cfg.EXC_HEADER_DECISION_PENDING,
                             f"{d['id']}: header {d['cell']} is {approved_value!r} in the approved template "
                             f"{ic.template_version}, {d['alternative_value']!r} in {d['alternative_source']}; "
                             f"output keeps {output_value!r} until Jake decides (Q5: a header issue fails the "
                             "whole import)", "template")


# --------------------------------------------------------------------------- rows

def _group_dependents(cols: list[SourceColumn]) -> list[list[SourceColumn]]:
    """Split repeated ASSIGNMENT_DEPENDENT* columns into one group per child."""
    groups: list[list[SourceColumn]] = []
    for col in cols:
        if col.status != "dependent":
            continue
        if not groups or any(g.target == col.target for g in groups[-1]):
            groups.append([])
        groups[-1].append(col)
    return groups


def _key_text(value: Any) -> str | None:
    """Join-key text of an identifier cell; leading zeros are kept (never guessed)."""
    if is_blank(value):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def _dependent_columns(ws, row: int) -> list[SourceColumn]:
    """ASSIGNMENT_DEPENDENT* technical keys in one header row of a sheet."""
    cols = []
    for c in range(1, ws.max_column + 1):
        raw = ws.cell(row, c).value
        key = norm_key(raw)
        if key and key.startswith(cfg.FAMILY_PREFIXES[1]):
            cols.append(SourceColumn(idx=c, key=raw, ui=None, topia=None, status="dependent", target=key))
    return cols


def _find_labels(ws, last_row: int, labels: tuple[str, ...]) -> dict[str, int]:
    """Column of each normalised label within header rows 1..last_row."""
    found: dict[str, int] = {}
    for r in range(1, last_row + 1):
        for c in range(1, ws.max_column + 1):
            label = norm_label(ws.cell(r, c).value)
            if label in labels and label not in found:
                found[label] = c
    return found


def _dependent_rows(ws, groups: list[list[SourceColumn]], first_row: int) -> list[tuple[int, int]]:
    """(row, populated child groups) for every row holding dependent data."""
    out = []
    for r in range(first_row, ws.max_row + 1):
        n = sum(1 for g in groups if any(not is_blank(ws.cell(r, c.idx).value) for c in g))
        if n:
            out.append((r, n))
    return out


def read_dependent_sources(src_wb, layout: TargetLayout, report: Report) -> dict[str, list[dict]]:
    """Dependents of the keyed Topia source (business_rules.json
    $.dependent_source), indexed by Assignment ID key text.

    Every other sheet with ASSIGNMENT_DEPENDENT columns is scanned: without
    Employee ID + Assignment ID key columns it is HELD_DEPENDENTS_UNKEYED
    (assessment D3, e.g. Topia 'Dependent Import'), with them but not
    configured it is HELD_DEPENDENTS_NOT_CONFIGURED. Held dependents are
    counted, reported and never attached."""
    dep = report.rules.dependent_source
    key_labels = (dep.employee_label, dep.assignment_label) if dep is not None else ("employee id", "assignment id")
    index: dict[str, list[dict]] = defaultdict(list)

    for name in src_wb.sheetnames:
        if name == cfg.SOURCE_SHEET or (dep is not None and name == dep.sheet):
            continue
        ws = src_wb[name]
        header_row = next((r for r in range(1, min(ws.max_row, 10) + 1) if _dependent_columns(ws, r)), None)
        if header_row is None:
            continue
        groups = _group_dependents(_dependent_columns(ws, header_row))
        found = _dependent_rows(ws, groups, header_row + 1)
        if not found:
            continue
        keys = _find_labels(ws, header_row, key_labels)
        code = "HELD_DEPENDENTS_NOT_CONFIGURED" if len(keys) == 2 else "HELD_DEPENDENTS_UNKEYED"
        report.hold_dependents(code, name, [f"row {r}" for r, _ in found], sum(n for _, n in found),
                               f"dependent columns in header row {header_row}")

    if dep is None or not dep.configured:
        return index
    if dep.sheet not in src_wb.sheetnames:
        report.hold_dependents("HELD_DEPENDENTS_SHEET_MISSING", dep.sheet, [], 0)
        return index
    ws = src_wb[dep.sheet]
    groups = _group_dependents(_dependent_columns(ws, dep.header_row))
    if not groups:
        msg = f"dependent_source.sheet {dep.sheet!r} has no ASSIGNMENT_DEPENDENT columns in header_row {dep.header_row}"
        report.config_issues.append(("$.dependent_source.parameters.header_row", msg))
        report.warn("configuration", cfg.FAMILY_PREFIXES[1], msg)
        return index
    found = _dependent_rows(ws, groups, dep.data_start_row)
    keys = _find_labels(ws, dep.data_start_row - 1, key_labels)
    if len(keys) < 2:
        if found:
            missing = ", ".join(lbl for lbl in key_labels if lbl not in keys)
            report.hold_dependents("HELD_DEPENDENTS_UNKEYED", dep.sheet, [f"row {r}" for r, _ in found],
                                   sum(n for _, n in found), f"key column(s) not found: {missing}")
        return index
    if not dep.approved_for_import:
        if found:
            report.hold_dependents("HELD_DEPENDENTS_GATED", dep.sheet, [f"row {r}" for r, _ in found],
                                   sum(n for _, n in found), f"status {dep.status} ({dep.q}, owner {dep.owner})")
        return index

    emp_c, asg_c = keys[dep.employee_label], keys[dep.assignment_label]
    no_id: list[tuple[int, int]] = []
    for r, n in found:
        assignment = _key_text(ws.cell(r, asg_c).value)
        if assignment is None:
            no_id.append((r, n))
            continue
        recs = []
        for group in groups:
            rec = {c.target: convert(c.target, ws.cell(r, c.idx).value, layout, report, f"{dep.sheet}!{c.letter}{r}")
                   for c in group}
            if any(v is not None for v in rec.values()):
                recs.append(rec)
        index[assignment].append({"row": r, "employee": _key_text(ws.cell(r, emp_c).value), "records": recs})
    if no_id:
        report.hold_dependents("HELD_DEPENDENTS_NO_ASSIGNMENT_ID", dep.sheet, [f"row {r}" for r, _ in no_id],
                               sum(n for _, n in no_id))
    return index


def _dependent_signature(rec: dict) -> tuple:
    """Comparable form of one dependent record (populated fields only)."""
    return tuple(sorted((k, _canonical(v)) for k, v in rec.items() if v is not None))


def check_assignment_dates(entries: list[tuple], report: Report) -> None:
    """UAT-06 (BR-17/Q17, Q23): warn on a missing start date, an end date
    before its start date and assignments of one Employee Number whose date
    ranges overlap (a missing end date = open-ended). Warnings only: no date
    is derived, moved or blanked and no row is held."""
    start_h, end_h = cfg.ASSIGNMENT_START_HEADER, cfg.ASSIGNMENT_END_HEADER
    by_employee: dict = defaultdict(list)
    for r, employee, start, end in entries:
        if not start:
            report.warn(f"source row {r}", start_h, "no assignment start date; left blank, nothing derived (Q23 pending)")
            continue
        if end and _ordinal(end) < _ordinal(start):
            report.warn(f"source row {r}", end_h, "end date before start date; both kept unchanged, verify (Q23 pending)")
        by_employee[employee].append((_ordinal(start), _ordinal(end) if end else None, r))
    for spans in by_employee.values():
        spans.sort()
        for i, (s1, e1, r1) in enumerate(spans):
            for s2, _e2, r2 in spans[i + 1:]:
                if e1 is None or s2 <= e1:
                    report.warn(f"source row {r2}", start_h,
                                f"assignment dates overlap source row {r1} of the same Employee Number"
                                f"{' (open-ended)' if e1 is None else ''}; dates kept unchanged, verify "
                                "(BR-17/Q17, Q23 pending)")


def build_records(src_ws, cols: list[SourceColumn], layout: TargetLayout, report: Report,
                  keyed_dependents: dict[str, list[dict]] | None = None) -> list[dict]:
    mapped = [c for c in cols if c.status == "mapped"]
    addresses = [c for c in cols if c.status == "address"]
    contacts = [c for c in cols if c.status == "contact"]
    dependents = _group_dependents(cols)
    cancel_col = next((c for c in cols if c.rule == cfg.RULE_CANCELLATION_PHASE), None)
    override_cols = [c for c in cols if c.rule in cfg.COLUMN_OVERRIDE_RULES]
    start_header = report.rules.params(cfg.RULE_DEFAULT_DATES)["start_header"]

    # Identity and duplicate controls first: held rows never reach the import
    # workbook, and the Employee command INSERT goes to the first *accepted*
    # assignment of an employee.
    screen_rows(src_ws, cols, report)

    out_rows: list[dict] = []
    seen_employees: set = set()
    keyed_dependents = keyed_dependents or {}
    # Assignment IDs that occur on more than one source assignment row (held
    # or not) cannot take keyed dependents unambiguously (Q1/Q18).
    id_counts: Counter = Counter(ident for ident, _ in report.row_identity.values()
                                 if ident is not None) if keyed_dependents else Counter()
    consumed: set = set()
    assignment_dates: list[tuple] = []  # (source row, employee, start, end) of accepted rows

    for r in range(cfg.SOURCE_DATA_START_ROW, src_ws.max_row + 1):
        raw = {c.idx: src_ws.cell(r, c.idx).value for c in cols}
        if all(is_blank(v) for v in raw.values()):
            continue
        if r in report.held_rows:
            # Listed in Exceptions and Held Rows; never imported. Keyed
            # dependents of a held assignment are held with it, never
            # attached elsewhere.
            _hold_keyed_dependents_of_held_row(r, keyed_dependents, id_counts, consumed, report)
            continue
        report.current_row = r
        row_ref = f"source row {r}"
        main = {}
        for col in mapped:
            main[col.target] = convert(col.target, raw[col.idx], layout, report, f"source {col.letter}{r}")

        employee = main.get(cfg.COL_EMPLOYEE_NUMBER)
        # Text, leading zeros kept; screen_rows held every row without one, so
        # there is no fallback to the Employee Number (BR-01). Keyed
        # dependents join on this row's own Assignment ID only.
        identifier = report.row_identity[r][0]
        join_key = identifier
        report.accepted.append({"row": r, "identifier": identifier, "employee": report.row_identity[r][1]})

        # Business rules (configuration-gated)
        apply_assignment_rules(main, raw[cancel_col.idx] if cancel_col else None, report)
        for col in override_cols:
            if not is_blank(raw[col.idx]):
                report.rule_hit(col.rule)
        start = main.get(start_header)
        validate_record(main, report, row_ref)
        for state_h, country_h in cfg.STATE_COUNTRY_HEADERS.items():
            if main.get(state_h) is not None:
                report.state_values[(state_h, main.get(country_h), str(main[state_h]))] += 1
        assignment_dates.append((r, employee, main.get(cfg.ASSIGNMENT_START_HEADER),
                                 main.get(cfg.ASSIGNMENT_END_HEADER)))

        # Child records
        addr_recs = []
        for col in addresses:
            value = raw[col.idx]
            if is_blank(value):
                continue
            lines = [ln.strip() for ln in str(value).splitlines() if ln.strip()]
            rec = {
                ADDR + "PRIMARY_ADDRESS": "Yes",
                ADDR + "LOCATION_TYPE": cfg.ADDRESS_COLUMNS[col.label],
                ADDR + "LINE_1": lines[0] if lines else None,
                ADDR + "LINE_2": lines[1] if len(lines) > 1 else None,
                ADDR + "LINE_3": ", ".join(lines[2:]) or None,
            }
            default_child_dates(rec, start, report)
            report.warn(row_ref, ADDR + "LINE_1", f"address from column {col.letter} loaded as free text; verify city/postal/country split")
            addr_recs.append(rec)

        contact_recs = []
        for col in contacts:
            value = raw[col.idx]
            if is_blank(value):
                continue
            loc, ctype = cfg.CONTACT_COLUMNS[col.label]
            rec = {
                CONTACT + "LOCATION_TYPE": loc,
                CONTACT + "CONTACT_TYPE": ctype,
                CONTACT + "NUMBER_ADDRESS": str(int(value)) if isinstance(value, float) and value.is_integer() else str(value).strip(),
                CONTACT + "PRIMARY_CONTACT": "Yes",
            }
            default_child_dates(rec, start, report)
            contact_recs.append(rec)

        dep_recs = []
        seen_deps: set = set()
        duplicates = 0
        for gi, group in enumerate(dependents):
            rec = {c.target: convert(c.target, raw[c.idx], layout, report, f"source {c.letter}{r}") for c in group}
            if any(v is not None for v in rec.values()):
                sig = _dependent_signature(rec)
                if sig in seen_deps:  # UAT-06: repeated record held, never written twice
                    report.duplicate_dependent_groups.setdefault(r, set()).add(gi)
                    duplicates += 1
                    continue
                seen_deps.add(sig)
                validate_family(rec, cfg.FAMILY_PREFIXES[1], report, row_ref)
                dep_recs.append(rec)
        for rec in _attach_keyed_dependents(join_key, employee, keyed_dependents, id_counts, consumed,
                                            bool(dep_recs), report, row_ref):
            sig = _dependent_signature(rec)
            if sig in seen_deps:
                report.keyed_children[r] = [e for e in report.keyed_children.get(r, []) if e["record"] is not rec]
                report.dependents_attached -= 1
                duplicates += 1
                continue
            seen_deps.add(sig)
            dep_recs.append(rec)
        if duplicates:
            report.hold_dependents("HELD_DEPENDENTS_DUPLICATE", cfg.SOURCE_SHEET, [f"source row {r}"], duplicates,
                                   f"assignment at {row_ref}")

        # Equus layout: first row INSERTs the assignment; extra child records
        # go on UPDATE rows keyed by the Assignment Identifier.
        n_rows = max(1, len(addr_recs), len(contact_recs), len(dep_recs))
        for i in range(n_rows):
            row = dict(main) if i == 0 else {}
            if i == 0:
                row[cfg.COL_EMPLOYEE_ACTION] = cfg.EMPLOYEE_ACTION_FIRST if employee not in seen_employees else cfg.EMPLOYEE_ACTION_REPEAT
                row[cfg.COL_ASSIGNMENT_COMMAND] = "INSERT"
                row[cfg.COL_ASSIGNMENT_NUMBER] = cfg.NEW_ASSIGNMENT_NUMBER
            else:
                row[cfg.COL_EMPLOYEE_ACTION] = cfg.EMPLOYEE_ACTION_REPEAT
                row[cfg.COL_ASSIGNMENT_COMMAND] = "UPDATE"
            row[cfg.COL_ASSIGNMENT_IDENTIFIER] = identifier
            for recs in (addr_recs, contact_recs, dep_recs):
                if i < len(recs):
                    row.update(recs[i])
            out_rows.append(row)
        seen_employees.add(employee)
        report.assignments += 1
    report.current_row = None
    check_assignment_dates(assignment_dates, report)

    unmatched = [e for key, entries in keyed_dependents.items() if key not in consumed for e in entries]
    if unmatched:
        sheet = report.rules.dependent_source.sheet
        report.hold_dependents("HELD_DEPENDENTS_UNMATCHED", sheet, [f"row {e['row']}" for e in unmatched],
                               sum(len(e["records"]) for e in unmatched))
    return out_rows


def _hold_keyed_dependents_of_held_row(r: int, keyed: dict[str, list[dict]], id_counts: Counter, consumed: set,
                                       report: Report) -> None:
    """Keyed dependents whose Assignment ID belongs to a held source row are
    held with it (UAT-03): never attached, reported once per identifier."""
    join_key = report.row_identity.get(r, (None, None))[0]
    if join_key is None or join_key not in keyed or join_key in consumed:
        return
    consumed.add(join_key)
    entries = keyed[join_key]
    sheet = report.rules.dependent_source.sheet
    rows, groups = [f"row {e['row']}" for e in entries], sum(len(e["records"]) for e in entries)
    if id_counts[join_key] > 1:
        report.hold_dependents("HELD_DEPENDENTS_AMBIGUOUS", sheet, rows, groups,
                               f"Assignment ID on {id_counts[join_key]} assignment rows")
    else:
        report.hold_dependents("HELD_DEPENDENTS_ASSIGNMENT_HELD", sheet, rows, groups,
                               f"assignment source row {r} held ({', '.join(report.held_rows[r])})")


def _attach_keyed_dependents(join_key: str | None, employee: Any, keyed: dict[str, list[dict]], id_counts: Counter,
                             consumed: set, has_sheet_dependents: bool, report: Report, row_ref: str) -> list[dict]:
    """Dependent records of the keyed source for one assignment row, joined
    by its own Assignment ID only (Q18); conflicting rows are held."""
    if join_key is None or join_key not in keyed:
        return []
    entries = keyed[join_key]
    sheet = report.rules.dependent_source.sheet
    first_visit = join_key not in consumed
    consumed.add(join_key)
    if id_counts[join_key] > 1:
        if first_visit:
            report.hold_dependents("HELD_DEPENDENTS_AMBIGUOUS", sheet, [f"row {e['row']}" for e in entries],
                                   sum(len(e["records"]) for e in entries),
                                   f"Assignment ID on {id_counts[join_key]} assignment rows")
        return []
    out, mismatched = [], []
    employee_key = _key_text(employee)
    for entry in entries:
        if entry["employee"] is not None and entry["employee"] != employee_key:
            mismatched.append(entry)
            continue
        if entry["employee"] is None:
            report.warn(f"{sheet}!row {entry['row']}", cfg.FAMILY_PREFIXES[1],
                        "dependent row has no Employee ID; attached by Assignment ID only")
        for rec in entry["records"]:
            validate_family(rec, cfg.FAMILY_PREFIXES[1], report, f"{sheet} row {entry['row']}")
            out.append(rec)
            report.keyed_children.setdefault(report.current_row, []).append(
                {"sheet": sheet, "row": entry["row"], "record": rec})
    if mismatched:
        report.hold_dependents("HELD_DEPENDENTS_EMPLOYEE_MISMATCH", sheet, [f"row {e['row']}" for e in mismatched],
                               sum(len(e["records"]) for e in mismatched), f"assignment at {row_ref}")
    if out and has_sheet_dependents:
        report.dependent_overlaps += 1
        report.warn(row_ref, cfg.FAMILY_PREFIXES[1], "dependents from both the assignment sheet and the keyed "
                                                     "dependent source; review for duplicates")
    report.dependents_attached += len(out)
    return out


# --------------------------------------------------------------------------- io

def write_output(template_bytes: bytes, rows: list[dict], layout_ws_name: str, report: Report) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb[layout_ws_name]
    layout = read_target_layout(ws)
    identity = getattr(report.rules, "identity", None)
    text_headers = set(identity.identifier_headers) if identity is not None else set()
    if ws.max_row >= cfg.TARGET_DATA_START_ROW:
        ws.delete_rows(cfg.TARGET_DATA_START_ROW, ws.max_row - cfg.TARGET_DATA_START_ROW + 1)
    for offset, row in enumerate(rows):
        r = cfg.TARGET_DATA_START_ROW + offset
        for header, value in row.items():
            if value is None:
                continue
            c = layout.headers.get(header)
            if c is None:
                report.warn(f"output row {r}", header, "no Equus column for value; dropped")
                continue
            cell = ws.cell(r, c, value)
            if isinstance(value, (date, datetime)):
                cell.number_format = DATE_FORMAT
            elif header in text_headers and isinstance(value, str):
                cell.number_format = "@"  # identifiers stay text (leading zeros)
            report.coverage[header] += 1
    report.rows_written = len(rows)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


#: Every non-blank source cell falls in exactly one of these categories
#: (UAT-05). "written" = written unchanged or by an approved / deterministic
#: conversion (date parse, number format, approved rule or lookup);
#: "failed_conversion" = unparseable date left blank (warned);
#: "replaced_by_rule" = source value overwritten by an *approved* rule.
#: "set_by_rule" (not a source category) counts output cells an approved rule
#: filled where the source cell was blank.
RECON_CATEGORIES = ("written", "blank_token", "provisional_conversion", "gated_rule", "held_row",
                    "excluded_colour", "pending_target", "failed_conversion", "replaced_by_rule", "unexplained")
#: Column statuses whose cells must reach the import workbook.
MIGRATED_STATUSES = ("mapped", "duplicate", "rule", "address", "contact", "dependent")
_CHILD_TABLES = {"address": ("ASSIGNMENT_MAILING_ADDRESS",), "contact": ("ASSIGNMENT_EMPLOYEE_CONTACT",),
                 "dependent": (cfg.FAMILY_PREFIXES[1], cfg.FAMILY_PREFIXES[1] + "_CF"),
                 "spouse": (cfg.FAMILY_PREFIXES[0], cfg.FAMILY_PREFIXES[0] + "_CF")}


def _present(raw: Any) -> bool:
    return raw is not None and str(raw).strip() != ""


def _text(value: Any) -> str:
    if isinstance(value, bool):
        return "Yes" if value else "No"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def _same_value(raw: Any, out: Any) -> bool:
    """True when the written cell carries the source value unchanged (dates by
    calendar day, numbers by value, text stripped)."""
    if out is None:
        return False
    if isinstance(out, (date, datetime)):
        src = raw if isinstance(raw, (date, datetime)) else to_date(raw)
        return src is not None and _ordinal(src) == _ordinal(out)
    if isinstance(raw, (date, datetime)):
        return False
    return _text(raw) == _text(out)


def _would_write(header: str, raw: Any, layout: TargetLayout) -> bool:
    """A source cell convert() keeps (non-blank, no placeholder, parseable date)."""
    if is_blank(raw) or is_dirty_placeholder(raw):
        return False
    return not layout.is_date(header) or to_date(raw.strip() if isinstance(raw, str) else raw) is not None


def _explain_changed(header: str, raw: Any, out: Any, report: Report) -> str | None:
    """Category of a written value that differs from its source cell, or None."""
    text = _text(raw)
    if (header, text, out) in report.provisional:
        return "provisional_conversion"  # Q12/Q3: rule country_name_to_iso3 not approved
    if header in cfg.COUNTRY_HEADERS and isinstance(out, str):
        if text.upper() == out and out in cfg.ISO3_CODES:
            return "written"
        if report.rule_enabled(cfg.RULE_COUNTRY_ISO3) and cfg.COUNTRY_ISO3.get(text.lower()) == out:
            return "written"
    value_maps = report.rules.params(cfg.RULE_VALUE_MAPS)["maps"]
    if report.rule_enabled(cfg.RULE_VALUE_MAPS) and value_maps.get(header, {}).get(text.lower()) == out:
        return "written"
    lookup = report.lookup_values(header)
    if text in lookup and lookup[text] == out:
        return "written"
    return None


def _table(header: Any) -> str:
    return str(header).split("!")[0] if header else ""


def _output_blocks(out_ws, headers: dict) -> tuple[dict, list, int, int, int]:
    """Group output rows by Assignment Identifier: an INSERT row followed by
    its UPDATE rows. Returns ({identifier: [[row, ...], ...]}, unexpected
    rows, INSERT count, UPDATE count, non-empty output rows)."""
    cmd_c, ident_c = headers.get(cfg.COL_ASSIGNMENT_COMMAND), headers.get(cfg.COL_ASSIGNMENT_IDENTIFIER)
    blocks: dict[str, list[list[int]]] = defaultdict(list)
    unexpected: list[int] = []
    inserts = updates = total = 0
    current: tuple[str, list[int]] | None = None
    for r in range(cfg.TARGET_DATA_START_ROW, out_ws.max_row + 1):
        if all(out_ws.cell(r, c).value is None for c in headers.values()):
            continue
        total += 1
        command = out_ws.cell(r, cmd_c).value
        ident = identifier_text(out_ws.cell(r, ident_c).value) if ident_c else None
        if command == "INSERT" and ident is not None:
            inserts += 1
            current = (ident, [r])
            blocks[ident].append(current[1])
        elif command == "UPDATE" and current is not None and ident == current[0]:
            updates += 1
            current[1].append(r)
        else:
            inserts += command == "INSERT"
            updates += command == "UPDATE"
            unexpected.append(r)
    return blocks, unexpected, inserts, updates, total


def reconcile(src_ws, cols: list[SourceColumn], output: bytes, report: Report) -> None:
    """Source-to-target reconciliation (UAT-05, master command section 8).

    Keyed on the Assignment Identifier, never on row position: every source
    data row is accepted (one INSERT row plus UPDATE rows for child records)
    or held. Every non-blank source cell of every column is counted in exactly
    one RECON_CATEGORIES bucket: mapped cells against the INSERT row; address,
    contact and dependent cells against the output child row they create (the
    k-th populated group of an assignment is on the k-th row of its block);
    held rows, red columns, gated rules and columns without an approved Equus
    target are explained categories. Anything else is unexplained.

    reconciliation_ok only when unexplained == 0 AND the counts balance
    (source rows = accepted + held; INSERT rows = accepted; UPDATE rows =
    expected expansion rows; child records expected = written per record)."""
    out_ws = openpyxl.load_workbook(io.BytesIO(output))[cfg.TARGET_SHEET]
    out_headers = {}
    for c in range(1, out_ws.max_column + 1):
        h = out_ws.cell(cfg.TARGET_HEADER_ROW, c).value
        if h is not None:
            out_headers[norm_key(h) or str(h).strip()] = c
    layout = read_target_layout(out_ws)
    all_rows = _data_rows(src_ws, cols)
    report.data_rows = len(all_rows)
    report.reconciliation, report.reconciliation_records = [], []
    reasons: list[str] = []
    if cfg.COL_ASSIGNMENT_COMMAND not in out_headers or cfg.COL_ASSIGNMENT_IDENTIFIER not in out_headers:
        report.reconciliation_ok = False  # header mismatch: nothing can be reconciled
        report.reconciliation_summary = {"result": "FAILED", "source_rows": len(all_rows),
                                         "reasons": ["output has no Assignment Command / Identifier header"]}
        return

    blocks, unexpected_rows, insert_count, update_count, output_rows = _output_blocks(out_ws, out_headers)
    accepted = {a["row"]: a["identifier"] for a in report.accepted}
    held = report.held_rows
    rule_targets = report.rules.applied_targets()

    def cell(r: int, header: str) -> Any:
        c = out_headers.get(header)
        return out_ws.cell(r, c).value if c is not None else None

    def block_of(r: int) -> list[int]:
        found = blocks.get(accepted.get(r)) or []
        return found[0] if len(found) == 1 else []

    # Field table: one row per source column.
    rows_by_col: dict[int, dict] = {}

    def field_row(col: SourceColumn) -> dict:
        if col.idx not in rows_by_col:
            rows_by_col[col.idx] = {
                "col": col.letter, "topia": col.topia or col.ui, "header": col.target, "group": col.status,
                "status": col.status, "colour": col.colour, "flags": ", ".join(col.flags), "source": 0,
                **{k: 0 for k in RECON_CATEGORIES}, "set_by_rule": 0,
                "source_date_total": None, "written_date_total": None, "q": set(), "owner": set(), "reason": set()}
        return rows_by_col[col.idx]

    def explain(row: dict, category: str, q: str, owner: str, reason: str) -> None:
        row[category] += 1
        row["q"].add(q)
        row["owner"].add(owner)
        row["reason"].add(reason)

    def column_disposition(col: SourceColumn) -> tuple[str, str, str, str]:
        """(category, q, owner, reason) of a column that is not migrated."""
        if col.status == "excluded":
            q, owner = cfg.READINESS_ITEMS["excluded_colour"]
            return "excluded_colour", q, owner, col.note
        if col.status == "held":
            rule = report.rules.rule(col.rule)
            return "gated_rule", rule.q, rule.owner, f"rule {rule.id} {rule.status}, not approved_for_import"
        if "MISSING_TARGET" in col.flags:
            q, owner = cfg.READINESS_ITEMS["missing_target"]
        elif col.label in cfg.NOT_MIGRATED_QUESTIONS:
            q, owner = cfg.NOT_MIGRATED_QUESTIONS[col.label]
        elif "PENDING" in col.flags:
            q, owner = cfg.READINESS_ITEMS["pending_column"]
        else:
            q, owner = cfg.READINESS_ITEMS["not_migrated"]
        return "pending_target", q, owner, f"no approved Equus target ({col.status}: {col.note})"

    def held_reason(r: int) -> tuple[str, str, str]:
        specs = [report.rules.identity.code(code) for code in held[r]]
        return ("/".join(sorted({s.q for s in specs})), "; ".join(sorted({s.owner for s in specs})),
                "row held: " + ", ".join(held[r]))

    def compare(row: dict, col: SourceColumn, header: str, raw: Any, out: Any) -> None:
        """One present, non-placeholder source cell against its written cell."""
        if _same_value(raw, out):
            row["written"] += 1
            if isinstance(out, (date, datetime)):
                row["source_date_total"] = (row["source_date_total"] or 0) + _ordinal(out)
                row["written_date_total"] = (row["written_date_total"] or 0) + _ordinal(out)
            return
        if out is None:
            if layout.is_date(header) and to_date(raw.strip() if isinstance(raw, str) else raw) is None:
                explain(row, "failed_conversion", "Q24", "Jake", "unparseable date left blank (warned)")
            else:
                row["unexplained"] += 1
            return
        category = _explain_changed(header, raw, out, report)
        if category == "provisional_conversion":
            rule = report.rules.rule(cfg.RULE_COUNTRY_ISO3)
            explain(row, category, rule.q, rule.owner, f"rule {rule.id} not approved; conversion provisional")
        elif category == "written":
            row["written"] += 1
        elif header in rule_targets:
            row["replaced_by_rule"] += 1  # source value replaced by an approved rule
        else:
            row["unexplained"] += 1
            if isinstance(out, (date, datetime)):
                src = to_date(raw)
                row["source_date_total"] = (row["source_date_total"] or 0) + _ordinal(src)
                row["written_date_total"] = (row["written_date_total"] or 0) + _ordinal(out)

    addresses = [c for c in cols if c.status == "address"]
    contacts = [c for c in cols if c.status == "contact"]
    dep_groups = _group_dependents(cols)
    spouse_cols = [c for c in cols if c.status == "mapped" and _table(c.target) in _CHILD_TABLES["spouse"]]
    bq, bowner = cfg.READINESS_ITEMS["blank_token"]

    for col in cols:
        if col.status in MIGRATED_STATUSES:
            field_row(col)

    record_unexplained_total = 0
    for r in all_rows:
        ident, employee = report.row_identity.get(r, (None, None))
        block = block_of(r) if r in accepted else []
        insert_row = block[0] if block else None
        rec = {"source_row": r, "identifier": ident, "employee": employee,
               "status": "held" if r in held else "accepted", "codes": ", ".join(held.get(r, [])),
               "output_rows": len(block), "insert_rows": 1 if block else 0, "update_rows": max(0, len(block) - 1),
               "unexplained": 0, "notes": []}
        expected = {"address": 0, "contact": 0, "dependent": 0, "spouse": 0}

        # Per-cell categorisation.
        child_pos = {"address": 0, "contact": 0, "dependent": 0}
        for col in cols:
            raw = src_ws.cell(r, col.idx).value
            if not _present(raw):
                continue
            row = field_row(col)
            row["source"] += 1
            if r in held:
                explain(row, "held_row", *held_reason(r))
                continue
            if col.status not in MIGRATED_STATUSES:
                explain(row, *column_disposition(col))  # the column is not migrated, whatever the cell holds
                continue
            if col.status == "dependent":
                continue  # per child group below
            if col.status not in ("address", "contact") and (is_blank(raw) or is_dirty_placeholder(raw)):
                explain(row, "blank_token", bq, bowner, "placeholder treated as blank (BLANK_TOKENS)")
                continue
            if col.status in ("mapped", "duplicate"):
                out = cell(insert_row, col.target) if insert_row else None
                compare(row, col, col.target, raw, out)
            elif col.status == "rule":
                p = report.rules.params(col.rule)
                if insert_row and cell(insert_row, p["target"]) == p["value"]:
                    row["written"] += 1  # consumed by the approved rule
                else:
                    row["unexplained"] += 1
            elif col.status in ("address", "contact"):
                if is_blank(raw):
                    explain(row, "blank_token", bq, bowner, "placeholder treated as blank (BLANK_TOKENS)")
                    continue
                kind = col.status
                k = child_pos[kind]
                child_pos[kind] += 1
                expected[kind] += 1
                orow = block[k] if k < len(block) else None
                if kind == "address":
                    lines = [ln.strip() for ln in str(raw).splitlines() if ln.strip()]
                    want = {ADDR + "LOCATION_TYPE": cfg.ADDRESS_COLUMNS[col.label],
                            ADDR + "LINE_1": lines[0] if lines else None,
                            ADDR + "LINE_2": lines[1] if len(lines) > 1 else None,
                            ADDR + "LINE_3": ", ".join(lines[2:]) or None}
                else:
                    loc, ctype = cfg.CONTACT_COLUMNS[col.label]
                    want = {CONTACT + "LOCATION_TYPE": loc, CONTACT + "CONTACT_TYPE": ctype,
                            CONTACT + "NUMBER_ADDRESS": _text(raw)}
                got = {h: (cell(orow, h) if orow else None) for h in want}
                if all((w is None and g is None) or (w is not None and _same_value(w, g)) for h, w in want.items()
                       for g in [got[h]]):
                    row["written"] += 1
                else:
                    row["unexplained"] += 1
            # dependent columns: per group below

        if r not in held:
            # Sheet dependents: the k-th group with data is on block row k.
            k = 0
            held_groups = report.duplicate_dependent_groups.get(r, set())
            for gi, group in enumerate(dep_groups):
                cells = [(c, src_ws.cell(r, c.idx).value) for c in group]
                if gi in held_groups:  # UAT-06: repeated dependent record held, no output row
                    q, owner, _ = cfg.DEPENDENT_HOLDS["HELD_DEPENDENTS_DUPLICATE"]
                    for c, v in cells:
                        if _present(v):
                            explain(field_row(c), "held_row", q, owner,
                                    "dependent record held: HELD_DEPENDENTS_DUPLICATE")
                    continue
                if not any(_would_write(c.target, v, layout) for c, v in cells):
                    for c, v in cells:  # placeholders / unparseable only: no record is created
                        if _present(v):
                            row = field_row(c)
                            if is_blank(v) or is_dirty_placeholder(v):
                                explain(row, "blank_token", bq, bowner, "placeholder treated as blank (BLANK_TOKENS)")
                            else:
                                explain(row, "failed_conversion", "Q24", "Jake", "unparseable date left blank (warned)")
                    continue
                orow = block[k] if k < len(block) else None
                k += 1
                expected["dependent"] += 1
                for c, v in cells:
                    if not _present(v):
                        continue
                    row = field_row(c)
                    if is_blank(v) or is_dirty_placeholder(v):
                        explain(row, "blank_token", bq, bowner, "placeholder treated as blank (BLANK_TOKENS)")
                    else:
                        compare(row, c, c.target, v, cell(orow, c.target) if orow else None)
            # Keyed dependents (business_rules.json $.dependent_source) follow the sheet dependents.
            for entry in report.keyed_children.get(r, []):
                orow = block[k] if k < len(block) else None
                k += 1
                expected["dependent"] += 1
                for header, value in entry["record"].items():
                    if value is None:
                        continue
                    key = ("keyed", entry["sheet"], header)
                    row = rows_by_col.setdefault(key, {
                        "col": f"{entry['sheet']} (keyed)", "topia": None, "header": header, "group": "keyed_dependent",
                        "status": "dependent", "colour": "", "flags": "", "source": 0,
                        **{cat: 0 for cat in RECON_CATEGORIES}, "set_by_rule": 0, "source_date_total": None,
                        "written_date_total": None, "q": set(), "owner": set(), "reason": set()})
                    row["source"] += 1
                    row["written" if _same_value(value, cell(orow, header) if orow else None) else "unexplained"] += 1
            expected["spouse"] = int(any(_would_write(c.target, src_ws.cell(r, c.idx).value, layout)
                                         for c in spouse_cols))

        # Source blank, output populated: only an approved rule may do that.
        if insert_row:
            for col in cols:
                if col.status != "mapped" or _present(src_ws.cell(r, col.idx).value):
                    continue
                if cell(insert_row, col.target) is not None:
                    field_row(col)["set_by_rule" if col.target in rule_targets else "unexplained"] += 1

        written = {kind: sum(1 for orow in (block[:1] if kind == "spouse" else block)
                             if any(_table(h) in tables and out_ws.cell(orow, c).value is not None
                                    for h, c in out_headers.items()))
                   for kind, tables in _CHILD_TABLES.items()}
        for kind in ("address", "contact", "dependent", "spouse"):
            rec[f"{kind}_expected"], rec[f"{kind}_written"] = expected[kind], written[kind]
            if expected[kind] != written[kind]:
                rec["unexplained"] += abs(expected[kind] - written[kind])
                rec["notes"].append(f"{kind} records expected {expected[kind]}, written {written[kind]}")
        if r in held:
            rec["expected_expansion_rows"] = 0  # nothing of a held row reaches the import workbook
        else:
            n_rows = max(1, expected["address"], expected["contact"], expected["dependent"])
            rec["expected_expansion_rows"] = n_rows - 1
            found = blocks.get(accepted.get(r)) or []
            if len(found) != 1:
                rec["unexplained"] += 1
                rec["notes"].append(f"{len(found)} INSERT blocks for this Assignment Identifier (expected 1)")
            elif len(block) != n_rows:
                rec["unexplained"] += abs(len(block) - n_rows)
                rec["notes"].append(f"output rows {len(block)}, expected {n_rows}")
        rec["notes"] = "; ".join(rec["notes"])
        record_unexplained_total += rec["unexplained"]
        report.reconciliation_records.append(rec)

    # Field rows: finalise and check that every source cell is in one bucket.
    field_unexplained = 0
    unbalanced = []
    for _, row in sorted(rows_by_col.items(), key=lambda kv: (1, str(kv[0])) if isinstance(kv[0], tuple)
                         else (0, f"{kv[0]:06d}")):
        for k in ("q", "owner", "reason"):
            row[k] = "; ".join(sorted(x for x in row[k] if x))
        if row["source"] != sum(row[k] for k in RECON_CATEGORIES):
            unbalanced.append(row["col"])
        field_unexplained += row["unexplained"]
        report.reconciliation.append(row)

    accepted_n, held_n = len(accepted), len(held)
    expansion = sum(r["expected_expansion_rows"] for r in report.reconciliation_records)
    duplicates = sum(len(v) - 1 for v in blocks.values() if len(v) > 1)
    orphan_blocks = sum(len(v) for k, v in blocks.items() if k not in set(accepted.values()))
    totals = {cat: sum(r[cat] for r in report.reconciliation) for cat in RECON_CATEGORIES}
    children = {kind: (sum(r[f"{kind}_expected"] for r in report.reconciliation_records),
                       sum(r[f"{kind}_written"] for r in report.reconciliation_records))
                for kind in ("address", "contact", "dependent", "spouse")}
    employees = {e for _, e in report.row_identity.values() if e is not None}
    checks = [
        (report.data_rows == accepted_n + held_n, f"source rows {report.data_rows} != accepted {accepted_n} + held "
                                                  f"{held_n}"),
        (insert_count == accepted_n == report.assignments,
         f"INSERT rows {insert_count}, accepted {accepted_n}, assignments {report.assignments} differ"),
        (update_count == expansion, f"UPDATE rows {update_count} != expected expansion rows {expansion}"),
        (report.rows_written == output_rows, f"rows written {report.rows_written} != output rows found {output_rows}"),
        (not unexpected_rows, f"{len(unexpected_rows)} output row(s) not linked to an accepted assignment"),
        (not duplicates, f"{duplicates} Assignment Identifier(s) INSERTed more than once"),
        (not orphan_blocks, f"{orphan_blocks} INSERT block(s) with an identifier of no accepted source row"),
        (not unbalanced, f"category counts do not add up for column(s) {', '.join(map(str, unbalanced))}"),
        (all(e == w for e, w in children.values()),
         "child records expected != written: " + ", ".join(f"{k} {e}/{w}" for k, (e, w) in children.items()
                                                            if e != w)),
    ]
    counts_ok = all(ok for ok, _ in checks)
    reasons = [msg for ok, msg in checks if not ok]
    unexplained = field_unexplained + record_unexplained_total + len(unexpected_rows) + duplicates + orphan_blocks
    if field_unexplained:
        cols_bad = [f"{r['col']} ({r['header']}) {r['unexplained']}" for r in report.reconciliation if r["unexplained"]]
        reasons.append(f"{field_unexplained} unexplained source cell(s): " + ", ".join(cols_bad))
    if record_unexplained_total:
        reasons.append(f"{record_unexplained_total} unexplained record-level difference(s), see Reconciliation "
                       "Records")
    report.reconciliation_ok = counts_ok and unexplained == 0
    report.reconciliation_summary = {
        "result": "OK" if report.reconciliation_ok else "FAILED",
        "source_rows": report.data_rows,
        "source_employees": len(employees),
        "source_assignment_identifiers": len({i for i, _ in report.row_identity.values() if i is not None}),
        "accepted_assignments": accepted_n,
        "accepted_employees": len({a["employee"] for a in report.accepted if a["employee"] is not None}),
        "held_rows": held_n,
        "held_dependent_records": sum(e["groups"] for e in report.dependent_exceptions),
        "insert_rows": insert_count,
        "update_rows": update_count,
        "expected_expansion_rows": expansion,
        "child_records_expected": {k: e for k, (e, _) in children.items()},
        "child_records_written": {k: w for k, (_, w) in children.items()},
        "rows_written": report.rows_written,
        "unexpected_output_rows": len(unexpected_rows),
        "duplicate_insert_blocks": duplicates,
        "source_cells_non_blank": sum(r["source"] for r in report.reconciliation),
        **{f"cells_{cat}": n for cat, n in totals.items()},
        "cells_set_by_rule": sum(r["set_by_rule"] for r in report.reconciliation),
        "record_unexplained": record_unexplained_total,
        "unexplained_total": unexplained,
        "counts_balance": counts_ok,
        "reasons": reasons,
    }


def _ordinal(value: Any) -> int:
    if isinstance(value, datetime):
        value = value.date()
    return value.toordinal() if isinstance(value, date) else 0


# --------------------------------------------------------------------------- readiness

def compute_readiness(report: Report) -> Readiness:
    """READY_FOR_UAT_EXECUTION only when nothing is gated, pending, held,
    provisional, unvalidated or unreconciled; NOT_READY otherwise. Reasons
    carry counts and headers, never source values (they are logged)."""
    reasons: list[dict] = []

    def add(category: str, item: str, q: str, owner: str, detail: str) -> None:
        reasons.append({"category": category, "item": item, "q": q, "owner": owner, "detail": detail})

    rules = report.rules
    for rule in rules.rules.values():
        n = len(report.rule_rows.get(rule.id, ()))
        if n and not rule.approved_for_import:
            effect = ("conversion performed for diagnostics only; results are provisional"
                      if rule.id in cfg.DIAGNOSTIC_RULES else "rule not applied")
            add("gated_rule", rule.id, rule.q, rule.owner,
                f"status {rule.status}, not approved_for_import; {n} source row(s) affected; {effect}")

    if report.provisional:
        rule = rules.rule(cfg.RULE_COUNTRY_ISO3)
        for (header, _text, iso), n in sorted(report.provisional.items(), key=str):
            add("provisional_conversion", header, rule.q, rule.owner,
                f"provisional conversion to {iso} ({n}x); see Rules sheet")

    for c in report.columns:
        where = f"source {c.letter} ({c.topia or c.ui})"
        if "HELD" in c.flags and c.has_data:
            rule = rules.rule(c.rule)
            add("held_column", where, rule.q, rule.owner, c.note)
        if "PENDING" in c.flags and c.has_data:
            q, owner = cfg.READINESS_ITEMS["pending_column"]
            add("pending_column", where, q, owner, f"yellow (pending clarification) column holds data; status {c.status}")
        if "MISSING_TARGET" in c.flags and c.has_data:
            q, owner = cfg.READINESS_ITEMS["missing_target"]
            add("missing_target", where, q, owner, "cyan custom field holds data but has no Equus column; not migrated")
        if "UNKNOWN_COLOUR" in c.flags:
            q, owner = cfg.READINESS_ITEMS["unknown_colour"]
            add("unknown_colour", where, q, owner, f"fill {c.colour} has no agreed meaning")

    seen: dict = defaultdict(set)
    unknown: dict = defaultdict(set)
    for (list_id, header, value) in report.reference_values:
        lst = report.reference_data[list_id]
        seen[list_id].add((header, value))
        if lst.usable and value not in lst.values:
            unknown[list_id].add((header, value))
    for list_id in sorted(seen):
        lst = report.reference_data[list_id]
        headers = ", ".join(sorted({h for h, _ in seen[list_id]}))
        if not lst.usable:
            add("reference_list_unavailable", list_id, lst.q, lst.owner,
                f"list {lst.state}: {len(seen[list_id])} distinct value(s) in {headers} cannot be validated")
        elif unknown[list_id]:
            add("reference_value_unknown", list_id, lst.q, lst.owner,
                f"{len(unknown[list_id])} distinct value(s) in {headers} not in the approved list")

    if report.state_values:  # UAT-06, Q12: states against the country_state list
        lst = report.reference_data.get(cfg.STATE_LIST_ID)
        q, owner = (lst.q, lst.owner) if lst is not None else ("Q12", "Jake (Equus)")
        states = {(h, c, s) for (h, c, s) in report.state_values}
        headers = ", ".join(sorted({h for h, _, _ in states}))
        bad = [k for k in states if report.state_in_list(k[1], k[2]) is False]
        if lst is None or not lst.usable:
            add("reference_list_unavailable", f"{cfg.STATE_LIST_ID} (state)", q, owner,
                f"list {lst.state if lst is not None else 'not configured'}: {len(states)} distinct state value(s) "
                f"in {headers} cannot be validated; sent unchanged")
        elif bad:
            add("reference_value_unknown", f"{cfg.STATE_LIST_ID} (state)", q, owner,
                f"{len(bad)} distinct state value(s) in {headers} not in the approved list for their country; "
                "sent unchanged")

    if report.unmapped_resources:
        lst = report.reference_data.get(cfg.RESOURCE_LIST_ID)
        q, owner = (lst.q, lst.owner) if lst is not None else cfg.READINESS_ITEMS["dropdown"]
        state = lst.state if lst is not None else "not configured"
        add("reference_lookup_missing", cfg.RESOURCE_LIST_ID, q, owner,
            f"list {state}: {len(report.unmapped_resources)} distinct resource value(s) have no approved Equus username")

    lookup_lists = [lst for lst in report.reference_data.values() if lst.kind == "lookup"]
    q, owner = cfg.READINESS_ITEMS["dropdown"]
    for header in sorted({h for h, _ in report.dropdown_values}):
        if any(lst.applies(header) for lst in lookup_lists):
            continue  # resource fields: covered by their lookup list above
        values = [v for (h, v) in report.dropdown_values if h == header]
        if header not in report.allowed_values:
            add("dropdown_list_missing", header, q, owner,
                f"{len(values)} distinct value(s) sent with no Equus list to validate them")
        else:
            bad = [v for v in values if report.allowed(header, v) == "NO"]
            if bad:
                add("dropdown_value_not_allowed", header, q, owner, f"{len(bad)} distinct value(s) not in the Equus list")

    # Identity / duplicate / header / structure findings (UAT-03). Any held
    # row or FILE finding makes the run NOT_READY.
    row_counts: Counter = Counter()
    row_specs: dict = {}
    for e in report.exceptions:
        if e["code"] == cfg.EXC_VALIDATION_WARNING:
            continue  # value-level warnings; listed in Exceptions / Warnings
        if e["code"] in cfg.ROW_CODES:
            row_counts[e["code"]] += 1
            row_specs[e["code"]] = e
        elif e["level"] == cfg.LEVEL_FILE:
            add("file_rejected", e["code"], e["q"], e["owner"], f"{e['message']} (FILE level)")
        elif e["code"] == cfg.EXC_HEADER_DECISION_PENDING:
            add("pending_header_decision", e["code"], e["q"], e["owner"], e["message"])
        else:
            add("template_version", e["code"], e["q"], e["owner"], e["message"])
    for code, n in row_counts.items():
        e = row_specs[code]
        add("row_held" if e["level"] == cfg.LEVEL_ROW_HELD else "file_rejected", code, e["q"], e["owner"],
            f"{n} source row(s) {code}; held out of the import workbook (level {e['level']}), "
            "see Exceptions and Held Rows")

    for e in report.dependent_exceptions:
        add("held_dependents", e["code"], e["q"], e["owner"],
            f"sheet {e['sheet']!r}: {e['rows']} dependent row(s), {e['groups']} dependent record(s) held; {e['detail']}")
    if report.dependent_overlaps:
        q, owner, _ = cfg.DEPENDENT_HOLDS["HELD_DEPENDENTS_NOT_CONFIGURED"]
        add("dependent_overlap", cfg.FAMILY_PREFIXES[1], q, owner,
            f"{report.dependent_overlaps} assignment(s) have dependents from both the assignment sheet and the "
            "keyed dependent source; review for duplicates")
    q, owner = cfg.FAMILY_LIST_OWNER
    for header in sorted({h for h, _ in report.family_values}):
        values = [v for (h, v) in report.family_values if h == header]
        not_in = [v for v in values if report.family_value_status(header, v) != "in list"]
        if not_in:
            state = "no Equus list" if header not in report.allowed_values else "Equus list"
            add("family_value_not_in_list", header, q, owner,
                f"{len(not_in)} of {len(values)} distinct value(s) not in list ({state}); sent unchanged")

    q, owner = cfg.READINESS_ITEMS["template_configuration"]
    for path, msg in report.config_issues:
        add("configuration", path, q, owner, msg)

    if not report.reconciliation_ok:
        q, owner = cfg.READINESS_ITEMS["reconciliation"]
        reasons_text = "; ".join(report.reconciliation_summary.get("reasons", []))
        add("reconciliation", "Reconciliation", q, owner, "source-to-output reconciliation did not balance"
            + (f": {reasons_text}" if reasons_text else ""))

    return Readiness(status=NOT_READY if reasons else READY, reasons=reasons)


# --------------------------------------------------------------------------- report

def build_report(report: Report, layout: TargetLayout) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Column Mapping"
    ws.append(["Source Col", "Colour", "Flags", "Row 2 (Equus technical)", "Row 3 (Equus UI)", "Row 4 (Topia)", "Status", "Equus Header", "Equus Col", "Note"])
    for c in report.columns:
        target_col = get_column_letter(layout.headers[c.target]) if c.target in layout.headers else None
        ws.append([c.letter, c.colour, ", ".join(c.flags), c.key, c.ui, c.topia, c.status, c.target, target_col, c.note])

    ws = wb.create_sheet("Governance")
    ws.append(["Source Col", "Row 4 (Topia)", "Colour", "Flag", "Has Data", "Status", "Equus Header"])
    for c in report.columns:
        for flag in c.flags:
            ws.append([c.letter, c.topia, c.colour, flag, "yes" if c.has_data else "no", c.status, c.target])
    ws.append([])
    ws.append(["Colour", "Columns"])
    for colour, n in sorted(Counter(c.colour for c in report.columns).items()):
        ws.append([colour, n])

    ws = wb.create_sheet("Equus Coverage")
    ws.append(["Equus Col", "Equus Header", "Row 1 Rule", "Cells Populated", "Source Columns"])
    sources = defaultdict(list)
    for c in report.columns:
        if c.target:
            sources[c.target].append(c.letter)
    for header, idx in sorted(layout.headers.items(), key=lambda kv: kv[1]):
        ws.append([get_column_letter(idx), header, layout.rules.get(header), report.coverage.get(header, 0), ", ".join(sources.get(header, []))])

    ws = wb.create_sheet("Dropdown Values")
    ws.append(["Equus Header", "Value Sent", "Occurrences", "Allowed?"])
    for (header, value), n in sorted(report.dropdown_values.items(), key=lambda kv: (kv[0][0], str(kv[0][1]))):
        ws.append([header, value, n, report.allowed(header, value)])

    ws = wb.create_sheet("Reconciliation")
    ws.append(RECONCILIATION_HEADER)
    for r in report.reconciliation:
        ws.append([r.get(k) for k in RECONCILIATION_KEYS])
    ws.append([])
    summary = report.reconciliation_summary
    for label, value in [("Result", "OK" if report.reconciliation_ok else "FAILED"),
                         ("Source data rows", report.data_rows),
                         ("Assignments (INSERT rows)", report.assignments),
                         ("Held source rows", len(report.held_rows)),
                         ("Rows written", report.rows_written),
                         ("Source SHA-256", report.source_sha256)]:
        ws.append([label, value])
    for key, value in summary.items():
        if key in ("result", "reasons"):
            continue
        ws.append([key, json.dumps(value) if isinstance(value, dict) else value])
    for reason in summary.get("reasons", []):
        ws.append(["reason", reason])

    ws = wb.create_sheet("Reconciliation Records")
    ws.append(RECORDS_HEADER)
    for r in report.reconciliation_records:
        ws.append([r.get(k) for k in RECORDS_KEYS])
        ws.cell(ws.max_row, 2).number_format = "@"

    _rules_sheet(wb.create_sheet("Rules"), report)
    _family_sheet(wb.create_sheet("Family"), report)
    _readiness_sheet(wb.create_sheet("Readiness"), report)
    _exceptions_sheet(wb.create_sheet("Exceptions"), report)
    _held_rows_sheet(wb.create_sheet("Held Rows"), report)
    _accepted_sheet(wb.create_sheet(ACCEPTED_SHEET), report)
    _template_sheet(wb.create_sheet("Template Check"), report)

    ws = wb.create_sheet("Warnings")
    ws.append(["Where", "Field", "Message"])
    for w in report.warnings:
        ws.append(list(w))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# The first eleven columns keep the pre-UAT-05 layout of the Reconciliation sheet.
RECONCILIATION_COLUMNS = [
    ("Source Col", "col"), ("Row 4 (Topia)", "topia"), ("Equus Header", "header"), ("Source Non-blank", "source"),
    ("Written", "written"), ("Blank Token", "blank_token"), ("Failed Conversion", "failed_conversion"),
    ("Set By Rule", "set_by_rule"), ("Unexplained", "unexplained"), ("Source Date Total", "source_date_total"),
    ("Written Date Total", "written_date_total"), ("Group", "group"), ("Colour", "colour"), ("Flags", "flags"),
    ("Provisional Conversion", "provisional_conversion"), ("Gated Rule", "gated_rule"), ("Held Row", "held_row"),
    ("Excluded Colour", "excluded_colour"), ("Pending Target", "pending_target"),
    ("Replaced By Rule", "replaced_by_rule"), ("Q", "q"), ("Owner", "owner"), ("Reason", "reason"),
]
RECONCILIATION_HEADER = [h for h, _ in RECONCILIATION_COLUMNS]
RECONCILIATION_KEYS = [k for _, k in RECONCILIATION_COLUMNS]
RECORDS_COLUMNS = [
    ("Source Row", "source_row"), ("Assignment Identifier", "identifier"), ("Employee Number", "employee"),
    ("Status", "status"), ("Codes", "codes"), ("Output Rows", "output_rows"), ("INSERT Rows", "insert_rows"),
    ("UPDATE Rows", "update_rows"), ("Expected Expansion Rows", "expected_expansion_rows"),
    ("Addresses Expected", "address_expected"), ("Addresses Written", "address_written"),
    ("Contacts Expected", "contact_expected"), ("Contacts Written", "contact_written"),
    ("Spouse Expected", "spouse_expected"), ("Spouse Written", "spouse_written"),
    ("Dependents Expected", "dependent_expected"), ("Dependents Written", "dependent_written"),
    ("Unexplained", "unexplained"), ("Notes", "notes"),
]
RECORDS_HEADER = [h for h, _ in RECORDS_COLUMNS]
RECORDS_KEYS = [k for _, k in RECORDS_COLUMNS]

RULES_HEADER = ["Rule ID", "Q", "Title", "Status", "Owner", "Approved For Import", "Applied?", "Rows Affected", "Notes"]
READINESS_HEADER = ["Category", "Item", "Q", "Owner", "Detail"]
EXCEPTIONS_HEADER = ["Source Row", "Assignment Identifier", "Employee Number", "Code", "Level", "Message", "Owner",
                     "Q"]
ACCEPTED_SHEET = "Accepted Identifiers"
ACCEPTED_HEADER = ["Source Row", "Assignment Identifier", "Employee Number"]


def _exceptions_sheet(ws, report: Report) -> None:
    ws.append(EXCEPTIONS_HEADER)
    for e in report.exceptions:
        ws.append([e["source_row"], e["identifier"], e["employee"], e["code"], e["level"], e["message"],
                   e["owner"], e["q"]])


def _held_rows_sheet(ws, report: Report) -> None:
    """Source data of every held row, in source column order, so a corrected
    batch can be built from it (docs/uat/REPROCESSING_PROCEDURE.md)."""
    ws.append(["Source Row", "Codes"] + [letter for letter, _ in report.source_headers])
    ws.append(["(source column)", "(row 4 / 3 / 2 label)"] + [label for _, label in report.source_headers])
    for r in sorted(report.held_rows):
        cells = list(report.held_row_values.get(r, []))  # raw values, unchanged
        ws.append([r, ", ".join(report.held_rows[r])] + cells)
        for value, cell in zip(cells, ws[ws.max_row][2:]):
            if isinstance(value, str):
                cell.number_format = "@"
            elif isinstance(value, (date, datetime)):
                cell.number_format = DATE_FORMAT


def _accepted_sheet(ws, report: Report) -> None:
    ws.append(ACCEPTED_HEADER)
    for a in report.accepted:
        cell_row = [a["row"], a["identifier"], a["employee"]]
        ws.append(cell_row)
        ws.cell(ws.max_row, 2).number_format = "@"


def _template_sheet(ws, report: Report) -> None:
    tc = report.template_check
    for label, value in [("Template version", tc.get("template_version")),
                         ("Approved template SHA-256", tc.get("approved_sha256")),
                         ("Supplied template SHA-256", tc.get("supplied_sha256")),
                         ("Header rows compared", ", ".join(str(r) for r in tc.get("header_rows", ()))),
                         ("Result", tc.get("result")),
                         ("File status", "FAILED - import workbook withheld" if report.file_failed else "ACCEPTED")]:
        ws.append([label, value])
    ws.append([])
    ws.append(["Differing Cell", "Approved", "Found"])
    for d in tc.get("diffs", []):
        ws.append([d["cell"], d["approved"], d["found"]])
    ws.append([])
    ws.append(["Pending Header Decision", "Cell", "Approved Template", "Output", "Alternative", "Alternative Source",
               "Status"])
    for d in tc.get("decisions", []):
        ws.append([d["id"], d["cell"], d["template_value"], d["output_value"], d["alternative_value"],
                   d["alternative_source"], "PENDING (owner Jake, Q5) - header not changed"])


def accepted_identifiers(report_bytes: bytes) -> set[str]:
    """Identifiers accepted by an earlier run, read from its report workbook
    (sheet 'Accepted Identifiers'); pass them as `previously_accepted` to
    transform() for a controlled re-run (docs/uat/REPROCESSING_PROCEDURE.md)."""
    ws = openpyxl.load_workbook(io.BytesIO(report_bytes), read_only=True)[ACCEPTED_SHEET]
    rows = ws.iter_rows(min_row=1, values_only=True)
    header = next(rows, None)
    if header is None or list(header[:len(ACCEPTED_HEADER)]) != ACCEPTED_HEADER:
        raise ValueError(f"sheet {ACCEPTED_SHEET!r} does not have the expected header {ACCEPTED_HEADER}")
    return {identifier_text(r[1]) for r in rows if r and identifier_text(r[1]) is not None}


def rule_applied(report: Report, rule: cfg.BusinessRule) -> str:
    n = len(report.rule_rows.get(rule.id, ()))
    if rule.approved_for_import:
        return "yes" if n else "yes (no rows affected)"
    if rule.id in cfg.DIAGNOSTIC_RULES and n:
        return "provisional (diagnostic only)"
    return "no (gated)"


def _rules_sheet(ws, report: Report) -> None:
    rules = report.rules
    for label, value in [("Business rules file", rules.file_name), ("Version", rules.version),
                         ("Effective version", rules.effective_version), ("Generated at", rules.generated_at)]:
        ws.append([label, value])
    ws.append([])
    ws.append(RULES_HEADER)
    for rule in rules.rules.values():
        ws.append([rule.id, rule.q, rule.title, rule.status, rule.owner, "yes" if rule.approved_for_import else "no",
                   rule_applied(report, rule), len(report.rule_rows.get(rule.id, ())), rule.notes])
    ws.append([])
    ws.append(["Reference List", "Q", "Status", "Owner", "Approved For Import", "Version", "Values", "State",
               "Applies To"])
    for lst in report.reference_data.values():
        ws.append([lst.list_id, lst.q, lst.status, lst.owner, "yes" if lst.approved_for_import else "no",
                   lst.version, len(lst.values), lst.state, ", ".join(lst.applies_to)])
    ws.append([])
    ws.append(["Provisional Conversion", "Equus Header", "Source Value", "Converted", "Occurrences"])
    for (header, text, iso), n in sorted(report.provisional.items(), key=str):
        ws.append(["provisional conversion", header, text, iso, n])


FAMILY_EXCEPTION_HEADER = ["Code", "Sheet", "Dependent Rows", "Dependent Records", "Q", "Owner", "Detail",
                           "Source Rows"]


def _family_sheet(ws, report: Report) -> None:
    """Spouse/dependent evidence: dependent source, held dependents,
    DATA_QUALITY notes (Q10/Q20) and relationship values (Q14)."""
    dep = report.rules.dependent_source
    for label, value in [
        ("Dependent source sheet", (dep.sheet or "not configured") if dep else "section absent"),
        ("Dependent source status", f"{dep.status}, approved_for_import {'yes' if dep.approved_for_import else 'no'}, "
                                    f"{dep.q}, owner {dep.owner}, version {dep.version}" if dep else ""),
        ("Join key", f"Assignment ID ({dep.assignment_label!r}); Employee ID ({dep.employee_label!r}) "
                     "cross-check only" if dep else ""),
        ("Dependents attached from keyed source", report.dependents_attached),
    ]:
        ws.append([label, value])
    ws.append([])
    ws.append(FAMILY_EXCEPTION_HEADER)
    for e in report.dependent_exceptions:
        ws.append([e["code"], e["sheet"], e["rows"], e["groups"], e["q"], e["owner"], e["detail"], e["source_rows"]])
    ws.append([])
    ws.append(["DATA_QUALITY", "Where", "Record", "Populated Fields", "Missing"])
    for n in report.data_quality:
        ws.append([cfg.FAMILY_NAME_NOTE, n["where"], n["record"], n["fields"], n["missing"]])
    ws.append([])
    ws.append(["Family Value", "Equus Header", "Value Sent", "Occurrences", "List Check"])
    for (header, value), n in sorted(report.family_values.items(), key=str):
        ws.append(["relationship", header, value, n, report.family_value_status(header, value)])


def _readiness_sheet(ws, report: Report) -> None:
    readiness = report.readiness
    tc = report.template_check
    for label, value in [("Status", readiness.status),
                         ("Business rules version", report.business_rules_version),
                         ("Reasons", len(readiness.reasons)),
                         ("Equus UAT execution", "INCOMPLETE - not executed from this environment; the best "
                                                 f"status this tool can compute is {READY}"),
                         ("File status", "FAILED - import workbook withheld" if report.file_failed else "ACCEPTED"),
                         ("Rows accepted", len(report.accepted)),
                         ("Rows held", len(report.held_rows)),
                         ("TEMPLATE_VERSION", tc.get("template_version")),
                         ("Template SHA-256 (supplied)", tc.get("supplied_sha256"))]:
        ws.append([label, value])
    ws.append([])
    ws.append(READINESS_HEADER)
    for r in readiness.reasons:
        ws.append([r["category"], r["item"], r["q"], r["owner"], r["detail"]])


def load_allowed_values() -> dict[str, list[str]] | None:
    path = Path(__file__).parent / cfg.ALLOWED_VALUES_FILE
    return json.loads(path.read_text()) if path.exists() else None


def transform(source_bytes: bytes, template_bytes: bytes, strict_pending: bool = False,
              allowed_values: dict[str, list[str]] | None = None,
              business_rules: Any = None, reference_data: Any = None,
              previously_accepted: Any = None) -> tuple[bytes, bytes, Report]:
    """Return (equus_workbook_bytes, report_workbook_bytes, report).

    The import workbook holds accepted rows only; held rows are listed in the
    report's Exceptions and Held Rows sheets. When a FILE-level finding exists
    (report.file_failed, e.g. header mismatch with the approved template) the
    import workbook is withheld: equus_workbook_bytes is b"" and only the
    report is produced. previously_accepted: identifiers accepted by an
    earlier run (see accepted_identifiers()); such rows are held as
    ALREADY_ACCEPTED so a corrected re-run imports only previously held rows.

    strict_pending: refuse to build the output while yellow (pending) columns
    hold data. Independent of report.readiness.
    allowed_values: Equus dropdown lists; defaults to ALLOWED_VALUES_FILE if bundled.
    business_rules: path to a business_rules.json (or a loaded BusinessRules);
    defaults to the shipped file. reference_data: directory of reference lists
    (or a loaded dict); defaults to the shipped reference_data/.
    Invalid configuration raises mapping_config.ConfigurationError before any
    workbook is read."""
    rules = business_rules if isinstance(business_rules, cfg.BusinessRules) else cfg.load_business_rules(business_rules)
    refs = reference_data if isinstance(reference_data, dict) else cfg.load_reference_data(reference_data)

    try:
        src_wb = openpyxl.load_workbook(io.BytesIO(source_bytes), data_only=True)
    except Exception as exc:  # zipfile.BadZipFile / openpyxl InvalidFileException
        raise SourceError(f"source is not a readable .xlsx workbook ({type(exc).__name__})") from exc
    if cfg.SOURCE_SHEET not in src_wb.sheetnames:
        raise SourceError(f"Source sheet {cfg.SOURCE_SHEET!r} not found; got {src_wb.sheetnames}")
    src_ws = src_wb[cfg.SOURCE_SHEET]

    try:
        tpl_wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    except Exception as exc:  # zipfile.BadZipFile / openpyxl InvalidFileException
        raise TemplateError(f"Equus template is not a readable .xlsx workbook ({type(exc).__name__})") from exc
    layout = check_template_structure(tpl_wb)

    report = Report(source_sha256=hashlib.sha256(source_bytes).hexdigest(), rules=rules, reference_data=refs)
    report.previously_accepted = {t for t in (identifier_text(v) for v in (previously_accepted or ())) if t}
    if allowed_values is None:
        allowed_values = load_allowed_values() or {}
    report.allowed_values = {h: {str(v).lower() for v in vals} for h, vals in allowed_values.items()}
    check_template(layout, report)
    check_configuration(layout, report)
    report.columns = build_column_plan(src_ws, layout, rules)
    pending = [c for c in report.columns if "PENDING" in c.flags and c.has_data]
    if strict_pending and pending:
        raise PendingFieldsError("Pending (yellow) columns hold data: " +
                                 ", ".join(f"{c.letter} {c.topia or c.ui}" for c in pending))
    keyed_dependents = read_dependent_sources(src_wb, layout, report)
    rows = build_records(src_ws, report.columns, layout, report, keyed_dependents)
    validate_row_structure(rows, report)
    report.summarise_values()
    out = write_output(template_bytes, rows, cfg.TARGET_SHEET, report)
    check_template_headers(out, template_bytes, report)
    reconcile(src_ws, report.columns, out, report)
    for where, fld, msg in report.warnings:  # every finding carries a code and a level
        report.add_exception(cfg.EXC_VALIDATION_WARNING, f"{fld}: {msg}", where)
    report.readiness = compute_readiness(report)
    report_bytes = build_report(report, layout)
    if report.file_failed:
        out = b""  # whole file withheld (FILE level); the report keeps every diagnostic
    return out, report_bytes, report
