"""Transform a Topia Full Historical Past Assignment export into the Equus
EAFH import layout. Pure functions: bytes in, bytes out, no AWS calls.

Business rules are read from the versioned configuration
(mapping_config.load_business_rules / load_reference_data). A rule changes
output only while its entry is approved_for_import; otherwise the affected
rows are counted and reported in the Rules and Readiness sheets."""

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
    reconciliation: list[dict] = field(default_factory=list)
    reconciliation_ok: bool = False
    # Configuration in force and what the rules did.
    rules: Any = None  # mapping_config.BusinessRules
    reference_data: dict = field(default_factory=dict)  # list id -> mapping_config.ReferenceList
    rule_rows: dict = field(default_factory=lambda: defaultdict(set))  # rule id -> source rows affected
    provisional: Counter = field(default_factory=Counter)  # (header, source text, converted) -> n
    reference_values: Counter = field(default_factory=Counter)  # (list id, header, value) -> n
    config_issues: list[tuple[str, str]] = field(default_factory=list)  # (config path, message)
    identifier_fallbacks: int = 0
    missing_employee_numbers: int = 0
    current_row: Any = None
    readiness: Readiness | None = None
    # Spouse / dependent records (Q10, Q11, Q18, Q19, Q20).
    data_quality: list[dict] = field(default_factory=list)  # where, record, fields, missing
    family_values: Counter = field(default_factory=Counter)  # (header, value) -> n
    dependent_exceptions: list[dict] = field(default_factory=list)  # code, sheet, rows, groups, q, owner, detail
    dependents_attached: int = 0  # child records attached from the keyed dependent source
    dependent_overlaps: int = 0  # assignments with dependents from the main sheet and the keyed source

    @property
    def business_rules_version(self) -> str:
        return self.rules.version if self.rules is not None else ""

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


# --------------------------------------------------------------------------- layout

def read_target_layout(ws) -> TargetLayout:
    headers, rules = {}, {}
    for c in range(1, ws.max_column + 1):
        raw = ws.cell(cfg.TARGET_HEADER_ROW, c).value
        if raw is None:
            continue
        header = norm_key(raw) or str(raw).strip()
        if header in headers:
            raise ValueError(f"Duplicate Equus header {header!r} in template")
        headers[header] = c
        rules[header] = norm_label(ws.cell(cfg.TARGET_RULE_ROW, c).value)
    return TargetLayout(headers, rules)


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


def build_records(src_ws, cols: list[SourceColumn], layout: TargetLayout, report: Report,
                  keyed_dependents: dict[str, list[dict]] | None = None) -> list[dict]:
    mapped = [c for c in cols if c.status == "mapped"]
    addresses = [c for c in cols if c.status == "address"]
    contacts = [c for c in cols if c.status == "contact"]
    dependents = _group_dependents(cols)
    cancel_col = next((c for c in cols if c.rule == cfg.RULE_CANCELLATION_PHASE), None)
    override_cols = [c for c in cols if c.rule in cfg.COLUMN_OVERRIDE_RULES]
    ident_col = next((c for c in cols if c.label == cfg.ASSIGNMENT_IDENTIFIER_LABEL), None)
    start_header = report.rules.params(cfg.RULE_DEFAULT_DATES)["start_header"]

    out_rows: list[dict] = []
    seen_employees: set = set()
    keyed_dependents = keyed_dependents or {}
    # Assignment IDs that occur on more than one assignment row cannot take
    # keyed dependents unambiguously (Q1/Q18).
    id_counts: Counter = Counter()
    if ident_col is not None and keyed_dependents:
        for r in range(cfg.SOURCE_DATA_START_ROW, src_ws.max_row + 1):
            if not all(is_blank(src_ws.cell(r, c.idx).value) for c in cols):
                id_counts[_key_text(src_ws.cell(r, ident_col.idx).value)] += 1
    consumed: set = set()

    for r in range(cfg.SOURCE_DATA_START_ROW, src_ws.max_row + 1):
        raw = {c.idx: src_ws.cell(r, c.idx).value for c in cols}
        if all(is_blank(v) for v in raw.values()):
            continue
        report.current_row = r
        row_ref = f"source row {r}"
        main = {}
        for col in mapped:
            main[col.target] = convert(col.target, raw[col.idx], layout, report, f"source {col.letter}{r}")

        employee = main.get(cfg.COL_EMPLOYEE_NUMBER)
        identifier = convert("ASSIGNMENT_CF!CF_TEXT029", raw[ident_col.idx], layout, report, f"source {ident_col.letter}{r}") if ident_col else None
        # Keyed dependents join only on this row's own Assignment ID, never on
        # the Employee Number fallback below.
        join_key = _key_text(identifier)
        if identifier is None:
            identifier = employee
            report.identifier_fallbacks += 1
            report.warn(row_ref, cfg.COL_ASSIGNMENT_IDENTIFIER, "no Topia Assignment ID; using Employee Number")
        if employee is None:
            report.missing_employee_numbers += 1
            report.warn(row_ref, cfg.COL_EMPLOYEE_NUMBER, "missing Employee Number (Talent Link ID)")

        # Business rules (configuration-gated)
        apply_assignment_rules(main, raw[cancel_col.idx] if cancel_col else None, report)
        for col in override_cols:
            if not is_blank(raw[col.idx]):
                report.rule_hit(col.rule)
        start = main.get(start_header)
        validate_record(main, report, row_ref)

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
        for group in dependents:
            rec = {c.target: convert(c.target, raw[c.idx], layout, report, f"source {c.letter}{r}") for c in group}
            if any(v is not None for v in rec.values()):
                validate_family(rec, cfg.FAMILY_PREFIXES[1], report, row_ref)
                dep_recs.append(rec)
        dep_recs.extend(_attach_keyed_dependents(join_key, employee, keyed_dependents, id_counts, consumed,
                                                 bool(dep_recs), report, row_ref))

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

    unmatched = [e for key, entries in keyed_dependents.items() if key not in consumed for e in entries]
    if unmatched:
        sheet = report.rules.dependent_source.sheet
        report.hold_dependents("HELD_DEPENDENTS_UNMATCHED", sheet, [f"row {e['row']}" for e in unmatched],
                               sum(len(e["records"]) for e in unmatched))
    return out_rows


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
            report.coverage[header] += 1
    report.rows_written = len(rows)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def reconcile(src_ws, cols: list[SourceColumn], output: bytes, report: Report) -> None:
    """Compare every mapped source cell with the cell actually written to the
    output workbook. Each difference must have a known reason; anything else
    is counted as unexplained and fails the run."""
    out_ws = openpyxl.load_workbook(io.BytesIO(output))[cfg.TARGET_SHEET]
    out_headers = {out_ws.cell(cfg.TARGET_HEADER_ROW, c).value: c for c in range(1, out_ws.max_column + 1)}
    command_col = out_headers[cfg.COL_ASSIGNMENT_COMMAND]
    insert_rows = [r for r in range(cfg.TARGET_DATA_START_ROW, out_ws.max_row + 1)
                   if out_ws.cell(r, command_col).value == "INSERT"]
    source_rows = [r for r in range(cfg.SOURCE_DATA_START_ROW, src_ws.max_row + 1)
                   if not all(is_blank(src_ws.cell(r, c.idx).value) for c in cols)]
    report.data_rows = len(source_rows)
    counts_ok = len(source_rows) == len(insert_rows) == report.assignments
    rule_targets = report.rules.applied_targets()

    report.reconciliation = []
    for col in (c for c in cols if c.status == "mapped"):
        out_c = out_headers[col.target]
        row = {"col": col.letter, "topia": col.topia or col.ui, "header": col.target, "source": 0, "written": 0,
               "blank_token": 0, "failed_conversion": 0, "set_by_rule": 0, "unexplained": 0,
               "source_date_total": 0, "written_date_total": 0}
        is_date_col = False
        for sr, orow in zip(source_rows, insert_rows):
            raw, out = src_ws.cell(sr, col.idx).value, out_ws.cell(orow, out_c).value
            present = raw is not None and str(raw).strip() != ""
            row["source"] += present
            row["written"] += out is not None
            if present and out is None:
                if is_blank(raw) or is_dirty_placeholder(raw):
                    row["blank_token"] += 1
                elif isinstance(raw, str) and to_date(raw) is None and col.target.endswith("DATE"):
                    row["failed_conversion"] += 1
                else:
                    row["unexplained"] += 1
            elif not present and out is not None:
                row["set_by_rule" if col.target in rule_targets else "unexplained"] += 1
            elif present and isinstance(out, (date, datetime)):
                is_date_col = True
                src_date, out_date = to_date(raw), out
                row["source_date_total"] += _ordinal(src_date)
                row["written_date_total"] += _ordinal(out_date)
                if _ordinal(src_date) != _ordinal(out_date):
                    row["unexplained"] += 1
        if not is_date_col:
            row["source_date_total"] = row["written_date_total"] = None
        report.reconciliation.append(row)
    report.reconciliation_ok = counts_ok and all(r["unexplained"] == 0 for r in report.reconciliation)


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

    q, owner = cfg.READINESS_ITEMS["identifier"]
    if report.identifier_fallbacks:
        add("identifier", cfg.COL_ASSIGNMENT_IDENTIFIER, q, owner,
            f"{report.identifier_fallbacks} row(s) without a Topia Assignment ID used the Employee Number")
    if report.missing_employee_numbers:
        add("identifier", cfg.COL_EMPLOYEE_NUMBER, q, owner,
            f"{report.missing_employee_numbers} row(s) without an Employee Number")

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
        add("reconciliation", "Reconciliation", q, owner, "source-to-output reconciliation did not balance")

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
    ws.append(["Source Col", "Row 4 (Topia)", "Equus Header", "Source Non-blank", "Written", "Blank Token",
               "Failed Conversion", "Set By Rule", "Unexplained", "Source Date Total", "Written Date Total"])
    for r in report.reconciliation:
        ws.append([r["col"], r["topia"], r["header"], r["source"], r["written"], r["blank_token"],
                   r["failed_conversion"], r["set_by_rule"], r["unexplained"],
                   r["source_date_total"], r["written_date_total"]])
    ws.append([])
    for label, value in [("Result", "OK" if report.reconciliation_ok else "FAILED"),
                         ("Source data rows", report.data_rows),
                         ("Assignments (INSERT rows)", report.assignments),
                         ("Rows written", report.rows_written),
                         ("Source SHA-256", report.source_sha256)]:
        ws.append([label, value])

    _rules_sheet(wb.create_sheet("Rules"), report)
    _family_sheet(wb.create_sheet("Family"), report)
    _readiness_sheet(wb.create_sheet("Readiness"), report)

    ws = wb.create_sheet("Warnings")
    ws.append(["Where", "Field", "Message"])
    for w in report.warnings:
        ws.append(list(w))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


RULES_HEADER = ["Rule ID", "Q", "Title", "Status", "Owner", "Approved For Import", "Applied?", "Rows Affected", "Notes"]
READINESS_HEADER = ["Category", "Item", "Q", "Owner", "Detail"]


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
    for label, value in [("Status", readiness.status),
                         ("Business rules version", report.business_rules_version),
                         ("Reasons", len(readiness.reasons)),
                         ("Equus UAT execution", "INCOMPLETE - not executed from this environment; the best "
                                                 f"status this tool can compute is {READY}")]:
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
              business_rules: Any = None, reference_data: Any = None) -> tuple[bytes, bytes, Report]:
    """Return (equus_workbook_bytes, report_workbook_bytes, report).

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

    src_wb = openpyxl.load_workbook(io.BytesIO(source_bytes), data_only=True)
    if cfg.SOURCE_SHEET not in src_wb.sheetnames:
        raise ValueError(f"Source sheet {cfg.SOURCE_SHEET!r} not found; got {src_wb.sheetnames}")
    src_ws = src_wb[cfg.SOURCE_SHEET]

    tpl_wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    layout = read_target_layout(tpl_wb[cfg.TARGET_SHEET])

    report = Report(source_sha256=hashlib.sha256(source_bytes).hexdigest(), rules=rules, reference_data=refs)
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
    report.summarise_values()
    out = write_output(template_bytes, rows, cfg.TARGET_SHEET, report)
    reconcile(src_ws, report.columns, out, report)
    report.readiness = compute_readiness(report)
    return out, build_report(report, layout), report
