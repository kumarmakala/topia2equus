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

    def summarise_values(self) -> None:
        """One warning per distinct unmapped resource / disallowed dropdown value."""
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
    values = {k: v for k, v in rec.items() if k.split("!")[0] in (prefix, prefix + "_CF") and v is not None}
    if values and not (values.get(prefix + "!FIRST_NAME") and values.get(prefix + "!LAST_NAME")):
        report.warn(row_ref, prefix, f"record has {', '.join(sorted(k.split('!')[1] for k in values))} "
                                     "but no first/last name; Equus may reject or create an unnamed person")


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


def build_records(src_ws, cols: list[SourceColumn], layout: TargetLayout, report: Report) -> list[dict]:
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
    return out_rows


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
    rows = build_records(src_ws, report.columns, layout, report)
    report.summarise_values()
    out = write_output(template_bytes, rows, cfg.TARGET_SHEET, report)
    reconcile(src_ws, report.columns, out, report)
    report.readiness = compute_readiness(report)
    return out, build_report(report, layout), report
