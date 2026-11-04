"""Transform a Topia Source export into the Equus EAFH import layout.

Pure functions: bytes in, bytes out, no AWS calls. The mapping rules themselves
live in mapping_config.py - read the "HOW THE MAPPING WORKS" note there first.

The pipeline is five small steps, each a function you can read on its own:

    1. read_target_layout   learn the Equus columns from the template (row 3)
    2. build_source_index   read the Topia column names (row 1 of the source)
    3. plan_source_columns  decide what each Topia column becomes in Equus
    4. build_output_rows    walk the source rows and build the output records
    5. write_output         drop the records into a copy of the Equus template

build_report then writes a human-readable mapping report alongside the output.
Nothing depends on a fixed Excel column position, so reordered / extra / missing
/ duplicate columns are all handled and reported rather than silently mishandled.
"""

from __future__ import annotations

import io
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter

import mapping_config as cfg

KEY_RE = re.compile(r"^[A-Z0-9_]+![A-Z0-9_]+$")
DATE_FORMAT = "MM/DD/YYYY"
_DATE_PATTERNS = ("%m/%d/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y",
                  "%Y-%m-%d %H:%M:%S", "%m/%d/%y", "%d/%m/%Y")

ADDR = "ASSIGNMENT_MAILING_ADDRESS!"
CONTACT = "ASSIGNMENT_EMPLOYEE_CONTACT!"


# --------------------------------------------------------------------------- header helpers

def norm_label(value: Any) -> str:
    """Normalise a header for lookup: collapse whitespace, strip, lower-case."""
    return re.sub(r"\s+", " ", str(value)).strip().lower() if value is not None else ""


def canon(value: Any) -> str:
    """Normalise then apply documented source aliases (exact renames only)."""
    n = norm_label(value)
    return cfg.SOURCE_ALIASES.get(n, n)


def target_key(value: Any) -> str:
    """Row-3 target header: keep TABLE!FIELD keys verbatim, else trimmed text."""
    if isinstance(value, str) and KEY_RE.match(value.strip().upper()):
        return value.strip().upper()
    return str(value).strip()


def is_blank(value: Any) -> bool:
    """True when a source value is empty or a documented placeholder (e.g. "N/A")."""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip().lower() in cfg.BLANK_TOKENS
    return False


# --------------------------------------------------------------------------- data classes

@dataclass
class SourceColumn:
    """One Topia column and what the plan decided to do with it."""
    idx: int
    raw: Any
    norm: str
    canon: str
    color: str | None = None
    status: str = "unmapped"   # mapped|address|contact|dependent|identity|rule|ignored|excluded|duplicate|unmapped
    target: str | None = None
    dep_index: int | None = None
    dep_target: str | None = None
    note: str = ""

    @property
    def letter(self) -> str:
        """Excel column letter for this 1-based index (for the report)."""
        return get_column_letter(self.idx)


@dataclass
class TargetLayout:
    """The Equus template's columns, keyed by row-3 DB field.

    The is_* helpers tell convert() how to clean a value, using either the
    row-1 instruction rule or the DB field's own name as a fallback.
    """
    headers: dict[str, int]          # DB header -> column index
    rules: dict[str, str]            # DB header -> lower-cased row-1 rule
    ui: dict[str, str]               # DB header -> row-2 UI label

    def is_date(self, h: str) -> bool:
        """Does this Equus field hold a date?"""
        return "date format" in self.rules.get(h, "") or h.endswith("_DATE")

    def is_country(self, h: str) -> bool:
        """Does this Equus field expect an ISO-3 country code?"""
        return ("country code" in self.rules.get(h, "")
                or h.endswith(("_COUNTRY", "_CITIZENSHIP", "COUNTRY_CODE")))

    def is_resource(self, h: str) -> bool:
        """Is this a 'resource' (person reference) field needing an Equus username?"""
        return "RESOURCE" in h

    def is_dropdown(self, h: str) -> bool:
        """Does this Equus field expect a value from a fixed dropdown list?"""
        r = self.rules.get(h, "")
        return "dropdown" in r or "yesno" in r


@dataclass
class Report:
    """Everything the mapping report and the handler summary need to know."""
    columns: list[SourceColumn] = field(default_factory=list)
    warnings: list[tuple[Any, str, str]] = field(default_factory=list)
    dropdown_values: Counter = field(default_factory=Counter)
    coverage: Counter = field(default_factory=Counter)
    missing_source: list[str] = field(default_factory=list)
    assignments: int = 0
    rows_written: int = 0

    def warn(self, row: Any, fld: str, msg: str) -> None:
        """Record a non-fatal data-quality note against a row/field."""
        self.warnings.append((row, fld, msg))


# --------------------------------------------------------------------------- step 1: target layout

def read_target_layout(ws) -> TargetLayout:
    """Read the Equus template header rows into a TargetLayout (step 1).

    Row 3 is the DB field (the key we fill), row 2 the UI label, row 1 the
    instruction hint. A repeated DB field in the template is a template bug and
    raises, because we would not know which column to fill.
    """
    headers, rules, ui = {}, {}, {}
    for c in range(1, ws.max_column + 1):
        raw = ws.cell(cfg.TARGET_HEADER_ROW, c).value
        if raw is None:
            continue
        header = target_key(raw)
        if header in headers:
            raise ValueError(f"Duplicate Equus header {header!r} in template")
        headers[header] = c
        rules[header] = norm_label(ws.cell(cfg.TARGET_RULE_ROW, c).value)
        ui[header] = str(ws.cell(cfg.TARGET_UI_ROW, c).value or "")
    return TargetLayout(headers, rules, ui)


# --------------------------------------------------------------------------- step 2: source index

def _cell_color(cell) -> str | None:
    """Return a cell's fill colour as an ARGB string, or None if unfilled."""
    try:
        if cell.fill and cell.fill.patternType:
            rgb = cell.fill.fgColor.rgb
            return rgb if isinstance(rgb, str) else None
    except Exception:
        pass
    return None


def build_source_index(ws) -> list[SourceColumn]:
    """Read row 1 of the Topia Source sheet into SourceColumn records (step 2)."""
    cols = []
    for c in range(1, ws.max_column + 1):
        raw = ws.cell(cfg.SOURCE_HEADER_ROW, c).value
        if raw is None or str(raw).strip() == "":
            continue
        cols.append(SourceColumn(idx=c, raw=raw, norm=norm_label(raw),
                                 canon=canon(raw), color=_cell_color(ws.cell(cfg.SOURCE_HEADER_ROW, c))))
    return cols


# --------------------------------------------------------------------------- step 3: column plan

def _claim_target(col: SourceColumn, tgt: str, note: str,
                  layout: TargetLayout, claimed: dict[str, SourceColumn]) -> None:
    """Point a source column at an Equus DB field, guarding two problems.

    If the field is not in the template it is "unmapped"; if another column has
    already claimed it this one is "duplicate" (first column wins). Otherwise it
    is "mapped" and the target is recorded so later columns cannot reuse it.
    """
    if tgt not in layout.headers:
        col.status, col.note = "unmapped", f"target {tgt} not in Equus template"
    elif tgt in claimed:
        col.status, col.target = "duplicate", tgt
        col.note = f"{tgt} already filled from column {claimed[tgt].letter}"
    else:
        col.status, col.target, col.note = "mapped", tgt, note
        claimed[tgt] = col


def _plan_dependent(col: SourceColumn, layout: TargetLayout) -> None:
    """Classify a "Child N <field>" column as a dependent field, or unmapped."""
    m = cfg.DEPENDENT_HEADER_RE.match(col.canon)
    dep_field = norm_label(m.group(2))
    tgt = cfg.DEPENDENT_FIELD_MAP.get(dep_field)
    if tgt and tgt in layout.headers:
        col.status, col.dep_index, col.dep_target = "dependent", int(m.group(1)), tgt
        col.note = f"dependent {m.group(1)} -> {tgt}"
    else:
        col.status = "unmapped"
        col.note = f"dependent field {dep_field!r} not in DEPENDENT_FIELD_MAP/template"


def _classify_column(col: SourceColumn, layout: TargetLayout,
                     claimed: dict[str, SourceColumn]) -> None:
    """Decide what one (non-duplicate) Topia column becomes in Equus.

    The order below IS the precedence - the first rule that matches wins, so
    the special tables (exclude / address / contact / dependent / identity) are
    checked before the plain FIELD_MAP lookup. See mapping_config for each table.
    """
    key = col.canon
    if key in cfg.EXCLUDE_FIELDS:
        col.status, col.note = "excluded", "listed in mapping_config.EXCLUDE_FIELDS"
    elif key in cfg.ADDRESS_COLUMNS:
        col.status, col.note = "address", f"LOCATION_TYPE = {cfg.ADDRESS_COLUMNS[key]}"
    elif key in cfg.CONTACT_COLUMNS:
        col.status, col.note = "contact", "LOCATION/CONTACT = %s / %s" % cfg.CONTACT_COLUMNS[key]
    elif key.startswith(cfg.IGNORED_LABEL_PREFIXES):
        col.status, col.note = "ignored", "describes a constant address location type"
    elif key == cfg.CANCELLATION_REASON_LABEL:
        col.status, col.note = "rule", f"non-empty -> ASSIGNMENT!PHASE = {cfg.CANCELLED_PHASE}"
    elif cfg.DEPENDENT_HEADER_RE.match(key):
        _plan_dependent(col, layout)
    elif key == cfg.EMPLOYEE_NUMBER_SOURCE:
        col.status, col.target = "identity", cfg.COL_EMPLOYEE_NUMBER
        col.note = "Talent Link ID -> Employee Number (control)"
        # In the rare case it also appears in FIELD_MAP, fill that field too.
        if key in cfg.FIELD_MAP:
            _claim_target(SourceColumn(col.idx, col.raw, col.norm, col.canon),
                          cfg.FIELD_MAP[key], "also mapped", layout, claimed)
    elif key in cfg.FIELD_MAP:
        _claim_target(col, cfg.FIELD_MAP[key], "FIELD_MAP", layout, claimed)
    else:
        col.status, col.note = "unmapped", "no Equus field in FIELD_MAP"


def _is_repeat_source_header(col: SourceColumn, first_seen: dict[str, SourceColumn]) -> bool:
    """True if this Topia header already appeared and is not a multi-value column.

    Address / contact / dependent headers legitimately repeat, so they are not
    treated as duplicates here.
    """
    key = col.canon
    if key not in first_seen:
        return False
    return (key not in cfg.ADDRESS_COLUMNS and key not in cfg.CONTACT_COLUMNS
            and not cfg.DEPENDENT_HEADER_RE.match(key))


def _flag_missing_required(cols: list[SourceColumn], report: Report) -> None:
    """Report any required identity source column that is absent entirely."""
    present = {c.canon for c in cols}
    for req in (cfg.EMPLOYEE_NUMBER_SOURCE, cfg.ASSIGNMENT_IDENTIFIER_SOURCE):
        if req not in present:
            report.missing_source.append(req)


def plan_source_columns(cols: list[SourceColumn], layout: TargetLayout, report: Report) -> None:
    """Classify every source column in place (step 3).

    One pass over the columns: skip repeated headers, then hand each column to
    _classify_column. `claimed` stops two columns filling the same Equus field.
    """
    claimed: dict[str, SourceColumn] = {}
    first_seen: dict[str, SourceColumn] = {}
    for col in cols:
        if _is_repeat_source_header(col, first_seen):
            col.status = "duplicate"
            col.note = f"duplicate source header of column {first_seen[col.canon].letter}"
            continue
        first_seen.setdefault(col.canon, col)
        _classify_column(col, layout, claimed)
    _flag_missing_required(cols, report)


# --------------------------------------------------------------------------- value conversion

def to_date(value: Any) -> date | datetime | None:
    """Parse any supported source date form into a date/datetime, or None."""
    if isinstance(value, datetime):
        return value.date() if value.time() == datetime.min.time() else value
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            return openpyxl.utils.datetime.from_excel(value)
        except Exception:
            return None
    text = str(value).strip()
    for fmt in _DATE_PATTERNS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def convert(header: str, value: Any, layout: TargetLayout, report: Report, row: Any) -> Any:
    """Clean one value for its Equus field: blanks, dates, countries, dropdowns.

    Unknown dropdown/country values are passed through unchanged and recorded in
    the report rather than guessed at.
    """
    if is_blank(value):
        return None
    if isinstance(value, str):
        value = value.strip()

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
        return iso

    if isinstance(value, bool):
        value = "Yes" if value else "No"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if "CF_TEXT" in header and not isinstance(value, str):
        value = str(value)

    if header in cfg.VALUE_MAPS and isinstance(value, str):
        value = cfg.VALUE_MAPS[header].get(value.lower(), value)
    if layout.is_resource(header) and isinstance(value, str):
        value = cfg.RESOURCE_MAP.get(value, value)
    if layout.is_dropdown(header) or layout.is_resource(header):
        report.dropdown_values[(header, value)] += 1
    return value


def _id_value(value: Any) -> Any:
    """Normalise an identifier for the control columns (no ISO/date logic)."""
    if is_blank(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str):
        return value.strip()
    return value


# --------------------------------------------------------------------------- step 4: build rows

def _dependent_groups(cols: list[SourceColumn]) -> dict[int, list[SourceColumn]]:
    """Group the dependent columns by child number ({1: [...], 2: [...]})."""
    groups: dict[int, list[SourceColumn]] = defaultdict(list)
    for col in cols:
        if col.status == "dependent":
            groups[col.dep_index].append(col)
    return dict(sorted(groups.items()))


def _map_main_fields(raw: dict[int, Any], mapped: list[SourceColumn],
                     layout: TargetLayout, report: Report, row_ref: str) -> dict[str, Any]:
    """Build the assignment's main field dict from every plainly-mapped column."""
    main: dict[str, Any] = {}
    for col in mapped:
        main[col.target] = convert(col.target, raw[col.idx], layout, report, row_ref)
    return main


def _resolve_identity(raw: dict[int, Any], emp_col: SourceColumn | None,
                      ident_col: SourceColumn | None, report: Report, row_ref: str) -> tuple[Any, Any]:
    """Work out (employee number, assignment identifier) for this source row.

    Employee Number comes from the Talent Link ID. Assignment Identifier comes
    from the legacy Topia ID, falling back to the Employee Number when absent so
    the row still links to its assignment (see clarification Q1/Q17).
    """
    employee = _id_value(raw[emp_col.idx]) if emp_col else None
    if employee is None:
        report.warn(row_ref, cfg.COL_EMPLOYEE_NUMBER,
                    "missing Employee Number (Talent Link ID); row written with blank key")

    identifier = _id_value(raw[ident_col.idx]) if ident_col else None
    if identifier is None:
        identifier = employee
        if emp_col:
            report.warn(row_ref, cfg.COL_ASSIGNMENT_IDENTIFIER,
                        f"no {cfg.ASSIGNMENT_IDENTIFIER_SOURCE!r}; Assignment Identifier falls back to Employee Number")
    return employee, identifier


def _apply_derived_rules(main: dict[str, Any], raw: dict[int, Any],
                         cancel_col: SourceColumn | None) -> Any:
    """Apply the small set of derived rules to the main dict; return the begin date.

    Rules (all documented in mapping_config):
      * a non-empty Cancellation Reason forces ASSIGNMENT!PHASE = "Cancelled";
      * INTERNATIONAL_RELO is derived from home vs host country when missing;
      * salary / family-size effective dates default to the assignment start.
    The returned begin date seeds address/contact child records.
    """
    if cancel_col and not is_blank(raw[cancel_col.idx]):
        main["ASSIGNMENT!PHASE"] = cfg.CANCELLED_PHASE

    if cfg.DERIVE_INTERNATIONAL_RELO:
        home, host = main.get("ASSIGNMENT!FROM_COUNTRY"), main.get("ASSIGNMENT!TO_COUNTRY")
        if main.get("ASSIGNMENT!INTERNATIONAL_RELO") is None and home and host:
            main["ASSIGNMENT!INTERNATIONAL_RELO"] = "Yes" if home != host else "No"

    start = main.get("ASSIGNMENT!START_DATE")
    if cfg.DEFAULT_DATES_FROM_START and start:
        if main.get("ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_AMOUNT") is not None:
            main["ASSIGNMENT_COMPENSATION_SALARY!EFFECTIVE_DATE"] = \
                main.get("ASSIGNMENT_COMPENSATION_SALARY!EFFECTIVE_DATE") or start
        if any(main.get(k) is not None for k in
               ("ASSIGNMENT_FAMILY_SIZE!HOST_FAMILY_SIZE", "ASSIGNMENT_FAMILY_SIZE!HOME_FAMILY_SIZE")):
            main["ASSIGNMENT_FAMILY_SIZE!EFFECTIVE_DATE"] = \
                main.get("ASSIGNMENT_FAMILY_SIZE!EFFECTIVE_DATE") or start

    return start if cfg.DEFAULT_DATES_FROM_START else None


def _build_address_records(raw: dict[int, Any], addresses: list[SourceColumn],
                           begin: Any, report: Report, row_ref: str) -> list[dict]:
    """Turn each non-empty address text column into one mailing-address record.

    The address is loaded as free-text lines (we have no structured city/postal
    split from Topia - clarification Q12); a warning flags this for review.
    """
    recs = []
    for col in addresses:
        value = raw[col.idx]
        if is_blank(value):
            continue
        lines = [ln.strip() for ln in str(value).splitlines() if ln.strip()]
        recs.append({
            ADDR + "PRIMARY_ADDRESS": "Yes",
            ADDR + "LOCATION_TYPE": cfg.ADDRESS_COLUMNS[col.canon],
            ADDR + "BEGIN_DATE": begin,
            ADDR + "LINE_1": lines[0] if lines else None,
            ADDR + "LINE_2": lines[1] if len(lines) > 1 else None,
            ADDR + "LINE_3": ", ".join(lines[2:]) or None,
        })
        report.warn(row_ref, ADDR + "LINE_1",
                    f"address from column {col.letter} loaded as free text; verify city/postal/country split")
    return recs


def _build_contact_records(raw: dict[int, Any], contacts: list[SourceColumn], begin: Any) -> list[dict]:
    """Turn each non-empty phone column into one employee-contact record."""
    recs = []
    for col in contacts:
        value = raw[col.idx]
        if is_blank(value):
            continue
        loc, ctype = cfg.CONTACT_COLUMNS[col.canon]
        num = str(int(value)) if isinstance(value, float) and value.is_integer() else str(value).strip()
        recs.append({
            CONTACT + "LOCATION_TYPE": loc,
            CONTACT + "CONTACT_TYPE": ctype,
            CONTACT + "NUMBER_ADDRESS": num,
            CONTACT + "START_DATE": begin,
            CONTACT + "PRIMARY_CONTACT": "Yes",
        })
    return recs


def _build_dependent_records(raw: dict[int, Any], dependents: dict[int, list[SourceColumn]],
                             layout: TargetLayout, report: Report, row_ref: str) -> list[dict]:
    """Build one dependent record per child block that has at least one value."""
    recs = []
    for group in dependents.values():
        rec = {c.dep_target: convert(c.dep_target, raw[c.idx], layout, report, row_ref) for c in group}
        if any(v is not None for v in rec.values()):
            recs.append(rec)
    return recs


def _emit_assignment_rows(main: dict[str, Any], employee: Any, identifier: Any,
                          seen_employees: set, child_record_lists: list[list[dict]]) -> list[dict]:
    """Expand one assignment into its Equus import rows.

    Row 0 INSERTs the assignment and carries all the main fields; any extra
    address/contact/dependent child records go on follow-on UPDATE rows keyed by
    the Assignment Identifier (clarification Q18/Q19). The employee is INSERTed
    the first time it is seen and IGNOREd on later rows so it is not duplicated.
    """
    n_rows = max([1] + [len(recs) for recs in child_record_lists])
    rows: list[dict] = []
    for i in range(n_rows):
        row = dict(main) if i == 0 else {}
        if i == 0:
            row[cfg.COL_EMPLOYEE_COMMAND] = (cfg.EMPLOYEE_COMMAND_FIRST
                                             if employee not in seen_employees else cfg.EMPLOYEE_COMMAND_REPEAT)
            row[cfg.COL_ASSIGNMENT_COMMAND] = cfg.ASSIGNMENT_COMMAND_INSERT
            row[cfg.COL_ASSIGNMENT_NUMBER] = cfg.NEW_ASSIGNMENT_NUMBER
        else:
            row[cfg.COL_EMPLOYEE_COMMAND] = cfg.EMPLOYEE_COMMAND_REPEAT
            row[cfg.COL_ASSIGNMENT_COMMAND] = cfg.ASSIGNMENT_COMMAND_UPDATE
        row[cfg.COL_ASSIGNMENT_IDENTIFIER] = identifier
        for recs in child_record_lists:
            if i < len(recs):
                row.update(recs[i])
        rows.append(row)
    return rows


def build_output_rows(src_ws, cols: list[SourceColumn], layout: TargetLayout, report: Report) -> list[dict]:
    """Walk every source assignment row and build the Equus output rows (step 4).

    Each non-empty source row yields one INSERT row plus any child UPDATE rows.
    The per-row work is delegated to the small _map/_resolve/_apply/_build/_emit
    helpers above so this function reads as the high-level recipe.
    """
    by_canon = {c.canon: c for c in cols}
    mapped = [c for c in cols if c.status == "mapped"]
    addresses = [c for c in cols if c.status == "address"]
    contacts = [c for c in cols if c.status == "contact"]
    dependents = _dependent_groups(cols)
    cancel_col = next((c for c in cols if c.status == "rule"), None)
    emp_col = next((c for c in cols if c.status == "identity"), None)
    ident_col = by_canon.get(cfg.ASSIGNMENT_IDENTIFIER_SOURCE)
    read_idx = [c.idx for c in cols]

    out_rows: list[dict] = []
    seen_employees: set = set()

    for r in range(cfg.SOURCE_DATA_START_ROW, src_ws.max_row + 1):
        raw = {idx: src_ws.cell(r, idx).value for idx in read_idx}
        if all(is_blank(v) for v in raw.values()):
            continue
        row_ref = f"source row {r}"

        main = _map_main_fields(raw, mapped, layout, report, row_ref)
        employee, identifier = _resolve_identity(raw, emp_col, ident_col, report, row_ref)
        main[cfg.COL_EMPLOYEE_NUMBER] = employee
        begin = _apply_derived_rules(main, raw, cancel_col)

        child_lists = [
            _build_address_records(raw, addresses, begin, report, row_ref),
            _build_contact_records(raw, contacts, begin),
            _build_dependent_records(raw, dependents, layout, report, row_ref),
        ]
        out_rows.extend(_emit_assignment_rows(main, employee, identifier, seen_employees, child_lists))
        seen_employees.add(employee)
        report.assignments += 1
    return out_rows


# --------------------------------------------------------------------------- step 5: output / report io

def relabel_headers_from_reference(ws) -> None:
    """Rewrite the output header rows (row 2 UI, row 3 DB) to the reference spelling.

    Keyed by the template's row-3 DB field. Runs AFTER data is written, so cell
    placement is unaffected. Columns absent from HEADER_MAP keep their jake
    header. The row-1 instruction is never written - it is dropped from the
    output entirely (see delete_instruction_row).
    """
    if not cfg.RELABEL_HEADERS_FROM_REFERENCE:
        return
    for c in range(1, ws.max_column + 1):
        db = ws.cell(cfg.TARGET_HEADER_ROW, c).value
        if db is None:
            continue
        entry = cfg.HEADER_MAP.get(str(db).strip())
        if not entry:
            continue
        ref_db, ref_ui = entry
        if ref_ui is not None:
            ws.cell(cfg.TARGET_UI_ROW, c, ref_ui)
        if cfg.RELABEL_DB_FIELDS and ref_db is not None:
            ws.cell(cfg.TARGET_HEADER_ROW, c, ref_db)


def delete_instruction_row(ws) -> None:
    """Remove the Equus instruction row (row 1) from the output workbook.

    The UI labels then shift up to row 1 and the DB fields to row 2. Does
    nothing when cfg.DELETE_INSTRUCTION_ROW is off.
    """
    if cfg.DELETE_INSTRUCTION_ROW:
        ws.delete_rows(cfg.TARGET_RULE_ROW, 1)


def write_output(template_bytes: bytes, rows: list[dict], report: Report) -> bytes:
    """Drop the output rows into a fresh copy of the Equus template (step 5).

    Values are placed by their DB field, dates get the US number format, then the
    headers are relabelled and the instruction row removed. Values whose DB field
    is not in the template are reported and skipped, never written blindly.
    """
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb[cfg.TARGET_SHEET]
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
    relabel_headers_from_reference(ws)   # headers -> reference, data already placed
    delete_instruction_row(ws)           # drop row 1; must be the last layout change
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_report(report: Report, layout: TargetLayout) -> bytes:
    """Write the five-sheet mapping report workbook (what went where and why)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Column Mapping"
    ws.append(["Source Col", "Topia Field", "Normalised", "Colour", "Status",
               "Equus Header", "Equus Col", "Equus UI (row 2)", "Rule (row 1)", "Note"])
    for c in report.columns:
        tcol = get_column_letter(layout.headers[c.target]) if c.target in layout.headers else None
        tgt = c.target or c.dep_target
        ws.append([c.letter, c.raw, c.canon, _colour_name(c.color), c.status, tgt, tcol,
                   layout.ui.get(tgt or "", ""), layout.rules.get(tgt or "", ""), c.note])

    ws = wb.create_sheet("Equus Coverage")
    ws.append(["Equus Col", "Equus Header", "Rule (row 1)", "Cells Populated", "Source Columns"])
    sources = defaultdict(list)
    for c in report.columns:
        tgt = c.target or c.dep_target
        if tgt:
            sources[tgt].append(c.letter)
    for header, idx in sorted(layout.headers.items(), key=lambda kv: kv[1]):
        ws.append([get_column_letter(idx), header, layout.rules.get(header, ""),
                   report.coverage.get(header, 0), ", ".join(sources.get(header, []))])

    ws = wb.create_sheet("Dropdown Values")
    ws.append(["Equus Header", "Value Sent", "Occurrences"])
    for (header, value), n in sorted(report.dropdown_values.items(), key=lambda kv: (kv[0][0], str(kv[0][1]))):
        ws.append([header, value, n])

    ws = wb.create_sheet("Warnings")
    ws.append(["Where", "Field", "Message"])
    for w in report.warnings:
        ws.append(list(w))

    ws = wb.create_sheet("Summary")
    counts = Counter(c.status for c in report.columns)
    ws.append(["Metric", "Value"])
    ws.append(["Source columns", len(report.columns)])
    for status in ("mapped", "identity", "address", "contact", "dependent", "rule",
                   "ignored", "excluded", "duplicate", "unmapped"):
        ws.append([f"  {status}", counts.get(status, 0)])
    ws.append(["Assignments read", report.assignments])
    ws.append(["Output rows written", report.rows_written])
    ws.append(["Warnings", len(report.warnings)])
    ws.append(["Missing required source columns", ", ".join(report.missing_source) or "none"])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


_COLOUR_NAMES = {
    "FFFF0000": "red", "FFFFFF00": "yellow", "FF00FFFF": "cyan",
    "FFB45F06": "orange", "FF00B050": "green",
}


def _colour_name(rgb: str | None) -> str:
    """Translate a known fill ARGB to a friendly colour name for the report."""
    if not rgb:
        return ""
    return _COLOUR_NAMES.get(rgb.upper(), rgb)


# --------------------------------------------------------------------------- entry point

def transform(source_bytes: bytes, template_bytes: bytes) -> tuple[bytes, bytes, Report]:
    """Run the full migration: Topia Source bytes -> Equus workbook + report.

    Returns (equus_workbook_bytes, report_workbook_bytes, report). This is the
    five-step pipeline described at the top of the module.
    """
    src_wb = openpyxl.load_workbook(io.BytesIO(source_bytes), data_only=True)
    if cfg.SOURCE_SHEET not in src_wb.sheetnames:
        # Fall back to the first sheet if the exact name differs.
        src_ws = src_wb[src_wb.sheetnames[0]]
    else:
        src_ws = src_wb[cfg.SOURCE_SHEET]

    tpl_wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    if cfg.TARGET_SHEET not in tpl_wb.sheetnames:
        raise ValueError(f"Template sheet {cfg.TARGET_SHEET!r} not found; got {tpl_wb.sheetnames}")
    layout = read_target_layout(tpl_wb[cfg.TARGET_SHEET])

    report = Report()
    report.columns = build_source_index(src_ws)
    plan_source_columns(report.columns, layout, report)
    rows = build_output_rows(src_ws, report.columns, layout, report)
    out = write_output(template_bytes, rows, report)
    return out, build_report(report, layout), report
