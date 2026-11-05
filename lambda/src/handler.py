"""
AWS Lambda: map a Topia export workbook to an Equus EAFH import workbook.
"""

import json
import logging
import os
import posixpath
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote_plus

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font


# LOGGING
logger = logging.getLogger()
logger.setLevel(os.getenv("LOG_LEVEL", "INFO"))

XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MAX_SOURCE_BYTES = 50 * 1024 * 1024


# LOAD CONFIG (once per container / cold start)
CONFIG = json.loads((Path(__file__).with_name("mapping.json")).read_text(encoding="utf-8"))


COLUMN_SLOT_GROUPS = [
    {
        "target_key": "location_target_columns",
        "slots_key": "location_slots",
        "source_key": "location_source_column",
    },
    {
        "target_key": "phone_target_columns",
        "slots_key": "phone_slots",
        "source_key": "phone_source_column",
    },
]
CHILD_TARGET_KEY = "child_target_columns"
CHILD_SLOTS_KEY = "child_slots"
CHILD_FIELDS = [("child", "child_source_column"), ("gender", "gender_source_column")]

_s3 = None



# BASIC HELPERS
def s3_client():
    """Return a lazily-created, cached boto3 S3 client. Input: none. Return: client."""
    global _s3
    if _s3 is None:
        import boto3

        _s3 = boto3.client("s3")
    return _s3


def normalize_header(value):
    """Collapse a header to trimmed, single-spaced text. Input: cell. Return: str."""
    if value is None:
        return ""
    return " ".join(str(value).split())


def is_blank(value):
    """Return True when a value is None or whitespace-only. Input: any. Return: bool."""
    return value is None or (isinstance(value, str) and value.strip() == "")


def is_empty_row(row):
    """Return True when every cell in a row is blank. Input: row. Return: bool."""
    return all(is_blank(value) for value in row)


def get_value(row, index):
    """Return row[index], or None when the index is missing/out of range."""
    if index is None or index >= len(row):
        return None
    return row[index]



# VALUE / DROPDOWN MAPPING

def normalize_mapping_key(value, case_sensitive=False, trim=True):
    """Normalize a lookup key for value_mappings. Input: value+flags. Return: str/None."""
    if value is None:
        return None
    value = str(value)
    if trim:
        value = value.strip()
    if not case_sensitive:
        value = value.casefold()
    return value


def compile_value_mappings():
    """Pre-build the value_mappings lookup tables. Input: none. Return: dict by field."""
    result = {}
    for field, configuration in CONFIG.get("value_mappings", {}).items():
        case_sensitive = configuration.get("case_sensitive", False)
        trim = configuration.get("trim", True)
        values = {
            normalize_mapping_key(source, case_sensitive, trim): destination
            for source, destination in configuration.get("values", {}).items()
        }
        result[field] = {
            "case_sensitive": case_sensitive,
            "trim": trim,
            "on_unmapped": configuration.get("on_unmapped", "error"),
            "default": configuration.get("default"),
            "values": values,
        }
    return result


VALUE_MAPPINGS = compile_value_mappings()


def map_value(value, target_field, source_row_number, slot_number=None):
    """Translate a raw value for its Equus field per value_mappings.

    Input: the value, its Equus target, the source row number (for errors) and
    optional slot number. Return: the translated value (or the original when the
    field has no mapping); raises when an unmapped value has on_unmapped="error".
    """
    rule = VALUE_MAPPINGS.get(target_field)
    if rule is None:
        return value
    if is_blank(value):
        return None
    key = normalize_mapping_key(value, rule["case_sensitive"], rule["trim"])
    if key in rule["values"]:
        return rule["values"][key]
    action = rule["on_unmapped"]
    if action == "keep":
        return value
    if action == "blank":
        return None
    if action == "default":
        return rule["default"]
    raise ValueError(
        f"Unmapped value. Row={source_row_number}, Field={target_field}, "
        f"Value={value}, Slot={slot_number}"
    )



# EVENT PARSING / FILTERING
def extract_targets(event):
    """Pull the objects to process from an S3 event or a direct invoke.

    Input: the Lambda ``event`` dict. 
    Return: list of (bucket, key, size) tuples.
    """
    if "Records" in event:
        targets = []
        for record in event["Records"]:
            if not record.get("eventName", "").startswith("ObjectCreated"):
                continue
            obj = record["s3"]["object"]
            targets.append((record["s3"]["bucket"]["name"], unquote_plus(obj["key"]), obj.get("size")))
        return targets
    return [(event["source_bucket"], event["source_key"], None)]


def skip_reason(bucket, key, size, dest_bucket, dest_prefix):
    """Return why an object should be skipped, or None to process it.

    Input: object identity/size plus the destination. 
    Return: string or None.
    """
    prefix = os.getenv("SOURCE_PREFIX", "")
    suffix = os.getenv("SOURCE_SUFFIX", ".xlsx").lower()
    name = posixpath.basename(key)
    if not key.startswith(prefix):
        return f"key outside prefix {prefix!r}"
    if not key.lower().endswith(suffix):
        return f"key does not end with {suffix!r}"
    if name.startswith("~$"):
        return "Excel lock file"
    if bucket == dest_bucket and key.startswith(dest_prefix):
        return "object is one of our own outputs"
    if size is not None and size > MAX_SOURCE_BYTES:
        return f"object larger than {MAX_SOURCE_BYTES} bytes"
    return None


def create_output_key(source_key, dest_prefix):
    """Build the DEST key: ``<prefix><stem>_EQUUS_<timestamp>.xlsx``.

    Input: the source S3 key and destination prefix.
    Return: the output S3 key.
    """
    stem = posixpath.splitext(posixpath.basename(source_key))[0]
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return f"{dest_prefix}{stem}_EQUUS_{stamp}.xlsx"



# OUTPUT PLAN (columns follow output_column_order)
def resolve_child_slots(header_index, child_targets):
    """Resolve the child row-expansion slots to source column indexes.

    Input: the {header -> index} map and child target fields.
    Return: the slots whose source columns are present (absent ones are skipped silently).
    """
    slots = []
    for slot in CONFIG.get(CHILD_SLOTS_KEY, []):
        cells = [
            (child_targets[role], header_index.get(normalize_header(slot.get(source_key))))
            for role, source_key in CHILD_FIELDS
            if role in child_targets
        ]
        if any(index is not None for _, index in cells):
            slots.append({"slot_number": slot.get("slot_number"), "cells": cells})
    return slots


def build_plan(header_index):
    """Build the ordered output columns from output_column_order and the slots.

    Input: the {header -> index} map. Return: a plan with the ordered column
    descriptors (flat source columns and child row-columns), the header labels,
    the child target fields and resolved child slots. EMPLOYEE!SYSTEM_NUMBER is
    forced to the first column; location/phone targets expand to one column per
    slot; child targets become row-expansion columns.
    """
    employee_items = list(CONFIG["employee_column_mapping"].items())  # (source, target)
    employee_id_target = CONFIG["employee_id"]["target_column"]
    group_targets = CONFIG["output_column_order"][len(employee_items):]

    # Location / phone target field -> its present slot source indexes.
    column_group_sources = {}
    for group in COLUMN_SLOT_GROUPS:
        targets = CONFIG.get(group["target_key"])
        if not targets:
            continue
        target_field = next(iter(targets.values()))
        column_group_sources[target_field] = [
            index
            for slot in CONFIG.get(group["slots_key"], [])
            if (index := header_index.get(normalize_header(slot.get(group["source_key"])))) is not None
        ]

    child_targets = CONFIG.get(CHILD_TARGET_KEY, {})
    child_target_fields = [child_targets[role] for role, _ in CHILD_FIELDS if role in child_targets]

    employee_columns = [
        {"kind": "flat", "target": target, "index": header_index.get(normalize_header(source))}
        for source, target in employee_items
    ]
    employee_columns.sort(key=lambda column: 0 if column["target"] == employee_id_target else 1)

    columns = list(employee_columns)
    for target in group_targets:
        if target in child_target_fields:
            columns.append({"kind": "child", "target": target, "index": None})
        elif target in column_group_sources:
            columns += [{"kind": "flat", "target": target, "index": i} for i in column_group_sources[target]]
        else:
            columns.append({"kind": "flat", "target": target, "index": None})

    return {
        "columns": columns,
        "header": [column["target"] for column in columns],
        "employee_id_target": employee_id_target,
        "child_target_fields": child_target_fields,
        "child_slots": resolve_child_slots(header_index, child_targets),
    }


def prepare_mapping(worksheet):
    """Read the header row, validate required columns and locate the Employee ID.

    Input: the source worksheet. Return: a dict with the header row number, the
    {header -> index} map and the Employee ID source index.
    """
    header_row = int(CONFIG.get("header_row", 1))
    raw_headers = next(
        worksheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True), None
    )
    if raw_headers is None:
        raise ValueError("Header row is empty")
    headers = [normalize_header(header) for header in raw_headers]
    logger.info("ACTUAL SOURCE HEADERS: %s", [h for h in headers if h])
    header_index = {header: position for position, header in enumerate(headers) if header}

    required = set(CONFIG.get("required_source_columns", []))
    employee_id_source = CONFIG["employee_id"]["source_column"]
    required.add(employee_id_source)
    missing = [column for column in required if normalize_header(column) not in header_index]
    if missing:
        raise ValueError(f"Missing required source columns: {sorted(missing)}")

    for source_column in CONFIG["employee_column_mapping"]:
        if normalize_header(source_column) not in header_index:
            logger.warning("Optional source column not found: '%s'", source_column)

    return {
        "header_row": header_row,
        "header_index": header_index,
        "employee_id_index": header_index[normalize_header(employee_id_source)],
    }



# OUTPUT CELLS / HEADER
def create_header(worksheet, columns):
    """Build the bold Equus header cells. Input: worksheet + column labels.
    Return: a list of bold header cells for row 1.
    """
    bold = Font(bold=True)
    cells = []
    for column in columns:
        cell = WriteOnlyCell(worksheet, column)
        cell.font = bold
        cells.append(cell)
    return cells


def create_output_cell(worksheet, target_column, value):
    """Build one output cell, forcing text/date number formats where configured.

    Input: worksheet, the Equus target column and its value. Return: a cell,
    a raw value, or None for blanks.
    """
    if value is None:
        return None
    if target_column in CONFIG.get("text_output_columns", []):
        text = str(int(value)) if isinstance(value, float) and value.is_integer() else str(value).strip()
        cell = WriteOnlyCell(worksheet, text)
        cell.number_format = "@"
        return cell
    if target_column in CONFIG.get("date_output_columns", []):
        cell = WriteOnlyCell(worksheet, value)
        cell.number_format = "yyyy-mm-dd"
        return cell
    return value



# TRANSFORMATION
def collect_child_records(source_row, child_slots, source_number):
    """Read the populated child slots into {target_field: value} records.

    Input: the source row, resolved child slots and the row number. Return: one
    record per slot that has at least one populated source cell.
    """
    records = []
    for slot in child_slots:
        record, populated = {}, False
        for target_field, index in slot["cells"]:
            raw = get_value(source_row, index)
            if not is_blank(raw):
                populated = True
            record[target_field] = map_value(raw, target_field, source_number, slot["slot_number"])
        if populated:
            records.append(record)
    return records


def transform(source_path, output_path):
    """Read the Topia workbook and write the mapped Equus workbook to disk.

    Input: source and output file paths. Return: a statistics dict. Output:
    row 1 = Equus fields (SYSTEM_NUMBER first); one primary row per assignment,
    plus one additional row per extra child (sharing the Employee ID).
    """
    source_workbook = load_workbook(source_path, read_only=True, data_only=True, keep_links=False)
    target_workbook = Workbook(write_only=True)
    try:
        sheet_name = CONFIG.get("sheet_name")
        source_ws = source_workbook[sheet_name] if sheet_name else source_workbook[source_workbook.sheetnames[0]]
        mapping = prepare_mapping(source_ws)
        plan = build_plan(mapping["header_index"])
        columns = plan["columns"]
        employee_id_target = plan["employee_id_target"]

        target_ws = target_workbook.create_sheet(CONFIG.get("output_sheet_name", "Equus"))
        target_ws.freeze_panes = "A2"
        target_ws.append(create_header(target_ws, plan["header"]))

        stats = {"sourceRows": 0, "targetRows": 0, "primaryRows": 0, "additionalRows": 0}

        for source_number, source_row in enumerate(
            source_ws.iter_rows(min_row=mapping["header_row"] + 1, values_only=True),
            start=mapping["header_row"] + 1,
        ):
            if is_empty_row(source_row):
                continue
            stats["sourceRows"] += 1
            employee_id_value = get_value(source_row, mapping["employee_id_index"])
            if is_blank(employee_id_value):
                raise ValueError(f"Employee ID is blank. SourceRow={source_number}")

            children = collect_child_records(source_row, plan["child_slots"], source_number)
            first_child = children[0] if children else {}

            # Primary row: flat columns from the source + the first child record.
            primary = []
            for column in columns:
                if column["kind"] == "child":
                    primary.append(create_output_cell(target_ws, column["target"], first_child.get(column["target"])))
                else:
                    value = map_value(get_value(source_row, column["index"]), column["target"], source_number)
                    primary.append(create_output_cell(target_ws, column["target"], value))
            target_ws.append(primary)
            stats["targetRows"] += 1
            stats["primaryRows"] += 1

            # Additional rows: Employee ID + that child's columns only.
            for record in children[1:]:
                row_cells = []
                for column in columns:
                    if column["kind"] == "child":
                        value = record.get(column["target"])
                    elif column["target"] == employee_id_target:
                        value = employee_id_value
                    else:
                        value = None
                    row_cells.append(create_output_cell(target_ws, column["target"], value))
                target_ws.append(row_cells)
                stats["targetRows"] += 1
                stats["additionalRows"] += 1

            if stats["sourceRows"] % 1000 == 0:
                logger.info("PROGRESS SourceRows=%s TargetRows=%s", stats["sourceRows"], stats["targetRows"])

        target_workbook.save(output_path)
        logger.info("Transformation completed: %s", json.dumps(stats))
        return stats
    finally:
        # The write-only target is finalized by save(); only the streaming
        # source workbook needs an explicit close.
        source_workbook.close()



# LAMBDA HANDLER

def lambda_handler(event, context):
    """Lambda entry point: map every accepted Topia object into the DEST bucket.

    Input: the S3/direct-invoke ``event`` and Lambda ``context``. Return: a dict
    with a per-source ``results`` summary (also logged) for direct-invoke callers.
    """
    logger.info("Received event: %s", json.dumps(event))
    dest_bucket = os.environ["DEST_BUCKET"]
    dest_prefix = os.getenv("DEST_PREFIX", "equus/")
    results = []

    for bucket, key, size in extract_targets(event):
        logger.info("Input object: s3://%s/%s", bucket, key)
        reason = skip_reason(bucket, key, size, dest_bucket, dest_prefix)
        if reason:
            logger.info("Skipping s3://%s/%s: %s", bucket, key, reason)
            results.append({"source": f"s3://{bucket}/{key}", "skipped": reason})
            continue

        with tempfile.TemporaryDirectory() as workdir:
            source_path = Path(workdir) / "source.xlsx"
            output_path = Path(workdir) / "output.xlsx"
            s3_client().download_file(bucket, key, str(source_path))
            logger.info("Downloaded source file")
            stats = transform(source_path, output_path)
            out_key = create_output_key(key, dest_prefix)
            s3_client().upload_file(
                str(output_path),
                dest_bucket,
                out_key,
                ExtraArgs={"ContentType": XLSX_CONTENT_TYPE, "ServerSideEncryption": "AES256"},
            )
            logger.info("Created: s3://%s/%s", dest_bucket, out_key)

        summary = {"source": f"s3://{bucket}/{key}", "output": f"s3://{dest_bucket}/{out_key}", **stats}
        logger.info(json.dumps(summary))
        results.append(summary)

    return {"results": results}
