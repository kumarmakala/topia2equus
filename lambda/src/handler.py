"""AWS Lambda entry point: Topia workbook in SOURCE bucket -> Equus import
workbook (+ mapping report + run manifest) in DEST bucket.

Triggers:
  * S3 ObjectCreated notification on the source bucket (filtered to
    SOURCE_PREFIX / SOURCE_SUFFIX - see template.yaml).
  * Direct invoke: {"source_bucket": "...", "source_key": "..."}.

Environment:
  DEST_BUCKET      required  destination bucket
  DEST_PREFIX      optional  key prefix for outputs (default "equus/")
  SOURCE_PREFIX    optional  only process keys under this prefix (default "")
  SOURCE_SUFFIX    optional  only process keys ending with this (default ".xlsx")
  STRICT_PENDING   optional  "true" = fail while yellow (pending) columns hold data
  TEMPLATE_S3_URI  optional  s3://bucket/key of the Equus template; defaults
                             to the template bundled with the function

Import readiness (transform.Readiness) is independent of STRICT_PENDING. The
import workbook is named <stem>_EQUUS_<stamp>.xlsx only when readiness is
READY_FOR_UAT_EXECUTION; otherwise <stem>_EQUUS_DRAFT_<stamp>.xlsx (a
transformation artefact, not an import file). The status is written to the
S3 object metadata ("readiness", "import-ready") and the log summary.

A FILE-level finding (e.g. output headers differ from the approved template,
Q5) withholds the import workbook: only reports/<stem>_FILE_REJECTED_<stamp>.xlsx
(and the run manifest) is written and the invocation fails with a TEMPLATE
classification (INTERNAL for an engine row-structure defect). Held rows
(ROW_HELD) are never in the import workbook; they are listed in the report's
Exceptions and Held Rows sheets (docs/uat/REPROCESSING_PROCEDURE.md).

Audit trail (UAT-07): every processed source leaves
<prefix>reports/<stem>_RUN_<stamp>.json - source bucket/key/etag/sha256,
template sha256 + version, business rules version + rule set, code version,
readiness status and reasons, counts and timestamps - also when the run fails.
Failures are raised as MigrationError with a classification code
(FAILURE_CODES) in the message and in the manifest.

Secure logging: log lines carry S3 keys, counts, versions and codes only,
never cell values, names or e-mail addresses (tests/test_handler_audit.py).
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import posixpath
import traceback
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote_plus

import mapping_config as cfg
from mapping_config import ConfigurationError
from transform import READY, PendingFieldsError, SourceError, TemplateError, transform

log = logging.getLogger()
log.setLevel(logging.INFO)

XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
JSON_CONTENT_TYPE = "application/json"
BUNDLED_TEMPLATE = Path(__file__).parent / "templates" / "equus_eafh_template.xlsx"
MAX_SOURCE_BYTES = 50 * 1024 * 1024
MANIFEST_VERSION = "1"

# Failure classification (raised message prefix "[CODE]" and manifest).
FAIL_CONFIG = "CONFIG"                  # business_rules.json / reference_data rejected at load
FAIL_TEMPLATE = "TEMPLATE"              # Equus template unreadable, sheet or header row mismatch (Q5)
FAIL_RECONCILIATION = "RECONCILIATION"  # source-to-output reconciliation did not balance
FAIL_PENDING = "PENDING"                # STRICT_PENDING=true and yellow columns hold data
FAIL_SOURCE = "SOURCE"                  # source workbook unreadable or without the source sheet
FAIL_INTERNAL = "INTERNAL"              # anything else (message withheld from the log: may quote data)
FAILURE_CODES = (FAIL_CONFIG, FAIL_TEMPLATE, FAIL_RECONCILIATION, FAIL_PENDING, FAIL_SOURCE, FAIL_INTERNAL)
# UAT-03 FILE-level exception codes that are template header findings (Q5).
FILE_TEMPLATE_CODES = frozenset({cfg.EXC_HEADER_MISMATCH, cfg.EXC_TEMPLATE_HASH_DIFFERS})

EQUUS_EXECUTION = ("INCOMPLETE - no Equus import was performed or observed by this function; Equus UAT "
                   "execution, import results and post-import audit are separate, authorised steps")

_s3 = None


class MigrationError(RuntimeError):
    """A classified run failure. `code` is one of FAILURE_CODES."""

    def __init__(self, code: str, message: str, manifest_key: str | None = None):
        self.code = code
        self.manifest_key = manifest_key
        super().__init__(f"[{code}] {message}")


def s3_client():
    global _s3
    if _s3 is None:
        import boto3
        _s3 = boto3.client("s3")
    return _s3


def _read(bucket: str, key: str) -> bytes:
    return s3_client().get_object(Bucket=bucket, Key=key)["Body"].read()


def _read_with_etag(bucket: str, key: str) -> tuple[bytes, str]:
    obj = s3_client().get_object(Bucket=bucket, Key=key)
    return obj["Body"].read(), obj["ETag"].strip('"')


def _code_version(template: bytes) -> str:
    """Short hash of the mapping code, business configuration and template, so
    a mapping or rule-approval change gives new output keys while a duplicate
    event for the same source does not."""
    base = Path(__file__).parent
    h = hashlib.sha256(template)
    for name in ("transform.py", "mapping_config.py", "business_rules.json"):
        h.update((base / name).read_bytes())
    for path in sorted((base / "reference_data").glob("*.json")):
        h.update(path.name.encode())
        h.update(path.read_bytes())
    return h.hexdigest()[:8]


def _code_sha256() -> str:
    """Hash of the Python code only (handler, transform, mapping_config)."""
    base = Path(__file__).parent
    h = hashlib.sha256()
    for name in ("handler.py", "transform.py", "mapping_config.py"):
        path = base / name
        if path.exists():
            h.update(name.encode())
            h.update(path.read_bytes())
    return h.hexdigest()


def _error_code(exc: Exception) -> str | None:
    return getattr(exc, "response", {}).get("Error", {}).get("Code")


def _exists(bucket: str, key: str) -> bool:
    try:
        s3_client().head_object(Bucket=bucket, Key=key)
        return True
    except Exception as exc:  # botocore ClientError; not imported so local fakes work
        if _error_code(exc) in ("404", "NoSuchKey", "NotFound"):
            return False
        raise


def _put_once(bucket: str, key: str, body: bytes, meta: dict, content_type: str = XLSX_CONTENT_TYPE) -> bool:
    """Conditional write: False if the key already exists (concurrent duplicate)."""
    try:
        s3_client().put_object(Bucket=bucket, Key=key, Body=body, ContentType=content_type,
                               Metadata=meta, IfNoneMatch="*")
        return True
    except Exception as exc:
        if _error_code(exc) in ("PreconditionFailed", "412"):
            return False
        raise


_template_meta: dict[str, dict] = {}  # sha256 -> {"source", "version"} of templates loaded by _load_template


def _load_template() -> bytes:
    """Template bytes; records where they came from for _load_template_info.
    The version is the S3 VersionId (or ETag) of TEMPLATE_S3_URI, or
    'bundled:<sha256[:12]>' for the bundled copy: Equus has issued no template
    version identifier (D2)."""
    uri = os.environ.get("TEMPLATE_S3_URI")
    if uri:
        bucket, _, key = uri.removeprefix("s3://").partition("/")
        obj = s3_client().get_object(Bucket=bucket, Key=key)
        body = obj["Body"].read()
        version = obj.get("VersionId") or ("etag:" + str(obj.get("ETag", "")).strip('"'))
        _template_meta[hashlib.sha256(body).hexdigest()] = {"source": f"s3://{bucket}/{key}", "version": version}
        return body
    body = BUNDLED_TEMPLATE.read_bytes()
    sha = hashlib.sha256(body).hexdigest()
    _template_meta[sha] = {"source": f"bundled:templates/{BUNDLED_TEMPLATE.name}", "version": f"bundled:{sha[:12]}"}
    return body


def _load_template_info() -> tuple[bytes, dict]:
    """Template bytes plus source, sha256 and version (manifest 'template')."""
    body = _load_template()
    sha = hashlib.sha256(body).hexdigest()
    known = _template_meta.get(sha, {"source": "unrecorded", "version": f"sha256:{sha[:12]}"})
    return body, {"source": known["source"], "sha256": sha, "version": known["version"]}


def _targets(event: dict) -> list[tuple[str, str, int | None]]:
    if "Records" in event:
        out = []
        for rec in event["Records"]:
            if not rec.get("eventName", "").startswith("ObjectCreated"):
                continue
            obj = rec["s3"]["object"]
            out.append((rec["s3"]["bucket"]["name"], unquote_plus(obj["key"]), obj.get("size")))
        return out
    return [(event["source_bucket"], event["source_key"], None)]


def _accept(bucket: str, key: str, size: int | None, dest_bucket: str, dest_prefix: str) -> str | None:
    """Defence in depth behind the S3 notification filter. Returns a skip reason."""
    prefix = os.environ.get("SOURCE_PREFIX", "")
    suffix = os.environ.get("SOURCE_SUFFIX", ".xlsx").lower()
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


# --------------------------------------------------------------------------- audit

def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def classify_failure(exc: Exception) -> str:
    if isinstance(exc, ConfigurationError):
        return FAIL_CONFIG
    if isinstance(exc, TemplateError):
        return FAIL_TEMPLATE
    if isinstance(exc, PendingFieldsError):
        return FAIL_PENDING
    if isinstance(exc, SourceError):
        return FAIL_SOURCE
    return FAIL_INTERNAL


def _safe_message(code: str, exc: Exception) -> str:
    """Messages of the classified errors are built by this code from file
    paths, JSON paths, column letters and header labels - never cell values.
    Any other exception may quote data, so only its type is reported."""
    if code == FAIL_INTERNAL:
        return f"unexpected {type(exc).__name__}; message withheld (may contain source data)"
    return str(exc)


def _frames(exc: BaseException) -> str:
    """Traceback as file:line:function only - no source lines, no values."""
    return " <- ".join(f"{Path(f.filename).name}:{f.lineno}:{f.name}"
                       for f in reversed(traceback.extract_tb(exc.__traceback__)))


def _file_sha256(path) -> str | None:
    try:
        return hashlib.sha256(Path(path).read_bytes()).hexdigest()
    except OSError:
        return None


def _configuration_section(rules, refs) -> dict:
    dep = rules.dependent_source
    return {
        "business_rules": {
            "file": rules.file_name,
            "sha256": _file_sha256(rules.source),
            "version": rules.version,
            "effective_version": rules.effective_version,
            "generated_at": rules.generated_at,
            "rules": [{"id": r.id, "q": r.q, "status": r.status, "approved_for_import": r.approved_for_import,
                       "owner": r.owner} for r in rules.rules.values()],
            "dependent_source": None if dep is None else {
                "q": dep.q, "status": dep.status, "approved_for_import": dep.approved_for_import,
                "owner": dep.owner, "configured": dep.configured},
        },
        "reference_data": [{"list_id": lst.list_id, "q": lst.q, "status": lst.status,
                            "approved_for_import": lst.approved_for_import, "owner": lst.owner,
                            "version": lst.version, "state": lst.state, "values": len(lst.values)}
                           for lst in refs.values()],
    }


def _counts(report) -> dict:
    """Counts only. 'held' is what this run did not migrate: source rows held
    out of the import workbook (UAT-03 ROW_HELD codes, report.held_rows, or
    otherwise without an INSERT row), dependent rows/records held with an
    exception code, and HELD columns that hold data (their values stay out of
    the import). exceptions_by_code: readiness reason categories (lower case),
    dependent exception codes and UAT-03 exception codes (upper case)."""
    held_dep_rows = sum(e["rows"] for e in report.dependent_exceptions)
    held_dep_records = sum(e["groups"] for e in report.dependent_exceptions)
    held_assignment_rows = max(len(report.held_rows), report.data_rows - report.assignments, 0)
    exceptions = Counter(r["category"] for r in report.readiness.reasons) if report.readiness else Counter()
    for e in report.dependent_exceptions:
        exceptions[e["code"]] += 1
    for e in report.exceptions:
        exceptions[e["code"]] += 1
    return {
        "source_rows": report.data_rows,
        "accepted": report.assignments,
        "held_rows": held_assignment_rows + held_dep_rows,
        "held": {
            "assignment_rows": held_assignment_rows,
            "dependent_rows": held_dep_rows,
            "dependent_records": held_dep_records,
            "columns_with_data": sum(1 for c in report.columns if "HELD" in c.flags and c.has_data),
        },
        "rows_written": report.rows_written,
        "warnings": len(report.warnings),
        "unmapped_source_columns": sum(c.status == "unmapped" for c in report.columns),
        "exceptions_by_code": dict(sorted(exceptions.items())),
    }


def _output_section(import_key: str | None, ready: bool, report_key: str | None) -> dict:
    if import_key is None:
        kind, statement = "NONE", "No import file was written by this run. Nothing may be imported into Equus."
    elif ready:
        kind = "IMPORT_CANDIDATE"
        statement = ("Readiness READY_FOR_UAT_EXECUTION: candidate for a controlled, authorised Equus UAT import "
                     "only after docs/uat/PRE_IMPORT_CHECKLIST.md is complete. No business sign-off is implied.")
    else:
        kind = "DRAFT_NOT_IMPORT_READY"
        statement = ("NOT import-ready: transformation artefact (*_EQUUS_DRAFT_*) for review only. Do not import "
                     "into Equus; resolve every readiness reason first.")
    return {"import_file": import_key, "kind": kind, "import_ready": bool(import_key) and ready,
            "statement": statement, "mapping_report": report_key}


def _readiness_section(report) -> dict:
    readiness = report.readiness
    return {"status": readiness.status, "computed": True, "reason_count": len(readiness.reasons),
            "reasons": readiness.reasons, "equus_uat_execution": EQUUS_EXECUTION}


# --------------------------------------------------------------------------- handler

def lambda_handler(event, context):
    dest_bucket = os.environ["DEST_BUCKET"]
    dest_prefix = os.environ.get("DEST_PREFIX", "equus/")
    template, template_info = _load_template_info()
    results = []

    for bucket, key, size in _targets(event):
        skip = _accept(bucket, key, size, dest_bucket, dest_prefix)
        if skip:
            log.info("Skipping s3://%s/%s: %s", bucket, key, skip)
            results.append({"source": f"s3://{bucket}/{key}", "skipped": skip})
            continue
        results.append(_process(bucket, key, template, template_info, dest_bucket, dest_prefix, context))

    return {"results": results}


def _process(bucket: str, key: str, template: bytes, template_info: dict, dest_bucket: str, dest_prefix: str,
             context) -> dict:
    started_at = _now()
    body, etag = _read_with_etag(bucket, key)
    # S3 delivers notifications at least once: keys are derived from the
    # source ETag + code version, so a duplicate maps to the same objects.
    stem = posixpath.splitext(posixpath.basename(key))[0]
    code_version = _code_version(template)
    stamp = f"{etag[:12]}_{code_version}"
    ready_key = f"{dest_prefix}{stem}_EQUUS_{stamp}.xlsx"
    draft_key = f"{dest_prefix}{stem}_EQUUS_DRAFT_{stamp}.xlsx"
    report_key = f"{dest_prefix}reports/{stem}_MAPPING_REPORT_{stamp}.xlsx"
    manifest_key = f"{dest_prefix}reports/{stem}_RUN_{stamp}.json"
    source_ref = f"s3://{bucket}/{key}"
    done = next((k for k in (ready_key, draft_key) if _exists(dest_bucket, k)), None)
    if done:
        log.info("Skipping %s: already processed as %s", source_ref, done)
        return {"source": source_ref, "skipped": "already processed"}

    strict = os.environ.get("STRICT_PENDING", "false").lower() == "true"
    manifest = {
        "manifest_version": MANIFEST_VERSION,
        "run": {"status": "RUNNING", "started_at": started_at, "finished_at": None,
                "aws_request_id": getattr(context, "aws_request_id", None),
                "function_version": getattr(context, "function_version", None)},
        "source": {"bucket": bucket, "key": key, "etag": etag, "sha256": hashlib.sha256(body).hexdigest(),
                   "size_bytes": len(body)},
        "template": dict(template_info),
        "configuration": None,
        "code": {"version": code_version, "sha256": _code_sha256()},
        "strict_pending": strict,
        "idempotency": {"stamp": stamp, "import_key": ready_key, "draft_key": draft_key},
        "readiness": None,
        "output": _output_section(None, False, None),
        "counts": None,
        "failure": None,
    }
    meta = {"source-bucket": bucket, "source-key": key[:1024], "source-etag": etag,
            "code-version": code_version, "template-sha256": template_info["sha256"],
            "run-manifest": manifest_key[:1024]}

    def fail(code: str, message: str, exc: Exception | None = None, report=None, report_written: str | None = None):
        manifest["run"].update(status="FAILED", finished_at=_now())
        manifest["failure"] = {"code": code, "message": message}
        if report is not None and report.readiness is not None:
            manifest["readiness"] = _readiness_section(report)
            manifest["counts"] = _counts(report)
        else:
            manifest["readiness"] = {"status": "NOT_READY", "computed": False, "reason_count": 1,
                                     "reasons": [{"category": "failure", "item": code, "q": "", "owner": "",
                                                  "detail": message}],
                                     "equus_uat_execution": EQUUS_EXECUTION}
        manifest["output"] = _output_section(None, False, report_written)
        _put_once(dest_bucket, manifest_key, _manifest_bytes(manifest), meta, JSON_CONTENT_TYPE)
        log.error(json.dumps({"source": source_ref, "failure": code, "manifest": f"s3://{dest_bucket}/{manifest_key}",
                              **({"frames": _frames(exc)} if exc is not None and code == FAIL_INTERNAL else {})}))
        raise MigrationError(code, f"{message} [source {source_ref}; manifest s3://{dest_bucket}/{manifest_key}]",
                             manifest_key) from (None if code == FAIL_INTERNAL else exc)

    log.info("Migrating %s", source_ref)
    try:
        rules = cfg.load_business_rules()
        refs = cfg.load_reference_data()
        manifest["configuration"] = _configuration_section(rules, refs)
        meta.update({"business-rules-version": rules.version,
                     "business-rules-effective-version": rules.effective_version})
        output, report_bytes, report = transform(body, template, strict_pending=strict,
                                                 business_rules=rules, reference_data=refs)
    except Exception as exc:
        code = classify_failure(exc)
        fail(code, _safe_message(code, exc), exc)

    readiness = report.readiness
    ready = readiness.status == READY
    # Never name a file as import-ready while it is not.
    out_key = ready_key if ready else draft_key
    meta.update({
        "readiness": readiness.status,
        "readiness-reasons": str(len(readiness.reasons)),
        "import-ready": "true" if ready else "false",
        "business-rules-version": report.business_rules_version,
        "rows-held": str(len(report.held_rows)),
        "file-status": "FAILED" if report.file_failed else "ACCEPTED",
    })
    manifest["readiness"] = _readiness_section(report)
    manifest["counts"] = _counts(report)
    if report.file_failed:
        # FILE-level finding (UAT-03, e.g. header mismatch with the approved
        # template, Q5): withhold the import workbook, keep the report.
        rejected_key = f"{dest_prefix}reports/{stem}_FILE_REJECTED_{stamp}.xlsx"
        meta["import-ready"] = "false"
        _put_once(dest_bucket, rejected_key, report_bytes, meta)
        file_codes = sorted({e["code"] for e in report.exceptions if e["level"] == cfg.LEVEL_FILE})
        code = FAIL_TEMPLATE if set(file_codes) & FILE_TEMPLATE_CODES else FAIL_INTERNAL
        fail(code, f"File-level validation failed ({', '.join(file_codes)}) for {source_ref}; "
                   f"see s3://{dest_bucket}/{rejected_key}", report=report, report_written=rejected_key)
    if not report.reconciliation_ok:
        # Keep the evidence, withhold the import file, and fail the
        # invocation so the event lands on the on-failure queue.
        failed_key = f"{dest_prefix}reports/{stem}_RECONCILIATION_FAILED_{stamp}.xlsx"
        meta["import-ready"] = "false"
        _put_once(dest_bucket, failed_key, report_bytes, meta)
        fail(FAIL_RECONCILIATION, f"Reconciliation failed for {source_ref}; see s3://{dest_bucket}/{failed_key}",
             report=report, report_written=failed_key)

    if not _put_once(dest_bucket, out_key, output, meta):
        log.info("Skipping %s: concurrent duplicate already wrote %s", source_ref, out_key)
        return {"source": source_ref, "skipped": "already processed"}
    _put_once(dest_bucket, report_key, report_bytes, meta)
    manifest["run"].update(status="SUCCEEDED", finished_at=_now())
    manifest["output"] = _output_section(out_key, ready, report_key)
    _put_once(dest_bucket, manifest_key, _manifest_bytes(manifest), meta, JSON_CONTENT_TYPE)

    counts = manifest["counts"]
    summary = {
        "source": source_ref,
        "output": f"s3://{dest_bucket}/{out_key}",
        "report": f"s3://{dest_bucket}/{report_key}",
        "manifest": f"s3://{dest_bucket}/{manifest_key}",
        "assignments": report.assignments,
        "rows_held": len(report.held_rows),
        "rows_written": report.rows_written,
        "source_rows": counts["source_rows"],
        "held_rows": counts["held_rows"],
        "unmapped_source_columns": counts["unmapped_source_columns"],
        "warnings": len(report.warnings),
        "readiness": readiness.status,
        "readiness_reasons": len(readiness.reasons),
        "import_ready": ready,
        "business_rules_version": report.business_rules_version,
        "exceptions_by_code": counts["exceptions_by_code"],
    }
    log.info(json.dumps(summary))
    return summary


def _manifest_bytes(manifest: dict) -> bytes:
    return json.dumps(manifest, indent=2, sort_keys=False, default=str).encode("utf-8")
