"""AWS Lambda entry point: Topia workbook in SOURCE bucket -> Equus import
workbook (+ mapping report) in DEST bucket.

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
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import posixpath
from pathlib import Path
from urllib.parse import unquote_plus

from transform import transform

log = logging.getLogger()
log.setLevel(logging.INFO)

XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
BUNDLED_TEMPLATE = Path(__file__).parent / "templates" / "equus_eafh_template.xlsx"
MAX_SOURCE_BYTES = 50 * 1024 * 1024

_s3 = None


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
    """Short hash of the mapping code and template, so a mapping change gives
    new output keys while a duplicate event for the same source does not."""
    h = hashlib.sha256(template)
    for name in ("transform.py", "mapping_config.py"):
        h.update((Path(__file__).parent / name).read_bytes())
    return h.hexdigest()[:8]


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


def _put_once(bucket: str, key: str, body: bytes, meta: dict) -> bool:
    """Conditional write: False if the key already exists (concurrent duplicate)."""
    try:
        s3_client().put_object(Bucket=bucket, Key=key, Body=body, ContentType=XLSX_CONTENT_TYPE,
                               Metadata=meta, IfNoneMatch="*")
        return True
    except Exception as exc:
        if _error_code(exc) in ("PreconditionFailed", "412"):
            return False
        raise


def _load_template() -> bytes:
    uri = os.environ.get("TEMPLATE_S3_URI")
    if uri:
        bucket, _, key = uri.removeprefix("s3://").partition("/")
        return _read(bucket, key)
    return BUNDLED_TEMPLATE.read_bytes()


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


def lambda_handler(event, context):
    dest_bucket = os.environ["DEST_BUCKET"]
    dest_prefix = os.environ.get("DEST_PREFIX", "equus/")
    template = _load_template()
    results = []

    for bucket, key, size in _targets(event):
        skip = _accept(bucket, key, size, dest_bucket, dest_prefix)
        if skip:
            log.info("Skipping s3://%s/%s: %s", bucket, key, skip)
            results.append({"source": f"s3://{bucket}/{key}", "skipped": skip})
            continue

        body, etag = _read_with_etag(bucket, key)
        # S3 delivers notifications at least once: keys are derived from the
        # source ETag + code version, so a duplicate maps to the same objects.
        stem = posixpath.splitext(posixpath.basename(key))[0]
        stamp = f"{etag[:12]}_{_code_version(template)}"
        out_key = f"{dest_prefix}{stem}_EQUUS_{stamp}.xlsx"
        report_key = f"{dest_prefix}reports/{stem}_MAPPING_REPORT_{stamp}.xlsx"
        meta = {"source-bucket": bucket, "source-key": key[:1024], "source-etag": etag}
        if _exists(dest_bucket, out_key):
            log.info("Skipping s3://%s/%s: already processed as %s", bucket, key, out_key)
            results.append({"source": f"s3://{bucket}/{key}", "skipped": "already processed"})
            continue

        log.info("Migrating s3://%s/%s", bucket, key)
        strict = os.environ.get("STRICT_PENDING", "false").lower() == "true"
        output, report_bytes, report = transform(body, template, strict_pending=strict)
        if not report.reconciliation_ok:
            # Keep the evidence, withhold the import file, and fail the
            # invocation so the event lands on the on-failure queue.
            failed_key = f"{dest_prefix}reports/{stem}_RECONCILIATION_FAILED_{stamp}.xlsx"
            _put_once(dest_bucket, failed_key, report_bytes, meta)
            raise RuntimeError(f"Reconciliation failed for s3://{bucket}/{key}; "
                               f"see s3://{dest_bucket}/{failed_key}")
        if not _put_once(dest_bucket, out_key, output, meta):
            log.info("Skipping s3://%s/%s: concurrent duplicate already wrote %s", bucket, key, out_key)
            results.append({"source": f"s3://{bucket}/{key}", "skipped": "already processed"})
            continue
        _put_once(dest_bucket, report_key, report_bytes, meta)

        summary = {
            "source": f"s3://{bucket}/{key}",
            "output": f"s3://{dest_bucket}/{out_key}",
            "report": f"s3://{dest_bucket}/{report_key}",
            "assignments": report.assignments,
            "rows_written": report.rows_written,
            "unmapped_source_columns": sum(c.status == "unmapped" for c in report.columns),
            "warnings": len(report.warnings),
        }
        log.info(json.dumps(summary))
        results.append(summary)

    return {"results": results}
