"""UAT-07: handler hardening - run manifest, audit metadata, failure
classification and secure logging (lambda/src/handler.py).

AC-1  a run writes <prefix>reports/<stem>_RUN_<stamp>.json with every listed
      field and the readiness; a duplicate event writes nothing new.
AC-2  a deliberately broken business_rules.json ends the run with a
      CONFIG-classified error and no import file.
AC-3  captured logs of the fixture run contain no first/last name or e-mail
      from the fixture.
Fixture: tests/fixtures/topia_sample.xlsx (synthetic data only); markers
below are synthetic too."""

import copy
import hashlib
import io
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))

import handler  # noqa: E402
import mapping_config as cfg  # noqa: E402
import transform as tmod  # noqa: E402
from transform import NOT_READY, READY  # noqa: E402

SOURCE = ROOT / "tests" / "fixtures" / "topia_sample.xlsx"
TEMPLATE = ROOT / "lambda" / "src" / "templates" / "equus_eafh_template.xlsx"
SHIPPED_RULES = ROOT / "lambda" / "src" / "business_rules.json"
SHEET = "Full Historical Past Assignment"
SRC_KEY = "topia/incoming/SYNTHETIC.xlsx"
EVENT = {"Records": [{"eventName": "ObjectCreated:Put",
                      "s3": {"bucket": {"name": "src"}, "object": {"key": SRC_KEY, "size": 1}}}]}

# Topia labels (row 4) of personal name / e-mail columns. Cost centre,
# policy and group names are not personal and are not listed.
PERSONAL_LABEL_RE = re.compile(r"(full|first|last|middle) name|e-?mail|(advisor|manager|initiator) .*name", re.I)
# Synthetic markers written into a copy of the fixture so a leak is unambiguous.
MARKERS = {"First Name": "Zqxfirstmarker", "Last Name": "Zqxlastmarker",
           "Email": "zqx.primary.marker@example.invalid", "Alternate Email": "zqx.alt.marker@example.invalid",
           "Spouse First Name": "Zqxspousefirst", "Spouse Last Name": "Zqxspouselast"}

MANIFEST_FIELDS = {"manifest_version", "run", "source", "template", "configuration", "code", "strict_pending",
                   "idempotency", "readiness", "output", "counts", "failure"}
COUNT_FIELDS = {"source_rows", "accepted", "held_rows", "held", "rows_written", "warnings",
                "unmapped_source_columns", "exceptions_by_code"}


# --------------------------------------------------------------------------- helpers

class S3Error(Exception):
    def __init__(self, code):
        super().__init__(code)
        self.response = {"Error": {"Code": code}}


class FakeS3:
    def __init__(self, objects):
        self.objects = objects
        self.puts = {}
        self.meta = {}
        self.content_types = {}

    def get_object(self, Bucket, Key):
        body = self.objects[(Bucket, Key)]
        return {"Body": io.BytesIO(body), "ETag": '"%s"' % hashlib.md5(body).hexdigest()}

    def head_object(self, Bucket, Key):
        if (Bucket, Key) not in self.puts:
            raise S3Error("404")
        return {}

    def put_object(self, Bucket, Key, Body, IfNoneMatch=None, Metadata=None, ContentType=None, **kwargs):
        if IfNoneMatch == "*" and (Bucket, Key) in self.puts:
            raise S3Error("PreconditionFailed")
        self.puts[(Bucket, Key)] = Body
        self.meta[(Bucket, Key)] = Metadata
        self.content_types[(Bucket, Key)] = ContentType


def _marked_source() -> bytes:
    wb = openpyxl.load_workbook(SOURCE)
    ws = wb[SHEET]
    labels = {ws.cell(4, c).value: c for c in range(1, ws.max_column + 1)}
    for label, marker in MARKERS.items():
        ws.cell(5, labels[label], marker)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _personal_values(source: bytes) -> set:
    """Every name / e-mail string of the fixture's data rows."""
    ws = openpyxl.load_workbook(io.BytesIO(source))[SHEET]
    out = set()
    for c in range(1, ws.max_column + 1):
        label = str(ws.cell(4, c).value or ws.cell(3, c).value or "")
        if not PERSONAL_LABEL_RE.search(label):
            continue
        for r in range(cfg.SOURCE_DATA_START_ROW, ws.max_row + 1):
            value = ws.cell(r, c).value
            if isinstance(value, str) and len(value.strip()) >= 3:
                out.add(value.strip())
    return out


def _leaks(text: str, values: set) -> list:
    return sorted(v for v in values if re.search(r"(?<![\w.@-])" + re.escape(v) + r"(?![\w.@-])", text))


def _keys(s3):
    return sorted(k for (_, k) in s3.puts)


def _manifest(s3):
    [key] = [k for k in _keys(s3) if k.endswith(".json")]
    return key, json.loads(s3.puts[("dst", key)])


def _import_files(s3):
    return [k for k in _keys(s3) if k.endswith(".xlsx") and not k.startswith("equus/reports/")]


def _setup(monkeypatch, source: bytes) -> FakeS3:
    s3 = FakeS3({("src", SRC_KEY): source})
    monkeypatch.setattr(handler, "_s3", s3)
    monkeypatch.setenv("DEST_BUCKET", "dst")
    monkeypatch.setenv("DEST_PREFIX", "equus/")
    monkeypatch.setenv("SOURCE_PREFIX", "topia/incoming/")
    monkeypatch.delenv("TEMPLATE_S3_URI", raising=False)
    monkeypatch.delenv("STRICT_PENDING", raising=False)
    return s3


@pytest.fixture
def fake_s3(monkeypatch):
    return _setup(monkeypatch, SOURCE.read_bytes())


def _run_failing(s3, code):
    with pytest.raises(handler.MigrationError) as info:
        handler.lambda_handler(EVENT, None)
    err = info.value
    assert err.code == code and str(err).startswith(f"[{code}] ")
    key, manifest = _manifest(s3)
    assert err.manifest_key == key
    assert manifest["run"]["status"] == "FAILED"
    assert manifest["failure"]["code"] == code
    assert manifest["output"]["import_file"] is None and manifest["output"]["import_ready"] is False
    assert manifest["output"]["kind"] == "NONE"
    assert manifest["readiness"]["status"] == NOT_READY
    assert _import_files(s3) == []
    return err, manifest


# --------------------------------------------------------------------------- AC-1 manifest

def test_run_writes_manifest_with_every_listed_field(fake_s3):
    summary = handler.lambda_handler(EVENT, None)["results"][0]
    keys = _keys(fake_s3)
    [out_key] = _import_files(fake_s3)
    manifest_key, m = _manifest(fake_s3)
    stamp = out_key.rsplit("_", 2)[-2] + "_" + out_key.rsplit("_", 1)[-1].removesuffix(".xlsx")
    assert manifest_key == f"equus/reports/SYNTHETIC_RUN_{stamp}.json"
    assert f"equus/reports/SYNTHETIC_MAPPING_REPORT_{stamp}.xlsx" in keys and len(keys) == 3
    assert fake_s3.content_types[("dst", manifest_key)] == "application/json"
    assert MANIFEST_FIELDS <= set(m)

    body = SOURCE.read_bytes()
    assert m["source"] == {"bucket": "src", "key": SRC_KEY, "etag": hashlib.md5(body).hexdigest(),
                           "sha256": hashlib.sha256(body).hexdigest(), "size_bytes": len(body)}
    tpl_sha = hashlib.sha256(TEMPLATE.read_bytes()).hexdigest()
    assert m["template"]["sha256"] == tpl_sha and m["template"]["version"] == f"bundled:{tpl_sha[:12]}"

    shipped = json.loads(SHIPPED_RULES.read_text(encoding="utf-8"))
    br = m["configuration"]["business_rules"]
    assert br["version"] == shipped["version"] and br["effective_version"] == shipped["effective_version"]
    assert br["sha256"] == hashlib.sha256(SHIPPED_RULES.read_bytes()).hexdigest()
    assert [r["id"] for r in br["rules"]] == [r["id"] for r in shipped["rules"]]
    assert all({"q", "status", "approved_for_import", "owner"} <= set(r) for r in br["rules"])
    assert {lst["list_id"] for lst in m["configuration"]["reference_data"]} == set(cfg.REFERENCE_LISTS)

    assert m["code"]["version"] == stamp.split("_")[1] and len(m["code"]["sha256"]) == 64
    assert m["idempotency"] == {"stamp": stamp, "import_key": f"equus/SYNTHETIC_EQUUS_{stamp}.xlsx",
                                "draft_key": f"equus/SYNTHETIC_EQUUS_DRAFT_{stamp}.xlsx"}

    # NOT_READY: the draft name, and the manifest says it is not import-ready.
    assert out_key == f"equus/SYNTHETIC_EQUUS_DRAFT_{stamp}.xlsx"
    assert m["readiness"]["status"] == NOT_READY and m["readiness"]["reason_count"] > 0
    assert len(m["readiness"]["reasons"]) == m["readiness"]["reason_count"]
    assert m["readiness"]["equus_uat_execution"].startswith("INCOMPLETE")
    assert m["output"]["import_file"] == out_key and m["output"]["kind"] == "DRAFT_NOT_IMPORT_READY"
    assert m["output"]["import_ready"] is False and "NOT import-ready" in m["output"]["statement"]
    assert m["output"]["mapping_report"] == f"equus/reports/SYNTHETIC_MAPPING_REPORT_{stamp}.xlsx"

    counts = m["counts"]
    assert COUNT_FIELDS <= set(counts)
    assert counts["source_rows"] == 2 and counts["accepted"] == 2
    assert counts["rows_written"] >= counts["accepted"] and counts["warnings"] > 0
    assert counts["held_rows"] == counts["held"]["assignment_rows"] + counts["held"]["dependent_rows"]
    categories = {r["category"] for r in m["readiness"]["reasons"]}
    assert categories <= set(counts["exceptions_by_code"])
    assert sum(counts["exceptions_by_code"].get(c, 0) for c in categories) == m["readiness"]["reason_count"]

    run = m["run"]
    assert run["status"] == "SUCCEEDED" and m["failure"] is None
    assert datetime.fromisoformat(run["started_at"]) <= datetime.fromisoformat(run["finished_at"])
    assert "SIGNED OFF" not in json.dumps(m).upper()

    # S3 object metadata carries readiness and configuration version.
    for key in (out_key, f"equus/reports/SYNTHETIC_MAPPING_REPORT_{stamp}.xlsx", manifest_key):
        meta = fake_s3.meta[("dst", key)]
        assert meta["readiness"] == NOT_READY and meta["import-ready"] == "false"
        assert meta["business-rules-version"] == shipped["version"]
        assert meta["business-rules-effective-version"] == shipped["effective_version"]
        assert meta["template-sha256"] == tpl_sha and meta["run-manifest"] == manifest_key
        assert meta["code-version"] == m["code"]["version"]
    assert summary["manifest"] == f"s3://dst/{manifest_key}" and summary["import_ready"] is False
    assert summary["held_rows"] == counts["held_rows"]


def test_duplicate_event_writes_nothing_new(fake_s3):
    handler.lambda_handler(EVENT, None)
    before = dict(fake_s3.puts)
    again = handler.lambda_handler(EVENT, None)["results"][0]
    assert again["skipped"] == "already processed"
    assert fake_s3.puts == before and len(before) == 3


def test_ready_run_manifest_marks_import_candidate_not_signed_off(fake_s3, monkeypatch):
    real = handler.transform

    def ready_transform(*args, **kwargs):
        out, report_bytes, report = real(*args, **kwargs)
        report.readiness = tmod.Readiness(status=READY, reasons=[])
        return out, report_bytes, report

    monkeypatch.setattr(handler, "transform", ready_transform)
    handler.lambda_handler(EVENT, None)
    [out_key] = _import_files(fake_s3)
    assert "_DRAFT_" not in out_key
    _, m = _manifest(fake_s3)
    assert m["readiness"]["status"] == READY
    assert m["output"]["kind"] == "IMPORT_CANDIDATE" and m["output"]["import_ready"] is True
    assert "PRE_IMPORT_CHECKLIST" in m["output"]["statement"]
    assert m["readiness"]["equus_uat_execution"].startswith("INCOMPLETE")
    assert fake_s3.meta[("dst", out_key)]["import-ready"] == "true"


# --------------------------------------------------------------------------- AC-2 / failure classes

def _broken_rules(tmp_path, mutate) -> Path:
    data = json.loads(SHIPPED_RULES.read_text(encoding="utf-8"))
    path = tmp_path / "business_rules.json"
    out = mutate(copy.deepcopy(data))
    path.write_text(out if isinstance(out, str) else json.dumps(out), encoding="utf-8")
    return path


@pytest.mark.parametrize("mutate,fragment", [
    (lambda d: json.dumps(d)[:-40], ""),  # truncated JSON
    (lambda d: {**d, "rules": [{**d["rules"][0], "status": "approved"}] + d["rules"][1:]}, "$.rules[0].status"),
    (lambda d: {**d, "rules": [{**d["rules"][0], "approved_for_import": "yes"}] + d["rules"][1:]},
     "$.rules[0].approved_for_import"),
    (lambda d: {k: v for k, v in d.items() if k != "version"}, "version"),
])
def test_broken_business_rules_is_config_failure_without_import_file(fake_s3, monkeypatch, tmp_path, mutate,
                                                                     fragment):
    monkeypatch.setattr(cfg, "BUSINESS_RULES_FILE", _broken_rules(tmp_path, mutate))
    err, m = _run_failing(fake_s3, handler.FAIL_CONFIG)
    assert isinstance(err.__cause__, cfg.ConfigurationError)
    assert fragment in str(err) and fragment in m["failure"]["message"]
    assert m["configuration"] is None and m["counts"] is None and m["readiness"]["computed"] is False
    # Only the manifest was written: no import file, no draft, no report.
    assert _keys(fake_s3) == [k for k in _keys(fake_s3) if k.endswith(".json")]


def _template_with(edit) -> bytes:
    wb = openpyxl.load_workbook(TEMPLATE)
    edit(wb["Sheet1"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _header_col(ws, header):
    return next(c for c in range(1, ws.max_column + 1) if ws.cell(3, c).value == header)


@pytest.mark.parametrize("edit,fragment", [
    # Renamed headers are UAT-03 FILE-level HEADER_MISMATCH findings (Q5).
    (lambda ws: ws.cell(3, _header_col(ws, "Assignment Command"), "Assignment Cmd"), "HEADER_MISMATCH"),
    (lambda ws: ws.cell(3, 1, "Employee Command"), "HEADER_MISMATCH"),  # D2: Jake's example A3 header
    (lambda ws: ws.cell(3, _header_col(ws, "EMPLOYEE!LAST_NAME"), "EMPLOYEE!FIRST_NAME"), "Duplicate Equus header"),
])
def test_template_header_mismatch_is_template_failure(fake_s3, monkeypatch, edit, fragment):
    body = _template_with(edit)
    info = {"source": "test", "sha256": hashlib.sha256(body).hexdigest(), "version": "test"}
    monkeypatch.setattr(handler, "_load_template_info", lambda: (body, info))
    err, m = _run_failing(fake_s3, handler.FAIL_TEMPLATE)
    assert fragment in str(err)
    assert m["template"]["sha256"] == info["sha256"]


def test_file_level_header_finding_is_template_failure(fake_s3, monkeypatch):
    """UAT-03 FILE-level finding (output headers differ from the approved
    template, Q5): the import workbook is withheld, the rejected-file report
    and the manifest are written and the failure is classified TEMPLATE."""
    body = _template_with(lambda ws: ws.cell(3, 6, "BROKEN"))
    monkeypatch.setattr(handler, "_load_template", lambda: body)
    err, m = _run_failing(fake_s3, handler.FAIL_TEMPLATE)
    assert "File-level validation failed" in str(err) and cfg.EXC_HEADER_MISMATCH in str(err)
    [rejected] = [k for k in _keys(fake_s3) if "FILE_REJECTED" in k]
    assert m["output"]["mapping_report"] == rejected and m["readiness"]["computed"] is True
    assert m["counts"]["exceptions_by_code"].get(cfg.EXC_HEADER_MISMATCH, 0) >= 1
    assert m["template"]["sha256"] == hashlib.sha256(body).hexdigest()
    assert fake_s3.meta[("dst", rejected)]["file-status"] == "FAILED"
    assert fake_s3.meta[("dst", rejected)]["import-ready"] == "false"


def test_unreadable_template_is_template_failure(fake_s3, monkeypatch):
    monkeypatch.setattr(handler, "_load_template_info", lambda: (b"not a workbook", {"source": "t", "sha256": "x",
                                                                                     "version": "t"}))
    _run_failing(fake_s3, handler.FAIL_TEMPLATE)


def test_strict_pending_is_pending_failure(fake_s3, monkeypatch):
    monkeypatch.setenv("STRICT_PENDING", "true")
    err, m = _run_failing(fake_s3, handler.FAIL_PENDING)
    assert "Pending (yellow) columns hold data" in str(err)
    assert m["strict_pending"] is True and m["configuration"]["business_rules"]["version"]


def test_reconciliation_failure_is_classified_and_keeps_evidence(fake_s3, monkeypatch):
    real = tmod.write_output

    def lossy(*args, **kwargs):
        data = real(*args, **kwargs)
        wb = openpyxl.load_workbook(io.BytesIO(data))
        ws = wb["Sheet1"]
        ws.cell(4, _header_col(ws, "EMPLOYEE!FIRST_NAME")).value = None
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    monkeypatch.setattr(tmod, "write_output", lossy)
    err, m = _run_failing(fake_s3, handler.FAIL_RECONCILIATION)
    assert "Reconciliation failed" in str(err)
    [failed] = [k for k in _keys(fake_s3) if "RECONCILIATION_FAILED" in k]
    assert m["output"]["mapping_report"] == failed
    assert m["readiness"]["computed"] is True
    assert "reconciliation" in {r["category"] for r in m["readiness"]["reasons"]}
    assert m["counts"]["source_rows"] == 2
    assert fake_s3.meta[("dst", failed)]["import-ready"] == "false"


def test_unreadable_source_is_source_failure(monkeypatch):
    s3 = _setup(monkeypatch, b"this is not an xlsx workbook")
    _run_failing(s3, handler.FAIL_SOURCE)


def test_unexpected_error_message_is_withheld(fake_s3, monkeypatch, caplog):
    caplog.set_level(logging.DEBUG)
    secret = MARKERS["Email"]

    def boom(*args, **kwargs):
        raise ValueError(f"cannot convert {secret!r}")

    monkeypatch.setattr(handler, "transform", boom)
    err, m = _run_failing(fake_s3, handler.FAIL_INTERNAL)
    assert err.__cause__ is None and err.__suppress_context__
    assert secret not in str(err) and secret not in json.dumps(m) and secret not in caplog.text
    assert "ValueError" in str(err) and "frames" in caplog.text


def test_failure_classification_table():
    assert handler.classify_failure(cfg.ConfigurationError("f", "$", "x")) == "CONFIG"
    assert handler.classify_failure(tmod.TemplateError("x")) == "TEMPLATE"
    assert handler.classify_failure(tmod.PendingFieldsError("x")) == "PENDING"
    assert handler.classify_failure(tmod.SourceError("x")) == "SOURCE"
    assert handler.classify_failure(KeyError("x")) == "INTERNAL"
    assert set(handler.FAILURE_CODES) == {"CONFIG", "TEMPLATE", "RECONCILIATION", "PENDING", "SOURCE", "INTERNAL"}


# --------------------------------------------------------------------------- AC-3 secure logging

def test_logs_contain_no_fixture_names_or_emails(monkeypatch, caplog):
    source = _marked_source()
    personal = _personal_values(source)
    assert set(MARKERS.values()) <= personal and len(personal) > len(MARKERS)
    # The source data really carries these values into the output, so a
    # clean log is not an accident of empty input.
    out, _, _ = tmod.transform(source, TEMPLATE.read_bytes())
    ws = openpyxl.load_workbook(io.BytesIO(out))["Sheet1"]
    written = {str(ws.cell(r, c).value) for r in range(4, ws.max_row + 1) for c in range(1, ws.max_column + 1)}
    assert {MARKERS["First Name"], MARKERS["Last Name"], MARKERS["Email"]} <= written

    s3 = _setup(monkeypatch, source)
    caplog.set_level(logging.DEBUG)
    handler.lambda_handler(EVENT, None)
    handler.lambda_handler(EVENT, None)  # duplicate path logs too
    assert "Migrating" in caplog.text and "already processed" in caplog.text
    assert _leaks(caplog.text, personal) == []
    _, m = _manifest(s3)
    assert _leaks(json.dumps(m), personal) == []


def test_failure_logs_and_messages_contain_no_fixture_personal_values(monkeypatch, caplog):
    source = _marked_source()
    personal = _personal_values(source)
    s3 = _setup(monkeypatch, source)
    monkeypatch.setenv("STRICT_PENDING", "true")
    caplog.set_level(logging.DEBUG)
    with pytest.raises(handler.MigrationError) as info:
        handler.lambda_handler(EVENT, None)
    assert _leaks(caplog.text + str(info.value), personal) == []
    _, m = _manifest(s3)
    assert _leaks(json.dumps(m), personal) == []
