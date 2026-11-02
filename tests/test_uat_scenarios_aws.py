"""UAT-06 AWS / import scenarios (master command section 7,
clarifications/UAT_MASTER_COMMAND.md:149-151; BR-05, Q5) against
lambda/src/handler.py with an in-memory S3 fake (no AWS, no Equus). Test ids
AW-xx are listed in docs/uat/TEST_RESULTS.md. Equus rejection, correction /
resubmission in Equus and the post-import audit are NOT EXECUTED here (see
TEST_RESULTS.md section 4).

Also: AC-2 (fixtures are synthetic) and AC-1 (every PASS row of
TEST_RESULTS.md names an existing test)."""

import hashlib
import io
import json
import re
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))
sys.path.insert(0, str(ROOT / "tests"))

import handler  # noqa: E402
import mapping_config as cfg  # noqa: E402
import transform as tmod  # noqa: E402
import uat_scenario_fixtures as fx  # noqa: E402
from transform import NOT_READY  # noqa: E402

SRC_KEY = "topia/incoming/SYN UAT 06.xlsx"
ENCODED_KEY = "topia/incoming/SYN+UAT+06.xlsx"  # S3 notifications URL-encode the key


def _event(key: str = ENCODED_KEY, name: str = "ObjectCreated:Put", n: int = 1) -> dict:
    return {"Records": [{"eventName": name, "s3": {"bucket": {"name": "src"}, "object": {"key": key, "size": 1}}}] * n}


class S3Error(Exception):
    def __init__(self, code):
        super().__init__(code)
        self.response = {"Error": {"Code": code}}


class FakeS3:
    """Conditional puts like S3 (IfNoneMatch='*'), plus one-shot failure injection."""

    def __init__(self, objects):
        self.objects = dict(objects)
        self.puts: dict = {}
        self.meta: dict = {}
        self.fail_put: list[str] = []  # key fragments: next matching put raises InternalError (once)
        self.fail_get = 0

    def get_object(self, Bucket, Key):
        if self.fail_get:
            self.fail_get -= 1
            raise S3Error("SlowDown")
        body = self.objects[(Bucket, Key)]
        return {"Body": io.BytesIO(body), "ETag": '"%s"' % hashlib.md5(body).hexdigest()}

    def head_object(self, Bucket, Key):
        if (Bucket, Key) not in self.puts:
            raise S3Error("404")
        return {}

    def put_object(self, Bucket, Key, Body, IfNoneMatch=None, Metadata=None, ContentType=None, **kwargs):
        for frag in list(self.fail_put):
            if frag in Key:
                self.fail_put.remove(frag)
                raise S3Error("InternalError")
        if IfNoneMatch == "*" and (Bucket, Key) in self.puts:
            raise S3Error("PreconditionFailed")
        self.puts[(Bucket, Key)] = Body
        self.meta[(Bucket, Key)] = Metadata

    # views
    def keys(self):
        return sorted(k for (_, k) in self.puts)

    def imports(self):
        return [k for k in self.keys() if k.endswith(".xlsx") and not k.startswith("equus/reports/")]

    def manifest(self):
        [key] = [k for k in self.keys() if k.endswith(".json")]
        return key, json.loads(self.puts[("dst", key)])


def _setup(monkeypatch, source: bytes, key: str = SRC_KEY) -> FakeS3:
    s3 = FakeS3({("src", key): source})
    monkeypatch.setattr(handler, "_s3", s3)
    monkeypatch.setenv("DEST_BUCKET", "dst")
    monkeypatch.setenv("DEST_PREFIX", "equus/")
    monkeypatch.setenv("SOURCE_PREFIX", "topia/incoming/")
    monkeypatch.delenv("TEMPLATE_S3_URI", raising=False)
    monkeypatch.delenv("STRICT_PENDING", raising=False)
    return s3


@pytest.fixture
def s3(monkeypatch):
    return _setup(monkeypatch, fx.BASE.read_bytes())


def _complete_set(s3: FakeS3) -> None:
    """Exactly one import workbook, its mapping report and a SUCCEEDED manifest that names both."""
    [imp] = s3.imports()
    stem_stamp = imp.removeprefix("equus/").removesuffix(".xlsx").replace("_EQUUS_DRAFT_", "|").replace("_EQUUS_", "|")
    stem, stamp = stem_stamp.split("|")
    assert s3.keys() == sorted([imp, f"equus/reports/{stem}_MAPPING_REPORT_{stamp}.xlsx",
                                f"equus/reports/{stem}_RUN_{stamp}.json"])
    _, m = s3.manifest()
    assert m["run"]["status"] == "SUCCEEDED" and m["output"]["import_file"] == imp
    assert m["output"]["mapping_report"] == f"equus/reports/{stem}_MAPPING_REPORT_{stamp}.xlsx"


# --------------------------------------------------------------------------- AW-01 valid trigger

def test_AW01_valid_s3_trigger_writes_draft_report_manifest(s3):
    [result] = handler.lambda_handler(_event(), None)["results"]
    _complete_set(s3)
    [imp] = s3.imports()
    assert "_EQUUS_DRAFT_" in imp and imp.startswith("equus/SYN UAT 06_")  # key URL-decoded
    assert result["readiness"] == NOT_READY and result["import_ready"] is False
    assert result["assignments"] == 2 and result["rows_held"] == 0
    meta = s3.meta[("dst", imp)]
    assert meta["import-ready"] == "false" and meta["file-status"] == "ACCEPTED" and meta["source-key"] == SRC_KEY
    rows = fx.output_rows(s3.puts[("dst", imp)])
    assert [r[fx.IDENT] for r in fx.inserts(rows)] == ["80001", "80002"]
    _, m = s3.manifest()
    assert m["readiness"]["equus_uat_execution"].startswith("INCOMPLETE")


@pytest.mark.parametrize("event", [
    _event(name="ObjectRemoved:Delete"),
    _event(key="other/prefix/SYN.xlsx"),
    _event(key="topia/incoming/~$SYN.xlsx"),
    _event(key="topia/incoming/SYN.csv"),
], ids=["not_object_created", "outside_prefix", "excel_lock_file", "wrong_suffix"])
def test_AW01b_events_outside_contract_are_skipped(s3, event):
    results = handler.lambda_handler(event, None)["results"]
    assert all("skipped" in r for r in results) and s3.puts == {}


# --------------------------------------------------------------------------- AW-02 duplicate event

def test_AW02_duplicate_event_and_duplicate_record_write_nothing_new(s3):
    first, second = handler.lambda_handler(_event(n=2), None)["results"]  # same record twice in one event
    assert "skipped" not in first and second["skipped"] == "already processed"
    before = dict(s3.puts)
    again = handler.lambda_handler(_event(), None)["results"][0]  # redelivered notification
    assert again["skipped"] == "already processed" and s3.puts == before
    _complete_set(s3)


# --------------------------------------------------------------------------- AW-03 / AW-04 retry and output failure

@pytest.mark.parametrize("failing", ["_MAPPING_REPORT_", "_EQUUS_DRAFT_", "_RUN_"])
def test_AW03_lambda_retry_after_failed_put_completes_without_duplicates(s3, failing):
    s3.fail_put.append(failing)
    with pytest.raises(S3Error):
        handler.lambda_handler(_event(), None)
    # AW-04 output failure: the error propagates (Lambda retry / on-failure queue),
    # and no import workbook exists without its mapping report.
    if s3.imports():
        assert any("_MAPPING_REPORT_" in k for k in s3.keys())
    assert not any(k.endswith(".json") for k in s3.keys())  # no manifest claims success
    # Retry of the same event (MaximumRetryAttempts: 1 in lambda/template.yaml).
    result = handler.lambda_handler(_event(), None)["results"][0]
    assert "skipped" not in result
    _complete_set(s3)
    # A further redelivery is a no-op.
    before = dict(s3.puts)
    assert handler.lambda_handler(_event(), None)["results"][0]["skipped"] == "already processed"
    assert s3.puts == before


def test_AW04_output_failure_on_first_put_leaves_nothing(s3):
    s3.fail_put.append("_MAPPING_REPORT_")  # first write of the run
    with pytest.raises(S3Error):
        handler.lambda_handler(_event(), None)
    assert s3.puts == {}


def test_AW03b_transient_read_failure_then_retry(s3):
    s3.fail_get = 1
    with pytest.raises(S3Error):
        handler.lambda_handler(_event(), None)
    assert s3.puts == {}
    handler.lambda_handler(_event(), None)
    _complete_set(s3)


# --------------------------------------------------------------------------- AW-05 / AW-06 invalid source / template

def _wrong_sheet() -> bytes:
    wb = fx.load()
    wb[fx.SHEET].title = "Some Other Sheet"
    return fx.save(wb)


@pytest.mark.parametrize("body,fragment", [(_wrong_sheet(), "not found"), (b"PK\x03\x04 corrupt", "not a readable")],
                         ids=["wrong_sheet", "corrupt"])
def test_AW05_invalid_source_is_source_failure_no_import(monkeypatch, body, fragment):
    s3 = _setup(monkeypatch, body)
    with pytest.raises(handler.MigrationError) as info:
        handler.lambda_handler(_event(), None)
    assert info.value.code == handler.FAIL_SOURCE and fragment in str(info.value)
    assert s3.imports() == []
    _, m = s3.manifest()
    assert m["run"]["status"] == "FAILED" and m["output"]["kind"] == "NONE"


def _template_without_sheet1() -> bytes:
    wb = openpyxl.load_workbook(fx.TEMPLATE)
    wb["Sheet1"].title = "Other"
    return fx.save(wb)


def _template_bad_header() -> bytes:
    wb = openpyxl.load_workbook(fx.TEMPLATE)
    wb["Sheet1"]["B3"] = "Assignment Cmd"
    return fx.save(wb)


@pytest.mark.parametrize("template", [_template_without_sheet1(), _template_bad_header(), b"not a workbook"],
                         ids=["no_sheet1", "header_mismatch", "unreadable"])
def test_AW06_invalid_template_is_template_failure_no_import(s3, monkeypatch, template):
    info = {"source": "test", "sha256": hashlib.sha256(template).hexdigest(), "version": "test"}
    monkeypatch.setattr(handler, "_load_template_info", lambda: (template, info))
    with pytest.raises(handler.MigrationError) as err:
        handler.lambda_handler(_event(), None)
    assert err.value.code == handler.FAIL_TEMPLATE
    assert s3.imports() == []
    _, m = s3.manifest()
    assert m["failure"]["code"] == "TEMPLATE" and m["output"]["import_ready"] is False


# --------------------------------------------------------------------------- AW-07 partial validation failure

def test_AW07_partial_validation_failure_held_rows_draft_still_produced(monkeypatch):
    wb = fx.load()
    fx.write_row(wb[fx.SHEET], 6, {"assignment id": None})
    s3 = _setup(monkeypatch, fx.save(wb))
    [result] = handler.lambda_handler(_event(), None)["results"]
    _complete_set(s3)
    [imp] = s3.imports()
    assert "_EQUUS_DRAFT_" in imp
    assert [r[fx.IDENT] for r in fx.inserts(fx.output_rows(s3.puts[("dst", imp)]))] == ["80001"]
    assert s3.meta[("dst", imp)]["rows-held"] == "1" and s3.meta[("dst", imp)]["import-ready"] == "false"
    _, m = s3.manifest()
    assert m["counts"]["held"]["assignment_rows"] == 1 and m["counts"]["accepted"] == 1
    assert any(r["item"] == cfg.EXC_MISSING_ASSIGNMENT_ID for r in m["readiness"]["reasons"])
    assert result["rows_held"] == 1


# --------------------------------------------------------------------------- AW-08 re-run without duplicates

def test_AW08_correction_batch_rerun_imports_only_corrected_rows(monkeypatch):
    """Run 1 holds row 6 (no Assignment ID). The correction batch (procedure
    5b, docs/uat/REPROCESSING_PROCEDURE.md) is built from run 1's Held Rows
    sheet with the identifier corrected at source and uploaded under a new
    key: run 2 writes only 80002; no identifier of run 1 is written again."""
    wb = fx.load()
    fx.write_row(wb[fx.SHEET], 6, {"assignment id": None})
    s3 = _setup(monkeypatch, fx.save(wb))
    handler.lambda_handler(_event(), None)
    report1 = next(v for (_, k), v in s3.puts.items() if "_MAPPING_REPORT_" in k)
    run1 = tmod.accepted_identifiers(report1)
    assert run1 == {"80001"}

    held = openpyxl.load_workbook(io.BytesIO(report1))["Held Rows"]
    batch = fx.load()
    ws = batch[fx.SHEET]
    fx.clear_data(ws)
    for c in range(3, held.max_column + 1):
        ws.cell(5, c - 2).value = held.cell(3, c).value
    fx.write_row(ws, 5, {"assignment id": fx.assignment_id(2)})  # correction at source
    key2 = "topia/incoming/SYN UAT 06 CORRECTION.xlsx"
    s3.objects[("src", key2)] = fx.save(batch)
    handler.lambda_handler(_event(key="topia/incoming/SYN+UAT+06+CORRECTION.xlsx"), None)
    report2 = next(v for (_, k), v in s3.puts.items() if "CORRECTION_MAPPING_REPORT_" in k)
    run2 = tmod.accepted_identifiers(report2)
    assert run2 == {"80002"} and not run1 & run2


def test_AW08b_same_bytes_new_key_detected_by_accepted_identifier_check(monkeypatch):
    """Finding F-02: idempotency is per source key + ETag, so the same bytes
    uploaded under a second key are transformed again. The reprocessing
    procedure step 6 (Accepted Identifiers of run 2 vs run 1) detects it; the
    handler has no previously_accepted input (procedure 5a is local only)."""
    s3 = _setup(monkeypatch, fx.BASE.read_bytes())
    handler.lambda_handler(_event(), None)
    s3.objects[("src", "topia/incoming/SYN COPY.xlsx")] = fx.BASE.read_bytes()
    handler.lambda_handler(_event(key="topia/incoming/SYN+COPY.xlsx"), None)
    reports = [v for (_, k), v in s3.puts.items() if "_MAPPING_REPORT_" in k]
    assert len(reports) == 2 and len(s3.imports()) == 2
    a, b = (tmod.accepted_identifiers(r) for r in reports)
    assert a & b == {"80001", "80002"}  # overlap -> must not be imported twice (procedure step 6)


# --------------------------------------------------------------------------- AC-2 synthetic fixtures

UAT_FIXTURES = (fx.BASE, fx.FAMILY, fx.LINKAGE)
SENSITIVE_LABEL = re.compile(r"name|e-?mail|\bid\b|cost cent|contact|address|phone|initiator|advisor|manager|"
                             r"relocation|leader|benefits|global|pre-hire|first|last|middle", re.I)
EMAIL = re.compile(r"[^@\s]+@[^@\s]+")
REAL_EXTRACT = [ROOT / "local_s3" / "source_s3" / "topia" / "incoming" / "FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx",
                ROOT / "FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx"]


def _data_cells(path: Path):
    """(sheet, column letter, label, value) of every non-empty data cell."""
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        first = fx.FIRST_DATA_ROW if ws.title == fx.SHEET else 2
        header = {c: (fx.norm(ws.cell(4, c).value) or fx.norm(ws.cell(3, c).value) or fx.norm(ws.cell(2, c).value))
                  if ws.title == fx.SHEET else fx.norm(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)}
        for row in ws.iter_rows(min_row=first):
            for cell in row:
                if cell.value is not None:
                    yield ws.title, cell.column_letter, header.get(cell.column, ""), cell.value


def _sensitive_values(path: Path, first_row: int = fx.FIRST_DATA_ROW) -> set:
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)[fx.SHEET]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    labels = [fx.norm(v) for v in rows[3]] if len(rows) > 3 else []
    out = set()
    for row in rows[first_row - 1:]:
        for i, v in enumerate(row):
            if v is None or isinstance(v, bool) or i >= len(labels) or not SENSITIVE_LABEL.search(labels[i] or ""):
                continue
            text = str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip()
            if len(text) >= 4:
                out.add(text.lower())
    return out


def test_uat_fixtures_are_synthetic():
    for path in UAT_FIXTURES:
        for sheet, letter, label, value in _data_cells(path):
            where = f"{path.name}!{sheet} column {letter}"
            if isinstance(value, str) and EMAIL.search(value):
                assert value.endswith("@example.invalid"), where
            if label in ("first name", "last name", "employee full name", "spouse first name", "spouse last name") \
                    or label.endswith("first name") or label.endswith("assignment_dependent!first_name"):
                assert str(value).lower().startswith(("synth", "testcase")), where
            if label == "employee id":
                assert 70000 <= int(value) < 70100, where
            if label == "assignment id" and isinstance(value, int):
                assert 80000 <= value < 80100 or 70000 <= value < 70100, where
            if "cost center" in label:
                assert str(value).startswith(("SYN-", "Synthetic")), where
    # No sensitive value of the older fixture's data rows (derived from the
    # real extract) is reused.
    older = _sensitive_values(fx.STRUCTURE)
    for path in UAT_FIXTURES:
        overlap = _sensitive_values(path) & older
        assert not overlap, f"{path.name}: {len(overlap)} value(s) shared with topia_sample data rows"


def test_uat_fixtures_share_nothing_with_real_extract_when_present():
    """Runs the comparison wherever the git-ignored real extract exists (the
    development checkout); reports counts only, never values."""
    present = [p for p in REAL_EXTRACT if p.exists()]
    for real in present:
        real_values = _sensitive_values(real)
        for path in UAT_FIXTURES:
            overlap = _sensitive_values(path) & real_values
            assert not overlap, f"{path.name}: {len(overlap)} sensitive value(s) shared with the real extract"
    # The synthetic patterns above hold either way (test_uat_fixtures_are_synthetic).


@pytest.mark.parametrize("path", UAT_FIXTURES, ids=lambda p: p.name)
def test_uat_fixture_matches_its_builder(path):
    a = openpyxl.load_workbook(path)
    b = openpyxl.load_workbook(io.BytesIO(fx.BUILDERS[path]()))
    assert a.sheetnames == b.sheetnames
    for name in a.sheetnames:
        assert [list(r) for r in a[name].iter_rows(values_only=True)] == \
               [list(r) for r in b[name].iter_rows(values_only=True)], f"{path.name}!{name}"


# --------------------------------------------------------------------------- AC-1 results document traceability

RESULTS = ROOT / "docs" / "uat" / "TEST_RESULTS.md"


def _result_rows():
    rows = []
    for line in RESULTS.read_text(encoding="utf-8").splitlines():
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if line.startswith("|") and len(cells) >= 5 and re.fullmatch(r"(EA|FM|MD|AW|NX)-\d{2}[a-z]?", cells[0]):
            rows.append(cells)
    return rows


def test_results_doc_every_pass_names_an_existing_test_and_not_executed_names_evidence():
    rows = _result_rows()
    assert len(rows) >= 40
    text = RESULTS.read_text(encoding="utf-8")
    assert "SIGNED OFF" not in text.upper().replace("SIGNED-OFF", "")
    for cells in rows:
        result = cells[-1]
        assert result in ("PASS", "NOT EXECUTED"), cells[0]
        refs = re.findall(r"(tests/test_uat_scenarios_\w+\.py)::(test_\w+)", " ".join(cells))
        if result == "PASS":
            assert refs, f"{cells[0]}: PASS without a test id"
            for file, name in refs:
                assert re.search(rf"^def {name}\(", (ROOT / file).read_text(encoding="utf-8"), re.M), \
                    f"{cells[0]}: {file}::{name} not found"
        else:
            assert cells[0].startswith("NX-") and not refs and len(cells[-2]) > 20, cells[0]
