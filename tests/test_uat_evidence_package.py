"""UAT-08 (OCNORA T-1393): consistency checks for the final UAT evidence
package under docs/uat/ (clarifications/UAT_MASTER_COMMAND.md sections 10,
12 and 15). Documentation tests only: no production code is exercised except
the default-configuration transform of the synthetic fixture
tests/fixtures/topia_sample.xlsx, to prove the field-mapping table still
matches the code.

AC-1: every section-12 deliverable exists or is marked INCOMPLETE with the
      missing input named.
AC-2: UAT_READINESS_ASSESSMENT.md marks G1-G12 Pass/Fail/Blocked with
      evidence and ends with exactly one of the four statuses, consistent
      with the gates.
AC-3: CHANGE_LOG.md lists every function changed in UAT-02..07 with
      before/after and an existing test id.
AC-4: no document claims an Equus import happened or that anyone signed off.
"""

import ast
import io
import re
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
UAT = ROOT / "docs" / "uat"
sys.path.insert(0, str(ROOT / "lambda" / "src"))

import transform as tmod  # noqa: E402

STATUSES = ("NOT READY FOR UAT", "READY FOR UAT EXECUTION", "UAT EXECUTED — SIGN-OFF PENDING", "UAT SIGNED OFF")
PRE_IMPORT_GATES = ("G1", "G2", "G3", "G4", "G5", "G9", "G10")
EQUUS_GATES = ("G6", "G7", "G8")
NEGATION = re.compile(r"(?i)\b(never|not|no|nothing|none|cannot|would|before|until|must|only|empty)\b|·")
TEST_REF = re.compile(r"(tests/[\w/]+\.py)::(test_\w+)")


def _read(name: str) -> str:
    return (UAT / name).read_text(encoding="utf-8")


def _cells(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip("|").split("|")]


def _test_exists(file: str, name: str) -> bool:
    path = ROOT / file
    return path.exists() and re.search(rf"^def {name}\(", path.read_text(encoding="utf-8"), re.M) is not None


def _assert_test_refs_exist(text: str, where: str) -> int:
    refs = TEST_REF.findall(text)
    for file, name in refs:
        assert _test_exists(file, name), f"{where}: {file}::{name} does not exist"
    return len(refs)


# --------------------------------------------------------------------------- AC-1 section 12 deliverables

def test_section12_deliverables_exist_or_are_marked_incomplete():
    text = _read("README.md")
    rows = [_cells(line) for line in text.splitlines()
            if line.startswith("|") and re.fullmatch(r"\d{1,2}", _cells(line)[0])]
    assert [int(r[0]) for r in rows] == list(range(1, 13)), "README must list the 12 deliverables of MC section 12"
    for number, deliverable, documents, status, missing in rows:
        docs = re.findall(r"`([\w\-]+\.md)`", documents)
        paths = re.findall(r"`((?:lambda|local)/[\w/.\-*]+)`", documents)
        assert docs or paths, f"deliverable {number} names no document"
        for doc in docs:
            assert (UAT / doc).exists(), f"deliverable {number}: docs/uat/{doc} missing"
        for path in paths:
            if "*" not in path:
                assert (ROOT / path).exists(), f"deliverable {number}: {path} missing"
        if "INCOMPLETE" in status:
            assert missing not in ("", "—", "-") and len(missing) > 10, f"deliverable {number}: missing input not named"
    by_number = {int(r[0]): r for r in rows}
    # No Equus access and no signatory: items 8 and 12 can only be INCOMPLETE.
    assert "INCOMPLETE" in by_number[8][3] and "Equus" in by_number[8][4]
    assert "INCOMPLETE" in by_number[12][3] and "Q31" in by_number[12][4]


def test_equus_execution_plan_is_not_executed_and_evidence_fields_are_empty():
    text = _read("EQUUS_UAT_EXECUTION_PLAN.md")
    assert "**STATUS: NOT EXECUTED.**" in text
    steps = re.findall(r"^## Step (\d) — ", text, re.M)
    assert steps == ["1", "2", "3", "4", "5"]
    in_evidence = False
    for line in text.splitlines():
        if line.startswith("| Evidence field |"):
            in_evidence = True
            continue
        if not line.startswith("|"):
            in_evidence = False
            continue
        if in_evidence and not line.startswith("|---"):
            field, value = _cells(line)
            assert value in ("", "NOT EXECUTED"), f"evidence field filled: {field}"


def test_business_signoff_record_is_an_empty_template():
    text = _read("BUSINESS_SIGNOFF_RECORD.md")
    assert "empty template" in text and "No decision has been recorded" in text
    for heading in ("Signatory", "Scope", "Date", "exceptions", "Decision"):
        assert re.search(rf"^## \d\. .*{heading}", text, re.M | re.I), heading
    for line in text.splitlines():
        if line.startswith("|") and not line.startswith("|---"):
            cells = _cells(line)
            if cells[0] in ("Field", "Exception ID"):
                continue
            assert all(c == "" for c in cells[1:]), f"sign-off record filled: {line}"


# --------------------------------------------------------------------------- AC-2 readiness assessment

def _gate_rows() -> dict[str, list[str]]:
    rows = {}
    for line in _read("UAT_READINESS_ASSESSMENT.md").splitlines():
        cells = _cells(line) if line.startswith("|") else []
        if cells and re.fullmatch(r"G\d{1,2}", cells[0]):
            assert cells[0] not in rows, f"{cells[0]} listed twice"
            rows[cells[0]] = cells
    return rows


def _gate_result(cells: list[str]) -> str:
    return re.sub(r"\*|\(.*\)", "", cells[2]).strip()


def test_readiness_assessment_marks_every_gate_with_evidence():
    rows = _gate_rows()
    assert sorted(rows, key=lambda g: int(g[1:])) == [f"G{i}" for i in range(1, 13)]
    for gate, cells in rows.items():
        assert _gate_result(cells) in ("Pass", "Fail", "Blocked"), gate
        evidence = cells[3]
        assert re.search(r"`[^`]+(\.py|\.md|\.json|\.xlsx)(:\d+[-\d]*)?`|JAKE!|EX!|::test_", evidence), \
            f"{gate}: no file / cell / test evidence"
        _assert_test_refs_exist(evidence, gate)
    # No Equus access: the Equus gates can never be Pass from local evidence.
    for gate in EQUUS_GATES + ("G11", "G12"):
        assert _gate_result(rows[gate]) != "Pass", gate


def test_readiness_assessment_has_exactly_one_final_status_consistent_with_gates():
    text = _read("UAT_READINESS_ASSESSMENT.md")
    finals = re.findall(r"^Final readiness status: \*\*(.+?)\*\*\s*$", text, re.M)
    assert len(finals) == 1, finals
    status = finals[0]
    assert status in STATUSES and status != "UAT SIGNED OFF"
    rows = _gate_rows()
    if any(_gate_result(rows[g]) != "Pass" for g in PRE_IMPORT_GATES):
        assert status == "NOT READY FOR UAT"
    if any(_gate_result(rows[g]) != "Pass" for g in EQUUS_GATES):
        assert status in ("NOT READY FOR UAT", "READY FOR UAT EXECUTION")
    # The section-15 answer and the index carry the same single status.
    assert re.search(rf"^\*\*{re.escape(status)}\.\*\*$", _read("FINAL_RESPONSE.md"), re.M)
    assert f"**Readiness status: {status}**" in _read("README.md")


# --------------------------------------------------------------------------- AC-3 change log

TASK_COMMITS = (  # (section title, commit, parent) per docs/uat/CHANGE_LOG.md section 0
    ("UAT-02", "3f92551", "e61d99f"),
    ("UAT-04", "66b7e25", "7d86669"),
    ("UAT-03", "ddf01f5", "66b7e25"),
    ("UAT-07", "452e9e0", "ddf01f5"),
    ("UAT-05", "b661f43", "ddf01f5"),
    ("UAT-06", "7b71503", "146d6ab"),
)
CODE_FILES = ("lambda/src/transform.py", "lambda/src/mapping_config.py", "lambda/src/handler.py",
              "local/reconcile.py", "tests/test_migration.py", "tests/test_business_rules.py",
              "tests/test_identity_controls.py")


def _change_log_sections() -> dict[str, str]:
    text = _read("CHANGE_LOG.md")
    parts = re.split(r"^## \d\. (UAT-0\d) .*$", text, flags=re.M)
    return dict(zip(parts[1::2], parts[2::2]))


def test_change_log_rows_have_before_after_and_existing_test_ids():
    sections = _change_log_sections()
    assert sorted(sections) == ["UAT-02", "UAT-03", "UAT-04", "UAT-05", "UAT-06", "UAT-07"]
    n = 0
    for task, body in sections.items():
        for line in body.splitlines():
            cells = _cells(line) if line.startswith("|") else []
            if not cells or not re.fullmatch(r"0\d-\d\d", cells[0]):
                continue
            n += 1
            assert len(cells) == 7, f"{task} {cells[0]}: expected 7 columns"
            _, file, element, before, after, reason, evidence = cells
            assert element and before and after and reason, f"{cells[0]}: empty cell"
            assert _assert_test_refs_exist(evidence, f"{task} {cells[0]}") >= 1, f"{cells[0]}: no test id"
    assert n >= 80


def _git(*args: str) -> str | None:
    try:
        out = subprocess.run(["git", "-C", str(ROOT), *args], capture_output=True, text=True, timeout=60)
    except (OSError, subprocess.SubprocessError):
        return None
    return out.stdout if out.returncode == 0 else None


def _definitions(source: str | None) -> dict[str, str]:
    """Functions, methods and classes. A class is represented by its
    non-method body (fields, decorators, bases) so that a changed method is
    attributed to the method and a changed field to the class."""
    found: dict[str, str] = {}
    if source is None:
        return found

    def walk(nodes, prefix=""):
        for node in nodes:
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                found[prefix + node.name] = ast.dump(node)
            elif isinstance(node, ast.ClassDef):
                shell = [ast.dump(n) for n in node.body if not isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]
                found[prefix + node.name] = repr((shell, [ast.dump(d) for d in node.decorator_list],
                                                  [ast.dump(b) for b in node.bases]))
                walk(node.body, prefix + node.name + ".")

    walk(ast.parse(source).body)
    return found


def test_change_log_lists_every_changed_function():
    if _git("cat-file", "-e", "3f92551^{commit}") is None or _git("cat-file", "-e", "e61d99f^{commit}") is None:
        pytest.skip("git history of the UAT-02..07 commits is not available in this checkout")
    sections = _change_log_sections()
    missing = []
    for task, commit, parent in TASK_COMMITS:
        body = sections[task]
        for file in CODE_FILES:
            before = _definitions(_git("show", f"{parent}:{file}"))
            after = _definitions(_git("show", f"{commit}:{file}"))
            for name, dump in after.items():
                if file.startswith("tests/") and name not in before:
                    continue  # a new test or helper in a test file is evidence, not a behaviour change
                if before.get(name) != dump and f"`{name}`" not in body:
                    missing.append(f"{task} {file}: {name}")
    assert not missing, "functions changed but not in docs/uat/CHANGE_LOG.md: " + ", ".join(missing)


# --------------------------------------------------------------------------- field mapping validation

def test_field_mapping_validation_matches_fixture_run():
    source = (ROOT / "tests" / "fixtures" / "topia_sample.xlsx").read_bytes()
    template = (ROOT / "lambda" / "src" / "templates" / "equus_eafh_template.xlsx").read_bytes()
    _, report_bytes, report = tmod.transform(source, template)
    rows = {}
    for line in _read("FIELD_MAPPING_VALIDATION.md").splitlines():
        cells = _cells(line) if line.startswith("|") else []
        if len(cells) == 9 and re.fullmatch(r"[A-Z]{1,2}", cells[0]):
            rows[cells[0]] = cells
    assert set(rows) == {c.letter for c in report.columns}, "every source column letter must be listed once"
    for col in report.columns:
        cells = rows[col.letter]
        assert cells[3] == col.status, col.letter
        header = re.match(r"`([^`]+)`", cells[4])
        assert (header.group(1) if header else None) == (col.target if col.status == "mapped" else None), col.letter
        assert cells[6], f"{col.letter}: no rule status"
        if col.status in ("held", "unmapped", "ignored", "excluded"):
            assert re.search(r"Q\d", cells[6]), f"{col.letter}: pending/held column without a Q number"
    assert report.readiness.status == tmod.NOT_READY
    assert openpyxl.load_workbook(io.BytesIO(report_bytes))["Column Mapping"].max_row == len(report.columns) + 1


# --------------------------------------------------------------------------- AC-4 no Equus or sign-off claims

CLAIM_PATTERNS = (
    re.compile(r"SIGNED[ -]OFF"),
    re.compile(r"(?i)\b(is|was|been|are|were|has|have|be)\s+signed[ -]off\b|\bsigned[ -]off\s+(by|on)\b"),
    re.compile(r"(?i)\buat executed\b"),
    re.compile(r"(?i)\bimport(ed)?\b.{0,40}\b(completed|succeeded|performed|happened|was run)\b"),
    re.compile(r"(?i)\buat (passed|accepted|approved)\b"),
)


def test_no_document_claims_an_equus_import_or_a_signoff():
    docs = sorted(UAT.glob("*.md"))
    assert len(docs) >= 18
    for path in docs:
        lines = path.read_text(encoding="utf-8").splitlines()
        for i, line in enumerate(lines):
            window = (lines[i - 1] if i else "") + " " + line
            if line.lstrip().startswith("- [ ]"):
                continue  # an unticked checklist item is a requirement, not a claim
            if line.startswith("|") and len(_cells(line)) == 2 and re.fullmatch(r"G\d{1,2}", _cells(line)[0]):
                continue  # gate legend row: the gate's definition (MC section 10), not a result
            for pattern in CLAIM_PATTERNS:
                if pattern.search(line):
                    assert NEGATION.search(window), f"{path.name}:{i + 1}: unqualified claim: {line.strip()}"
        text = "\n".join(lines)
        assert not re.search(r"(?i)status\W{0,6}\**\s*UAT SIGNED OFF", text), path.name
        assert not re.search(r"(?i)final readiness status: \*\*UAT", text), path.name
