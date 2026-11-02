"""UAT-09 (T-1394, PC-ADM-01): the RTM quotes Jake verbatim from the workbook.

Compares each Q1-Q43 row of docs/uat/RTM_Q1-Q43.md §1 with
clarifications/jake-1.xlsx, sheet 'datamapping clarrification shef'
(Qn on row n+1; status in column I, answer in column J). No answer text is
hard-coded here: the expected value is always read from the cell.
"""
import re
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
JAKE = ROOT / "clarifications" / "jake-1.xlsx"
RTM = ROOT / "docs" / "uat" / "RTM_Q1-Q43.md"
PENDING = ROOT / "docs" / "uat" / "PENDING_CLARIFICATIONS.md"
SHEET = "datamapping clarrification shef"


def _norm(text):
    return re.sub(r"\s+", " ", text).strip()


@pytest.fixture(scope="module")
def sheet():
    return openpyxl.load_workbook(JAKE, read_only=False, data_only=True)[SHEET]


@pytest.fixture(scope="module")
def matrix_answers():
    """Q number -> 'Jake's answer (cell)' column of the §1 matrix."""
    text = RTM.read_text(encoding="utf-8")
    section = text.split("## 1. Matrix", 1)[1].split("## 2.", 1)[0]
    answers = {}
    for line in section.splitlines():
        m = re.match(r"\| Q(\d+) \| ", line)
        if m:
            cells = line.split(" | ")
            answers[int(m.group(1))] = cells[2]
    assert sorted(answers) == list(range(1, 44))
    return answers


def test_row_rule_and_every_status_is_open(sheet):
    for n in range(1, 44):
        assert sheet[f"A{n + 1}"].value == f"Q{n}"
        assert sheet[f"I{n + 1}"].value == "Open"


@pytest.mark.parametrize("n", range(1, 21))
def test_q1_q20_quote_is_exact_column_j_text(sheet, matrix_answers, n):
    cell = sheet[f"J{n + 1}"].value
    assert cell, f"J{n + 1} unexpectedly empty"
    answer = matrix_answers[n]
    m = re.match(r'`JAKE!J(\d+)`: "(.*)"', answer)
    assert m, answer
    assert int(m.group(1)) == n + 1
    assert _norm(m.group(2)) == _norm(cell)


@pytest.mark.parametrize("n", range(21, 44))
def test_q21_q43_state_j_cell_empty(sheet, matrix_answers, n):
    assert sheet[f"J{n + 1}"].value is None
    assert f"J{n + 1} is empty (no answer)" in matrix_answers[n]


def test_no_condensed_markers_and_pc_adm_01_closed():
    rtm = RTM.read_text(encoding="utf-8")
    assert "[condensed per ASSESS]" not in rtm
    assert "[verbatim per ASSESS" not in rtm
    assert "Quotation limitation" not in rtm
    row = next(l for l in PENDING.read_text(encoding="utf-8").splitlines()
               if l.startswith("| PC-ADM-01 |"))
    assert re.search(r"CLOSED \d{4}-\d{2}-\d{2}", row)
