"""Builders of the synthetic UAT-04 family fixtures (no real values).

Run `./.venv/bin/python tests/family_fixtures.py` to regenerate
tests/fixtures/family_dependents_unkeyed.xlsx and
tests/fixtures/family_dependents_keyed.xlsx from tests/fixtures/topia_sample.xlsx.

- family_dependents_unkeyed.xlsx: the assignment sheet plus a 'Dependent
  Import' sheet laid out like the Topia one (row 1 = two groups of
  ASSIGNMENT_DEPENDENT technical keys, row 2 blank, data in rows 3-4, no
  Employee or Assignment key column): 2 dependent rows, 4 dependent groups.
- family_dependents_keyed.xlsx: source row 6 re-pointed to Employee ID 10001,
  so employee 10001 has two assignments (Assignment ID 90001 and 90002),
  plus a 'Dependents Keyed' sheet: row 1 = 'Employee ID', 'Assignment ID'
  and two repeated child groups; data from row 2 (see KEYED_ROWS).
"""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures"
BASE = FIXTURES / "topia_sample.xlsx"
UNKEYED = FIXTURES / "family_dependents_unkeyed.xlsx"
KEYED = FIXTURES / "family_dependents_keyed.xlsx"
SHEET = "Full Historical Past Assignment"
UNKEYED_SHEET = "Dependent Import"
KEYED_SHEET = "Dependents Keyed"

DEP = "ASSIGNMENT_DEPENDENT!"
DEP_CF = "ASSIGNMENT_DEPENDENT_CF!"
# Same order as the Topia 'Dependent Import' row 1 (A1:I1, repeated J1:R1).
UNKEYED_GROUP = [DEP + "FIRST_NAME", DEP + "LAST_NAME", DEP + "GENDER", DEP + "BIRTH_DATE", DEP + "RELATIONSHIP",
                 DEP + "STATUS", DEP_CF + "CF_TEXT001", DEP_CF + "CF_TEXT003", DEP_CF + "CF_TEXT006"]
KEYED_GROUP = [DEP + "FIRST_NAME", DEP + "LAST_NAME", DEP + "BIRTH_DATE", DEP + "RELATIONSHIP"]


def _child(first, last="Synthetic", born=datetime(2010, 1, 1), rel="Child"):
    return [first, last, born, rel]


NO_CHILD = [None, None, None, None]
BIRTH_DATE_ONLY = [None, None, datetime(2015, 5, 5), None]

# (Employee ID, Assignment ID, child group 1, child group 2)
KEYED_ROWS = [
    (10001, 90001, _child("Child1"), _child("Child2")),       # row 2
    (10001, 90001, _child("Child3"), BIRTH_DATE_ONLY),        # row 3: 4th dependent has only a birth date
    (10001, 90002, _child("ChildB1"), NO_CHILD),              # row 4: other assignment of the same employee
    (None, None, _child("ChildNoKey"), NO_CHILD),             # row 5: no Assignment ID -> held
    (99999, 90002, _child("ChildWrongEmp"), NO_CHILD),        # row 6: Employee ID conflict -> held
    (10001, 10001, _child("ChildByEmpNo"), NO_CHILD),         # row 7: Assignment ID = Employee Number -> unmatched
    (10001, 90001, NO_CHILD, NO_CHILD),                       # row 8: empty groups -> nothing
]


def _labels(ws) -> dict[str, int]:
    return {" ".join(str(ws.cell(4, c).value).split()).lower(): c
            for c in range(1, ws.max_column + 1) if ws.cell(4, c).value is not None}


def _save(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_unkeyed() -> bytes:
    wb = openpyxl.load_workbook(BASE)
    ws = wb.create_sheet(UNKEYED_SHEET)
    for i, key in enumerate(UNKEYED_GROUP * 2, 1):
        ws.cell(1, i, key)
    ws.cell(3, 1, "Child1")
    ws.cell(3, 10, "Child2")
    for g, first in enumerate(("Child3", "Child4")):
        for i, value in enumerate([first, "Synthetic", "Female", datetime(2012, 2, 2), "Child", "On Assignment",
                                   "Synthetic A", "Synthetic B", "Synthetic C"]):
            ws.cell(4, 1 + 9 * g + i, value)
    return _save(wb)


def build_keyed() -> bytes:
    wb = openpyxl.load_workbook(BASE)
    src = wb[SHEET]
    src.cell(6, _labels(src)["employee id"], 10001)
    ws = wb.create_sheet(KEYED_SHEET)
    for i, header in enumerate(["Employee ID", "Assignment ID"] + KEYED_GROUP * 2, 1):
        ws.cell(1, i, header)
    for r, (emp, asg, g1, g2) in enumerate(KEYED_ROWS, 2):
        for i, value in enumerate([emp, asg] + g1 + g2, 1):
            ws.cell(r, i, value)
    return _save(wb)


if __name__ == "__main__":
    UNKEYED.write_bytes(build_unkeyed())
    KEYED.write_bytes(build_keyed())
    print(f"wrote {UNKEYED.name}, {KEYED.name}")
