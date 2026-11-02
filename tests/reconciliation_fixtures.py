"""Builder of the synthetic UAT-05 reconciliation fixture (no real values).

Run `./.venv/bin/python tests/reconciliation_fixtures.py` to regenerate
tests/fixtures/reconciliation_children.xlsx from tests/fixtures/topia_sample.xlsx.

Source rows (sheet 'Full Historical Past Assignment'):
- row 5 (Assignment ID 90001): two addresses (FJ with four lines, FN), two
  contacts (FR, FS) and three dependents in appended Child 1-3 columns (the
  third has only a birth date, Q20)  -> 1 INSERT + 2 UPDATE rows expected.
- row 6 (Assignment ID 90002): one address (FL), a placeholder contact
  ('N/A' in FS -> blank token), Child 1 holds only a 'Not Provided'
  placeholder (no record), Child 2 a first name only -> 1 INSERT row.
Expected child records: address 3, contact 2, dependent 4; expansion rows 2.
"""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures"
BASE = FIXTURES / "topia_sample.xlsx"
CHILDREN = FIXTURES / "reconciliation_children.xlsx"
SHEET = "Full Historical Past Assignment"

DEP = "ASSIGNMENT_DEPENDENT!"
GROUP = [("FIRST_NAME", "First Name"), ("LAST_NAME", "Last Name"), ("BIRTH_DATE", "Date of Birth"),
         ("RELATIONSHIP", "Relationship")]

ADDRESSES = {
    5: {"personal home address with effective date": "1 Synthetic Lane\nUnit 2\nSynthville\n00000",
        "work host address with effective date": "9 Example Park"},
    6: {"personal host address with effective date": "2 Sample Road"},
}
CONTACTS = {
    5: {"phone number": "5550001", "mobile phone number": 5550002.0},
    6: {"mobile phone number": "N/A"},
}
DEPENDENTS = {
    5: [["SynthA", "Testcase", datetime(2010, 1, 1), "Child"],
        ["SynthB", "Testcase", datetime(2012, 2, 2), "Child"],
        [None, None, datetime(2015, 5, 5), None]],
    6: [["Not Provided", None, None, None],
        ["SynthD", None, None, None],
        [None, None, None, None]],
}
EXPECTED = {"address": 3, "contact": 2, "dependent": 4, "expansion_rows": 2,
            "rows": {5: {"address": 2, "contact": 2, "dependent": 3, "output_rows": 3},
                     6: {"address": 1, "contact": 0, "dependent": 1, "output_rows": 1}}}


def _labels(ws) -> dict[str, int]:
    return {" ".join(str(ws.cell(4, c).value).split()).lower(): c
            for c in range(1, ws.max_column + 1) if ws.cell(4, c).value is not None}


def build_children() -> bytes:
    wb = openpyxl.load_workbook(BASE)
    ws = wb[SHEET]
    labels = _labels(ws)
    for values in (ADDRESSES, CONTACTS):
        for r, cells in values.items():
            for label, value in cells.items():
                ws.cell(r, labels[label], value)
    start = ws.max_column + 1
    n_groups = max(len(g) for g in DEPENDENTS.values())
    for g in range(n_groups):
        for i, (key, label) in enumerate(GROUP):
            ws.cell(2, start + len(GROUP) * g + i, DEP + key)
            ws.cell(4, start + len(GROUP) * g + i, f"Child {g + 1} {label}")
    for r, groups in DEPENDENTS.items():
        for g, values in enumerate(groups):
            for i, value in enumerate(values):
                ws.cell(r, start + len(GROUP) * g + i, value)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


if __name__ == "__main__":
    CHILDREN.write_bytes(build_children())
    print(f"wrote {CHILDREN.relative_to(ROOT)}")
