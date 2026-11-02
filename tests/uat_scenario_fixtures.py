"""Builders and helpers of the synthetic UAT-06 scenario fixtures.

Every data value written here is invented: names start with 'Synth' or
'Testcase', e-mails end with '@example.invalid', Employee IDs are 7000x,
Assignment IDs 8000x, cost centres 'SYN-...'. Header rows 1-4 (labels, row-2
technical keys, colour fills) are taken from tests/fixtures/topia_sample.xlsx
(the tracked structural fixture); its data rows are discarded, so no value of
the real extract or of topia_sample's data rows is reused
(tests/test_uat_scenarios_aws.py::test_uat_fixtures_are_synthetic).

Run `./.venv/bin/python tests/uat_scenario_fixtures.py` to regenerate:
- tests/fixtures/uat_synthetic_base.xlsx   two employees, one assignment each
- tests/fixtures/uat_family.xlsx           family scenarios F-01..F-09 (FAMILY_ROWS)
- tests/fixtures/uat_family_linkage.xlsx   keyed dependent sheet for F-10 (LINKAGE_ROWS)
"""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures"
STRUCTURE = FIXTURES / "topia_sample.xlsx"
BASE = FIXTURES / "uat_synthetic_base.xlsx"
FAMILY = FIXTURES / "uat_family.xlsx"
LINKAGE = FIXTURES / "uat_family_linkage.xlsx"
SHEET = "Full Historical Past Assignment"
KEYED_SHEET = "Dependents Keyed"
FIRST_DATA_ROW = 5

DEP = "ASSIGNMENT_DEPENDENT!"
SPOUSE = "SPOUSE_DOMESTIC_PARTNER!"
DEP_GROUP = [("FIRST_NAME", "First Name"), ("LAST_NAME", "Last Name"), ("BIRTH_DATE", "Date of Birth"),
             ("RELATIONSHIP", "Relationship")]
N_DEP_GROUPS = 4


def employee_id(n: int) -> int:
    return 70000 + n


def assignment_id(n: int) -> int:
    return 80000 + n


def synthetic_row(n: int) -> dict:
    """Values by normalised Topia row-4 label for synthetic person `n`."""
    letter = chr(96 + n)  # 1 -> a
    return {
        "employee full name": f"TESTCASE, Synth{letter}",
        "first name": f"Synth{letter}",
        "last name": "Testcase",
        "employee id": employee_id(n),
        "email": f"synth.{letter}@example.invalid",
        "alternate email": f"synth.{letter}.alt@example.invalid",
        "gender": "Female" if n % 2 else "Male",
        "date of birth": datetime(1990, 1, n),
        "marital status": "Single",
        "nationality 1": "Synthnational",
        "nationality 2": "Not Provided",
        "nationality 3": "Not Provided",
        "citizenship 1": "Canada",
        "country of birth": "Canada",
        "person type": "employee",
        "person status": "employee",
        "business title": "Synthetic Title",
        "employment status": "Employed",
        "assignment id": assignment_id(n),
        "assignment type": "Synthetic Type",
        "assignment type category": "Synthetic Policy",
        "assignment status": "Active",
        "status category": "Active",
        "confidentiality type": "Synthetic Non-Confidential",
        "is confidential": False,
        "home pay installments": 12,
        "host pay installments": 12,
        "assignment start date": datetime(2020, 1, 1),
        "assignment end date": datetime(2021, 1, 1),
        "arrival date": datetime(2020, 1, 1),
        "home country": "Canada",
        "home city": "Synthcity",
        "home state": "Ontario",
        "host country": "Germany",
        "host city": "Synthburg",
        "host state": "Bavaria",
        "home cost center code": f"SYN-H{n:03d}_NVC",
        "home cost center name": "Synthetic Home Centre",
        "host cost center code": f"SYN-T{n:03d}_NVC",
        "host cost center name": "Synthetic Host Centre",
        "mobility cost center code": f"SYN-M{n:03d}_DVC",
        "mobility cost center name": "Synthetic Mobility Centre",
        "expense cost center code": f"SYN-E{n:03d}_NVC",
        "expense cost center name": "Synthetic Expense Centre",
        "home entity": "Synthetic Entity Home",
        "home division": "Synthetic Division",
        "home business line": "Synthetic Line",
        "home role region": "Synthetic Region",
        "home job title": "Synthetic Job",
        "home job level": 5,
        "home job profile": "SYN_PROFILE",
        "host entity": "Synthetic Entity Host",
        "host division": "Synthetic Division",
        "host business line": "Synthetic Line",
        "host role region": "Synthetic Region",
        "host job title": "Synthetic Job",
        "host job level": 6,
        "host job profile": "SYN_PROFILE",
        "home scheduled weekly hours": 40,
        "host scheduled weekly hours": 40,
        "host office location": "Synthetic Office",
        "host supervisory organization": "SYN-ORG",
        "end of assignment reason": "Synthetic Reason",
        "expected assignment outcome": "Synthetic Outcome",
        "vip flag": "No",
        "global id": 990000000000000 + n,
        "pre-hire id": f"SYNPH{n:03d}",
        "pii nationality 1": "Synthnational",
        "primary mobility contact": "Synthcoord One",
        "primary mobility email": "synth.coord@example.invalid",
        "group name": "Synthetic Group",
        "move initiator name": "Synthinit One",
        "move initiator email": "synth.init@example.invalid",
    }


# --------------------------------------------------------------------------- helpers shared by the tests

def norm(value) -> str:
    return " ".join(str(value).split()).lower() if value is not None else ""


def labels(ws) -> dict[str, int]:
    """Normalised row-4 label (row-3 label when row 4 is empty) -> column; first wins."""
    out: dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        label = norm(ws.cell(4, c).value) or norm(ws.cell(3, c).value)
        if label and label not in out:
            out[label] = c
    return out


def save(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def load(path: Path = BASE):
    return openpyxl.load_workbook(path)


def write_row(ws, r: int, values: dict) -> None:
    """Write values by label; a label missing from the sheet is an error."""
    cols = labels(ws)
    for label, value in values.items():
        if label not in cols:
            raise KeyError(f"label {label!r} not in fixture")
        ws.cell(r, cols[label]).value = value


def set_person(ws, r: int, n: int, **overrides) -> None:
    """Row r = synthetic person n (all known labels), then `overrides`
    (keyword names: label with spaces replaced by '_')."""
    values = synthetic_row(n)
    values.update({k.replace("_", " "): v for k, v in overrides.items()})
    write_row(ws, r, values)


def clear_data(ws) -> None:
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None


def add_dependent_groups(ws, n_groups: int = N_DEP_GROUPS) -> int:
    """Append n_groups repeated ASSIGNMENT_DEPENDENT column groups (Child k ...);
    returns the first appended column."""
    start = ws.max_column + 1
    for g in range(n_groups):
        for i, (key, label) in enumerate(DEP_GROUP):
            ws.cell(2, start + len(DEP_GROUP) * g + i, DEP + key)
            ws.cell(4, start + len(DEP_GROUP) * g + i, f"Child {g + 1} {label}")
    return start


def set_dependents(ws, r: int, start: int, groups: list[list]) -> None:
    for g, values in enumerate(groups):
        for i, value in enumerate(values):
            ws.cell(r, start + len(DEP_GROUP) * g + i).value = value


def child(first: str, born=datetime(2015, 1, 1), rel: str = "Child") -> list:
    return [first, "Testcase", born, rel]


# --------------------------------------------------------------------------- fixture builders

def build_base() -> bytes:
    wb = openpyxl.load_workbook(STRUCTURE)
    ws = wb[SHEET]
    clear_data(ws)
    set_person(ws, 5, 1)
    set_person(ws, 6, 2)
    return save(wb)


# source row -> (person n, overrides, spouse values by label, dependent groups); see module docstring.
FAMILY_ROWS = {
    5: (1, {}, {}, []),                                                                    # F-01 none
    6: (2, {}, {"spouse first name": "Synthspouse", "spouse last name": "Testcase",
                "spouse relationship": "Spouse"},
        [child("Synthkidb1")]),                                                            # F-02, F-05
    7: (3, {}, {"spouse relationship": "Partner", "spouse date of birth": datetime(1991, 3, 3)},
        [child("Synthkidc1"), child("Synthkidc2", datetime(2017, 2, 2))]),                 # F-03, F-06
    8: (4, {}, {},
        [child("Synthkidd1"), child("Synthkidd2", datetime(2016, 6, 6)),
         child("Synthkidd3", datetime(2018, 8, 8)), [None, None, datetime(2020, 2, 2), None]]),  # F-04, F-07
    9: (4, {"assignment id": assignment_id(5), "assignment start date": datetime(2022, 1, 1),
            "assignment end date": datetime(2023, 1, 1), "arrival date": datetime(2022, 1, 1)}, {},
        [child("Synthkide1", datetime(2019, 9, 9))]),                                      # F-08
    10: (6, {}, {}, [child("Synthkidf1"), child("Synthkidf1"), child("Synthkidf2")]),     # F-09
}
FAMILY_EXPECTED = {  # source row -> (dependent records written, spouse written, held duplicates)
    5: (0, False, 0), 6: (1, True, 0), 7: (2, True, 0), 8: (4, False, 0), 9: (1, False, 0), 10: (2, False, 1),
}


def build_family() -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(build_base()))
    ws = wb[SHEET]
    clear_data(ws)
    start = add_dependent_groups(ws)
    for r, (n, overrides, spouse, deps) in FAMILY_ROWS.items():
        set_person(ws, r, n, **{k.replace(" ", "_"): v for k, v in overrides.items()})
        write_row(ws, r, spouse)
        set_dependents(ws, r, start, deps)
    return save(wb)


# (Employee ID, Assignment ID, child group) rows of the keyed dependent sheet (F-10).
LINKAGE_ROWS = [
    (employee_id(1), assignment_id(1), child("Synthlinka")),   # row 2: linked to 80001
    (None, None, child("Synthlinkb")),                         # row 3: no Assignment ID -> held
    (employee_id(1), assignment_id(99), child("Synthlinkc")),  # row 4: unknown Assignment ID -> held
    (employee_id(9), assignment_id(2), child("Synthlinkd")),   # row 5: other Employee ID than 80002 -> held
    (employee_id(1), employee_id(1), child("Synthlinke")),     # row 6: Employee Number as Assignment ID -> held
]


def build_linkage() -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(build_base()))
    ws = wb.create_sheet(KEYED_SHEET)
    for i, header in enumerate(["Employee ID", "Assignment ID"] + [DEP + k for k, _ in DEP_GROUP], 1):
        ws.cell(1, i, header)
    for r, (emp, asg, group) in enumerate(LINKAGE_ROWS, 2):
        for i, value in enumerate([emp, asg] + group, 1):
            ws.cell(r, i, value)
    return save(wb)


BUILDERS = {BASE: build_base, FAMILY: build_family, LINKAGE: build_linkage}


# --------------------------------------------------------------------------- output readers for the tests

TEMPLATE = ROOT / "lambda" / "src" / "templates" / "equus_eafh_template.xlsx"
IDENT = "Assignment Identifier"


def output_rows(xlsx: bytes) -> list[dict]:
    """Import workbook data rows as {row-3 header: value} (non-empty cells)."""
    ws = openpyxl.load_workbook(io.BytesIO(xlsx))["Sheet1"]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    return [{h: ws.cell(r, c).value for c, h in enumerate(headers, 1) if ws.cell(r, c).value is not None}
            for r in range(4, ws.max_row + 1)
            if any(ws.cell(r, c).value is not None for c in range(1, ws.max_column + 1))]


def inserts(rows: list[dict]) -> list[dict]:
    return [r for r in rows if r.get("Assignment Command") == "INSERT"]


def block(rows: list[dict], identifier: str) -> list[dict]:
    return [r for r in rows if str(r.get(IDENT)) == identifier]


def sheet_table(report_bytes: bytes, sheet: str, header_row: int = 1) -> list[dict]:
    ws = openpyxl.load_workbook(io.BytesIO(report_bytes))[sheet]
    header = [c.value for c in ws[header_row]]
    return [dict(zip(header, r)) for r in ws.iter_rows(min_row=header_row + 1, values_only=True)
            if any(v is not None for v in r)]


def codes(report, code: str) -> list[dict]:
    return [e for e in report.exceptions if e["code"] == code]


def warnings_for(report, fragment: str) -> list[tuple]:
    return [w for w in report.warnings if fragment in w[2]]


def reasons(report, category: str | None = None) -> list[dict]:
    return [r for r in report.readiness.reasons if category is None or r["category"] == category]


if __name__ == "__main__":
    for path, build in BUILDERS.items():
        path.write_bytes(build())
        print(f"wrote {path.relative_to(ROOT)}")
