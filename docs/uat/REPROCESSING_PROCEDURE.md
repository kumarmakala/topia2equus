# Reprocessing procedure — held rows, corrections and re-runs (UAT-03, BR-05)

Task: OCNORA T-1388. Sources: `clarifications/UAT_MASTER_COMMAND.md` §3 BR-05 ("rejection and
correction reporting, controlled reprocessing without duplicating accepted rows"),
`clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md` Q5 row and D1/D2, Jake's Q5 answer
(`clarifications/jake-1.xlsx` sheet `datamapping clarrification shef`!J6: duplicates can create
duplicates, protections are not a fail-safe, row-level rejection, header issues fail the whole
file, test first).

**Status:** procedure IMPLEMENTED AND TESTED locally with synthetic data only. Equus import
behaviour (what Equus does with a re-submitted identifier) is **NOT EXECUTED / INCOMPLETE** — no
Equus access from this environment. Nothing here is SIGNED OFF.

## 1. What a run produces

| Artifact | Content | Code |
|---|---|---|
| Import workbook (`*_EQUUS_DRAFT_*.xlsx` while `NOT_READY`) | **accepted rows only** — every held row is absent | `lambda/src/transform.py:963` (`build_records` skips `report.held_rows`) |
| Report sheet **Exceptions** | one line per finding: Source Row, Assignment Identifier, Employee Number, Code, Level, Message, Owner, Q | `transform.py:1415`, `:1421` |
| Report sheet **Held Rows** | row 1 source column letters, row 2 source labels, then for every held row: source row number, codes, and the **raw source cells A..last column, unchanged** | `transform.py:1428` |
| Report sheet **Accepted Identifiers** | Source Row, Assignment Identifier (text), Employee Number of every row written to the import workbook | `transform.py:1443` |
| Report sheet **Template Check** | TEMPLATE_VERSION, approved / supplied template SHA-256, differing header cells, pending header decisions (D2) | `transform.py:1451` |
| `reports/*_FILE_REJECTED_*.xlsx` (handler) | report only; the import workbook is **withheld** when any FILE-level finding exists | `lambda/src/handler.py:188-194` |

Levels (`mapping_config.EXCEPTION_LEVELS`): **FILE** = whole import file withheld (`transform()`
returns `b""` for the workbook, `transform.py:1628`; handler raises so the event goes to the
on-failure queue); **ROW_HELD** = row excluded from the import workbook and listed in Exceptions
and Held Rows; **WARNING** = reported, value and row kept.

Codes and their configured level / Q / owner: `lambda/src/business_rules.json`
`identity_controls.exception_codes` (validated by `mapping_config._load_identity_controls`).

| Code | Level | Q | Owner (configured) | Meaning |
|---|---|---|---|---|
| MISSING_ASSIGNMENT_ID | ROW_HELD | Q1 | Alka (Topia source) / Jake (Equus identifier, D1) | blank / placeholder Assignment ID; **never** replaced by the Employee Number |
| MISSING_EMPLOYEE_NUMBER | ROW_HELD | Q1 | Alka (Topia source) | no Talent Link ID |
| DUPLICATE_ROW | ROW_HELD | Q5 | Alka (Topia extract) | every compared cell equals an earlier row; the earlier row is accepted |
| CONFLICTING_IDENTIFIER | ROW_HELD | Q1/Q5 | Alka / John (Topia) / Jake (D1) | same identifier on rows with different data; **all** such rows held |
| ALREADY_ACCEPTED | ROW_HELD | Q5 | Migration team (controlled reprocessing) | identifier accepted by an earlier run passed as `previously_accepted` |
| HEADER_MISMATCH | FILE | Q5 | Jake (Equus template) | output header rows 1-3 differ from the approved template |
| TEMPLATE_HASH_DIFFERS | WARNING | Q5 | Jake (Equus template) | same headers, different template bytes |
| HEADER_DECISION_PENDING | WARNING | Q5 | Jake (Equus template) | D2: A3 `A` (template) vs `Employee Command` (Jake's example); header unchanged |
| ROW_STRUCTURE_INVALID | FILE | Q17/Q18/Q19 | Migration team (engine defect) | generated rows break Jake's INSERT/UPDATE pattern |
| VALIDATION_WARNING | WARNING | Q5 | Alka / John (mapping review) | every existing value-level warning (Warnings sheet), value kept |

## 2. Procedure

1. **Run** the migration on the extract (S3 upload or `local/run_local.py`). Keep the report
   workbook of this run (run 1) — it is the record of what was accepted.
2. **If the run is FILE-level failed** (`*_FILE_REJECTED_*` report, `file-status = FAILED` in S3
   metadata): nothing may be imported. Resolve the FILE finding (e.g. obtain the approved template
   or fix the template header, owner Jake) and re-run the **whole** extract. No import happened, so
   there is nothing to de-duplicate.
3. **Import** run 1's accepted rows into Equus only through the controlled UAT procedure
   (master command §9: authorised environment, hash, batch id, accepted/rejected rows). Record
   which identifiers Equus actually accepted. *Not executed here.*
4. **Route** every Exceptions line to its Owner / Q. Corrections are made **at source (Topia)**
   by the owner — never by editing the import workbook and never by inventing a value (e.g. an
   Assignment ID is never filled with the Employee Number).
5. **Build the correction batch**, either
   - **(a) full re-extract:** the corrected full Topia extract, run with
     `transform(..., previously_accepted=transform.accepted_identifiers(run1_report_bytes))`.
     Every identifier accepted in run 1 is held as `ALREADY_ACCEPTED`; only the previously held
     (now corrected) identifiers reach the import workbook. Test:
     `tests/test_identity_controls.py::test_corrected_rerun_imports_only_previously_held_identifiers`.
   - **(b) held-rows batch:** copy the Held Rows sheet's data (columns C.. = source columns A..,
     in source order) under the source header rows 1-4, apply the corrections at source, and run
     that workbook. Test: `tests/test_identity_controls.py::test_correction_batch_built_from_held_rows_sheet`.
6. **Verify** the correction run's report before importing: no identifier from run 1's Accepted
   Identifiers sheet appears in the new import workbook (`Accepted Identifiers` of run 2 ∩ run 1 =
   ∅), and every previously held row is now either accepted or still held with a reason.
7. **Repeat** 4-6 until no row is held, or the remaining held rows are dispositioned in writing by
   their owner.

### Notes and limits
- `previously_accepted` is the tool's own record of accepted rows, **not** proof of what Equus
  imported. Before step 5(a) it must be reconciled with Equus's import results (step 3); a row
  accepted by the tool but rejected by Equus must be removed from the `previously_accepted` set so
  it is re-imported. *Equus behaviour on re-submitted identifiers is unobserved (Q5, Q26).*
- Employee command in a correction batch: the first accepted assignment of an employee *in that
  file* gets Employee command `INSERT` (unchanged behaviour, Q17). If the employee was already
  created by run 1, whether Equus needs `IGNORE` instead is **Q26 (INSERT/UPDATE/IGNORE
  semantics), OPEN, owner Jake** — not decided here; record it per correction batch.
- Held Rows contains personal data from the source (it is in the report workbook, like the
  Warnings sheet). Reports belong under `local_s3/dest_s3/` / the destination bucket and are never
  committed.
- The S3 handler does not read `previously_accepted` (its event contract is unchanged); controlled
  re-runs with (a) are operator-driven via `transform()` or a separately approved handler change.
