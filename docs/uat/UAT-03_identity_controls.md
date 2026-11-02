# UAT-03 — Identity, duplicate, header and row-level error controls (BR-01, BR-05, BR-17, BR-18)

Task: OCNORA T-1388. Sources: `clarifications/UAT_MASTER_COMMAND.md` §3 (BR-01, BR-05, BR-17,
BR-18, BR-19), §13; `clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md` D1, D2, D4 row 7, Q1/Q5/
Q17-Q19 rows. Jake's answers are evidence, **not approval** (every tracker Status, column I, is
Open). Equus UAT execution: **NOT EXECUTED** (no Equus access). Readiness can never be better
than READY FOR UAT EXECUTION; with any held row it is **NOT_READY**. Nothing is SIGNED OFF.

## 1. Change log

| File | Function / element | Before | After | Reason | Tests |
|---|---|---|---|---|---|
| `lambda/src/transform.py:935` | `build_records` | blank Assignment ID ⇒ identifier = Employee Number (old `transform.py:547-550`) | fallback **removed**; `screen_rows` runs first, held rows skipped (`:963`), identifier from `identifier_text` | BR-01 / D4 row 7 (forbidden fallback) | `test_missing_assignment_id_is_held_never_falls_back`, `test_no_identifier_fallback_left_in_engine` |
| `transform.py:583` | `identifier_text` (new) | identifier converted like a CF_TEXT value | text; leading zeros kept; numeric `90011.0` ⇒ `"90011"`; blank / placeholder ⇒ none | BR-01 (leading zeros) | `test_identifier_text_preserves_leading_zeros_and_drops_dot_zero`, `test_one_employee_three_assignments_three_inserts_text_identifiers` |
| `transform.py:618` | `screen_rows` (new) | none | MISSING_ASSIGNMENT_ID, MISSING_EMPLOYEE_NUMBER, DUPLICATE_ROW, CONFLICTING_IDENTIFIER, ALREADY_ACCEPTED | BR-05, BR-17 | AC-1, AC-2 tests |
| `transform.py:697` | `validate_row_structure` (new) | none | first row INSERT/INSERT/NEW, further rows IGNORE/UPDATE, blank Employee & Assignment Number, same identifier; violation = FILE | Q17-Q19, BR-18, task item 6 | `test_generated_rows_follow_jakes_example_structure`, `test_row_structure_violations_are_file_level` |
| `transform.py:753` | `check_template_headers` (new) | template rows copied, never checked | output rows 1-3 compared with the approved bundled template; TEMPLATE_VERSION + SHA-256 recorded; D2 reported as pending decision | BR-05, Q5, D2 | AC-4 tests |
| `transform.py:1122` | `write_output` | — | identifier cells get number format `@` (text) | leading zeros | AC-3 test |
| `transform.py:1151` | `reconcile` | source rows = all rows | accepted rows reconcile against INSERT rows; `data_rows == accepted + held` is required | every row accounted for | `test_every_case_accounts_for_every_source_row` |
| `transform.py:1217` | `compute_readiness` | `identifier` reasons for fallbacks | reasons `row_held`, `file_rejected`, `pending_header_decision`, `template_version` with Q / owner | AC-5 | `test_readiness_not_ready_whenever_a_row_is_held` |
| `transform.py:1415-1481` | report sheets | — | **Exceptions**, **Held Rows**, **Accepted Identifiers**, **Template Check**; Readiness gains File status, Rows accepted / held, TEMPLATE_VERSION, template SHA-256 | task items 3-5 | AC-1, AC-4 tests |
| `transform.py:1571` | `transform` | — | `previously_accepted=` parameter; value-level warnings become `VALIDATION_WARNING` exceptions; FILE finding ⇒ workbook `b""` | task item 4, 5 | `test_corrected_rerun_imports_only_previously_held_identifiers` |
| `lambda/src/handler.py:185-194` | `lambda_handler` | — | FILE finding ⇒ `reports/*_FILE_REJECTED_*` + `RuntimeError` (no import workbook); `rows-held`, `file-status` metadata; `rows_held` in summary | Q5 (header issues fail the file) | `test_handler_withholds_import_on_file_level_failure` |
| `lambda/src/business_rules.json` `identity_controls` | new section; file version 1.1.0 (UAT-04) ⇒ **1.2.0** | — | identifier label, identifier headers, duplicate-compare column statuses, approved template (version, file, SHA-256, header rows), pending header decision D2, ten exception codes with level / Q / owner | architecture requirement | `test_shipped_identity_controls` |
| `lambda/src/mapping_config.py:711,806` | `_load_identity_controls` | — | strict validation: unknown / missing code, disallowed level (row-identity codes can never be WARNING), template SHA-256 must match the bundled file, decision cell inside header rows, template file inside `lambda/src` | reject invalid configuration | `test_invalid_identity_configuration_rejected` (8 cases) |
| `mapping_config.py:1033-1036` | views | `ASSIGNMENT_IDENTIFIER_LABEL` constant | derived from JSON; `TEMPLATE_VERSION`, `APPROVED_TEMPLATE_SHA256` | — | — |
| `transform.py:1062`, `mapping_config.py:226` | `_hold_keyed_dependents_of_held_row` (new, re-integration on top of UAT-04) | — | keyed dependents (UAT-04 `dependent_source`) whose Assignment ID belongs to a held assignment row are held with it as `HELD_DEPENDENTS_ASSIGNMENT_HELD` (Q18, owner Alka / Topia), or `HELD_DEPENDENTS_AMBIGUOUS` when the Assignment ID is on more than one source row; never attached to another assignment | BR-01, Q18 | `test_keyed_dependents_of_a_held_assignment_are_held_with_it`, `tests/test_family_rules.py::test_duplicate_assignment_id_holds_keyed_dependents` |
| `OPEN_QUESTIONS.md` §1 | "Code today" | described the fallback | describes the hold | doc accuracy | — |

Rollback: revert the commit. The identifier fallback must **not** be restored (BR-01, task
rollback note). No existing test was removed or weakened; no existing assertion needed changing.

## 2. Assumptions (recorded instead of asking)

1. **Identifier column stays Topia column X "Assignment ID"** (assessment D1 "until resolved: keep
   col X"). D1 (Jake's text vs his example, `jake-1.xlsx` J2 vs ` Mny dependents Ex`!C4/E4/Y4) is
   still OPEN, owner Jake / Alka / John; it blocks G3.
2. **Missing Employee Number holds the row** (`MISSING_EMPLOYEE_NUMBER`, Q1). Previously only a
   warning; an assignment without a Talent Link ID cannot be linked to an employee (Q1, Q17) and
   holding it changes no value. Owner Alka.
3. **Duplicate comparison columns** = every source column whose plan status is mapped, address,
   contact, dependent, rule or held, plus the identifier and Employee Number columns
   (`identity_controls.duplicate_compare_statuses`). Red (excluded), ignored, unmapped and
   duplicate-target columns are not compared ("all mapped cells equal"). Values are compared after
   trimming text and treating placeholders as blank; numbers compare by value (`90001` = `90001.0`).
4. **Exact duplicates sharing the identifier**: the first row is accepted, later ones DUPLICATE_ROW.
   **Same identifier, different data**: every row of that identifier is CONFLICTING_IDENTIFIER
   (including rows that are also exact duplicates of each other), even across employees.
5. **Approved template = the bundled `lambda/src/templates/equus_eafh_template.xlsx`**, SHA-256
   `c90e9fc057e9fbcc5b1b38e8b4821dd266b92bd7557915f48f7925fc3f445be7`, labelled
   `TEMPLATE_VERSION = EAFH-bundled-1` because Equus has not supplied an official version number
   (D2, Q5, owner Jake). A changed bundled template is rejected at load time until the SHA-256 and
   version in the configuration are updated with an approved template.
6. **D2 is reported, not resolved**: A3 stays `A`; a supplied template with `Employee Command` in
   A3 is a HEADER_MISMATCH (FILE) because it differs from the approved template. If Jake confirms
   `Employee Command`, the template, its SHA-256/version and `mapping_config.COL_EMPLOYEE_ACTION`
   change together.
7. **Header comparison is exact** per cell (whitespace-only = empty) over rows 1-3 and every column
   of either workbook.
8. **FILE level ⇒ `transform()` returns `b""`** for the import workbook so no caller can use a
   withheld file; the report keeps all diagnostics.
9. **Row-structure violations are FILE level**: they can only come from an engine defect, and a
   wrong linkage (Q18) would corrupt Equus data.
10. **Value-level warnings** (pre-existing Warnings sheet) are listed as `VALIDATION_WARNING` /
    WARNING in Exceptions so every finding carries a code; they do not change readiness (as before).
11. **`previously_accepted`** is optional and operator-driven (see
    `docs/uat/REPROCESSING_PROCEDURE.md`); the S3 handler contract is unchanged.

## 3. Test evidence

`tests/test_identity_controls.py` (38 tests, synthetic fixtures `tests/fixtures/topia_sample.xlsx`
+ `tests/fixtures/identity_cases.json`), run with `./.venv/bin/python -m pytest -q
tests/test_identity_controls.py` ⇒ 38 passed; `tests/test_migration.py` +
`tests/test_business_rules.py` ⇒ 69 passed (unchanged). Full gate: run by the OCNORA runner.

| Criterion | Tests |
|---|---|
| AC-1 blank Assignment ID held, not in import, in Exceptions + Held Rows, no fallback | `::test_missing_assignment_id_is_held_never_falls_back` (blank and placeholder), `::test_no_identifier_fallback_left_in_engine` |
| AC-2 duplicates / conflicting identifiers | `::test_exact_duplicate_row_one_accepted_one_held`, `::test_same_identifier_different_data_holds_both_rows`, `::test_identifier_reused_by_other_employee_is_conflict` |
| AC-3 three assignments, INSERT/IGNORE/IGNORE, distinct ids, `'00123'` as text | `::test_one_employee_three_assignments_three_inserts_text_identifiers`, `::test_employee_command_insert_goes_to_first_accepted_assignment`, `::test_generated_rows_follow_jakes_example_structure` |
| AC-4 header mismatch FILE-failed with columns; D2 pending decision | `::test_header_mismatch_fails_file_and_lists_columns`, `::test_employee_command_header_in_template_is_mismatch_not_silently_accepted`, `::test_pending_header_decision_reported_header_unchanged`, `::test_handler_withholds_import_on_file_level_failure` |
| AC-5 NOT_READY whenever a row is held | `::test_readiness_not_ready_whenever_a_row_is_held`, `::test_every_case_accounts_for_every_source_row` |

**Real-extract run: NOT EXECUTED in this job** (the in-memory run against
`local_s3/source_s3/topia/incoming/…` was blocked by the runner's command policy). Counts of held
rows per code on the real extract are therefore not recorded here; they must be produced with
`local/run_local.py` and read from the report's Exceptions / Readiness sheets (never committed).

## 4. Gates affected
G3 (identity and linkage conflicts) stays **Blocked**: D1 and D2 are open. G4 adds the 38 tests
above. G6-G8 (Equus import, duplicate / re-import behaviour in Equus) **NOT EXECUTED**. Overall
status remains **NOT READY FOR UAT**.
