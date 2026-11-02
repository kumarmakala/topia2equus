# UAT-06 — Evidence-based UAT test results (master command section 7)

Task: OCNORA T-1391 (UAT-06). Sources: `clarifications/UAT_MASTER_COMMAND.md` section 7
(lines 137-153), `clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md` (sections 1-3),
Jake's answers in `clarifications/jake-1.xlsx` sheet `datamapping clarrification shef` column J
(row n+1 = Qn, e.g. Q5 = J6) and sheet ` Mny dependents Ex` (rows 4-7).

**Scope of this evidence:** local, automated tests of the transformation (`lambda/src/transform.py`)
and of the Lambda handler (`lambda/src/handler.py`, in-memory S3 fake) with **synthetic fixtures
only**. These tests simulate; they do not observe Equus. Every scenario that needs Equus is listed in
section 4 as **NOT EXECUTED** with the evidence Equus UAT must supply. Equus UAT execution, import
results and post-import audit are **INCOMPLETE**. A Jake answer is evidence, not approval (every
tracker Status cell in column I reads Open). Readiness is unchanged by this task:
**NOT READY FOR UAT** (open lookups Q2-Q4, Q12, Q14, Q15; identifier conflict D1; header decision D2;
unapproved rules D4 — see `docs/uat/RTM_Q1-Q43.md`).

## 1. Commands and counts

Interpreter: `./.venv/bin/python` (Python 3.12.15, openpyxl 3.1.5, pytest), run from the repository
root on 2026-11-02.

| Command | Collected | Passed | Failed | Skipped |
|---|---|---|---|---|
| `./.venv/bin/python -m pytest --collect-only -q tests` (full suite) | 257 | – | – | – |
| `./.venv/bin/python -m pytest -q tests/test_uat_scenarios_employee.py tests/test_uat_scenarios_family.py tests/test_uat_scenarios_mapping.py tests/test_uat_scenarios_aws.py tests/test_migration.py` | 102 | 102 | 0 | 0 |
| `./.venv/bin/python -m pytest -q -x tests/test_identity_controls.py tests/test_family_rules.py tests/test_reconciliation.py tests/test_handler_audit.py tests/test_business_rules.py` (existing suites, after the code changes of section 2) | 155 | 155 | 0 | 0 |

Full suite = 257 tests = 179 pre-existing (none removed, skipped or weakened; no existing assertion
was changed) + 78 new UAT-06 tests (`tests/test_uat_scenarios_*.py`: employee 14, family 18,
mapping 21, AWS / fixtures / traceability 25). The two runs above together execute all 257. The
gate command is `./.venv/bin/python -m pytest -q tests` (plus py_compile and JSON parsing), run
independently by the OCNORA runner; its result is authoritative.

Fixtures (new, synthetic, built by `tests/uat_scenario_fixtures.py`):
`tests/fixtures/uat_synthetic_base.xlsx` (2 employees 70001/70002, Assignment IDs 80001/80002),
`tests/fixtures/uat_family.xlsx` (rows 5-10, `FAMILY_ROWS`), `tests/fixtures/uat_family_linkage.xlsx`
(keyed dependent sheet, `LINKAGE_ROWS`). Header rows 1-4 (labels, row-2 technical keys, colour fills)
are copied from `tests/fixtures/topia_sample.xlsx`; every data value is invented. Each scenario
rewrites data rows by Topia row-4 label in memory.

| Check | Test |
|---|---|
| AC-2: e-mails `@example.invalid`, names `Synth*` / `Testcase`, Employee IDs 7000x, Assignment IDs 8000x, cost centres `SYN-*` / `Synthetic*`; no name / e-mail / id / cost-centre value shared with the data rows of `topia_sample.xlsx` | `tests/test_uat_scenarios_aws.py::test_uat_fixtures_are_synthetic` |
| AC-2: no such value shared with the real extract (compared wherever the git-ignored extract exists; counts only, never values) | `tests/test_uat_scenarios_aws.py::test_uat_fixtures_share_nothing_with_real_extract_when_present` |
| Committed fixtures equal their builder output | `tests/test_uat_scenarios_aws.py::test_uat_fixture_matches_its_builder` |
| AC-1: every PASS row below names an existing test; every NOT EXECUTED row names evidence | `tests/test_uat_scenarios_aws.py::test_results_doc_every_pass_names_an_existing_test_and_not_executed_names_evidence` |

## 2. Production code changes (bug fixes required by failing scenarios)

Each change is the smallest fix that made a section-7 scenario pass; behaviour is reported, never a
value invented. Base = commit `146d6ab`.

| Fix | Failing scenario on base | Change | Tests |
|---|---|---|---|
| B-01a | EA-10 overlapping / missing dates were not reported at all (no overlap check existed in `transform.py`) | `transform.check_assignment_dates` (`lambda/src/transform.py:991`), called from `build_records`: warns on missing start date, end before start, and overlapping date ranges of one Employee Number (missing end = open-ended). Warnings only (`VALIDATION_WARNING`); no date derived, moved or blanked; no row held. Headers in `mapping_config.ASSIGNMENT_START_HEADER` / `ASSIGNMENT_END_HEADER` (`lambda/src/mapping_config.py:166`) | EA-10a, EA-10b |
| B-01b | MD-04 '05/06/2024' was converted silently to 2024-05-06 | `_convert` (`lambda/src/transform.py:508`) warns when a text date could be read day-first or month-first; the value is still read only with the configured month/day order (`mapping_config.AMBIGUOUS_DATE_RE`, `DATE_ORDER`, `lambda/src/mapping_config.py:171`); Q24 stays OPEN | MD-04a, MD-04b |
| B-02 | MD-02 a state was never validated (BR-12, Q12 J13: country **and state** must match Equus values); the `country_state` list (country code -> states) was only used for country codes | `Report.state_in_list` / `state_values` (`lambda/src/transform.py:157,235`), state check in `summarise_values` (`:270`) and readiness (`:1773`): with an approved list a state not listed for its country is a warning + `reference_value_unknown` (`country_state (state)`); while the list is not approved every state is `reference_list_unavailable`. Values are sent unchanged. Header pairs in `mapping_config.STATE_COUNTRY_HEADERS` (`lambda/src/mapping_config.py:157`), checked against the template (`transform.py:345`) | MD-02a, MD-02b |
| B-03 | FM-09 a repeated dependent record was written twice (Q5 J6: Equus can create duplicates) | `build_records` (`lambda/src/transform.py:1119-1136`) holds a dependent record identical to an earlier one of the same assignment as `HELD_DEPENDENTS_DUPLICATE` (`mapping_config.DEPENDENT_HOLDS`, `lambda/src/mapping_config.py:253`; Q5/Q19, owner Alka / Topia); reconciliation counts its cells as `held_row` (`transform.py:1554-1562`); readiness `held_dependents` | FM-09, `test_family_fixture_accounts_for_every_row` |
| B-04 | AW-03 / AW-04: the handler wrote import workbook -> report -> manifest and treated an existing import workbook as "done". A failed report or manifest put left an import workbook without report / manifest, and the Lambda retry skipped it for ever. Verified: on the base handler `test_AW03...[_MAPPING_REPORT_]`, `test_AW03...[_RUN_]` and `test_AW04_output_failure_on_first_put_leaves_nothing` fail | `lambda/src/handler.py:371-374` done = import workbook **and** manifest exist; `:464-477` write order report -> import workbook -> manifest, every put conditional (`IfNoneMatch='*'`), a retry completes the missing objects and never writes one twice | AW-03, AW-03b, AW-04 |

Rollback: revert these hunks (`git diff 146d6ab -- lambda/src`) and delete the new tests / fixtures.

## 3. Scenario traceability (section 7) — executed locally

Result PASS = the named test passed in the run of section 1. Q cells refer to Jake's column J
(`datamapping clarrification shef`!J(n+1)); BR numbers to the master command section 3.

### 3.1 Employee / assignment

| ID | Scenario | Q / BR | Test | Expected (asserted) | Result |
|---|---|---|---|---|---|
| EA-01 | Single employee, one assignment | BR-01, Q1 (J2) | `tests/test_uat_scenarios_employee.py::test_EA01_single_employee_one_assignment` | 1 row INSERT/INSERT/NEW; Assignment Identifier = CF_TEXT029 = Topia Assignment ID (col X, D1 open) | PASS |
| EA-02 | One employee, multiple historical assignments | BR-17, Q17 (J18) | `tests/test_uat_scenarios_employee.py::test_EA02_one_employee_multiple_historical_assignments_never_merged` | 3 INSERT/NEW rows, Employee command INSERT, IGNORE, IGNORE; never merged; leading zeros kept | PASS |
| EA-03 | Multiple employees | BR-17 | `tests/test_uat_scenarios_employee.py::test_EA03_multiple_employees_each_insert` | one Employee INSERT per employee; reconciliation balances | PASS |
| EA-04 | Duplicate source rows | BR-05, Q5 (J6) | `tests/test_uat_scenarios_employee.py::test_EA04_duplicate_source_row_held_once_written` | first row written, repeat held DUPLICATE_ROW (Held Rows sheet), NOT_READY | PASS |
| EA-05 | Missing Employee Number | BR-01, Q1 | `tests/test_uat_scenarios_employee.py::test_EA05_missing_employee_number_held` | row held MISSING_EMPLOYEE_NUMBER, absent from import | PASS |
| EA-06 | Missing Assignment Identifier | BR-01, Q1, D1 | `tests/test_uat_scenarios_employee.py::test_EA06_missing_assignment_identifier_held_no_fallback` | row held MISSING_ASSIGNMENT_ID; Employee Number never used as identifier or CF_TEXT029 | PASS |
| EA-07a | Conflicting identifiers (same id, different data) | BR-05, Q1/Q5 | `tests/test_uat_scenarios_employee.py::test_EA07a_same_identifier_different_data_holds_all_rows` | all rows of the identifier held; message names columns, not values | PASS |
| EA-07b | Conflicting identifiers (id reused by another employee) | BR-05, Q1/Q5 | `tests/test_uat_scenarios_employee.py::test_EA07b_identifier_reused_by_other_employee_is_conflict` | both rows held CONFLICTING_IDENTIFIER | PASS |
| EA-08 | Previously imported assignment (tool record) | BR-05, Q5; Q26 open | `tests/test_uat_scenarios_employee.py::test_EA08_previously_imported_assignment_not_imported_again` | with `previously_accepted` every earlier identifier is held ALREADY_ACCEPTED; nothing re-imported. Equus side: NX-05 | PASS |
| EA-09 | Repeated run of the same file | BR-05, Q5 | `tests/test_uat_scenarios_employee.py::test_EA09_repeated_run_same_file_identical_output_modulo_timestamps` | import and report workbooks byte-identical except `docProps/core.xml` (timestamps); handler side: AW-02 | PASS |
| EA-10a | Overlapping dates | BR-17, Q17, Q23 open | `tests/test_uat_scenarios_employee.py::test_EA10a_overlapping_assignment_dates_warned_never_altered` | warning naming both source rows; dates unchanged; no hold (fix B-01a) | PASS |
| EA-10b | Missing / inverted dates | Q22/Q23 open | `tests/test_uat_scenarios_employee.py::test_EA10b_missing_and_inverted_dates_warned_nothing_derived` | missing start left blank + warning; end before start kept + warning (fix B-01a) | PASS |
| EA-11a | Existing employee with a new assignment (same file) | Q17 (J18), ` Mny dependents Ex` | `tests/test_uat_scenarios_employee.py::test_EA11a_existing_employee_new_assignment_same_file_ignore_insert` | Employee command IGNORE, Assignment Command INSERT, Assignment Number NEW | PASS |
| EA-11b | Existing employee from an earlier run with a new assignment | Q26 OPEN (owner Jake) | `tests/test_uat_scenarios_employee.py::test_EA11b_existing_employee_from_earlier_run_new_assignment_pending_q26` | old identifier held; new one INSERT/NEW; Employee command INSERT characterised (finding F-01) | PASS |

### 3.2 Spouse / dependents (fixture `tests/fixtures/uat_family.xlsx`)

| ID | Scenario | Q / BR | Test | Expected (asserted) | Result |
|---|---|---|---|---|---|
| FM-00 | Every family row accounted for; child-record counts reconcile | Q27, BR-19 | `tests/test_uat_scenarios_family.py::test_family_fixture_accounts_for_every_row` `tests/test_uat_scenarios_family.py::test_FM_rows_match_expected_counts` | dependents expected = written = 10; rows per assignment = max(1, dependents) | PASS |
| FM-01 | No spouse, no dependents | BR-19 | `tests/test_uat_scenarios_family.py::test_FM01_no_spouse_no_dependents_single_insert_row` | single INSERT row, no family cells | PASS |
| FM-02 | One dependent | BR-11, Q11 (J12) | `tests/test_uat_scenarios_family.py::test_FM02_one_dependent_on_insert_row` | dependent on the INSERT row | PASS |
| FM-03 | Two dependents | BR-19, Q19 (J20) | `tests/test_uat_scenarios_family.py::test_FM03_two_dependents_insert_then_update_same_identifier` | INSERT then UPDATE, IGNORE, blank Employee / Assignment Number, same identifier (Jake example rows 4-7) | PASS |
| FM-04 | Three or more dependents | BR-19, Q19 (J20) | `tests/test_uat_scenarios_family.py::test_FM04_four_dependents_no_capacity_limit` | 4 dependents = INSERT + 3 UPDATE rows; no capacity limit assumed | PASS |
| FM-05 | Spouse with name | BR-10 | `tests/test_uat_scenarios_family.py::test_FM05_spouse_with_name_written_on_insert_row` | spouse fields on the INSERT row; no data-quality note | PASS |
| FM-06 | Spouse without name | BR-10, Q10 (J11) | `tests/test_uat_scenarios_family.py::test_FM06_spouse_without_name_imported_with_data_quality_note_only` | written without names, DATA_QUALITY note, relationship reported against pending list (Q14) | PASS |
| FM-07 | Dependent with one populated field | BR-20, Q20 (J21) | `tests/test_uat_scenarios_family.py::test_FM07_dependent_with_one_populated_field_is_created` | record with birth date only created; no name invented | PASS |
| FM-08 | Multiple assignments, each with different dependents | BR-18, Q18 (J19) | `tests/test_uat_scenarios_family.py::test_FM08_multiple_assignments_each_with_own_dependents` | each block carries its own identifier and only its dependents | PASS |
| FM-09 | Repeated dependent records | Q5 (J6), Q19 | `tests/test_uat_scenarios_family.py::test_FM09_repeated_dependent_record_held_as_duplicate` | repeat held HELD_DEPENDENTS_DUPLICATE, first written once; reconciliation explains 4 held cells (fix B-03) | PASS |
| FM-10a | Missing or incorrect linkage | BR-18, Q18 (J19), Q11 D3 | `tests/test_uat_scenarios_family.py::test_FM10a_missing_or_incorrect_linkage_is_held_never_guessed` | no Assignment ID, unknown ID, Employee Number as ID and employee mismatch held; only the correctly keyed child attached (test-only approved configuration) | PASS |
| FM-10b | Keyed dependent source not approved | Q11 (J12), D3 | `tests/test_uat_scenarios_family.py::test_FM10b_keyed_sheet_without_approval_attaches_nothing` | shipped configuration: HELD_DEPENDENTS_NOT_CONFIGURED, nothing attached | PASS |

### 3.3 Mapping / data quality (fixture `tests/fixtures/uat_synthetic_base.xlsx`)

| ID | Scenario | Q / BR | Test | Expected (asserted) | Result |
|---|---|---|---|---|---|
| MD-01a | Country values: provisional conversion vs unknown value | BR-12, Q12 (J13), Q3 | `tests/test_uat_scenarios_mapping.py::test_MD01a_country_names_provisional_unknown_passed_through` | known name -> provisional ISO-3 (rule not approved); unknown passed through + warning; NOT_READY | PASS |
| MD-01b | Country against an approved list (test-only list) | BR-12, Q12 | `tests/test_uat_scenarios_mapping.py::test_MD01b_country_against_approved_list_unknown_reported` | unknown value `reference_value_unknown`, sent unchanged | PASS |
| MD-02a | State not in the approved list | BR-12, Q12 (J13) | `tests/test_uat_scenarios_mapping.py::test_MD02a_state_not_in_approved_list_reported_never_converted` | warning + `reference_value_unknown` `country_state (state)`; value unchanged (fix B-02) | PASS |
| MD-02b | State while the list is not approved | BR-12, Q12 | `tests/test_uat_scenarios_mapping.py::test_MD02b_state_unvalidated_while_list_not_approved` | `reference_list_unavailable` `country_state (state)` (fix B-02) | PASS |
| MD-03 | Currency blank, no derivation | BR-04, Q4 (J5) | `tests/test_uat_scenarios_mapping.py::test_MD03_currency_blank_never_derived_from_country` | no FROM/TO/DEFAULT/salary currency written although countries and salary are present | PASS |
| MD-04a | Ambiguous date '05/06/2024' | Q24 open | `tests/test_uat_scenarios_mapping.py::test_MD04a_ambiguous_text_date_read_with_configured_order_and_reported` | read only as month/day (2024-05-06) and reported as ambiguous (fix B-01b) | PASS |
| MD-04b | Unambiguous text date | Q24 open | `tests/test_uat_scenarios_mapping.py::test_MD04b_unambiguous_text_date_not_reported` | 06/13/2024 read, no ambiguity warning | PASS |
| MD-04c | Invalid dates (incl. day-first 13/06/2024) | Q24 open | `tests/test_uat_scenarios_mapping.py::test_MD04c_invalid_date_left_blank_warned_and_reconciled` | left blank + warning; reconciliation `failed_conversion` | PASS |
| MD-05 | Salary without effective date | Q21, Q22 open | `tests/test_uat_scenarios_mapping.py::test_MD05_salary_without_effective_date_no_default_exception` | no default effective date; `gated_rule` default_effective_dates_from_start (Q22/Q23) | PASS |
| MD-06 | Optional vs mandatory fields | BR-01, BR-20 | `tests/test_uat_scenarios_mapping.py::test_MD06_blank_optional_fields_accepted_mandatory_identity_held` | blank optional fields accepted; missing identity fields held. Equus mandatory list: NX-07 | PASS |
| MD-07 | Red / yellow / cyan / brown flags | Q8 (J9), Q30 | `tests/test_uat_scenarios_mapping.py::test_MD07_red_yellow_cyan_brown_flags` | red excluded; yellow PENDING reason; cyan without target MISSING_TARGET; brown no meaning; STRICT_PENDING refuses | PASS |
| MD-08 | Approved custom fields | Q1, Q29 open | `tests/test_uat_scenarios_mapping.py::test_MD08_approved_custom_field_targets_written_as_text` | Global ID, Pre-Hire ID, Assignment ID written to their CF columns as text | PASS |
| MD-09 | E-mail mapping | BR-13, Q13 (J14); Q41 | `tests/test_uat_scenarios_mapping.py::test_MD09_email_mapping_no_contact_records` | USER_ID_EMAIL + CF_TEXT008; no contact rows; resource e-mails not written (Q41) | PASS |
| MD-10 | Placeholders incl. 'Not ProvidedHome' | Q43 open | `tests/test_uat_scenarios_mapping.py::test_MD10_placeholders_blank_and_glued_placeholder_reported` | blank tokens blanked; glued placeholder blanked + warning; reconciliation `blank_token` | PASS |
| MD-11 | Confidential record (AG) | BR-06, Q6 (J7), Q28 | `tests/test_uat_scenarios_mapping.py::test_MD11_confidential_attribute_held_never_written` | never written; readiness `missing_target` (Q6/Q7/Q29); reconciliation `pending_target` with Q6 | PASS |
| MD-12 | Header mismatch (file level) | BR-05, Q5 (J6), D2 | `tests/test_uat_scenarios_mapping.py::test_MD12_header_mismatch_fails_whole_file` | whole import withheld (b""); HEADER_MISMATCH FILE level; report kept | PASS |
| MD-13 | Values outside Equus lists | BR-14, Q14 (J15) | `tests/test_uat_scenarios_mapping.py::test_MD13_value_outside_equus_list_reported_unchanged` | value unchanged; `dropdown_value_not_allowed` + warning | PASS |
| MD-14 | Unknown reference values (resources) | BR-02, Q2 (J3) | `tests/test_uat_scenarios_mapping.py::test_MD14_unknown_reference_values_resource_without_username` | name passed through; `reference_lookup_missing` resource_usernames | PASS |
| MD-15 | Numeric value outside review range (158 h) | BR-16, Q16 (J17) | `tests/test_uat_scenarios_mapping.py::test_MD15_weekly_hours_outside_range_kept_and_warned` | kept, warned, never capped | PASS |

### 3.4 AWS / import (handler with in-memory S3)

| ID | Scenario | Q / BR | Test | Expected (asserted) | Result |
|---|---|---|---|---|---|
| AW-01 | Valid S3 trigger | handler contract | `tests/test_uat_scenarios_aws.py::test_AW01_valid_s3_trigger_writes_draft_report_manifest` | URL-decoded key; DRAFT import + mapping report + manifest; import-ready false; Equus execution INCOMPLETE | PASS |
| AW-01b | Events outside the contract | handler contract | `tests/test_uat_scenarios_aws.py::test_AW01b_events_outside_contract_are_skipped` | delete event, other prefix, lock file, other suffix skipped; nothing written | PASS |
| AW-02 | Duplicate event | BR-05, Q5 | `tests/test_uat_scenarios_aws.py::test_AW02_duplicate_event_and_duplicate_record_write_nothing_new` | same record twice / redelivered: skipped, nothing new | PASS |
| AW-03 | Lambda retry after a simulated failure | BR-05, Q5 | `tests/test_uat_scenarios_aws.py::test_AW03_lambda_retry_after_failed_put_completes_without_duplicates` | failure at report, import or manifest put; retry completes exactly one set; further redelivery a no-op (fix B-04) | PASS |
| AW-03b | Retry after a transient read failure | BR-05 | `tests/test_uat_scenarios_aws.py::test_AW03b_transient_read_failure_then_retry` | nothing written on failure; retry completes | PASS |
| AW-04 | Output failure | BR-05 | `tests/test_uat_scenarios_aws.py::test_AW04_output_failure_on_first_put_leaves_nothing` `tests/test_uat_scenarios_aws.py::test_AW03_lambda_retry_after_failed_put_completes_without_duplicates` | put error propagates; first put failing leaves nothing; no import workbook without its report; no manifest claims success (fix B-04) | PASS |
| AW-05 | Invalid source (wrong sheet, corrupt) | BR-05 | `tests/test_uat_scenarios_aws.py::test_AW05_invalid_source_is_source_failure_no_import` | SOURCE failure, FAILED manifest, no import | PASS |
| AW-06 | Invalid template (no Sheet1, header mismatch, unreadable) | Q5 (J6), D2 | `tests/test_uat_scenarios_aws.py::test_AW06_invalid_template_is_template_failure_no_import` | TEMPLATE failure, no import | PASS |
| AW-07 | Partial validation failure | BR-05, Q5 | `tests/test_uat_scenarios_aws.py::test_AW07_partial_validation_failure_held_rows_draft_still_produced` | held row absent; DRAFT still produced with accepted rows; rows-held metadata and manifest counts | PASS |
| AW-08 | Re-run without duplicates (correction batch) | BR-05, Q5 | `tests/test_uat_scenarios_aws.py::test_AW08_correction_batch_rerun_imports_only_corrected_rows` | correction batch from Held Rows writes only the corrected identifier; run 1 ∩ run 2 = ∅ | PASS |
| AW-08b | Same bytes under a new key | BR-05, Q5 | `tests/test_uat_scenarios_aws.py::test_AW08b_same_bytes_new_key_detected_by_accepted_identifier_check` | processed again (finding F-02); overlap detected by the Accepted Identifiers check of the reprocessing procedure | PASS |

## 4. NOT EXECUTED — requires Equus UAT

No Equus access exists in this environment. These scenarios are not passed by any local test and
must be executed in the authorised Equus UAT environment (master command section 9). Evidence the
Equus UAT must supply for each:

| ID | Scenario | Q / BR | Why not executable here | Evidence Equus UAT must supply | Result |
|---|---|---|---|---|---|
| NX-01 | Row-level Equus rejection | BR-05, Q5 (J6) | Equus import engine not available | Import job id, file SHA-256, per-row accept/reject list with Equus error text for a batch seeded with known-bad rows (e.g. value outside an Equus list); proof that accepted rows were imported and rejected rows were not | NOT EXECUTED |
| NX-02 | File-level Equus rejection | BR-05, Q5 (J6), D2 | Equus import engine not available | Import attempt with a header deviation (e.g. A3 'Employee Command' vs 'A'); Equus job result showing the whole file rejected and zero records created; confirmation which A3 header Equus accepts | NOT EXECUTED |
| NX-03 | Correction and resubmission in Equus | BR-05, Q5, Q26 | Equus import engine not available | Rejected rows corrected at source and resubmitted per `docs/uat/REPROCESSING_PROCEDURE.md`; Equus job results for both batches; record counts before/after proving no duplicate employee, assignment or dependent | NOT EXECUTED |
| NX-04 | Post-import audit | BR-05, Q5, Q27 | No Equus data to query | Equus extract or audit report after import, reconciled per record against the run's Accepted Identifiers and mapping report (identifiers, dependents, spouses, compensation, dates, custom fields); duplicate audit by Employee Number + Assignment Identifier | NOT EXECUTED |
| NX-05 | Previously imported assignment re-submitted to Equus | Q5, Q26 OPEN | Equus behaviour on a re-submitted identifier / Employee INSERT for an existing employee is unobserved | Equus result for (a) an Assignment Identifier already in Equus re-sent with INSERT, (b) Employee command INSERT vs IGNORE for an employee created by an earlier batch (finding F-01) | NOT EXECUTED |
| NX-06 | Invalid Equus values accepted or rejected | BR-12, BR-14, Q2-Q4, Q12, Q14, Q15 | Equus lists not supplied; local tests only report values against configured lists | The Equus country/state, dropdown, currency, cost-centre and resource lists (versioned, approved), and Equus import results for values inside and outside them | NOT EXECUTED |
| NX-07 | Equus mandatory fields | BR-20, Q20 (J21), Q14 | No Equus mandatory-field definition provided; template row 1 carries no mandatory markers | Equus list of mandatory columns per record type and import results for rows missing each of them | NOT EXECUTED |

## 5. Findings and open points (reported, not resolved here)

| ID | Finding | Owner / Q | Evidence |
|---|---|---|---|
| F-01 | An employee created by an earlier batch gets Employee command INSERT on its first assignment in a later batch: the engine only knows earlier *identifiers* (`previously_accepted`), not earlier employees. Whether Equus needs IGNORE there is unknown. No rule invented. | Jake, Q26 (OPEN) | EA-11b; NX-05 |
| F-02 | Handler idempotency is per source key + ETag; the same bytes under another key are transformed again, and the handler has no `previously_accepted` input (procedure 5a is local only). Control: reprocessing procedure step 6 (Accepted Identifiers comparison) before any import. | Migration team, Q5 | AW-08b |
| F-03 | Equus mandatory-field and value lists are not available; local tests prove reporting, not Equus acceptance. | Jake, Q14 / Q12 / Q2-Q4 / Q15 | MD-01..MD-14; NX-06, NX-07 |
| F-04 | Ambiguous text dates are read month-first only (configured order) and reported; the accepted date format(s) per field are not confirmed. | Jake, Q24 (OPEN) | MD-04a..MD-04c |
| F-05 | A manifest-put failure after the import workbook was written leaves report + import workbook without manifest until the Lambda retry (MaximumRetryAttempts: 1, `lambda/template.yaml:57`) or a manual re-drive completes it; no object is deleted by the function. | Migration team | AW-03 |

## 6. Assumptions

- A1: "Repeated dependent record" = every populated field equal to an earlier dependent record of
  the same assignment; it is held for review (possible twins), not dropped (B-03).
- A2: "Overlapping dates" = date ranges of assignments of one Employee Number intersect; a missing
  end date is treated as open-ended for this warning only (B-01a).
- A3: "Ambiguous date" = text `d/d/yyyy` (or `-` / `.`) whose first two parts are both ≤ 12 and
  differ; native Excel dates are not ambiguous (B-01b).
- A4: Header rows 1-4 of the new fixtures (Topia column labels, technical keys, colour fills) are
  the schema of the tracked `tests/fixtures/topia_sample.xlsx`; they hold no employee data. All data
  rows are synthetic.
- A5: Test-only approved configuration (country/state list, keyed dependent source) is written to
  `tmp_path` inside a test; the shipped `business_rules.json` and `reference_data/*.json` are
  unchanged and approve nothing.
