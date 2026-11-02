# Defect and Exception Register

Evidence package item 9 of `clarifications/UAT_MASTER_COMMAND.md` §12 (lines 188-194). OCNORA task
T-1393 (UAT-08). Prepared 2026-11-02 at commit `7b71503`.

Every item below is **Open** unless its retest status says otherwise. A local retest proves engine
behaviour on synthetic fixtures only; **no item has an Equus retest** (Equus UAT execution:
**NOT EXECUTED**, no access from this environment). A Jake answer is evidence, not approval: every
tracker Status cell (`datamapping clarrification shef`!I2:I44) reads `Open`. Nothing is SIGNED OFF.

Severity scale: **Critical** (identity, data loss or confidentiality; blocks UAT) · **High** ·
**Medium** · **Low**. This is the scale of `docs/uat/CODE_REVIEW_REPORT.md`.

Retest status values:
- `LOCAL PASS (test id)`: the correction is implemented and the named test passes in the suite.
- `OPEN – decision`: waiting for an owner's decision or input; nothing to retest yet.
- `NOT EXECUTED – Equus`: can only be retested in Equus UAT.

## 1. Critical discrepancies D1–D5 (assessment §1)

| ID | Severity | Description and evidence | Impact | Owner | Correction (done / required) | Retest status |
|---|---|---|---|---|---|---|
| **D1** | Critical | Identifier. Jake's text (`JAKE!J2`): "The Topia ID should be entered for the Assignment Identifier and for ASSIGNMENT_CF!CF_TEXT029." His example contradicts it: `EX!E4` (Assignment Identifier) equals `EX!C4` (Employee Number), and `EX!Y4` (CF_TEXT029 "Topia File Number") differs from both. These three facts were checked with openpyxl on 2026-11-02; values are not reproduced. Source column F "Topia ID" is red and empty; column X "Assignment ID" is used | Every assignment's identity and child-row linkage (Q17, Q18). The wrong legacy id could reach CF_TEXT029 | Jake (field semantics) + Alka / John (which Topia column) | **Done (UAT-03):** the fallback to the Employee Number is removed; a missing identifier holds the row; X is kept as a documented assumption. **Required:** a written answer naming the source column for each field | Fallback removal: LOCAL PASS `tests/test_identity_controls.py::test_no_identifier_fallback_left_in_engine`, `tests/test_uat_scenarios_employee.py::test_EA06_missing_assignment_identifier_held_no_fallback`. The conflict itself: **OPEN – decision** |
| **D2** | Critical | Column A header. Bundled template A3 = `A`, Equus sample A3 = `A`, Jake's example `EX!A3` = `Employee Command`. Q5 (`JAKE!J6`): "Note that if there are issues in the header row, the full import fails." | A wrong header can make Equus reject the whole file | Jake (Equus template) | **Done (UAT-03):** header check against the approved template (TEMPLATE_VERSION `EAFH-bundled-1`, SHA-256 `c90e9fc0…45be7`); `HEADER_DECISION_PENDING` reported on every run. **Required:** the official current template | Control: LOCAL PASS `tests/test_identity_controls.py::test_pending_header_decision_reported_header_unchanged`, `tests/test_uat_scenarios_mapping.py::test_MD12_header_mismatch_fails_whole_file`. Equus header acceptance: **NOT EXECUTED – Equus** (NX-02) |
| **D3** | Critical | The real extract's "Dependent Import" sheet has no Employee or Assignment key | Dependents from the real extract are **not migrated**: 4 records in 2 rows are held (`docs/uat/RECONCILIATION_SUMMARY.md` §2) | Alka / Topia (keyed extract) | **Done (UAT-04):** unkeyed data held as `HELD_DEPENDENTS_UNKEYED`; a keyed source is configurable (ships unapproved). **Required:** a keyed dependent extract | Hold: LOCAL PASS `tests/test_family_rules.py::test_unkeyed_dependent_import_sheet_is_held_and_nothing_attached`. Migration of real dependents: **OPEN – decision** |
| **D4** | Critical (was) | Unapproved rules were active by default: Q40, Q34, Q22/Q23/Q37, Q35, Q12/Q3, Q32, Q1 fallback, `STRICT_PENDING` | Unapproved values in import files | Schneider (Q22, Q34, Q37, Q40); Jake (Q35, Q12, Q32); Alka / John (Q23) | **Done (UAT-02, UAT-03):** every rule is in `lambda/src/business_rules.json` with `approved_for_import: false`; output is named `_EQUUS_DRAFT_` while `NOT_READY`; the fallback is removed. **Required:** an approval per rule, then flip its flag | LOCAL PASS `tests/test_business_rules.py::test_shipped_configuration_declares_every_d4_rule_unapproved`, `tests/test_business_rules.py::test_default_output_has_no_unapproved_rule_values`. The rule decisions themselves are listed as E-01…E-06: **OPEN – decision** |
| **D5** | Critical | AG Confidentiality Type and DG Expected Assignment Outcome hold data and have no Equus column. ED PII Nationality 1 is excluded by configuration | Confidential and business data not migrated | Jake (Q6 template column); Schneider (Q7 SE-424, Q9, Q28) | **Done (UAT-02, UAT-05):** no longer silent. AG/DG give readiness `missing_target`, and all three are reconciled as `pending_target` with Q and owner. **Required:** a target column or an authorised exclusion per column | Reporting: LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD11_confidential_attribute_held_never_written`, `tests/test_reconciliation.py::test_pending_target_columns_are_explained`. Disposition: **OPEN – decision** |

## 2. Security and data-protection defects

| ID | Severity | Description and evidence | Impact | Owner | Correction | Retest status |
|---|---|---|---|---|---|---|
| **SEC-01** (CR-R01, confirmed) | **Critical** | The git-tracked file `FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx` (repository root, added in `ade394f`) has SHA-256 `006ba3b50df6cd60702692ff28da3b6a57f063c74a67f716859e56fd96ac18b8`. That is **byte-identical** to the real Topia extract used for reconciliation (`local_s3/source_s3/topia/incoming/…`, same hash in `docs/uat/RECONCILIATION_SUMMARY.md` §2). The tracked `DATA MIGRATION SE EAFH TEST FILE.xlsx` (SHA-256 `960b721b…30cc6`) is byte-identical to the Equus sample. Hashes computed with `shasum -a 256` on 2026-11-02; contents not opened. Also tracked: `backup.xlsx`, `Topia_Equus_Field_Mapping.xlsx`, and two outputs under `local_s3/dest_s3/equus/**` derived from the extract | Personal HR data (the project treats this workbook as personal data) is in the git history of every clone | Repository owner + Pratyush (data storage) | **Required:** remove from the index and rewrite history (`git filter-repo` or similar), rotate clones, and confirm removal. This needs explicit authorisation. **Not done:** this job may not rewrite history or edit `.gitignore` | **OPEN – decision** (blocks G9) |
| SEC-02 | High | Confidential-record visibility in Equus is undefined (Q28); confidentiality has no target (Q6) | Confidential assignments could be visible to the wrong users after import | Schneider security (Q28); Jake (Q6) | Security-group process and template column | **OPEN – decision**; Equus check **NOT EXECUTED – Equus** |
| SEC-03 | Low | Mapping reports (Warnings, Dropdown Values, Family sheets) quote source values (CR-H09) | Reports are personal data and must be handled like the source | Migration team | Handling rule in `docs/uat/PRE_IMPORT_CHECKLIST.md` §4; logs carry no values | Logs: LOCAL PASS `tests/test_handler_audit.py::test_logs_contain_no_fixture_names_or_emails`. Report handling: procedural |

## 3. Held and pending business exceptions (each blocks import-ready output)

| ID | Severity | Item (source cols → Equus) | Impact | Owner (Q) | Correction required | Retest status |
|---|---|---|---|---|---|---|
| E-01 | High | `international_relo_derivation` gated; `ASSIGNMENT!INTERNATIONAL_RELO` blank | Field empty for every assignment | Schneider (Q40) | Approved rule, or approval to leave blank | Gate: LOCAL PASS `tests/test_business_rules.py::test_enable_international_relo_restores_derivation`. **OPEN – decision** |
| E-02 | High | `cancellation_reason_sets_phase` gated; DH Cancellation Reason **held** | PHASE only from AC (yellow) | Jake + Schneider (Q34) | Status / phase / cancellation precedence | LOCAL PASS `tests/test_business_rules.py::test_enable_cancellation_phase_restores_cancelled`. **OPEN – decision** |
| E-03 | Critical | `default_effective_dates_from_start` gated. Salary and family-size `EFFECTIVE_DATE`, address `BEGIN_DATE`, contact `START_DATE` blank | Equus may require these dates (unknown, NX-07) | Jake (Q22); Alka / John (Q23); Schneider (Q37) | Written date rules | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD05_salary_without_effective_date_no_default_exception`. **OPEN – decision** |
| E-04 | High | `value_maps_employee_type_and_status` gated; S / U values sent as in the source | Values may not match Equus lists | Jake (Q35) | Crosswalk and lists | LOCAL PASS `tests/test_business_rules.py::test_enable_value_maps_restores_translation`. **OPEN – decision** |
| E-05 | Medium | `host_job_level_to_salary_grade` gated; CQ Host Job Level **held**; `TO_SALARY_GRADE` blank | Grade not migrated | Jake / Alka (Q32) | Written correction of the row-2 key | LOCAL PASS `tests/test_business_rules.py::test_enable_host_job_level_restores_salary_grade`. **OPEN – decision** |
| E-06 | High | `country_name_to_iso3` provisional (O, P, Q, AY, BD); 7 provisional cells in the real extract | Codes may not be Equus values (`XKX`, abbreviations; CR-M34) | Jake (Q12/Q3: Equus country list) | Approved country list | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD01a_country_names_provisional_unknown_passed_through`. **OPEN – decision**; Equus acceptance **NOT EXECUTED – Equus** (NX-06) |
| E-07 | High | Yellow columns S, AA, AB, AC, AK, AL, AO, AP (data present) and GO, GP, GQ, GR | Written to the DRAFT only; readiness `pending_column` | Alka / John; Q35, Q34, Q36, Q23, Q38, Q37, Q39 | Mapping review decision per column | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD07_red_yellow_cyan_brown_flags`. **OPEN – decision** |
| E-08 | Critical | Resource usernames (EI, EK, EM, EO, EQ, FD, FF–FI): `resource_usernames` list empty | Display names sent where Equus expects usernames | Jake (Q2) | Versioned list with approval | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD14_unknown_reference_values_resource_without_username`. **OPEN – decision** |
| E-09 | Critical | Currencies: FROM/TO/DEFAULT/salary currency blank; no derivation | Compensation incomplete | Jake (Q4), Schneider comp (Q21, Q38) | Approved source per field | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD03_currency_blank_never_derived_from_country`. **OPEN – decision** |
| E-10 | Critical | Dropdown lists missing for every dropdown header, including spouse relationship | No value can be validated against Equus | Jake (Q14, Q20) | Versioned `equus_lists.json` | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD13_value_outside_equus_list_reported_unchanged`. **OPEN – decision**; **NOT EXECUTED – Equus** (NX-06) |
| E-11 | Critical | Cost centres: `cost_centres` list empty; BL name holds a code | Code/name pairs unvalidated | Jake (Q15), Alka (Q42) | Code → name list | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD14_unknown_reference_values_resource_without_username` (reference-list pattern). **OPEN – decision** |
| E-12 | High | States: `country_state` list empty (BC, BH) | States unvalidated | Jake (Q12) | Equus state list | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD02b_state_unvalidated_while_list_not_approved`. **OPEN – decision** |
| E-13 | Medium | Resource e-mails EJ, EL, EN, EP, ER, FE not migrated (`pending_target`) | Data not migrated | Jake / Schneider (Q41) | Disposition | LOCAL PASS `tests/test_reconciliation.py::test_pending_target_columns_are_explained`. **OPEN – decision** |
| E-14 | Medium | Host pay FX/FY: target not in the template | Host pay not migrated | Jake + Schneider (Q21) | Compensation crosswalk | **OPEN – decision** |
| E-15 | Medium | 74 red columns excluded (58 cells with data in the real extract) | Data not migrated | Schneider (Q30) | Written disposition per column | LOCAL PASS `tests/test_reconciliation.py::test_gated_excluded_blank_and_provisional_categories`. **OPEN – decision** |
| E-16 | Medium | Placeholders blanked in all fields (M, N; glued placeholder in N6) | Possible loss if a token is real data | Alka (Q43) | Per-field pattern list | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD10_placeholders_blank_and_glued_placeholder_reported`. **OPEN – decision** |
| E-17 | Medium | Ambiguous text dates read month-first and warned (F-04) | Day/month swap possible | Jake (Q24) | Accepted formats per field | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD04a_ambiguous_text_date_read_with_configured_order_and_reported`. **OPEN – decision** |
| E-18 | Low | Host weekly hours 158 kept and warned | Implausible value imported | Schneider (Q16) | Source-owner confirmation | LOCAL PASS `tests/test_uat_scenarios_mapping.py::test_MD15_weekly_hours_outside_range_kept_and_warned`. **OPEN – decision** |
| E-19 | Medium | Every address and contact gets `PRIMARY_* = Yes` (CR-T17, PC-J-15) | Several "primary" records per assignment | Jake (no Q yet) | Rule for the primary flag | **OPEN – decision** |
| E-20 | High | Employee command for an employee created by an earlier batch is INSERT (F-01) | Possible duplicate employee or Equus rejection | Jake (Q26) | INSERT/UPDATE/IGNORE action matrix | Characterised: LOCAL PASS `tests/test_uat_scenarios_employee.py::test_EA11b_existing_employee_from_earlier_run_new_assignment_pending_q26`. **NOT EXECUTED – Equus** (NX-05) |
| E-21 | Medium | Same bytes under a new S3 key are transformed again (F-02) | A re-import risk if the procedure is skipped | Migration team (Q5) | Accepted Identifiers check, `docs/uat/REPROCESSING_PROCEDURE.md` step 6 | LOCAL PASS `tests/test_uat_scenarios_aws.py::test_AW08b_same_bytes_new_key_detected_by_accepted_identifier_check` |
| E-22 | Low | A manifest-put failure leaves report + import without a manifest until retry (F-05) | Temporary incomplete audit trail | Migration team | Lambda retry or manual re-drive completes it | LOCAL PASS `tests/test_uat_scenarios_aws.py::test_AW03_lambda_retry_after_failed_put_completes_without_duplicates` |
| E-23 | High | Source inventory: only a 2-row extract exists; cutoff unknown | Reconciliation of the final population not possible | Alka / Schneider (Q25) | Final extract with hash, cutoff and expected counts | **OPEN – decision** |
| E-24 | High | Reconciliation acceptance criteria not agreed | G5 cannot be signed | Schneider + Jake (Q27) | Agreed criteria | Engineering criteria: LOCAL PASS `tests/test_reconciliation.py::test_children_fixture_reconciles_with_zero_unexplained`. **OPEN – decision** |
| E-25 | High | No named UAT signatory | G11 and G12 cannot be obtained | Schneider PMO (Q31) | Named signatory and process | **OPEN – decision** |
| E-26 | Medium | Custom-field inventory unconfirmed (cyan → CF from template labels) | CF targets may be wrong | Jake / Alka (Q29); Renniece / Cathy (Q33) | Signed-off inventory | **OPEN – decision** |
| E-27 | Medium | Pay installments AK/AL meaning unknown (yellow) | Values may be misread | Schneider payroll (Q36) | Meaning and values | **OPEN – decision** |
| E-28 | Low | `CF_TEXT016` meaning; template Q–T label shift | Field unused | Jake (Q3) | Corrected template labels | Warning: LOCAL PASS `tests/test_migration.py::test_country_of_birth_converted_despite_shifted_template_rules`. **OPEN – decision** |
| E-29 | Medium | Equus mandatory-field list unknown | Rows may be rejected by Equus | Jake (Q14, Q20) | List of mandatory columns per record type | **NOT EXECUTED – Equus** (NX-07) |
| E-30 | Low | Q8, Q10, Q13, Q17, Q18, Q19 confirmed by Jake but not formally approved | Implemented on evidence only | Schneider (formal approval, PC-S-18) | Formal approval | Behaviour: LOCAL PASS (see `docs/uat/RTM_Q1-Q43.md`). **OPEN – decision** |

## 4. Code-review findings closed or reduced by UAT-02…07

Only findings with a passing test are listed as corrected. Every other CR item in
`docs/uat/CODE_REVIEW_REPORT.md` keeps its original status. This package has not re-reviewed them.

| CR | Correction | Test |
|---|---|---|
| CR-T09 / CR-M06 fallback | Identifier fallback removed (UAT-03) | `tests/test_identity_controls.py::test_no_identifier_fallback_left_in_engine` |
| CR-T10 | Source duplicate detection (UAT-03) | `tests/test_identity_controls.py::test_exact_duplicate_row_one_accepted_one_held` |
| CR-T11 | Missing Employee Number held, never treated as a known employee (UAT-03) | `tests/test_identity_controls.py::test_missing_employee_number_is_held` |
| CR-T14 / CR-M31 | Name warning reworded as a DATA_QUALITY note (UAT-04) | `tests/test_family_rules.py::test_spouse_without_names_is_written_with_data_quality_note_not_held` |
| CR-T16 / CR-H03 / CR-M03 | Header check against the approved template; D2 reported (UAT-03) | `tests/test_identity_controls.py::test_header_mismatch_fails_file_and_lists_columns` |
| CR-T18 / CR-T19 | Child-row and value-level reconciliation (UAT-05) | `tests/test_reconciliation.py::test_dropped_dependent_first_name_on_update_row_is_unexplained` |
| CR-T24 | UPDATE-row shape validated against Jake's example (UAT-03, UAT-04) | `tests/test_family_rules.py::test_jake_example_has_the_structure_this_test_encodes` |
| CR-T26 / CR-H05 | Readiness independent of `STRICT_PENDING`; PENDING failure keeps the manifest (UAT-02, UAT-07) | `tests/test_business_rules.py::test_readiness_independent_of_strict_pending` `tests/test_handler_audit.py::test_strict_pending_is_pending_failure` |
| CR-M25 | `RULE_TARGETS` removed (UAT-02) | `tests/test_business_rules.py::test_default_output_has_no_unapproved_rule_values` |
| CR-M32 / CR-T22 | Missing dropdown list = readiness exception (UAT-02) | `tests/test_uat_scenarios_mapping.py::test_MD13_value_outside_equus_list_reported_unchanged` |
| CR-M35 | Versioned, validated configuration (UAT-02) | `tests/test_business_rules.py::test_malformed_business_rules_fail_fast_naming_path` |
| CR-H02 | Code version hashes the JSON configuration (UAT-02) | `tests/test_business_rules.py::test_output_key_changes_when_rule_approval_changes` |
| CR-H07 | Write order and retry completion (UAT-06, B-04) | `tests/test_uat_scenarios_aws.py::test_AW03_lambda_retry_after_failed_put_completes_without_duplicates` |
| CR-H08 | DRAFT naming, metadata and run manifest (UAT-02, UAT-07) | `tests/test_handler_audit.py::test_run_writes_manifest_with_every_listed_field` |
| CR-T01 (partly) | Ambiguous dates are warned; the parse order is still unapproved (E-17) | `tests/test_uat_scenarios_mapping.py::test_MD04a_ambiguous_text_date_read_with_configured_order_and_reported` |

## 5. Counts

- Critical items open: D1, D2, D3, D5 (disposition), SEC-01, E-03, E-08, E-09, E-10, E-11 = **10**.
- D4 is corrected in code: every rule is gated. Its six rule decisions remain open as E-01…E-06.
- Items with an Equus-only retest: D2 (NX-02), E-06, E-10 (NX-06), E-20 (NX-05), E-29 (NX-07),
  SEC-02. All are **NOT EXECUTED**.
