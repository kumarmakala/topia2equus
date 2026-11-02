# UAT Readiness Assessment: gates G1–G12

Evidence package item 11 of `clarifications/UAT_MASTER_COMMAND.md` §12. Gates as defined in
§10 (lines 171-179). OCNORA task T-1393 (UAT-08). Assessed 2026-11-02 at commit `7b71503` (no code
change in this task).

Rules applied:
- Each gate is **Pass**, **Fail** or **Blocked**. Blocked means the gate cannot be evaluated or
  passed until a missing decision, input or access is supplied.
- A gate needing Equus is never Pass from local tests.
- A Jake answer is evidence, not approval: every tracker Status cell
  `datamapping clarrification shef`!I2:I44 reads `Open` (openpyxl, 2026-11-02).
- The final status is exactly one of the four values in §10. This assessment never assigns UAT
  SIGNED OFF.

## 1. Gate table

| Gate | Requirement (MC §10) | Result | Evidence (file:line, test id) | What closes it (owner) |
|---|---|---|---|---|
| G1 | Mappings approved or dispositioned | **Blocked** | No mapping is formally approved (`JAKE!I2:I44` all `Open`). Q21–Q43 have no answer (`JAKE!J22:J44` empty). Pending columns are listed in `docs/uat/FIELD_MAPPING_VALIDATION.md` §3. Classification: `docs/uat/RTM_Q1-Q43.md` §2 (7 CONFIRMED BY JAKE, 4 PARTIALLY CLARIFIED, 32 PENDING). Test: `tests/test_business_rules.py::test_default_readiness_not_ready_and_lists_every_gated_and_pending_item` | Approval or written disposition of Q1–Q43 (Schneider, Jake, Alka) |
| G2 | Equus custom fields and lookups available | **Blocked** | `lambda/src/reference_data/{resource_usernames,cost_centres,country_state}.json` have empty `values` and are not approved. No `lambda/src/equus_lists.json` exists. The confidentiality column is missing from the template (Q6, `JAKE!J7`: "This is pending with Equus. I will work on having the needed column/field added to the import template."). Tests: `tests/test_business_rules.py::test_shipped_reference_lists_are_empty_placeholders`, `tests/test_uat_scenarios_mapping.py::test_MD14_unknown_reference_values_resource_without_username` | Jake: Q2, Q3, Q4, Q12, Q14, Q15, Q20 lists; Q6 template column |
| G3 | Identity and linkage conflicts resolved | **Blocked** | D1 open: `JAKE!J2` vs `EX!C4`/`EX!E4`/`EX!Y4`, identifier label in `lambda/src/business_rules.json` `identity_controls`. D2 open: `HEADER_DECISION_PENDING` (`lambda/src/transform.py:804`). D3 open: unkeyed dependents held (`transform.py:911`). The controls work locally: `tests/test_identity_controls.py::test_missing_assignment_id_is_held_never_falls_back`, `tests/test_identity_controls.py::test_pending_header_decision_reported_header_unchanged`, `tests/test_family_rules.py::test_unkeyed_dependent_import_sheet_is_held_and_nothing_attached` | Jake + Alka / John (D1); Jake (D2); Alka / Topia (D3) |
| G4 | Automated tests pass with recorded results | **Pass** (local) | 257 tests collected (`./.venv/bin/python -m pytest --collect-only -q tests`, 2026-11-02), plus the new `tests/test_uat_evidence_package.py` of this task. Recorded runs: `docs/uat/TEST_RESULTS.md` §1 (102 + 155 = 257 passed). The runner gate passed on every UAT commit (`docs/uat/CHANGE_LOG.md` §0). The runner gate for this job is authoritative | Keep green; re-run on every configuration change |
| G5 | Reconciliation passes | **Blocked** | The local engineering reconciliation passes: real extract 222 non-blank cells, 0 unexplained, counts balance (`docs/uat/RECONCILIATION_SUMMARY.md` §2), `tests/test_reconciliation.py::test_children_fixture_reconciles_with_zero_unexplained`. Business acceptance criteria are unanswered (Q27, `JAKE!J28` empty). Only a 2-row extract exists (Q25). Real dependents are not migrated (D3: 4 held records). No Equus-side reconciliation exists (G7) | Schneider + Jake (Q27); Alka / Schneider (Q25); D3 |
| G6 | Controlled Equus UAT import completed | **Blocked** | **NOT EXECUTED**: no Equus access; entry criteria 0.1–0.6 unmet (`docs/uat/EQUUS_UAT_EXECUTION_PLAN.md` §0). Every run is `NOT_READY` / `_EQUUS_DRAFT_` (`lambda/src/handler.py:431-434`; `tests/test_uat_scenarios_aws.py::test_AW01_valid_s3_trigger_writes_draft_report_manifest`) | Authorised Equus UAT access and an import-ready candidate |
| G7 | Equus records independently validated | **Blocked** | **NOT EXECUTED**: no Equus records exist (`docs/uat/TEST_RESULTS.md` §4 NX-04) | G6, then plan step 3 |
| G8 | Duplicate, error and retry scenarios verified | **Blocked** | Local side passes: `tests/test_uat_scenarios_employee.py::test_EA04_duplicate_source_row_held_once_written`, `tests/test_uat_scenarios_aws.py::test_AW02_duplicate_event_and_duplicate_record_write_nothing_new`, `tests/test_uat_scenarios_aws.py::test_AW03_lambda_retry_after_failed_put_completes_without_duplicates`. Equus side NX-01…NX-05 **NOT EXECUTED**. Q5 (`JAKE!J6`): "There are some protections in place to prevent duplicate employees from being created, but I would not rely on this as a fail-safe." | Equus UAT steps 2–4 |
| G9 | Security and confidentiality checks pass | **Fail** | **SEC-01**: the tracked root file `FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx` has SHA-256 `006ba3b5…ac18b8`, identical to the real extract (`docs/uat/DEFECT_EXCEPTION_REGISTER.md` §2), so personal data is in git history. Q6 / Q28 confidentiality undecided; AG not migrated (D5). Logs are clean: `tests/test_handler_audit.py::test_logs_contain_no_fixture_names_or_emails` | Authorised history clean-up (repository owner + Pratyush); Q6 (Jake), Q28 (Schneider security) |
| G10 | No unresolved critical defects or unapproved critical exceptions | **Fail** | Open Critical items: D1, D2, D3, D5 disposition, SEC-01, E-03, E-08, E-09, E-10, E-11 (`docs/uat/DEFECT_EXCEPTION_REGISTER.md` §5). D4 itself is corrected (all rules gated: `tests/test_business_rules.py::test_shipped_configuration_declares_every_d4_rule_unapproved`) | Close or formally accept each critical item |
| G11 | Business reviewers approve the evidence | **Blocked** | Not obtained. No reviewer is named (Q31, `JAKE!J32` empty). `docs/uat/BUSINESS_SIGNOFF_RECORD.md` is an empty template | Schneider PMO (Q31) |
| G12 | Authorised signatory signs off | **Blocked** | Not obtained. No signatory is named (Q31, `JAKE!J32` empty). `docs/uat/BUSINESS_SIGNOFF_RECORD.md` is an empty template. The developer or AI never signs (`clarifications/UAT_MASTER_COMMAND.md:29`, principle 10) | Named signatory after G1–G11 |

Totals: Pass 1 (G4) · Fail 2 (G9, G10) · Blocked 9 (G1, G2, G3, G5, G6, G7, G8, G11, G12).

## 2. Why this status, and not another

- **READY FOR UAT EXECUTION** would need G1–G5, G9 and G10 to be Pass (the gates that precede an
  import) and an `IMPORT_CANDIDATE` run. Here G1, G2, G3 and G5 are Blocked and G9 and G10 are Fail.
  The tool itself computes `NOT_READY` for every run
  (`tests/test_business_rules.py::test_default_readiness_not_ready_and_lists_every_gated_and_pending_item`).
- **UAT EXECUTED — SIGN-OFF PENDING** would need G6–G8 executed in Equus. Not executed.
- **UAT SIGNED OFF** is never assigned by this assessment (MC §10, line 179).
- The open items named by the task and the master command are still open: Q2–Q4, Q14, Q15
  lookups (G2); Q6, Q7 and Q9 decisions (G1, G9); D1 and D2 (G3, G10).

## 3. Final status

Final readiness status: **NOT READY FOR UAT**

Supporting gates: G1, G2, G3, G5 Blocked; G9, G10 Fail (section 1).
