# UAT evidence package: index

Topia → Equus EAFH migration, final UAT candidate. Requirement: `clarifications/UAT_MASTER_COMMAND.md`.
Evidence base: `clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md` and Jake's workbook
`clarifications/jake-1.xlsx`. Index maintained by OCNORA task T-1393 (UAT-08), 2026-11-02, commit
`7b71503`.

**Readiness status: NOT READY FOR UAT** (`UAT_READINESS_ASSESSMENT.md` §3).
- Equus UAT execution, import results and post-import audit: **INCOMPLETE / NOT EXECUTED**. This
  environment has no Equus access.
- No document in this package records a sign-off. Sign-off belongs to the authorised business
  signatory (master command §0 principle 10).

## Section 12 deliverables

| # | Deliverable (MC §12) | Document | Status | Missing input |
|---|---|---|---|---|
| 1 | Requirements Traceability Matrix (Q1–Q43) | `RTM_Q1-Q43.md` (Tests column and §4 verbatim answers updated in UAT-08) | Complete as a record; approvals INCOMPLETE | Formal approval of every Q (tracker I = Open); answers to Q21–Q43 |
| 2 | Code Review Report | `CODE_REVIEW_REPORT.md` (UAT-01, base `e61d99f`) | Complete as of its base; findings closed since are listed in `DEFECT_EXCEPTION_REGISTER.md` §4 | — |
| 3 | Change Log | `CHANGE_LOG.md` | Complete (UAT-02…07) | — |
| 4 | Updated code | `lambda/src/transform.py`, `mapping_config.py`, `handler.py`, `business_rules.json`, `reference_data/*.json`; `local/reconcile.py` | Complete for the engineering scope; business rules gated OFF | Approvals and Equus lookups |
| 5 | Automated test results | `TEST_RESULTS.md` (UAT-06); per-task results in `UAT-02_business_rules_configuration.md`, `UAT-03_identity_controls.md`, `UAT-04_family_rules.md`, `RECONCILIATION_SUMMARY.md` §4 | Complete (local) | — |
| 6 | Field Mapping Validation | `FIELD_MAPPING_VALIDATION.md` | Complete as local evidence; NOT APPROVED | Mapping approvals (G1) |
| 7 | Source-to-Output Reconciliation | `RECONCILIATION_SUMMARY.md` | Local reconciliation OK; business criteria INCOMPLETE | Q27 criteria, Q25 final extract, D3 keyed dependents |
| 8 | Equus UAT Execution Evidence | `EQUUS_UAT_EXECUTION_PLAN.md` (runbook with empty evidence fields) | **INCOMPLETE: NOT EXECUTED** | Equus UAT access and authorisation; import-ready candidate; Equus lookups; D1/D2 |
| 9 | Defect and Exception Register | `DEFECT_EXCEPTION_REGISTER.md` | Complete (all items open unless retested) | Owner decisions |
| 10 | Pending Clarification Register | `PENDING_CLARIFICATIONS.md` | Complete; every item Open | Owner answers |
| 11 | UAT Readiness Assessment | `UAT_READINESS_ASSESSMENT.md` | Complete: G1–G12 assessed, final status NOT READY FOR UAT | — |
| 12 | Business Sign-off Record | `BUSINESS_SIGNOFF_RECORD.md` | **INCOMPLETE: empty template**, never filled by the developer or AI | Named signatory (Q31) and decision |

Section 15 response (A–H): `FINAL_RESPONSE.md`.

## Supporting documents

| Document | Content |
|---|---|
| `PRE_IMPORT_CHECKLIST.md` | Operator checklist for the run manifest before any import (UAT-07) |
| `REPROCESSING_PROCEDURE.md` | Correction batch without re-importing accepted rows (UAT-03) |
| `UAT-02_business_rules_configuration.md` | Rule gating and approval configuration |
| `UAT-03_identity_controls.md` | Identity, duplicate and header controls |
| `UAT-04_family_rules.md` | Spouse and dependent rules |

## UAT-08 checks

`./.venv/bin/python -m pytest -q tests/test_uat_evidence_package.py` checks that this package is
consistent:
- every §12 deliverable exists or is marked INCOMPLETE with its missing input;
- G1–G12 each have a result, and there is exactly one final status that agrees with the gates;
- the change log covers every changed function and names an existing test id;
- the field mapping table matches the code;
- no document claims an Equus import or a sign-off.

Result on 2026-11-02: **9 passed**. `test_results_doc_every_pass_names_an_existing_test_and_not_executed_names_evidence`
in `tests/test_uat_scenarios_aws.py` also passed against the unchanged `TEST_RESULTS.md`. The full
suite was not run in this job; the runner's gate runs it.

## Assumptions recorded by UAT-08

1. The integration branch is the current `HEAD` (`7b71503`). The change log is derived from the
   commits of UAT-02…07 listed in `CHANGE_LOG.md` §0.
2. G4 is assessed as Pass on the recorded results and the runner gate history. The runner's gate
   for this job re-executes the full suite and is authoritative.
3. G5 is Blocked, not Pass: the local reconciliation is OK, but its acceptance criteria (Q27) and
   the final extract (Q25) are missing, and real dependents are held (D3).
4. G9 is Fail, not Blocked. A tracked workbook is byte-identical to the real extract (SEC-01): a
   confirmed security defect, not a pending decision. Its contents were not opened.
5. The documentation-consistency test is test code, not production code. It is added under
   `tests/` so that each acceptance criterion has executable evidence.
