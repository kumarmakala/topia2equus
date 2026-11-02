# Equus UAT Execution Plan: operator runbook (master command §9, steps 1–5)

Evidence package item 8 of `clarifications/UAT_MASTER_COMMAND.md` §12 ("Equus UAT Execution
Evidence"). OCNORA task T-1393 (UAT-08). Prepared 2026-11-02.

> **STATUS: NOT EXECUTED.** No step of this plan has been run. This environment has no Equus
> access, and no Equus import, Equus job result or post-import audit exists. Every evidence field
> below is empty on purpose. The Equus UAT execution evidence (item 8) is therefore
> **INCOMPLETE**. Missing inputs: authorised Equus UAT access, Equus lookups (Q2–Q4, Q12, Q14, Q15),
> the official template (D2), the identifier decision (D1) and a named signatory (Q31).
> Master command §9 (line 168): "Never perform the import automatically without explicit
> authorisation and access; never mark an unexecuted scenario as passed."

## 0. Entry criteria (all must be true before step 1 starts; today none of the bold ones is)

| # | Criterion | Evidence source | Today |
|---|---|---|---|
| 0.1 | The run produced an `IMPORT_CANDIDATE` (`readiness.status = READY_FOR_UAT_EXECUTION`, no `_DRAFT_` in the name) | run manifest `equus/reports/<stem>_RUN_<stamp>.json` (`lambda/src/handler.py:431-474`) | **No.** Every run with the shipped configuration is `NOT_READY` (`tests/test_business_rules.py::test_default_readiness_not_ready_and_lists_every_gated_and_pending_item`) |
| 0.2 | D1 (identifier) and D2 (A3 header) resolved in writing | `docs/uat/DEFECT_EXCEPTION_REGISTER.md` §1 | **No** |
| 0.3 | Equus lookups delivered and loaded with `approved_for_import: true` | `lambda/src/reference_data/*.json`, `lambda/src/equus_lists.json` | **No.** The lists are empty and `equus_lists.json` does not exist |
| 0.4 | Authorised UAT environment, operator and Equus contact named in writing (never production) | written authorisation | **No** |
| 0.5 | Final source extract identified by hash, cutoff and expected counts (Q25) | Q25 answer | **No** |
| 0.6 | Security finding SEC-01 dispositioned | register §2 | **No** |

## Step 1 — Pre-import validation

Procedure: tick every box of `docs/uat/PRE_IMPORT_CHECKLIST.md` §2 (2a–2d). Any "no" means stop.
Also check the reconciliation sheet (`Result = OK`, every `Unexplained` = 0;
`docs/uat/RECONCILIATION_SUMMARY.md` §1).

| Evidence field | Value (operator fills) |
|---|---|
| Source S3 key / ETag | |
| Source SHA-256 (manifest `source.sha256`) | |
| Template SHA-256 / version (manifest `template`) | |
| Business rules version / SHA-256 | |
| Code version / SHA-256 | |
| Full test-suite command, counts, date | |
| Readiness status / reason count | |
| Reconciliation result / unexplained total | |
| Rows accepted / held / written | |
| Checklist completed by (name, date) | |
| Result (GO / NO-GO) | NOT EXECUTED |

## Step 2 — Controlled import into the authorised Equus UAT environment

Procedure: an authorised person imports exactly the `IMPORT_CANDIDATE` workbook named in the
manifest. Never a `_DRAFT_` file, and never production. Before upload, compute the file's SHA-256
and compare it with the S3 object. Record the batch. Do not re-upload after a partial failure
without step 4.

| Evidence field | Value (operator fills) |
|---|---|
| Equus environment / tenant (UAT) | |
| Authorised operator / authorisation reference | |
| File name and SHA-256 at upload | |
| Import batch / job id | |
| Start / end timestamp (UTC) | |
| Equus job result (file-level: accepted / failed) | |
| Rows accepted / rows rejected (Equus counts) | |
| Rejected-row file (attached, location) | |
| Equus error text per rejected row (attached) | |
| Result | NOT EXECUTED |

## Step 3 — Post-import verification

Procedure: extract the imported records from Equus. Reconcile them per record against the run's
**Accepted Identifiers** sheet and the mapping report. Check identifiers, Employee Number,
Assignment Identifier, CF_TEXT029, dependents, spouses, compensation, dates and custom fields. Run
the Equus duplicate audit report (Q5, `JAKE!J6`: "an audit report can be run to identify them"):
duplicates by Employee Number + Assignment Identifier.

| Evidence field | Value (operator fills) |
|---|---|
| Equus extract / audit report reference | |
| Employees expected / found in Equus | |
| Assignments expected / found | |
| Dependents / spouses expected / found | |
| Field mismatches (count, list attached) | |
| Duplicate audit result (count, list) | |
| Verified by (independent of the operator) | |
| Result | NOT EXECUTED |

## Step 4 — Error correction and resubmission

Procedure: follow `docs/uat/REPROCESSING_PROCEDURE.md`. Correct rejected or held rows **at the
source**, re-run the Lambda with a new source object, and pass the previous Accepted Identifiers
(`previously_accepted`). Confirm the correction batch contains **only** corrected identifiers before
importing it (scenario AW-08 locally: `tests/test_uat_scenarios_aws.py::test_AW08_correction_batch_rerun_imports_only_corrected_rows`).
Never delete Equus records without explicit authorisation (master command §11).

| Evidence field | Value (operator fills) |
|---|---|
| Rejected / held rows corrected (count, codes) | |
| Correction source SHA-256 and run manifest | |
| Overlap of correction batch with earlier accepted identifiers (must be 0) | |
| Correction import batch id / result | |
| Record counts before / after (no duplicates) | |
| Result | NOT EXECUTED |

## Step 5 — Business validation

Procedure: business reviewers inspect a sample of the imported records and the evidence package
(RTM, reconciliation, this plan's evidence). They record their decision in
`docs/uat/BUSINESS_SIGNOFF_RECORD.md`. The developer or AI never records it.

| Evidence field | Value (business reviewer fills) |
|---|---|
| Reviewers (names, roles) | |
| Records sampled (identifiers list location) | |
| Findings / exceptions raised | |
| Decision reference (sign-off record) | |
| Result | NOT EXECUTED |

## Scenarios that must be executed in Equus (from `docs/uat/TEST_RESULTS.md` §4)

| ID | Scenario | Step | Status |
|---|---|---|---|
| NX-01 | Row-level Equus rejection | 2, 4 | NOT EXECUTED |
| NX-02 | File-level Equus rejection (A3 header, D2) | 2 | NOT EXECUTED |
| NX-03 | Correction and resubmission without duplicates | 4 | NOT EXECUTED |
| NX-04 | Post-import audit and per-record reconciliation | 3 | NOT EXECUTED |
| NX-05 | Previously imported assignment / existing employee (Q26, F-01) | 2, 3 | NOT EXECUTED |
| NX-06 | Values inside and outside Equus lists (Q2–Q4, Q12, Q14, Q15) | 2 | NOT EXECUTED |
| NX-07 | Equus mandatory fields | 2 | NOT EXECUTED |

Gates fed by this plan: G6 (step 2), G7 (step 3), G8 (steps 2–4, NX-01…NX-05), G11 (step 5). All
are **Blocked** in `docs/uat/UAT_READINESS_ASSESSMENT.md`.
