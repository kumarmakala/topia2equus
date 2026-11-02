# Final Response: master command §15 (A–H)

Format of `clarifications/UAT_MASTER_COMMAND.md` §15 (lines 213-218). OCNORA task T-1393 (UAT-08).
Prepared 2026-11-02 at commit `7b71503`. Index of the evidence package: `docs/uat/README.md`.

## A. Executive summary

- **Built (UAT-02…07).** The Topia → Equus transformation now:
  - holds every unapproved business rule behind versioned configuration (all switched off);
  - holds rows with missing, duplicate or conflicting identifiers;
  - no longer falls back from a missing Assignment Identifier to the Employee Number;
  - follows Jake's INSERT-then-UPDATE dependent layout;
  - reconciles every source cell, child rows included;
  - leaves an auditable run manifest for each Lambda run.
- **Tests.** 257 automated tests on synthetic fixtures; every UAT commit passed the runner's gate.
- **What blocks UAT.** Business decisions and Equus inputs are still missing:
  - the identifier conflict (D1) and the A3 header (D2);
  - a keyed dependent source (D3);
  - Equus lookups (Q2–Q4, Q12, Q14, Q15);
  - decisions on confidential and custom fields (Q6, Q7, Q9, Q28);
  - answers to Q21–Q43, which are all unanswered.
- **Security.** The real Topia extract is tracked in git history (SEC-01).
- **Equus.** No import has been executed or observed.
- **Readiness: NOT READY FOR UAT.**

## B. Business requirement status Q1–Q43: Jake's answers vs formally approved rules

Jake's exact words are in `docs/uat/RTM_Q1-Q43.md` §4, read with openpyxl from
`clarifications/jake-1.xlsx` (`datamapping clarrification shef`, column J; Qn = row n+1). **Formally
approved rules: none.** Every Status cell I2:I44 reads `Open`.

| Q | Topic (col B) | Jake answered? | Classification (RTM §1) | Formally approved | Implemented as |
|---|---|---|---|---|---|
| Q1 | Identity | yes (J2) | PARTIALLY CLARIFIED (D1) | No | Employee Number = E; `NEW`; identifier and CF_TEXT029 from X; no fallback; missing → held |
| Q2 | Identity (usernames) | yes (J3: values to follow) | PENDING EQUUS CONFIGURATION | No | Empty list; names sent unchanged and reported |
| Q3 | Equus template | yes (J4: values to follow) | PENDING EQUUS CONFIGURATION | No | Q–T shift warned; CF_TEXT016 blank |
| Q4 | Currency | yes (J5: values to follow) | PENDING EQUUS CONFIGURATION | No | Currencies blank; nothing derived |
| Q5 | Equus import | yes (J6) | CONFIRMED BY JAKE | No | Duplicate, conflict and header controls; FILE-level withholding; reprocessing |
| Q6 | Custom fields | yes (J7: pending Equus) | PENDING EQUUS CONFIGURATION | No | AG held (`missing_target`) |
| Q7 | Custom fields | yes (J8: pending Schneider) | PENDING BUSINESS APPROVAL | No | DG held (`missing_target`) |
| Q8 | Source colour legend | yes (J9) | CONFIRMED BY JAKE | No | Brown = no meaning |
| Q9 | Custom fields / privacy | yes (J10) | PENDING BUSINESS APPROVAL | No | ED not migrated; reported with Q9 |
| Q10 | Family | yes (J11) | CONFIRMED BY JAKE | No | Nameless spouse written plus a DATA_QUALITY note |
| Q11 | Family | yes (J12) | PARTIALLY CLARIFIED (D3) | No | Unkeyed dependents held; keyed source configurable, unapproved |
| Q12 | Addresses | yes (J13) | PARTIALLY CLARIFIED | No | ISO-3 provisional; state and country lists empty; city free text |
| Q13 | Contacts | yes (J14) | CONFIRMED BY JAKE | No | G → USER_ID_EMAIL, H → CF_TEXT008; no e-mail contacts |
| Q14 | Reference values | yes (J15: lists to follow) | PENDING EQUUS CONFIGURATION | No | Every dropdown reported `dropdown_list_missing` |
| Q15 | Cost centres | yes (J16: values to follow) | PENDING EQUUS CONFIGURATION | No | Empty list; code-like names warned |
| Q16 | Working hours | yes (J17: values to follow) | PENDING BUSINESS APPROVAL | No | 158 kept and warned |
| Q17 | Historical assignments | yes (J18) | CONFIRMED BY JAKE | No | One INSERT/NEW per assignment; never merged |
| Q18 | Dependent linkage | yes (J19) | CONFIRMED BY JAKE | No | Every row carries its own identifier; cross-linkage held |
| Q19 | Additional children | yes (J20) | CONFIRMED BY JAKE | No | INSERT + n UPDATE rows, as in ` Mny dependents Ex` rows 4–7 |
| Q20 | Dependent requirements | yes (J21) | PARTIALLY CLARIFIED | No | One populated field creates a record; lookups pending |
| Q21–Q43 | Compensation, dates, formats, inventory, import actions, reconciliation, confidentiality, custom fields, scope, approvals, grades, employment type, lifecycle, references, installments, family size, currency, RMC, derived values, resource fields, CF_TEXT024, missing values | **no** (J22:J44 empty) | PENDING BUSINESS APPROVAL or PENDING EQUUS CONFIGURATION (RTM §2) | No | Unapproved rules gated off (Q22, Q23, Q32, Q34, Q35, Q37, Q40); everything else reported with its Q and owner |

## C. Actual code changes

Full table (file, function, before, after, reason, test evidence) for UAT-02…07:
`docs/uat/CHANGE_LOG.md` §1–§6. There are 92 rows, and every changed function is listed (checked by
`tests/test_uat_evidence_package.py::test_change_log_lists_every_changed_function`).

Key changes:
- **UAT-02 (`3f92551`):** `business_rules.json`, plus loaders `load_business_rules` /
  `load_reference_data`. Rules are applied only when approved: `apply_assignment_rules`,
  `default_child_dates`, `_convert`. Readiness comes from `compute_readiness`, and files are named
  DRAFT while not ready (`handler.lambda_handler`).
- **UAT-04 (`66b7e25`):** `validate_family` (DATA_QUALITY note), `read_dependent_sources`,
  `_attach_keyed_dependents`, Family sheet.
- **UAT-03 (`ddf01f5`):** the Employee Number fallback is removed in `build_records`;
  `screen_rows`, `identifier_text`, `validate_row_structure`, `check_template_headers`; Exceptions
  and Held Rows sheets.
- **UAT-07 (`452e9e0`):** `_process`, run manifest, `classify_failure`, secure logging.
- **UAT-05 (`b661f43`):** `reconcile` keyed on the identifier, with child rows and per-cell
  categories; `local/reconcile.py`.
- **UAT-06 (`7b71503`):** `check_assignment_dates`, ambiguous-date warning, state validation,
  duplicate dependents held, handler write order (B-01a…B-04).
- **UAT-08 (this task):** no code change. Added documents and
  `tests/test_uat_evidence_package.py`.

## D. Test execution results

| Command | Result | Source |
|---|---|---|
| `./.venv/bin/python -m pytest --collect-only -q tests` (2026-11-02, before this task's new test file) | 257 collected | this task |
| UAT-06 recorded runs (two commands covering all 257) | 102 + 155 = 257 passed, 0 failed, 0 skipped | `docs/uat/TEST_RESULTS.md` §1 |
| Runner gate (`verify`) per UAT commit | exit 0 on `3f92551`, `66b7e25`, `ddf01f5`, `452e9e0`, `b661f43`, `7b71503` | commit messages, `docs/uat/CHANGE_LOG.md` §0 |
| `./.venv/bin/python -m pytest -q tests/test_uat_evidence_package.py` (this task) | see `docs/uat/README.md` "UAT-08 checks" | this task |

Scenario coverage (master command §7): 56 scenarios executed locally (PASS) and 7 Equus scenarios
NX-01…NX-07 **NOT EXECUTED** (`docs/uat/TEST_RESULTS.md` §3–§4). The runner's gate for this job is
authoritative for the full suite.

## E. Source-to-target reconciliation

Real extract (SHA-256 `006ba3b5…ac18b8`, counts only; `docs/uat/RECONCILIATION_SUMMARY.md` §2):

| Measure | Value |
|---|---|
| Rows | 2 source rows, 2 employees, 2 accepted assignments, 0 held rows |
| Cells | 222 non-blank: written 128, blank_token 4, provisional 7, gated_rule 4, excluded_colour 58, pending_target 21, **unexplained 0** |
| Spouses | 2/2 written |
| Dependents | **4 records held** (D3) |
| Result | Reconciliation OK |

- The synthetic children fixture reconciles with 0 unexplained cells. Addresses (3/3), contacts
  (2/2) and dependents (4/4) are all written (`RECONCILIATION_SUMMARY.md` §3).
- Field-by-field disposition: `docs/uat/FIELD_MAPPING_VALIDATION.md`.
- **Limits:**
  - These are engineering criteria; the business criteria (Q27) are unanswered.
  - Only a 2-row extract exists (Q25).
  - No Equus-side reconciliation was done (NOT EXECUTED).

## F. Critical UAT blockers

| Priority | Blocker | Impact | Missing decision / input | Owner | Evidence |
|---|---|---|---|---|---|
| 1 | D1 identifier conflict | Identity and linkage of every assignment | Which Topia column feeds Assignment Identifier and which feeds CF_TEXT029 | Jake + Alka / John | `JAKE!J2` vs `EX!C4/E4/Y4`; register D1 |
| 2 | SEC-01 real extract in git | Personal data exposure | Authorised history clean-up | Repository owner + Pratyush | SHA-256 match; register §2 |
| 3 | D2 A3 header | Whole-file rejection risk (Q5) | Official template | Jake | `EX!A3` vs template A3; register D2 |
| 4 | Equus lookups Q2, Q3, Q4, Q12, Q14, Q15, Q20 | Values cannot be validated; usernames and currencies missing | Versioned lists | Jake | register E-06, E-08…E-12 |
| 5 | D3 keyed dependents | Real dependents not migrated | Keyed extract | Alka / Topia | register D3 |
| 6 | Q6, Q7, Q9, Q28 confidential and custom fields | Confidential data not migrated; visibility undefined | Template column; SE-424; Schneider decisions | Jake; Schneider | register D5, SEC-02 |
| 7 | Q22 / Q23 / Q37 dates; Q21 compensation; Q34, Q35, Q40 rules | Effective dates, phase, status and international flag blank or unvalidated | Written rules | Jake, Schneider, Alka / John | register E-01…E-05, E-14 |
| 8 | Q25 inventory, Q27 criteria, Q31 signatory | G5, G11 and G12 cannot close | Final extract; criteria; named signatory | Alka / Schneider; Schneider + Jake; Schneider PMO | register E-23…E-25 |
| 9 | Equus UAT access | G6–G8 cannot run | Authorised UAT environment | Schneider / Jake | `docs/uat/EQUUS_UAT_EXECUTION_PLAN.md` §0 |

## G. Gates G1–G12

Full evidence: `docs/uat/UAT_READINESS_ASSESSMENT.md` §1.

| Gate | Result | Gate | Result |
|---|---|---|---|
| G1 | Blocked | G7 | Blocked (NOT EXECUTED) |
| G2 | Blocked | G8 | Blocked (Equus side NOT EXECUTED) |
| G3 | Blocked | G9 | **Fail** (SEC-01; Q6, Q28) |
| G4 | Pass (local) | G10 | **Fail** (open critical items) |
| G5 | Blocked (local reconciliation OK; Q27, Q25, D3) | G11 | Blocked (not obtained) |
| G6 | Blocked (NOT EXECUTED) | G12 | Blocked (not obtained) |

## H. Final readiness recommendation

**NOT READY FOR UAT.**

Supporting evidence:
- G1, G2, G3 and G5 are Blocked; G9 and G10 are Fail (`docs/uat/UAT_READINESS_ASSESSMENT.md`).
- The tool computes `NOT_READY` for every run with the shipped configuration
  (`tests/test_business_rules.py::test_default_readiness_not_ready_and_lists_every_gated_and_pending_item`).
- No Equus import was executed, so the two statuses after execution do not apply.
- Sign-off belongs to the authorised business signatory alone (`docs/uat/BUSINESS_SIGNOFF_RECORD.md`
  is an empty template).

Next step: close blockers 1–5 of section F. Then re-run the gate assessment and, if every pre-import
gate passes, proceed to `docs/uat/EQUUS_UAT_EXECUTION_PLAN.md` step 1.
