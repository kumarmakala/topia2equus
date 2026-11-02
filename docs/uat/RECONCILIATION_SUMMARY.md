# UAT-05 — Source-to-target reconciliation summary

Prepared 2026-11-02 for OCNORA task T-1390 (UAT-05) on base commit `ddf01f5`.
Master command section 8 (`clarifications/UAT_MASTER_COMMAND.md`) and evidence
assessment section 3, row Q27 (`clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md`:
"Count equality on INSERT rows only") are the requirement. Q27 (reconciliation
criteria, owner Schneider + Jake) is still **OPEN**: the criteria below are the
engineering criteria of this task, not approved business acceptance criteria.

This document contains **counts, column letters, Equus headers, exception codes and
hashes only**. No source value, name, identifier or e-mail is reproduced.

**Equus UAT execution: NOT EXECUTED / INCOMPLETE** — no Equus access from this
environment. A local reconciliation that says OK is gate G5 evidence for the local
transformation only; it is not an Equus import result (master command sections 9–10).
Readiness computed by the tool for both runs below: **NOT_READY** (never SIGNED OFF).

## 1. What the reconciliation now does

Code: `lambda/src/transform.py` `reconcile()` (line 1256) with helpers
`_output_blocks()` (1227), `RECON_CATEGORIES` (1163), `MIGRATED_STATUSES` (1166);
report sheets "Reconciliation" (1769) and "Reconciliation Records" (1789), column
layouts `RECONCILIATION_COLUMNS` (1814) / `RECORDS_COLUMNS` (1825); readiness reason
detail (1728). Keyed dependents attached from a configured keyed source are recorded
per assignment in `report.keyed_children` (1112). Configuration (attribution only,
no rule): `lambda/src/mapping_config.py` `READINESS_ITEMS` entries `excluded_colour`,
`not_migrated`, `blank_token` (348–350) and `NOT_MIGRATED_QUESTIONS` (360).

Before (base `ddf01f5`, `transform.py:1151-1206`): mapped columns only, accepted source
rows zipped **by position** with INSERT rows; UPDATE rows (addresses, contacts,
dependents) never compared.

After:
- **Keyed on the Assignment Identifier**: output rows are grouped into blocks (one INSERT
  row plus its UPDATE rows with the same identifier); each accepted source row is matched
  to its block through `report.accepted` / `row_identity`. A missing block, a second INSERT
  of the same identifier, an INSERT with an identifier of no accepted row, or a row not
  linked to any block is unexplained.
- **Every non-blank source cell of every column** falls in exactly one category:

| Category | Meaning | Q / owner reported |
|---|---|---|
| written | written unchanged, or by a deterministic / approved conversion (date parse, number format, approved rule or lookup) | — |
| blank_token | placeholder (`BLANK_TOKENS`, `BLANK_PREFIX_TOKENS`) left blank | Q43 / Alka |
| provisional_conversion | country name converted by `country_name_to_iso3`, rule not approved | Q12/Q3 / rule owner |
| gated_rule | column held by an unapproved rule (Host Job Level, Cancellation Reason) | rule Q / owner |
| held_row | cell of a source row held by an identity/duplicate control (UAT-03) | exception code Q / owner |
| excluded_colour | red column (not required) | Q30 / Schneider |
| pending_target | column with data and no approved Equus target (cyan MISSING_TARGET, yellow pending, FIELD_OVERRIDES None, unmapped) | Q6/Q7/Q29, Q9, Q34, Q41, yellow list or Q29/Q30 |
| failed_conversion | unparseable date left blank (warned) | Q24 / Jake |
| replaced_by_rule | source value overwritten by an **approved** rule | — |
| unexplained | anything else — fails the run | — |

  `set_by_rule` (outside the partition) counts output cells an approved rule filled where
  the source was blank; a populated output cell with a blank source and no approved rule
  is unexplained.
- **Child rows**: the k-th populated address / contact / dependent group of an assignment
  is compared with the k-th row of its block (INSERT = 0, UPDATE = 1..n), field by field
  (address `LOCATION_TYPE`, `LINE_1..3`; contact `LOCATION_TYPE`, `CONTACT_TYPE`,
  `NUMBER_ADDRESS`; every dependent column). Spouse fields are mapped columns on the
  INSERT row and are counted as one child record per assignment.
- **Per record** (sheet "Reconciliation Records"): source row, identifier, employee,
  accepted/held + codes, output rows (INSERT + UPDATE), expected expansion rows, and
  address / contact / spouse / dependent records expected vs written.
- **Summary**: source rows, employees, accepted assignments, held rows, held dependent
  records, INSERT / UPDATE rows, expected expansion rows (Σ over records of
  `max(1, addresses, contacts, dependents) − 1`), child records expected vs written,
  rows written, cells per category, unexplained total, `counts_balance`.
- **`reconciliation_ok`** only when `unexplained_total == 0` **and** counts balance:
  source rows = accepted + held; INSERT rows = accepted = assignments; UPDATE rows =
  expected expansion rows; rows written = output rows found; no unexpected / duplicate /
  orphan blocks; every column's categories add up to its non-blank cells; child records
  expected = written. Handler behaviour is unchanged: a FAILED reconciliation withholds the
  import file (`handler.py` `RECONCILIATION_FAILED` report).

CLI: `local/reconcile.py` runs `transform()` on a workbook, writes the full report
workbook and a value-free JSON evidence file to `local_s3/dest_s3/equus/reconciliation/`
(git-ignored, never committed), prints the summary, exits 0 = OK / 1 = FAILED.

## 2. Real extract — counts only

Command: `./.venv/bin/python local/reconcile.py` (default workbook
`local_s3/source_s3/topia/incoming/FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx`,
the runner's copy of the main-checkout extract).
Source SHA-256 `006ba3b50df6cd60702692ff28da3b6a57f063c74a67f716859e56fd96ac18b8`;
business rules version 1.2.0 (shipped, nothing approved for import); template
`lambda/src/templates/equus_eafh_template.xlsx`.

**Result: reconciliation OK** — `unexplained_total` 0, `counts_balance` true.
Readiness: NOT_READY (gated rules, provisional conversions, pending/missing-target
columns, D2 header decision pending, unkeyed dependents held, no Equus lists).

| Measure | Count |
|---|---|
| Source data rows | 2 |
| Distinct employees (source / accepted) | 2 / 2 |
| Distinct Assignment Identifiers (source) | 2 |
| Accepted assignments (INSERT rows) | 2 |
| Held source rows | 0 (no identity/duplicate exception) |
| Addresses / contacts / dependents expected (main sheet) | 0 / 0 / 0 |
| Spouse records expected / written | 2 / 2 |
| Expected expansion rows / UPDATE rows written | 0 / 0 |
| Rows written | 2 |
| Unexpected output rows / duplicate INSERT blocks | 0 / 0 |
| Held dependent records (sheet "Dependent Import", `HELD_DEPENDENTS_UNKEYED`, Q11, owner Alka / Topia, D3) | 4 records in 2 rows |
| Non-blank source cells | 222 |
| written | 128 |
| blank_token | 4 |
| provisional_conversion | 7 |
| gated_rule | 4 |
| held_row | 0 |
| excluded_colour | 58 |
| pending_target | 21 |
| failed_conversion / replaced_by_rule / set_by_rule | 0 / 0 / 0 |
| **unexplained** | **0** |

Exception codes of the run: `HEADER_DECISION_PENDING` 1 (D2, column A3), `VALIDATION_WARNING` 20.

Columns not (fully) written, by category (column letter → Equus header, cells):
- blank_token: **M** → `EMPLOYEE_CF!CF_TEXT005` 2; **N** → `EMPLOYEE_CF!CF_TEXT013` 2
  (N6 is the glued placeholder covered by `BLANK_PREFIX_TOKENS`, `mapping_config.py:162`).
- provisional_conversion (Q12/Q3, `country_name_to_iso3` not approved): **O**
  `EMPLOYEE!PRIMARY_CITIZENSHIP` 2, **P** `EMPLOYEE!SECONDARY_CITIZENSHIP` 1,
  **AY** `ASSIGNMENT!FROM_COUNTRY` 2, **BD** `ASSIGNMENT!TO_COUNTRY` 2.
- gated_rule: **CQ** Host Job Level 2 (Q32); **DH** Cancellation Reason 2 (Q34).
- pending_target: **AG** 2 and **DG** 2 (cyan, no Equus column; Q6/Q7/Q29), **ED** 2
  (FIELD_OVERRIDES None; Q9, Schneider), **AB** 2 (Q34), **EJ** 2, **EL** 1, **EN** 2,
  **EP** 1, **ER** 1, **FE** 2 (resource e-mails; Q41), **FK**, **FM**, **FO**, **FQ** 1
  each ("Location Type - …" descriptors of the constant `LOCATION_TYPE`; Q29/Q30).
- excluded_colour (red, Q30): R, W, AD, AE, AF, AH, AI, AJ, AM, AV, AW, AZ, BA, BE, BF,
  CG, CH, CI, CK, CL, CW, CY, CZ, DA, DI, DJ, DL, DM (2 each), DV, EV (1 each) = 58 cells.

Written but still pending: 7 yellow columns (S, AA, AC, AK, AL, AO, AP; 14 cells) are
written to the draft workbook and keep readiness NOT_READY (`pending_column`).

Known data facts reproduced: 2 source rows, 2 employees, 2 assignments; addresses,
contacts and dependents empty in the main sheet; AG, DG, ED populated and pending
(not written, explained); N blank-token. Test: `tests/test_reconciliation.py::test_real_extract_known_facts`
(skipped where the git-ignored extract is absent).

## 3. Synthetic fixture — full reconciliation

Fixture `tests/fixtures/reconciliation_children.xlsx`
(SHA-256 `8593e7865207535f6592dd05e244794053508fee968b252806b48309052b66bd`), built by
`tests/reconciliation_fixtures.py` from `tests/fixtures/topia_sample.xlsx`: row 5 with two
addresses (FJ four lines, FN), two contacts (FR, FS) and three dependents (the third has only
a birth date, Q20); row 6 with one address (FL), placeholder contact (FS `N/A`), a dependent
group holding only `Not Provided` and one dependent with a first name only.
Command: `./.venv/bin/python local/reconcile.py tests/fixtures/reconciliation_children.xlsx`.

**Result: reconciliation OK**, `unexplained_total` 0, `counts_balance` true; readiness NOT_READY.

| Measure | Count |
|---|---|
| Source rows / employees / accepted assignments / held rows | 2 / 2 / 2 / 0 |
| Address records expected / written | 3 / 3 |
| Contact records expected / written | 2 / 2 |
| Dependent records expected / written | 4 / 4 |
| Spouse records expected / written | 2 / 2 |
| INSERT rows / UPDATE rows | 2 / 2 |
| Expected expansion rows | 2 (row 5: 3 output rows; row 6: 1) |
| Rows written | 4 |
| Non-blank source cells | 239 |
| written / blank_token / provisional / gated / held_row / excluded / pending_target | 143 / 6 / 7 / 4 / 0 / 58 / 21 |
| unexplained | 0 |

Child-row blank tokens: FS (contact, row 6) 1, GY (`ASSIGNMENT_DEPENDENT!FIRST_NAME`,
Child 1 of row 6) 1 — explained, no record created.

## 4. Test evidence (`tests/test_reconciliation.py`, 22 tests, all passing)

| Criterion | Test id(s) |
|---|---|
| AC-1 fixture with addresses, contacts, three dependents: zero unexplained, expansion rows = UPDATE rows | `test_children_fixture_reconciles_with_zero_unexplained`, `test_children_fixture_per_record_child_counts`, `test_children_fixture_child_cells_reconciled_field_by_field` |
| Keyed on identifier, not position | `test_reconciliation_is_keyed_on_identifier_not_row_position` |
| AC-2 dependent first name dropped on an UPDATE row is unexplained | `test_dropped_dependent_first_name_on_update_row_is_unexplained` (INSERT variant: `test_dropped_insert_first_name_still_caught_on_fixture`; existing `tests/test_migration.py::test_reconciliation_catches_silent_loss`) |
| Child-row and record-level losses / duplicates | `test_dropped_address_line_and_changed_contact_are_unexplained`, `test_lost_update_row_is_caught_at_record_level`, `test_duplicated_and_unknown_output_rows_are_caught` |
| AC-3 held rows and pending targets (AG, DG, ED) are explained | `test_held_rows_are_an_explained_category[duplicate_row]`, `[missing_identifier]`, `test_pending_target_columns_are_explained[...]` |
| Other categories | `test_gated_excluded_blank_and_provisional_categories`, `test_approved_rule_replacing_a_source_value_is_explained` |
| Report / CLI, no values in evidence | `test_report_workbook_has_field_and_record_sheets`, `test_cli_evidence_has_no_source_values`, `test_cli_main_writes_report_and_returns_status` |
| AC-4 real extract counts | `test_real_extract_known_facts` + section 2 |

Existing suites re-run after the change: `tests/test_migration.py`,
`tests/test_identity_controls.py`, `tests/test_family_rules.py`,
`tests/test_business_rules.py` — 138 passed; no assertion changed.

## 5. Assumptions and limits (recorded instead of asking)

1. "pending_target" covers every column that holds data but has no approved Equus target
   (cyan, yellow, FIELD_OVERRIDES None, uncoloured unmapped, "Location Type" descriptors).
   They are *explained* for reconciliation but remain open decisions with the Q / owner
   shown; reconciliation OK does not mean these columns are dispositioned (G1 stays Blocked).
2. Red columns are explained as excluded by colour per Alka's legend; the scope decision
   itself is Q30 (Schneider), still open.
3. Placeholder cells in red / held / pending columns are counted under the column's
   disposition, not as blank_token; blank_token is used for migrated columns only.
4. Address and contact cells are compared to what the engine writes today (free-text lines,
   no placeholder stripping beyond `BLANK_TOKENS`); no new address rule was introduced.
5. Keyed dependents (shipped `dependent_source` is unconfigured and unapproved) are
   reconciled against the converted record attached to the assignment, not the raw keyed
   sheet cell; held keyed / unkeyed dependents are reported as held dependent records.
6. Output cells that no source column maps (constants such as `PRIMARY_ADDRESS`, commands,
   child dates set by an approved rule) are not partitioned; a mapped target populated with
   a blank source is unexplained unless an approved rule fills it.
7. Q27 (reconciliation acceptance criteria) is unanswered; Schneider + Jake must approve
   these criteria before G5 can be signed off by the business.

## 6. Gate impact

G5 (reconciliation passes): **local evidence available** — the real extract and the
synthetic fixture reconcile with zero unexplained cells and balanced counts; business
acceptance of the criteria (Q27) is pending and the Equus-side reconciliation (G6–G8) is
NOT EXECUTED. Overall status remains **NOT READY FOR UAT** (G1–G3, G9, G10 open per the
evidence assessment section 5).
