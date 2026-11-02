# UAT-04 — Spouse and dependent rules (Q10, Q11, Q18, Q19, Q20) and a keyed dependent source

Task: OCNORA T-1389 (job-773b81d8062c). Sources: `clarifications/UAT_MASTER_COMMAND.md` (BR-10,
BR-11, BR-18, BR-19, BR-20, §13 "never fabricate dependent names"),
`clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md` (§2 rows Q10, Q11, Q18-Q20; D3).
Every tracker Status (column I) is still **Open**: Jake's answers below are evidence, not approval.
Equus UAT execution is **NOT EXECUTED** (no Equus access from this environment); nothing here is
signed off. Readiness of the real extract after this change: **NOT_READY** (see §5).

## 1. Jake's words (read with openpyxl from `clarifications/jake-1.xlsx`)

| Q | Cell | Jake (verbatim) | Classification |
|---|---|---|---|
| Q10 | `datamapping clarrification shef`!J11 | "Spouse/Parnter records can be imported without names." | CONFIRMED BY JAKE |
| Q11 | !J12 | "The latest import template John provided includes dependents. A separate import will not need to be run. Since they will be in the same import, the Assignment Identifier will link the dependents to the same assignment." | PARTIALLY CLARIFIED (D3: no keyed Topia source) |
| Q18 | !J19 | "They are associated with the Assignment. Each spouse/dependet row should be accompanied by an Assignment Identifier. This links different rows to the same assignment." | CONFIRMED BY JAKE |
| Q19 | !J20 | "Please see the example I provided on the 'MNY Dependents EX' sheet/tab I added. As many dependents as are needed can be added as long as the rows contain the Assignment Identifier" | CONFIRMED BY JAKE |
| Q20 | !J21 | "While not recommended, only one dependent field is needed to create a dependent record. Lookup values will be provided as mentioned above." | CONFIRMED BY JAKE (technical) / PENDING EQUUS CONFIGURATION (lookups, Q14) |

Jake's example ` Mny dependents Ex` rows 4-7 (structure only, no values copied):
A4:A7 = INSERT, IGNORE, IGNORE, IGNORE; B4:B7 = INSERT, UPDATE, UPDATE, UPDATE; C (Employee
Number) and D (Assignment Number = NEW) only on row 4; E (Assignment Identifier) identical on all
four rows; spouse EA:EF only on row 4; dependent EG:EP on every row. A3 reads `Employee Command`
there and `A` in the bundled template (D2, unchanged by this task).

## 2. What changed

| File:line | Change | Reason |
|---|---|---|
| `lambda/src/mapping_config.py:183-222` | `FAMILY_NAME_FIELDS`, `FAMILY_NAME_NOTE`, `FAMILY_LIST_HEADERS` (spouse and dependent RELATIONSHIP), `FAMILY_LIST_OWNER` (Q14, Jake), `DEPENDENT_HOLDS` (code → Q, owner, meaning) | Business configuration stays in `mapping_config.py` |
| `lambda/src/mapping_config.py:537-613` | `DependentSource` + strict loader `_load_dependent_source` for the optional `$.dependent_source` section of `business_rules.json` (`_TOP_OPTIONAL`, `:460`); rejects `approved_for_import: true` with `sheet: null`, `data_start_row <= header_row`, non-positive rows, identical or non-normalised key labels, unknown `child_groups`, unknown keys, `sheet` = assignment sheet | Architecture: invalid configuration rejected at load time |
| `lambda/src/mapping_config.py:840` | `DEPENDENT_SOURCE` read-only view | |
| `lambda/src/business_rules.json:6-22` | `dependent_source`: `Q11/Q18`, `partially_clarified`, `approved_for_import: false`, owner Alka / Topia (D3) + Jake, `sheet: null`, `key_columns {employee: "employee id", assignment: "assignment id"}`, `child_groups: repeated_column_groups`; file version 1.0.0 → **1.1.0** (the section inherits the file version, as rules do) | Spec: `DEPENDENT_SOURCE = {sheet, key_columns, child_groups}`; nothing keyed exists yet, so it ships unconfigured |
| `lambda/src/transform.py:483` `validate_family` | Replaces "Equus may reject or create an unnamed person" with `DATA_QUALITY: record has no first/last name; importable per Q10/Q20, business review recommended (has …; missing …)`; records the note in `Report.data_quality`; counts relationship values. Never holds, drops or invents a name | Q10, Q20; CR-T14 |
| `lambda/src/transform.py:179` `Report.family_value_status`, `:205` `summarise_values` | Relationship values checked against the bundled Equus lists (`mapping_config.ALLOWED_VALUES_FILE`); without a list every value is `not in list (no Equus list; Q14 pending)`; values are sent unchanged | Q14 / Q20 lookups pending |
| `lambda/src/transform.py:603` `read_dependent_sources` | Scans every non-assignment sheet with `ASSIGNMENT_DEPENDENT*` keys: no Employee ID + Assignment ID columns → `HELD_DEPENDENTS_UNKEYED`; keyed but not the configured sheet → `HELD_DEPENDENTS_NOT_CONFIGURED`. Configured sheet: missing → `HELD_DEPENDENTS_SHEET_MISSING`; no key columns → `HELD_DEPENDENTS_UNKEYED`; not approved → `HELD_DEPENDENTS_GATED`; rows without Assignment ID → `HELD_DEPENDENTS_NO_ASSIGNMENT_ID`. Otherwise indexes child groups by Assignment ID | D3; "unkeyed dependent data is held, never guessed" |
| `lambda/src/transform.py:678` `build_records`, `:803` `_attach_keyed_dependents` | Joins keyed dependents on the row's **own** Assignment ID (computed before the Employee-Number fallback, so the fallback can never link dependents). Employee ID must equal the assignment's Employee Number (else `HELD_DEPENDENTS_EMPLOYEE_MISMATCH`); an Assignment ID on several assignment rows → `HELD_DEPENDENTS_AMBIGUOUS`; keys matching no assignment → `HELD_DEPENDENTS_UNMATCHED`; dependents from both sources on one assignment → warning + readiness reason `dependent_overlap` | Q18 cross-assignment isolation; Q1 no fallback |
| `lambda/src/transform.py:1011-1027` `compute_readiness` | New reason categories `held_dependents` (one per exception, counts only), `dependent_overlap`, `family_value_not_in_list` (Q14, Jake) | Every held item visible with Q and owner |
| `lambda/src/transform.py:1142` `_family_sheet` | New report sheet **Family**: dependent source status, held-dependent table, DATA_QUALITY notes, relationship values with list check | Evidence |
| `OPEN_QUESTIONS.md` §10 | Spouse/dependent text updated to Q10/Q11/Q18-Q20 answers | |
| `tests/test_family_rules.py` (new), `tests/family_fixtures.py` (new), `tests/fixtures/family_dependents_unkeyed.xlsx`, `tests/fixtures/family_dependents_keyed.xlsx` (new, synthetic) | See §4 | |

Unchanged on purpose: the INSERT/UPDATE row layout (`transform.py:776` `n_rows = max(...)`, unbounded)
already matched Jake's example; this task adds the proof. The Assignment Identifier fallback to the
Employee Number (`build_records`, D1 / BR-01) is untouched here (separate task); keyed dependents
never use it.

## 3. Behaviour

| Situation | Output | Report |
|---|---|---|
| Spouse with relationship, no names (fixture GW5/GW6) | Written on the INSERT row as sent | DATA_QUALITY note; not a hold; no readiness reason for names |
| Dependent with only a birth date | Dependent record written | DATA_QUALITY note |
| Dependent group with no value | No record, no extra row | — |
| N dependents on one assignment | INSERT row + N-1 UPDATE rows (`A` = IGNORE, blank Employee Number, blank Assignment Number, same Assignment Identifier); spouse only on the INSERT row | — |
| Two assignments of one employee | Each assignment's family rows carry its own Assignment Identifier; second assignment INSERT row has `A` = IGNORE | — |
| Topia `Dependent Import` (no key) | Nothing attached | `HELD_DEPENDENTS_UNKEYED` (Q11, Alka / Topia, D3) with row and record counts; NOT_READY |
| Keyed sheet configured + approved | Attached by Assignment ID | Family sheet: count attached; holds for bad keys |
| Relationship value with no Equus list | Sent unchanged | `not in list (no Equus list; Q14 pending)`; readiness `family_value_not_in_list` |

## 4. Test evidence (`tests/test_family_rules.py`, synthetic fixtures only)

| AC | Tests |
|---|---|
| AC-1 | `test_spouse_without_names_is_written_with_data_quality_note_not_held`, `test_dependent_with_only_birth_date_is_written_and_empty_group_creates_no_record`, `test_relationship_values_checked_against_equus_list_never_converted` |
| AC-2 | `test_jake_example_has_the_structure_this_test_encodes` (reads Jake's example structure), `test_four_sheet_dependents_match_jake_example_structure`, `test_four_keyed_dependents_match_jake_example_structure`, `test_any_number_of_dependents_gives_insert_plus_updates[1,2,3,7]` |
| AC-3 | `test_two_assignments_of_one_employee_never_share_sheet_dependents`, `test_keyed_dependents_stay_with_their_own_assignment`, `test_keyed_rows_with_missing_conflicting_or_unknown_keys_are_held`, `test_keyed_dependents_never_join_through_identifier_fallback`, `test_duplicate_assignment_id_holds_keyed_dependents` |
| AC-4 | `test_unkeyed_dependent_import_sheet_is_held_and_nothing_attached`, `test_unkeyed_sheet_stays_held_even_with_a_keyed_source_configured`, `test_keyed_sheet_joins_by_assignment_id_when_configured_and_approved`, `test_keyed_sheet_not_configured_or_not_approved_is_held`, `test_configured_keyed_sheet_without_key_columns_is_held_unkeyed` |
| Config | `test_shipped_dependent_source_is_unconfigured_and_unapproved`, `test_invalid_dependent_source_rejected_at_load[...]` (9 cases) |

Run: `./.venv/bin/python -m pytest -q tests/test_family_rules.py` → 30 passed;
`tests/test_migration.py tests/test_business_rules.py` → 69 passed (no existing assertion changed;
`test_validation_warnings_keep_values` still finds "no first/last name" in the new note).
The full suite is run by the runner's gate.

## 5. Real-extract run (counts only; no values recorded)

Source SHA-256 `006ba3b50df6cd60702692ff28da3b6a57f063c74a67f716859e56fd96ac18b8`, default
(shipped) configuration: 2 assignments, 2 rows written, reconciliation OK, readiness NOT_READY.
Family items: `HELD_DEPENDENTS_UNKEYED` on sheet `Dependent Import`, 2 dependent rows (rows 3, 4),
4 dependent records, 0 attached; 2 spouse DATA_QUALITY notes (populated: RELATIONSHIP[, STATUS];
missing FIRST_NAME, LAST_NAME); `family_value_not_in_list` for `SPOUSE_DOMESTIC_PARTNER!RELATIONSHIP`
(2 distinct values, no Equus list). Dependents from the real extract are therefore **still not
migrated** — correctly held until Topia supplies a keyed extract.

## 6. Assumptions and open points

1. **Keyed layout (assumption):** a keyed sheet has `ASSIGNMENT_DEPENDENT*` technical keys in
   `header_row` and the key-column labels somewhere in rows 1..`data_start_row`-1; child groups
   repeat (`repeated_column_groups`), which also covers one child per row. The real keyed layout
   is unknown until Alka / Topia deliver it (D3); adjust `$.dependent_source.parameters` then.
2. **Gating (assumption):** linking by Assignment Identifier is CONFIRMED BY JAKE but not approved
   (tracker Open), so the keyed join ships `approved_for_import: false` and `sheet: null`; a
   keyed sheet present while unapproved is `HELD_DEPENDENTS_GATED`.
3. **Employee ID cross-check:** a keyed row whose Employee ID differs from the assignment's
   Employee Number is held; a blank Employee ID is attached by Assignment ID with a warning
   (the Assignment ID is the join key per Q18). Leading zeros are kept as text, never normalised.
4. **Equus lists location:** the task text says `reference_data/equus_lists.json`; the existing
   loader reads `lambda/src/equus_lists.json` (`mapping_config.ALLOWED_VALUES_FILE`) and
   `reference_data/` accepts only the three declared lists, so the existing location was kept.
   No list exists yet (Q14), so every relationship value is "not in list".
5. Not done (non-goals): no relationship/gender lookups invented; no family-size calculation (Q37).
6. Open: Q14 lists (Jake); keyed dependent extract (Alka / Topia, D3); D1 identifier conflict and
   D2 A3 header remain blockers outside this task. Status: **NOT READY FOR UAT**.
