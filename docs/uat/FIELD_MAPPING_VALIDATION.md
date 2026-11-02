# Field Mapping Validation: every Topia source column → Equus header

Evidence package item 6 of `clarifications/UAT_MASTER_COMMAND.md` §12 (line 190). OCNORA task
T-1393 (UAT-08). Prepared 2026-11-02 at commit `7b71503` (code unchanged by this task).

**Status of this deliverable: COMPLETE as local evidence; NOT APPROVED.** No mapping below is
formally approved: every tracker Status cell `datamapping clarrification shef`!I2:I44 reads
`Open` (read with openpyxl on 2026-11-02). G1 (mappings approved or dispositioned) stays
**Blocked**. Nothing here has been observed in Equus; whether Equus accepts any value is
**NOT EXECUTED** (see `docs/uat/EQUUS_UAT_EXECUTION_PLAN.md`).

## 1. How this table was produced

- Run: `transform.transform(source, template)` with the **shipped** configuration
  (`lambda/src/business_rules.json` version 1.2.0, nothing `approved_for_import`; empty
  `lambda/src/reference_data/*.json`; no `equus_lists.json`), executed with `./.venv/bin/python`.
- Source: the synthetic fixture `tests/fixtures/topia_sample.xlsx`
  (SHA-256 `4e2084405bf253eca8f0126094d057f9f2acf88d4c8545335cbd5e30b89dd68f`). Its header rows 1–4
  (row 2 = Equus technical key, row 3 = Equus UI label, row 4 = Topia label, colour fills) are the
  layout of the real extract; its data values are synthetic. The real extract was **not** used for
  this table, and no value from any data row is reproduced here (only labels, letters and counts).
- Template: `lambda/src/templates/equus_eafh_template.xlsx`
  (SHA-256 `c90e9fc057e9fbcc5b1b38e8b4821dd266b92bd7557915f48f7925fc3f445be7`, `TEMPLATE_VERSION`
  `EAFH-bundled-1`), 146 headers in row 3.
- Columns come from the mapping report's **Column Mapping** sheet (`lambda/src/transform.py:1866-1870`,
  plan built by `build_column_plan`, `transform.py:356-456`). The "Fixture run" column comes from the
  **Reconciliation** sheet (`reconcile`, `transform.py:1360`; categories `RECON_CATEGORIES`,
  `transform.py:1267`). Conversion rules are those of `_convert` (`transform.py:491-555`),
  `mapping_config.FIELD_OVERRIDES` (`lambda/src/mapping_config.py:89`), `ADDRESS_COLUMNS` (`:126`),
  `CONTACT_COLUMNS` (`:136`), `COUNTRY_HEADERS` (`:144`), `STATE_COUNTRY_HEADERS` (`:157`),
  `BLANK_TOKENS` / `BLANK_PREFIX_TOKENS` (`:179`, `:183`) and `NUMERIC_RANGES` (`:191`).
- The same counts, taken from the real extract with `local/reconcile.py`, are in
  `docs/uat/RECONCILIATION_SUMMARY.md` §2. Their category totals are identical to the fixture run
  below (222 non-blank cells; 0 unexplained).
- The Q / decision column cites Jake's column J, using the cell rule `JAKE!J(n+1)` = Qn. `EX!` is
  the sheet ` Mny dependents Ex`.

## 2. Summary of the fixture run

| Measure | Value |
|---|---|
| Source columns in the plan (B…GX) | 205 |
| Plan status | mapped 106 · excluded (red) 74 · unmapped 12 · ignored 5 · address 4 · held 2 · contact 2 |
| Colour (rows 1–4) | peach 93 · red 74 · cyan 20 · yellow 12 · brown = `none (Q8: no meaning)` 6 |
| Non-blank source cells | 222 |
| written | 128 |
| blank_token (Q43) | 4 |
| provisional_conversion (Q12/Q3) | 7 |
| gated_rule (Q32, Q34) | 4 |
| held_row | 0 |
| excluded_colour (Q30) | 58 |
| pending_target (Q6/Q7/Q29, Q9, Q34, Q41, Q29/Q30) | 21 |
| failed_conversion / unexplained | 0 / 0 |
| Reconciliation result | OK (`reconciliation_ok = True`) |
| Readiness computed by the tool | `NOT_READY`, 54 reasons |

Test evidence for this run:
- `tests/test_reconciliation.py::test_gated_excluded_blank_and_provisional_categories`
- `tests/test_reconciliation.py::test_pending_target_columns_are_explained`
- `tests/test_business_rules.py::test_default_output_has_no_unapproved_rule_values`
- `tests/test_business_rules.py::test_default_readiness_not_ready_and_lists_every_gated_and_pending_item`
- `tests/test_uat_evidence_package.py::test_field_mapping_validation_matches_fixture_run`. This
  test re-runs the fixture and checks that every source column letter, plan status and Equus header
  in the table below still matches the code.

## 3. Pending, held and not-migrated columns (call-outs)

| Group | Columns | What happens | Q / owner |
|---|---|---|---|
| **Identifier conflict** | X → `ASSIGNMENT_CF!CF_TEXT029` **and** `Assignment Identifier` | Written as text with leading zeros kept. A missing value holds the row (no fallback) | Q1, **D1 OPEN**: Jake + Alka / John |
| Gated rule, column HELD | CQ Host Job Level; DH Cancellation Reason | Not migrated; reconciliation `gated_rule` | Q32 (Jake / Alka); Q34 (Jake + Schneider) |
| Missing Equus target (D5) | AG Confidentiality Type; DG Expected Assignment Outcome | Not migrated; readiness `missing_target`; reconciliation `pending_target` | Q6 (Jake, Equus template) + Q28; Q7 (Schneider, SE-424) |
| Excluded pending decision (D5) | ED PII Nationality 1 | Not migrated; reconciliation `pending_target` with Q9 | Q9: Schneider |
| Yellow = pending | S, AA, AB, AC, AK, AL, AO, AP, GO, GP, GQ, GR | Written to the **DRAFT** workbook only (except AB, which has no target). Readiness `pending_column` | Q35, Q34, Q36, Q23, Q38, Q37, Q39; Alka / John (mapping review) |
| Resource e-mails | EJ, EL, EN, EP, ER, FE | Not migrated (`pending_target`) | Q41: Jake / Schneider |
| "Location Type - …" descriptors | FK, FM, FO, FQ | Not migrated (they describe a constant) | Q29/Q30 |
| Target not in template | FX, FY (host pay) | Not migrated | Q21: Jake + Schneider |
| Unmapped note column | DN | Not migrated; no data in fixture | Q30 disposition (Schneider) |
| Provisional conversion | O, P, Q, AY, BD (country → ISO-3) | Converted, but every value is listed as provisional | Q12/Q3: Jake (Equus country list) |
| Lookup list empty | Resource names (EI, EK, EM, EO, EQ, FD, FF–FI); cost centres (BI, BK, BM, BO); states (BC, BH) | Sent unchanged; readiness `reference_lookup_missing` / `reference_list_unavailable` | Q2, Q15, Q12: Jake (Equus) / Alka |
| Dropdown without Equus list | every row-1 "dropdown" target, incl. GW spouse relationship | Sent unchanged; readiness `dropdown_list_missing`, `family_value_not_in_list` | Q14: Jake |
| Red = not required | 74 columns (`excluded`) | Not migrated | Q30: Schneider |
| Dependents | sheet "Dependent Import" of the real extract (no key) | Not migrated; `HELD_DEPENDENTS_UNKEYED` (4 records in 2 rows, real extract) | Q11, **D3 OPEN**: Alka / Topia |

## 4. Column-by-column validation (205 source columns)

| Src col | Topia label (row 4) | Colour | Plan status | Equus header (template col) | Conversion rule | Rule status / owner | Q / decision | Fixture run: non-blank cells -> category |
|---|---|---|---|---|---|---|---|---|
| B | First Name | peach | mapped | `EMPLOYEE!FIRST_NAME` (F) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| C | Last Name | peach | mapped | `EMPLOYEE!LAST_NAME` (G) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| D | Middle Name | peach | mapped | `EMPLOYEE!MIDDLE_NAME` (H) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 1: written 1 |
| E | Employee ID | peach | mapped | `Employee Number` (C) | as read; placeholders -> blank (Q43) | key alias (alias of EMPLOYEE!SYSTEM_NUMBER; CR-M07, Q29); not formally approved (tracker I = Open) | Q1 CONFIRMED BY JAKE (J2: Talent Link ID = Employee Number) | 2: written 2 |
| F | Topia ID | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| G | Email | peach | mapped | `EMPLOYEE!USER_ID_EMAIL` (I) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q13 CONFIRMED BY JAKE (J14: primary e-mail) | 2: written 2 |
| H | Alternate Email | peach | mapped | `EMPLOYEE_CF!CF_TEXT008` (J) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q13 CONFIRMED BY JAKE (J14: secondary e-mail) | 2: written 2 |
| I | Gender | peach | mapped | `EMPLOYEE!GENDER` (K) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| J | Date of Birth | peach | mapped | `EMPLOYEE!BIRTH_DATE` (L) | date parse, configured order M/D/Y; ambiguous text warned (Q24); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| K | Marital Status | peach | mapped | `EMPLOYEE!MARITAL_STATUS` (M) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| L | Nationality 1 | peach | mapped | `EMPLOYEE_CF!CF_TEXT002` (N) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| M | Nationality 2 | peach | mapped | `EMPLOYEE_CF!CF_TEXT005` (O) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: blank_token 2 |
| N | Nationality 3 | cyan | mapped | `EMPLOYEE_CF!CF_TEXT013` (P) | number -> text; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q3/Q29 (custom field from template label) | 2: blank_token 2 |
| O | Citizenship 1 | peach | mapped | `EMPLOYEE!PRIMARY_CITIZENSHIP` (R) | country name -> ISO-3, **provisional** (`country_name_to_iso3`); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q3/Q12 | 2: provisional 2 |
| P | Citizenship 2 | peach | mapped | `EMPLOYEE!SECONDARY_CITIZENSHIP` (S) | country name -> ISO-3, **provisional** (`country_name_to_iso3`); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q3/Q12 | 1: provisional 1 |
| Q | Country of Birth | peach | mapped | `EMPLOYEE!BIRTH_COUNTRY` (T) | country name -> ISO-3, **provisional** (`country_name_to_iso3`); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q3/Q12 (template Q-T label shift warned) | no data |
| R | Person Type | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| S | Person Status | yellow | mapped | `EMPLOYEE!TYPE` (U) | value map gated OFF (Q35); source value sent; Equus list check: no list (Q14); placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; row-2 key = template header; not formally approved (tracker I = Open) | Q35 | 2: written 2 |
| T | Business Title | peach | mapped | `ASSIGNMENT_CF!CF_TEXT014` (V) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| U | Employment Status | peach | mapped | `EMPLOYEE!EMPLOYMENT_STATUS` (W) | value map gated OFF (Q35); source value sent; Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q35 | 1: written 1 |
| V | Work Location | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| W | Preferred Pronouns | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| X | Assignment ID | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT029` (Y) + `Assignment Identifier` (E) | identifier_text (text, leading zeros kept); same value -> `Assignment Identifier`; number -> text; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q1 PARTIALLY CLARIFIED; **D1 OPEN** (J2 vs EX!C4/E4/Y4); also feeds Assignment Identifier | 2: written 2 |
| Y | Assignment Ref No | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| Z | Assignment Type | peach | mapped | `ASSIGNMENT!TYPE` (Z) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| AA | Assignment Type Category | yellow | mapped | `ASSIGNMENT!RELOCATION_POLICY_ID` (AA) | Equus list check: no list (Q14); placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; row-2 key = template header; not formally approved (tracker I = Open) | Q35 | 2: written 2 |
| AB | Assignment Status | yellow | unmapped | - | none (not migrated) | pending_target: yellow, row 2 is a note; Q34, Jake + Schneider |  | 2: pending_target 2 |
| AC | Status Category | yellow | mapped | `ASSIGNMENT!PHASE` (BV) | Equus list check: no list (Q14); placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q34 | 2: written 2 |
| AD | Assignment Number | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AE | Created Date | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AF | Status Last Updated | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AG | Confidentiality Type | cyan | unmapped | - | none (not migrated) | HELD (missing_target): cyan, no Equus column; Q6 PENDING EQUUS CONFIGURATION, Jake; Q28 (D5) |  | 2: pending_target 2 |
| AH | Is Confidential | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AI | Tax Equalized | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AJ | Exempt Accruals | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AK | Home Pay Installments | yellow | mapped | `ASSIGNMENT_CF!CF_NUMBER001` (AB) | as read; placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; row-2 key = template header; not formally approved (tracker I = Open) | Q36 | 2: written 2 |
| AL | Host Pay Installments | yellow | mapped | `ASSIGNMENT_CF!CF_NUMBER002` (AC) | as read; placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; row-2 key = template header; not formally approved (tracker I = Open) | Q36 | 2: written 2 |
| AM | Home Leaves Per Year | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AN | Do Not Contact | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| AO | Assignment Start Date | yellow | mapped | `ASSIGNMENT!START_DATE` (AD) | date parse, configured order M/D/Y; ambiguous text warned (Q24); placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; row-2 key = template header; not formally approved (tracker I = Open) | Q23 | 2: written 2 |
| AP | Assignment End Date | yellow | mapped | `ASSIGNMENT!ACTUAL_END_DATE` (AE) | date parse, configured order M/D/Y; ambiguous text warned (Q24); placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; row-2 key = template header; not formally approved (tracker I = Open) | Q23 | 2: written 2 |
| AQ | Departure Date | peach | mapped | `ASSIGNMENT!HOME_COUNTRY_DEPARTURE_DATE` (AF) | date parse, configured order M/D/Y; ambiguous text warned (Q24); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | no data |
| AR | Arrival Date | peach | mapped | `ASSIGNMENT!HOST_COUNTRY_ARRIVAL_DATE` (AG) | date parse, configured order M/D/Y; ambiguous text warned (Q24); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| AS | Move In Temp Date | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| AT | Move Out Date | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| AU | Move Out Temp Date | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| AV | Payroll Start Date | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AW | Payroll End Date | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| AX | Move In Date | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| AY | Home Country | peach | mapped | `ASSIGNMENT!FROM_COUNTRY` (AH) | country name -> ISO-3, **provisional** (`country_name_to_iso3`); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q12 | 2: provisional 2 |
| AZ | Home Region | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| BA | Home Continent | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| BB | Home City | peach | mapped | `ASSIGNMENT!FROM_CITY` (AI) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q12 (city free text, J13) | 2: written 2 |
| BC | Home State | peach | mapped | `ASSIGNMENT!FROM_STATE_PROVINCE` (AJ) | state checked against `country_state` (list empty); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q12 (state list) | 2: written 2 |
| BD | Host Country | peach | mapped | `ASSIGNMENT!TO_COUNTRY` (AL) | country name -> ISO-3, **provisional** (`country_name_to_iso3`); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q12 | 2: provisional 2 |
| BE | Host Region | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| BF | Host Continent | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| BG | Host City | peach | mapped | `ASSIGNMENT!TO_CITY` (AM) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q12 (city free text, J13) | 2: written 2 |
| BH | Host State | peach | mapped | `ASSIGNMENT!TO_STATE_PROVINCE` (AN) | state checked against `country_state` (list empty); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q12 (state list) | 2: written 2 |
| BI | Home Cost Center Code | peach | mapped | `ASSIGNMENT_CF!CF_TEXT005` (AP) | number -> text; `cost_centres` list empty (Q15); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q15 | 2: written 2 |
| BJ | Home Cost Center Name | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT023` (AQ) | number -> text; code-like name warned; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q15/Q29 | 2: written 2 |
| BK | Host Cost Center Code | peach | mapped | `ASSIGNMENT_CF!CF_TEXT012` (AR) | number -> text; `cost_centres` list empty (Q15); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q15 | 2: written 2 |
| BL | Host Cost Center Name | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT024` (AS) | number -> text; code-like name warned; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q15/Q42/Q29 | 2: written 2 |
| BM | Mobility Cost Center Code | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT027` (AT) | number -> text; `cost_centres` list empty (Q15); Equus list check: no list (Q14); placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q15/Q29 | 2: written 2 |
| BN | Mobility Cost Center Name | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT028` (AU) | number -> text; code-like name warned; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q15/Q29 | 2: written 2 |
| BO | Expense Cost Center Code | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT032` (AV) | number -> text; `cost_centres` list empty (Q15); placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q15/Q29 | 2: written 2 |
| BP | Expense Cost Center Name | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT033` (AW) | number -> text; code-like name warned; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q15/Q29 | 2: written 2 |
| BQ | Home Entity | peach | mapped | `ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_1` (AX) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!FROM_COMPANY_LEVEL1; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| BR | Home Entity Group | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| BS | Home Organization | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| BT | Home Division | peach | mapped | `ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_4` (BA) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!FROM_COMPANY_LEVEL4; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| BU | Home Business Line | peach | mapped | `ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_3` (AZ) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!FROM_COMPANY_LEVEL3; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| BV | Home Role Region | peach | mapped | `ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_2` (AY) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!FROM_COMPANY_LEVEL2; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| BW | Home Job Title | peach | mapped | `ASSIGNMENT!FROM_TITLE` (BB) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| BX | Home Job Level | peach | mapped | `ASSIGNMENT!FROM_SALARY_GRADE` (BC) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| BY | Home Job Profile | peach | mapped | `ASSIGNMENT_CF!CF_TEXT002` (BD) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| BZ | Home Position | peach | mapped | `ASSIGNMENT_CF!CF_TEXT008` (BE) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | no data |
| CA | Host Position | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT025` (BF) | number -> text; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q29 | 2: written 2 |
| CB | Home Employment Type | peach | mapped | `ASSIGNMENT_CF!CF_TEXT007` (BG) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 1: written 1 |
| CC | Host Employment Type | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT026` (BH) | number -> text; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q33/Q29 | no data |
| CD | Home FTE | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| CE | Home Office Location | peach | mapped | `ASSIGNMENT_CF!CF_TEXT006` (BI) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| CF | Home Project | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| CG | Home Executive Leadership | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| CH | Home Director Flag | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| CI | Home Critical Role | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| CJ | Host Entity | peach | mapped | `ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_1` (BJ) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!TO_COMPANY_LEVEL1; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| CK | Host Entity Group | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| CL | Host Organization | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| CM | Host Division | peach | mapped | `ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_4` (BM) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!TO_COMPANY_LEVEL4; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| CN | Host Business Line | peach | mapped | `ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_3` (BL) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!TO_COMPANY_LEVEL3; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| CO | Host Role Region | peach | mapped | `ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_2` (BK) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!TO_COMPANY_LEVEL2; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| CP | Host Job Title | peach | mapped | `ASSIGNMENT!TO_TITLE` (BN) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| CQ | Host Job Level | peach | held | - | none (not migrated) | HELD: rule `host_job_level_to_salary_grade` (Q32) open, approved_for_import = false; Jake / Alka |  | 2: gated_rule 2 |
| CR | Host Job Profile | peach | mapped | `ASSIGNMENT_CF!CF_TEXT009` (BP) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| CS | Host FTE | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| CT | Home Scheduled Weekly Hours | cyan | mapped | `ASSIGNMENT_CF!CF_NUMBER004` (BQ) | range 0-80 warn only; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q29/Q16 | 2: written 2 |
| CU | Host Scheduled Weekly Hours | peach | mapped | `ASSIGNMENT_CF!CF_NUMBER003` (BR) | range 0-80 warn only; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q16 (158 h kept + warned) | 2: written 2 |
| CV | Host Office Location | peach | mapped | `ASSIGNMENT_CF!CF_TEXT013` (BS) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| CW | Host Project | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| CX | Host Supervisory Organization | peach | mapped | `ASSIGNMENT_CF!CF_TEXT011` (BT) | number -> text; Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| CY | Host Executive Leadership | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| CZ | Host Director Flag | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| DA | Host Critical Role | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| DB | Travel Type | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DC | Visa Type | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DD | Duration Type | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DE | Location Type | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DF | End of Assignment Reason | cyan | mapped | `ASSIGNMENT_CF!CF_TEXT030` (BU) | number -> text; Equus list check: no list (Q14); placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q29/Q34 | 2: written 2 |
| DG | Expected Assignment Outcome | cyan | unmapped | - | none (not migrated) | HELD (missing_target): cyan, no Equus column; Q7 PENDING BUSINESS APPROVAL, Schneider (SE-424) (D5) |  | 2: pending_target 2 |
| DH | Cancellation Reason | cyan | held | - | none (not migrated) | HELD: rule `cancellation_reason_sets_phase` (Q34) open, approved_for_import = false; Jake + Schneider |  | 2: gated_rule 2 |
| DI | Start Date Confirmed | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| DJ | End Date Confirmed | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| DK | VIP Flag | peach | mapped | `ASSIGNMENT!VIP` (BW) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | 2: written 2 |
| DL | Billable Resource | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| DM | Out of Policy Request | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 2: excluded 2 |
| DN | Rotation Schedule | peach | unmapped | - | none (not migrated) | unmapped: row 2 is a note, not a field; disposition Q30 (Schneider) |  | no data |
| DO | Remote Work Driver | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DP | Allocation % Home | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DQ | Allocation % Host | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DR | Mobility Budget Amount | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DS | Mobility Budget Currency | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DT | Mobility Budget Remaining | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DU | Mobility Budget Closed Date | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DV | Max Approved Duration | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 1: excluded 1 |
| DW | Global ID | cyan | mapped | `EMPLOYEE_CF!CF_TEXT014` (BX) | number -> text; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q29 | 2: written 2 |
| DX | Home Payroll ID | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DY | Host Payroll ID | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| DZ | Host Employee ID | peach | mapped | `EMPLOYEE_CF!CF_TEXT001` (BY) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q29 | 1: written 1 |
| EA | Contingent Worker ID | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EB | Requisition ID | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EC | Pre-Hire ID | cyan | mapped | `EMPLOYEE_CF!CF_TEXT015` (BZ) | number -> text; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q29 | 2: written 2 |
| ED | PII Nationality 1 | cyan | ignored | - | none (not migrated) | HELD (pending_target): `FIELD_OVERRIDES` None; Q9 PENDING BUSINESS APPROVAL, Schneider (D5) |  | 2: pending_target 2 |
| EE | Country of Permanent Residence | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EF | Military Status | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EG | Disability | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EH | Home Owner Status | peach | mapped | `ASSIGNMENT!HOMEOWNER` (CA) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | no data |
| EI | Primary Mobility Contact | peach | mapped | `ASSIGNMENT!COORDINATOR_RESOURCE_ASSIGNMENT_ID` (CB) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | 2: written 2 |
| EJ | Primary Mobility Email | peach | unmapped | - | none (not migrated) | pending_target: no technical field; Q41, Jake / Schneider |  | 2: pending_target 2 |
| EK | HR Advisor Home Name | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE1_ID` (CC) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | 1: written 1 |
| EL | HR Advisor Home Email | peach | unmapped | - | none (not migrated) | pending_target: no technical field; Q41, Jake / Schneider |  | 1: pending_target 1 |
| EM | HR Advisor Host Name | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE2_ID` (CD) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | 2: written 2 |
| EN | HR Advisor Host Email | peach | unmapped | - | none (not migrated) | pending_target: no technical field; Q41, Jake / Schneider |  | 2: pending_target 2 |
| EO | Line Manager Home Name | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE3_ID` (CE) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | 1: written 1 |
| EP | Line Manager Home Email | peach | unmapped | - | none (not migrated) | pending_target: no technical field; Q41, Jake / Schneider |  | 1: pending_target 1 |
| EQ | Line Manager Host Name | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE4_ID` (CF) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | 1: written 1 |
| ER | Line Manager Host Email | peach | unmapped | - | none (not migrated) | pending_target: no technical field; Q41, Jake / Schneider |  | 1: pending_target 1 |
| ES | Policy Name | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| ET | Policy Type | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EU | Group Name | peach | mapped | `ASSIGNMENT!GROUP_NAME` (CG) | Equus list check: no list (Q14); placeholders -> blank (Q43) | key alias (alias of ASSIGNMENT!GROUP; CR-M07, Q29); not formally approved (tracker I = Open) |  | 2: written 2 |
| EV | Tax Calculation Method | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | 1: excluded 1 |
| EW | Tax Filing Status | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EX | Tax Treaty | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EY | Residual Tax | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| EZ | Hypo Tax Do Not Calculate | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| FA | Marginal Hypo Tax Rate | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| FB | Effective Hypo Tax Federal Rate | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| FC | Effective Hypo Tax State Rate | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| FD | Move Initiator Name | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE11_ID` (CH) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | 2: written 2 |
| FE | Move Initiator Email | peach | unmapped | - | none (not migrated) | pending_target: no technical field; Q41, Jake / Schneider |  | 2: pending_target 2 |
| FF | Relocation Contact | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE12_ID` (CI) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | no data |
| FG | Mobility Host Contact (IMP) | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE10_ID` (CJ) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | no data |
| FH | HR Leader Host (Host HR SVP) | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE6_ID` (CK) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | no data |
| FI | Employee Benefits Manager - Host | peach | mapped | `ASSIGNMENT!ASSIGNMENT_ADDTL_RESOURCE13_ID` (CL) | display name sent unchanged; `resource_usernames` list empty (Q2); Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q2 | no data |
| FJ | Personal Home Address with Effective date | peach | address | - | free text -> LINE_1..LINE_3; constant LOCATION_TYPE = Departure Residence; PRIMARY_ADDRESS = Yes; BEGIN_DATE default gated (Q22/Q23) | LOCATION_TYPE = Departure Residence; not formally approved (tracker I = Open) | Q12 (free-text address); PC-J-15 (PRIMARY flag) | no data |
| FK | Location Type - Personal Home Address | peach | ignored | - | none (not migrated) | pending_target: describes constant LOCATION_TYPE; Q29/Q30, Jake / Alka, Schneider |  | 1: pending_target 1 |
| FL | Personal Host Address with Effective date | peach | address | - | free text -> LINE_1..LINE_3; constant LOCATION_TYPE = Destination Residence; PRIMARY_ADDRESS = Yes; BEGIN_DATE default gated (Q22/Q23) | LOCATION_TYPE = Destination Residence; not formally approved (tracker I = Open) | Q12; PC-J-15 | no data |
| FM | Location Type - Personal Host Address | peach | ignored | - | none (not migrated) | pending_target: describes constant LOCATION_TYPE; Q29/Q30, Jake / Alka, Schneider |  | 1: pending_target 1 |
| FN | Work Host Address with Effective date | peach | address | - | free text -> LINE_1..LINE_3; constant LOCATION_TYPE = Destination Office; PRIMARY_ADDRESS = Yes; BEGIN_DATE default gated (Q22/Q23) | LOCATION_TYPE = Destination Office; not formally approved (tracker I = Open) | Q12; PC-J-15 | no data |
| FO | Location Type - Work Host Address | peach | ignored | - | none (not migrated) | pending_target: describes constant LOCATION_TYPE; Q29/Q30, Jake / Alka, Schneider |  | 1: pending_target 1 |
| FP | Work Home Address with Effective date | peach | address | - | free text -> LINE_1..LINE_3; constant LOCATION_TYPE = Departure Office; PRIMARY_ADDRESS = Yes; BEGIN_DATE default gated (Q22/Q23) | LOCATION_TYPE = Departure Office; not formally approved (tracker I = Open) | Q12; PC-J-15 | no data |
| FQ | Location Type - Work Home Address | peach | ignored | - | none (not migrated) | pending_target: describes constant LOCATION_TYPE; Q29/Q30, Jake / Alka, Schneider |  | 1: pending_target 1 |
| FR | Phone Number | peach | contact | - | NUMBER_ADDRESS; constants Employee Preferred / Telephone; PRIMARY_CONTACT = Yes; START_DATE default gated (Q22/Q23) | LOCATION_TYPE/CONTACT_TYPE = Employee Preferred / Telephone; not formally approved (tracker I = Open) | Q13/Q14; PC-J-15 | no data |
| FS | Mobile Phone Number | peach | contact | - | NUMBER_ADDRESS; constants Personal Mobile / Telephone; PRIMARY_CONTACT = Yes; START_DATE default gated (Q22/Q23) | LOCATION_TYPE/CONTACT_TYPE = Personal Mobile / Telephone; not formally approved (tracker I = Open) | Q13/Q14; PC-J-15 | no data |
| FT | Talent Profile | peach | mapped | `UT_GOVERNANCE_QUESTIONNAIRE2!TALENT_RATING` (DE) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | no data |
| FU | Primary reasons for selecting this employee for the role (Business Justification) | peach | mapped | `UT_GOVERNANCE_QUESTIONNAIRE2!BUSINESS_JUSTIFICATION` (DF) | as read; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q7 (kept separate from DG, CR-M16) | no data |
| FV | Home Annualized Base Pay Currency | peach | mapped | `ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_CURRENCY` (DG) | as read; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q21/Q4 | no data |
| FW | Home Annualized Bse Pay Amount | peach | mapped | `ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_AMOUNT` (DH) | as read; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q21 | no data |
| FX | Host Annualized Base Pay Amount | peach | unmapped | - | none (not migrated) | unmapped: ASSIGNMENT!OFFER_SALARY not in Equus template; Q21 OPEN, Jake + Schneider |  | no data |
| FY | Host Annualized Base Pay Currency | peach | unmapped | - | none (not migrated) | unmapped: ASSIGNMENT!OFFER_SALARY_CURRENCY not in Equus template; Q21 OPEN, Jake + Schneider |  | no data |
| FZ | Home Profit Sharing Plan Frequency | cyan | mapped | `UT_CALCULATION_INPUT!HOME_PROFIT_SHARING_PLAN_FREQUENCY` (DJ) | Equus list check: no list (Q14); placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q21/Q29 | no data |
| GA | Host Profit Sharing Plan Frequency | cyan | mapped | `UT_CALCULATION_INPUT!HOST_PROFIT_SHARING_PLAN_FREQUENCY` (DK) | Equus list check: no list (Q14); placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q21/Q29 | no data |
| GB | Home Other Benefits Deduction Currency | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| GC | Home Other Benefits Deduction Amount | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| GD | Host Other Benefits Deduction Currency | peach | mapped | `UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS_CURRENCY` (DL) | as read; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q21/Q4 | no data |
| GE | Host Other Benefits Deduction Amount | peach | mapped | `UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS` (DM) | as read; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q21 | no data |
| GF | Home Target Bonus % | peach | mapped | `ASSIGNMENT_COMPENSATION_POLICY_CF!CF_NUMBER001` (DN) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q21 | no data |
| GG | Host Target Bonus % | peach | mapped | `ASSIGNMENT_COMPENSATION_POLICY_CF!CF_NUMBER002` (DO) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q21 | no data |
| GH | Home Family Allowance Currency | peach | mapped | `UT_CALCULATION_INPUT!FAMILY_ALLOWANCE_CURRENCY` (DP) | as read; placeholders -> blank (Q43) | `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q21/Q4 | no data |
| GI | Host Family Allowance Currency | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| GJ | Home HFM Code | peach | mapped | `ASSIGNMENT_CF!CF_TEXT004` (DR) | number -> text; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | no data |
| GK | Home Annual Vacation Days | peach | mapped | `ASSIGNMENT_COMPENSATION_POLICY!VACATION_DAYS_PER_YEAR` (DS) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | no data |
| GL | Housing Index Override (Unfurnished Apartment, 2 Bedroom - Best Area Min) | red | excluded | - | none (not migrated) | EXCLUDED: red = not required (Alka legend); Q30 OPEN, Schneider |  | no data |
| GM | Mobility Driver | peach | mapped | `ASSIGNMENT!RELOCATION_REASON` (DT) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | no data |
| GN | Original Hire Date | peach | mapped | `EMPLOYEE!HIRE_DATE` (DU) | date parse, configured order M/D/Y; ambiguous text warned (Q24); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) |  | no data |
| GO | Default Currency | yellow | mapped | `ASSIGNMENT!DEFAULT_CURRENCY` (DY) | as read; placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; row-2 key = template header; not formally approved (tracker I = Open) | Q38 | no data |
| GP | Host Family Size - do we need to calculate based on Spouse/Dependent at Host? | yellow | mapped | `ASSIGNMENT_FAMILY_SIZE!HOST_FAMILY_SIZE` (DV) | as read; placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q37 | no data |
| GQ | Home Family Size | yellow | mapped | `ASSIGNMENT_FAMILY_SIZE!HOME_FAMILY_SIZE` (DW) | as read; placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; `FIELD_OVERRIDES` target (template label; Q29 OPEN); not formally approved (tracker I = Open) | Q37 | no data |
| GR | RMC they use | yellow | mapped | `UT_CALCULATION_INPUT!RMC` (DZ) | as read; placeholders -> blank (Q43) | PENDING (yellow): written to DRAFT only, readiness `pending_column`, Alka / John; row-2 key = template header; not formally approved (tracker I = Open) | Q39 | no data |
| GS | Spouse First Name | none (Q8: no meaning) | mapped | `SPOUSE_DOMESTIC_PARTNER!FIRST_NAME` (EB) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q8 (brown = no meaning) + Q10 | no data |
| GT | Spouse Last Name | none (Q8: no meaning) | mapped | `SPOUSE_DOMESTIC_PARTNER!LAST_NAME` (EA) | as read; placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q8 + Q10 | no data |
| GU | Spouse Gender | none (Q8: no meaning) | mapped | `SPOUSE_DOMESTIC_PARTNER!GENDER` (EE) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q8 + Q14 | no data |
| GV | Spouse Date of Birth | none (Q8: no meaning) | mapped | `SPOUSE_DOMESTIC_PARTNER!BIRTH_DATE` (EF) | date parse, configured order M/D/Y; ambiguous text warned (Q24); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q8 | no data |
| GW | Spouse Relationship | none (Q8: no meaning) | mapped | `SPOUSE_DOMESTIC_PARTNER!RELATIONSHIP` (EC) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q8 + Q10 + Q14 (relationship list) | 2: written 2 |
| GX | Spouse At Host | none (Q8: no meaning) | mapped | `SPOUSE_DOMESTIC_PARTNER!STATUS` (ED) | Equus list check: no list (Q14); placeholders -> blank (Q43) | row-2 key = template header; not formally approved (tracker I = Open) | Q8 + Q14 | 1: written 1 |

Row count: 205 (B…GX). Column A of the source holds row labels and is not read
(`mapping_config.SOURCE_FIRST_COL = 2`, `lambda/src/mapping_config.py:33`).

## 5. Equus headers with no source column (40 of 146)

These headers are filled by the engine, filled from child records, or left blank. None of them is
derived from a guess.

| Template cols | Headers | How they are filled | Q / owner |
|---|---|---|---|
| A, B, D, E | `A` (Employee command), `Assignment Command`, `Assignment Number`, `Assignment Identifier` | Engine: INSERT/IGNORE, INSERT/UPDATE, `NEW`, identifier from source X (`build_records`, `transform.py:1016`; structure checked by `validate_row_structure`, `transform.py:748`) | Q1, Q17, Q19 CONFIRMED BY JAKE; Q26 OPEN; **D2 OPEN** (A3 header) |
| Q | `EMPLOYEE_CF!CF_TEXT016` | Blank (meaning unknown) | Q3: Jake |
| X | `ASSIGNMENT!INTERNATIONAL_RELO` | Blank: rule `international_relo_derivation` gated OFF | Q40: Schneider |
| AK, AO | `ASSIGNMENT!FROM_CURRENCY`, `TO_CURRENCY` | Blank: no source and no approved derivation | Q4: Jake |
| BO | `ASSIGNMENT!TO_SALARY_GRADE` | Blank: Host Job Level (CQ) is held | Q32: Jake / Alka |
| CM–CX | `ASSIGNMENT_MAILING_ADDRESS!*` | From address columns FJ/FL/FN/FP (child rows). `BEGIN_DATE` stays blank while the date rule is gated. `CITY`, `COUNTRY_CODE`, `STATE_PROVINCE_CODE` and `POSTAL_CODE` are not split out of the free text | Q12, Q22/Q23; PC-J-15 |
| CY–DD | `ASSIGNMENT_EMPLOYEE_CONTACT!*` | From contact columns FR/FS (child rows). `START_DATE` stays blank while gated | Q13, Q14, Q22/Q23 |
| DI | `ASSIGNMENT_COMPENSATION_SALARY!EFFECTIVE_DATE` | Blank: default from the start date is gated OFF | Q22: Jake |
| DQ | `UT_CALCULATION_INPUT!FAMILY_ALLOWANCE` | Blank: no source column | Q21: Jake + Schneider |
| DX | `ASSIGNMENT_FAMILY_SIZE!EFFECTIVE_DATE` | Blank: gated OFF | Q37: Schneider |
| EG–EP | `ASSIGNMENT_DEPENDENT!*`, `ASSIGNMENT_DEPENDENT_CF!*` | From dependent column groups in the main sheet, or from an approved keyed dependent source (none configured). The real extract's unkeyed "Dependent Import" sheet is held | Q11, Q18–Q20; **D3 OPEN**: Alka / Topia |

## 6. Assumptions (recorded instead of asking)

1. "The mapping report of the fixture run" means the default-configuration run on
   `tests/fixtures/topia_sample.xlsx`. That fixture is the only tracked workbook with the full
   205-column source layout, and the runner allows it in documentation because its data rows are
   synthetic.
2. "Rule status" reports the configuration state (gated, held, excluded, pending) and the Q number.
   It does not repeat "not formally approved" as a separate classification, because no mapping is
   formally approved. Jake's confirmations are cited as evidence (CONFIRMED BY JAKE), never as
   approval.

