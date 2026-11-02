# Requirements Traceability Matrix: Q1–Q43

Evidence package item 1 of `clarifications/UAT_MASTER_COMMAND.md` §12. OCNORA task T-1386 (UAT-01).
Prepared 2026-11-02 against base commit `e61d99f`. This change is documentation only: no code
or tests were changed.

**Readiness status: NOT READY FOR UAT.** No row below is SIGNED OFF or UAT VERIFIED. The Equus
import, its results and the post-import audit are **INCOMPLETE / NOT EXECUTED**, because this
environment has no Equus access.

## 0. How to read this matrix

### Sources
| Ref | Source | Notes |
|---|---|---|
| JAKE | `clarifications/jake-1.xlsx`, sheet `datamapping clarrification shef`, column J "Equus Response"; column I "Status" | SHA-256 `30d0666c…1c09`, as recorded in `clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md:5` |
| EX | `clarifications/jake-1.xlsx`, sheet ` Mny dependents Ex` (note the leading space), rows 4–7, 146 columns | Jake's worked example |
| ASSESS | `clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md` | Jake's answers condensed and classified, with D1–D5 |
| MC | `clarifications/UAT_MASTER_COMMAND.md` | BR-01…BR-20, §4 focus areas, gates G1–G12 |
| OQ | `OPEN_QUESTIONS.md` | Cell evidence from the real extract (SRC) and the Equus sample (EQ) |
| SRC | Real Topia extract, sheet "Full Historical Past Assignment" (row 2 = Equus key, row 3 = UI label, row 4 = Topia label, data rows 5–6) | Personal data. Only column letters and counts are cited here |
| EQ | Equus sample `DATA MIGRATION SE EAFH TEST FILE.xlsx` | Personal or test data. Only cell coordinates are cited |
| TPL | `lambda/src/templates/equus_eafh_template.xlsx`, Sheet1 rows 1–3 | |

### Cell references and quotations (read this first)
- **Row rule.** Question Qn is on tracker row n+1, so Jake's answer is `JAKE!J{n+1}` and its
  status is `JAKE!I{n+1}`. The rule comes from the cells that ASSESS names explicitly: J2 = Q1
  (`ASSESS:20`), J6 = Q5 (`ASSESS:43`), J12/J19/J20 = Q11/Q18/Q19 (`ASSESS:48`). Every I cell
  reads **Open** (`ASSESS:15`).
- **Quotation limitation.** The runner sandbox blocked every way of opening the workbook in this
  job: openpyxl, unzip, tar and a read-only subagent were all denied, and the Read tool cannot
  open binary .xlsx files. As a result:
  - The answer texts below come from ASSESS §1–§2, which was prepared from the workbook on
    2026-11-02.
  - Where ASSESS quotes Jake word for word (J2 and J6), the text is marked **[verbatim per ASSESS]**.
  - Every other answer is ASSESS's condensed wording, marked **[condensed per ASSESS]**.
  - A reviewer must re-read column J with openpyxl and replace each condensed quote with the
    exact cell text. This is tracked as `PC-ADM-01` in `docs/uat/PENDING_CLARIFICATIONS.md`.
- Q21–Q43 have no answer. ASSESS §3 and MC §2 both state this.

### Classification vocabulary
Each row uses exactly one value from MC §2, never a combination: CONFIRMED BY JAKE ·
PARTIALLY CLARIFIED · PENDING BUSINESS APPROVAL · PENDING EQUUS CONFIGURATION ·
IMPLEMENTED AND TESTED · UAT VERIFIED · SIGNED OFF.
- For unanswered questions (Q21–Q43), the value names who must answer: business or Equus. The
  answer cell reads *no answer*.
- Two classifications differ deliberately from ASSESS, which gave combined values:
  - **Q1 is PARTIALLY CLARIFIED.** Jake's text is clear, but his own example contradicts it (D1).
  - **Q20 is PARTIALLY CLARIFIED.** The technical rule is confirmed, but the lookups are pending.
- No row is IMPLEMENTED AND TESTED. For that, the code must implement the stated rule and a test
  must assert it, and no Jake answer is formal approval yet (every I cell reads Open).

### Test ids
All tests below are in `tests/test_migration.py`: 22 functions, collected as 24 items at base,
per `ASSESS:10`. Later UAT tasks append their own test ids to the "Tests" column.
**⚠ encodes unapproved behaviour** means the test asserts a rule that nobody has approved. Its
assertion may change only for an approved decision that cites the Q number.

### Gate legend (MC §10)
| Gate | Requirement |
|---|---|
| G1 | Mappings approved or dispositioned |
| G2 | Equus custom fields and lookups available |
| G3 | Identity and linkage conflicts resolved |
| G4 | Automated tests pass |
| G5 | Reconciliation passes |
| G6 | Controlled Equus UAT import completed |
| G7 | Equus records independently validated |
| G8 | Duplicate, error and retry scenarios verified |
| G9 | Security and confidentiality checks pass |
| G10 | No unresolved critical defects or unapproved exceptions |
| G11 | Business reviewers approve the evidence |
| G12 | Authorised sign-off |

Every row also blocks G6, G7, G11 and G12 until UAT runs. The Gates column lists only the
additional, specific gates each question blocks.

Escalations D1–D5 are recorded in `docs/uat/PENDING_CLARIFICATIONS.md` §1. Code findings
(CR-xx) are in `docs/uat/CODE_REVIEW_REPORT.md`.

## 1. Matrix (43 rows)

| Q | Topic | Jake's answer (cell) | Classification | Code locations | Affects UAT data? (evidence) | Tests | Decision owner | Closure evidence required | Gates blocked |
|---|---|---|---|---|---|---|---|---|---|
| Q1 | Identity: Employee Number, Assignment Number, Assignment Identifier, CF_TEXT029 (BR-01) | `JAKE!J2` [verbatim per ASSESS]: "Topia ID should be entered for the Assignment Identifier and for ASSIGNMENT_CF!CF_TEXT029." [condensed per ASSESS]: Employee Number = Talent Link ID; Assignment Number = "NEW"; Identifier links the rows. Contradicted by EX!C4/E4/Y4 (D1) | PARTIALLY CLARIFIED | `mapping_config.py:54` (identifier label = "assignment id"), `:68` (same column → CF_TEXT029), `:61` (SYSTEM_NUMBER → Employee Number), `:49` (NEW); `transform.py:362`, `:376-382` (**fallback to Employee Number, forbidden by BR-01**: CR-T09), `:446`, `:450` | Yes. SRC X populated in both data rows; SRC F (red "Topia ID") empty in both; SRC E populated (`ASSESS:26-28`) | `test_core_fields`, `test_child_records_become_update_rows`. No test covers a missing or duplicate identifier | Jake (Equus semantics) + Alka/John (which Topia column) | Written answer to D1 naming the source column for Assignment Identifier and for CF_TEXT029; corrected example or confirmation; no-fallback test | G1, G3, G10 |
| Q2 | Resource usernames (BR-02) | `JAKE!J3` [condensed per ASSESS]: "I will provide the values … with the other lookup fields" | PENDING EQUUS CONFIGURATION | `mapping_config.py:186-188` (`RESOURCE_MAP = {}`); `transform.py:90`, `:316-320` (names passed through and only counted), `:121-122` | Yes. SRC EI/EK/EM/EO/EQ/FD hold display names; EQ CC4/CH4 hold usernames (`OPEN_QUESTIONS.md:27`) | `test_unmapped_resources_listed_once_per_value` | Jake (Equus) | Versioned resource-username list with approval metadata; unknown name raised as an exception, not passed through | G1, G2 |
| Q3 | Citizenship and birth country; template Q–T shift (BR-03) | `JAKE!J4` [condensed per ASSESS]: same as Q2 (values to be provided) | PENDING EQUUS CONFIGURATION | `mapping_config.py:125-132` (COUNTRY_HEADERS), `:69` (Nationality 3 → CF_TEXT013); `transform.py:161-169` (shift warning), `:292-300` | Yes. Template Q–T shift is warned on every run (TPL Q2:T3, `OPEN_QUESTIONS.md:35`). Whether SRC citizenship columns hold data is not evidenced in the repo: UNVERIFIED | `test_country_of_birth_converted_despite_shifted_template_rules`, `test_conversions_and_rules` | Jake (Equus) | Approved meaning of `EMPLOYEE_CF!CF_TEXT016`; corrected template labels; approved country list | G1, G2 |
| Q4 | Currencies: home, host, salary, default (BR-04) | `JAKE!J5` [condensed per ASSESS]: same as Q2 | PENDING EQUUS CONFIGURATION | No code writes FROM/TO_CURRENCY: no SRC column carries them (`OPEN_QUESTIONS.md:40-41`). `mapping_config.py:86` (home pay currency override) | Yes. EQ AK/AO are currency columns; no SRC currency column | none | Jake (Equus) + Schneider comp | Approved currency source per field, or approval to leave blank; currency list | G1, G2 |
| Q5 | Duplicates, row and file errors, header failure (BR-05) | `JAKE!J6` [verbatim per ASSESS, excerpt]: "if there are issues in the header row, the full import fails." [condensed per ASSESS]: duplicates can create duplicates; protections are not a fail-safe; audit report; row-level rejection; test first | CONFIRMED BY JAKE | `handler.py:144-170` (S3 idempotency only); `transform.py:136-138` (duplicate template header only); no source-duplicate, header-version or row-classification code (CR-T10, CR-T16, CR-H05) | Yes. Applies to every run; header risk D2 | `test_duplicate_event_is_skipped`, `test_duplicate_s3_delivery_with_moto`, `test_output_key_changes_with_mapping_version`, `test_template_headers_preserved_and_sample_rows_removed` | Jake (Equus behaviour); developer (controls) | Source-duplicate, header and row-error controls with tests; UAT evidence of Equus duplicate and rejection behaviour (G8) | G3, G8, G10 |
| Q6 | Confidentiality type (BR-06) | `JAKE!J7` [condensed per ASSESS]: pending with Equus; column to be added to template | PENDING EQUUS CONFIGURATION | `transform.py:245-246` (MISSING_TARGET flag; value not written) (D5) | Yes. SRC AG holds data in both rows (`ASSESS:67`) | `test_governance_flags` (asserts MISSING_TARGET) | Jake (Equus) + Schneider security | Template version with the confidentiality column; HELD-exception reporting until then | G1, G2, G9 |
| Q7 | Expected Assignment Outcome (BR-07) | `JAKE!J8` [condensed per ASSESS]: pending Schneider; ticket SE-424; lookup values, and does it differ from reason? | PENDING BUSINESS APPROVAL | `transform.py:245-246` (MISSING_TARGET; value not written) (D5); `mapping_config.py:84-85` (business justification override, a separate field) | Yes. SRC DG holds data in both rows (`ASSESS:67`) | `test_governance_flags` | Schneider business (SE-424) | SE-424 decision: target field, lookup values, relation to reason/justification | G1, G2 |
| Q8 | Brown fill on spouse columns (BR-08) | `JAKE!J9` [condensed per ASSESS]: brown has no meaning | CONFIRMED BY JAKE | `mapping_config.py:26-31` (brown FFB45F06 not listed); `transform.py:155-157`, `:247-248` (flagged UNKNOWN_COLOUR) | Yes. SRC GS–GX row 4 brown (`OPEN_QUESTIONS.md:75-77`) | `test_governance_flags` (asserts UNKNOWN_COLOUR on "Spouse First Name": **⚠ contradicts Jake's answer; change only for an approved Q8**) | Jake (answered); Schneider (formal approval) | Formal approval of Q8; configuration entry "FFB45F06 = no meaning (Q8)" with test | G1 |
| Q9 | PII Nationality 1 (BR-09) | `JAKE!J10` [condensed per ASSESS]: John confirmed with Anand it is not needed; Schneider must decide | PENDING BUSINESS APPROVAL | `mapping_config.py:96-98` (`"pii nationality 1": None`, permanently excluded); `transform.py:218-221` (D5) | UNVERIFIED. SRC ED is a cyan column; whether it holds data is not evidenced in the repo | none | Schneider business | Written Schneider decision (include, exclude or remap); until then reported as HELD, not silently excluded | G1, G9 |
| Q10 | Spouse or partner without names (BR-10) | `JAKE!J11` [condensed per ASSESS]: spouse/partner can be imported without names | CONFIRMED BY JAKE | `transform.py:335-339` (warning text claims "Equus may reject…": CR-T14); `mapping_config.py:179-180` (comment "need a name") | Yes. SRC GW5/GW6 hold a relationship without a name (`OPEN_QUESTIONS.md:83`) | `test_validation_warnings_keep_values` | Jake (answered); Schneider (formal approval) | Formal approval; warning reworded as a data-quality note; "Partner" value in the Q14 list | G1, G2 |
| Q11 | Dependents in the same workbook (BR-11) | `JAKE!J12` [condensed per ASSESS]: latest template includes dependents; same import; Assignment Identifier links | PARTIALLY CLARIFIED | `mapping_config.py:14` (only SOURCE_SHEET is read); `transform.py:232-235`, `:344-353`, `:431-436`, `:612-614` (D3) | Yes. SRC "Dependent Import" sheet A1:R4 has no Employee or Assignment key and is never read (`ASSESS:50-52`), so dependents are not migrated today | `test_child_records_become_update_rows` (synthetic dependent columns in the main sheet) | Alka/Topia (keyed extract); Jake (layout) | A keyed dependent source (employee + assignment key) and its column map | G1, G3, G5 |
| Q12 | Country, state, city values (BR-12) | `JAKE!J13` [condensed per ASSESS]: country and state must match system values (to be provided); city is free text | PARTIALLY CLARIFIED | `mapping_config.py:190-243` (hand-written ISO-3 table, CR-M34); `transform.py:292-300` (unknown country passed through), `:400-415` (address free text → LINE_1-3, PRIMARY_ADDRESS = "Yes") | Yes. SRC FJ/FL/FN/FP addresses are free text (`OPEN_QUESTIONS.md:87`); home/host countries in both rows | `test_conversions_and_rules`, `test_country_of_birth_converted_despite_shifted_template_rules`, `test_child_records_become_update_rows` | Jake (Equus lists) | Equus country and state lists, versioned; unknown values held, not passed through | G1, G2 |
| Q13 | E-mail fields (BR-13) | `JAKE!J14` [condensed per ASSESS]: no e-mail contact type; USER_ID_EMAIL primary, CF_TEXT008 secondary | CONFIRMED BY JAKE | No special code. SRC G/H map by row-2 key through `transform.py:236-237`; `mapping_config.py:113-116` creates phone contacts only | Yes. SRC G/H map to these headers on every run (`ASSESS:88`); whether they hold data in both rows is UNVERIFIED in the repo | none (`ASSESS:88`: test to add) | Jake (answered); Schneider (formal approval) | Formal approval; test asserting G→USER_ID_EMAIL, H→CF_TEXT008 and no e-mail contact rows | G1 |
| Q14 | Dropdown lists (BR-14) | `JAKE!J15` [condensed per ASSESS]: "Yes, will provide" | PENDING EQUUS CONFIGURATION | `mapping_config.py:182-184` (optional `equus_lists.json`; a missing file means no check), `:154-159`; `transform.py:314-315`, `:321-322`, `:599-601` (no schema validation, CR-T21) | Yes. Every dropdown value is sent unchecked today (no list file exists) | `test_allowed_dropdown_values_checked`, `test_conversions_and_rules` | Jake (Equus) | Versioned list file with approval metadata; unknown value raised as an exception | G1, G2 |
| Q15 | Cost centres, code vs name (BR-15) | `JAKE!J16` [condensed per ASSESS]: same as Q2 | PENDING EQUUS CONFIGURATION | `mapping_config.py:71-76`, `:169-177`; `transform.py:328-331` | Yes. SRC BL5/BL6 (host cost-centre name) hold a code (`OPEN_QUESTIONS.md:92`) | `test_validation_warnings_keep_values` | Jake (Equus) + Schneider (cost-centre list) | Approved code-to-name list; names never manufactured | G1, G2 |
| Q16 | Host weekly hours = 158 (BR-16) | `JAKE!J17` [condensed per ASSESS]: same as Q2 | PENDING BUSINESS APPROVAL | `mapping_config.py:161-167` (0–80 review range); `transform.py:309-312` (warn, keep) | Yes. SRC CU6 (`OPEN_QUESTIONS.md:91`) | `test_validation_warnings_keep_values` | Schneider business (source owner) | Source-owner confirmation or correction of CU6; approved range | G1 |
| Q17 | Multiple historical assignments per employee (BR-17) | `JAKE!J18` [condensed per ASSESS]: NEW + INSERT creates; Employee Number links assignments; Identifier links rows | CONFIRMED BY JAKE | `mapping_config.py:45-49`; `transform.py:365`, `:444`, `:455` (`seen_employees` also records `None`: CR-T11) | No in the current 2-row extract (SRC E has two distinct values, `ASSESS:27-28`). Yes for the full extract (Q25) | `test_core_fields`. No multi-assignment test | Jake (answered); Schneider (formal approval) | Multi-assignment tests (same employee, several identifiers); UAT import evidence | G3, G8 |
| Q18 | Dependent and spouse linkage by Assignment Identifier (BR-18) | `JAKE!J19` [condensed per ASSESS]: spouse and dependents belong to the assignment; each row carries the Assignment Identifier | CONFIRMED BY JAKE | `transform.py:440-454` (every output row carries the identifier, `:450`) | Yes for address and contact UPDATE rows. Dependents: not migrated (D3) | `test_child_records_become_update_rows` | Jake (answered); Schneider (formal approval) | Cross-assignment isolation test; child-row reconciliation (CR-T18) | G3, G5 |
| Q19 | Many dependents: INSERT then UPDATE rows (BR-19) | `JAKE!J20` [condensed per ASSESS]: see ` Mny dependents Ex`; as many dependents as needed, with the Identifier | CONFIRMED BY JAKE | `transform.py:440` (unbounded `n_rows`), `:441-454`, `:344-353` | No dependents in the UAT data today (D3). The example layout is in EX rows 4–7 | `test_child_records_become_update_rows` (2 dependents) | Jake (answered); Schneider (formal approval) | Generated rows validated against EX rows 4–7 (commands, identifier, blank Employee Number on UPDATE rows); 3+ dependent test | G3 |
| Q20 | Minimum dependent data (BR-20) | `JAKE!J21` [condensed per ASSESS]: not recommended, but one dependent field creates a dependent; lookups to follow | PARTIALLY CLARIFIED | `transform.py:431-436` (record kept if any value), `:335-339` (name warning) | No dependents in the UAT data today (D3) | none for dependents with one field | Jake (answered and lookups); Schneider (data quality) | Lookups; test for a dependent with one populated field; no entirely empty records | G1, G2 |
| Q21 | Compensation: home/host annualised pay → SALARY_AMOUNT/CURRENCY, OFFER_SALARY | *no answer* (`JAKE!J22`) | PENDING BUSINESS APPROVAL | `mapping_config.py:86-92` (pay, profit-sharing, benefits and allowance overrides); SRC FX/FY not in template | Empty in the 2-row extract; yes for the full extract (`ASSESS:101`) | `test_report_lists_unmapped_columns` (host pay unmapped) | Jake + Schneider comp | Approved header crosswalk for every pay field | G1, G10 |
| Q22 | Blank salary data and salary effective date | *no answer* (`JAKE!J23`) | PENDING BUSINESS APPROVAL | `mapping_config.py:143-145` (`DEFAULT_DATES_FROM_START = True`), `:136-141`; `transform.py:390-395` (D4) | Yes (`ASSESS:102`) | none (rule only reconciled as "set_by_rule" in `test_reconciliation_balances`) | Jake | Written effective-date rule; until then the default stays OFF for import-ready output | G1, G10 |
| Q23 | Start and actual end date semantics; address and contact begin dates | *no answer* (`JAKE!J24`) | PENDING BUSINESS APPROVAL | SRC AO/AP map by row-2 key (`transform.py:236-237`); `transform.py:396`, `:409`, `:427` (begin/start dates default to the assignment start) | Yes. SRC AO/AP yellow and populated in both rows (`OPEN_QUESTIONS.md:60`) | `test_governance_flags` (Assignment Start Date PENDING) | Alka/John | Planned-vs-actual rule; begin-date rule for addresses and contacts | G1, G10 |
| Q24 | Date formats and parsing | *no answer* (`JAKE!J25`) | PENDING EQUUS CONFIGURATION | `transform.py:22-23` (US output; 6 parse patterns incl. `%m/%d/%y`), `:84`, `:254-267`, `:477-479` | Yes. Every date column (`ASSESS:104`) | `test_conversions_and_rules` (MM/DD/YYYY), `test_core_fields` | Jake (Equus) | Accepted formats per field; ambiguity rule (no silent d/m vs m/d conversion) | G1 |
| Q25 | Source inventory and cutoff | *no answer* (`JAKE!J26`) | PENDING BUSINESS APPROVAL | `mapping_config.py:14-19` (sheet and row layout); `transform.py:611-614`, `:619` (SHA-256 recorded) | Yes. Only a 2-row extract is available (`ASSESS:105`) | `test_reconciliation_balances` (records SHA-256) | Alka/Topia + Schneider | Final extract identified by hash, cutoff date, expected employee and assignment counts | G5 |
| Q26 | INSERT/UPDATE/IGNORE semantics | *no answer* (`JAKE!J27`) | PENDING EQUUS CONFIGURATION | `mapping_config.py:47-48`; `transform.py:443-449` (IGNORE for a repeated employee and on UPDATE rows) | Yes (`ASSESS:106`) | `test_child_records_become_update_rows` | Jake (Equus) | Action matrix per row type; UAT import evidence | G3, G8 |
| Q27 | Reconciliation acceptance criteria | *no answer* (`JAKE!J28`) | PENDING BUSINESS APPROVAL | `transform.py:487-532` (INSERT rows only; non-date values not compared: CR-T18, CR-T19); `mapping_config.py:136-141` | Yes (`ASSESS:107`) | `test_reconciliation_balances`, `test_reconciliation_catches_silent_loss`, `test_handler_raises_and_writes_report_only_on_reconciliation_failure` | Schneider + Jake | Agreed criteria covering child rows, values and counts | G5 |
| Q28 | Visibility of confidential records | *no answer* (`JAKE!J29`) | PENDING BUSINESS APPROVAL | No code touches it. SRC AG is dropped at `transform.py:245-246` | Yes. SRC AG populated (`ASSESS:108`) | none | Schneider security | Security-group process; approved access control | G9 |
| Q29 | Cyan custom-field inventory | *no answer* (`JAKE!J30`) | PENDING EQUUS CONFIGURATION | `mapping_config.py:67-95` (CF numbers taken from template labels), `:58-62` | Yes (`ASSESS:109`) | `test_conversions_and_rules` (CF_TEXT014), `test_core_fields` (CF_TEXT029) | Jake/Alka | Signed-off custom-field inventory (cyan column → CF field) | G1, G2 |
| Q30 | Scope and exclusions (red rule) | *no answer* (`JAKE!J31`) | PENDING BUSINESS APPROVAL | `mapping_config.py:21-31`, `:109-110`; `transform.py:144-158`, `:188-191` | Yes. 74 red columns in the fixture layout (`test_governance_flags`) | `test_red_column_excluded_even_with_valid_row2_key`, `test_governance_flags` | Schneider business | Written disposition of every red column | G1 |
| Q31 | Approval authority | *no answer* (`JAKE!J32`) | PENDING BUSINESS APPROVAL | No code touches this: it is governance, not a source column | Gate G12 only | none | Schneider PMO | Named signatory and approval process | G11, G12 |
| Q32 | Job grade / Host Job Level header | *no answer* (`JAKE!J33`) | PENDING BUSINESS APPROVAL | `mapping_config.py:79` (override to TO_SALARY_GRADE, against SRC row-2 TO_TITLE) (D4) | Yes. Populated in both rows (`ASSESS:112`) | `test_conversions_and_rules` (TO_SALARY_GRADE: **⚠ encodes unapproved behaviour**) | Jake/Alka | Written correction of the row-2 key | G1, G10 |
| Q33 | Host Employment Type → CF_TEXT026 | *no answer* (`JAKE!J34`) | PENDING BUSINESS APPROVAL | `mapping_config.py:78` | No. No data in the extract (`ASSESS:113`) | none | Renniece/Cathy (per ASSESS) | Confirmation of the target | G1 |
| Q34 | Status, phase and cancellation precedence | *no answer* (`JAKE!J35`) | PENDING BUSINESS APPROVAL | `mapping_config.py:70` (Status Category → PHASE), `:118-120`; `transform.py:216-217`, `:385-386` (non-empty reason ⇒ "Cancelled", overwriting the mapped value) (D4) | Yes. Both rows cancelled (`ASSESS:114`); SRC AB unmapped (`OPEN_QUESTIONS.md:57`) | `test_conversions_and_rules` (PHASE == "Cancelled": **⚠ encodes unapproved behaviour**) | Jake + Schneider | Precedence rule between status category, assignment status and cancellation reason | G1, G10 |
| Q35 | Person Status and policy crosswalk | *no answer* (`JAKE!J36`) | PENDING EQUUS CONFIGURATION | `mapping_config.py:154-159` (`employee→Employee`, `employed→Active`); `transform.py:314-315`; SRC AA passed through | Yes. SRC S and AA populated (`OPEN_QUESTIONS.md:55-56`) | `test_conversions_and_rules` (Active/Employee: **⚠ encodes unapproved behaviour**) | Jake (Equus) | Approved lookup lists and crosswalk | G1, G2, G10 |
| Q36 | Pay installments | *no answer* (`JAKE!J37`) | PENDING BUSINESS APPROVAL | SRC AK/AL map by row-2 key (`transform.py:236-237`) → CF_NUMBER001/002 | Yes. Populated, including a fractional value (`ASSESS:116`) | none | Schneider payroll | Meaning and allowed values | G1 |
| Q37 | Family size and its effective date | *no answer* (`JAKE!J38`) | PENDING BUSINESS APPROVAL | `mapping_config.py:93-95`; `transform.py:394-395` (effective date defaults to start) (D4) | No. No data in the extract (`ASSESS:117`) | none | Schneider | Counting rule; effective-date rule | G1, G10 |
| Q38 | Default currency | *no answer* (`JAKE!J39`) | PENDING EQUUS CONFIGURATION | No special code. SRC GO maps by row-2 key (`transform.py:236-237`) | No. No data (`OPEN_QUESTIONS.md:61`) | none | Jake | Rule and source | G1 |
| Q39 | RMC | *no answer* (`JAKE!J40`) | PENDING BUSINESS APPROVAL | No special code. SRC GR maps by row-2 key (`transform.py:236-237`) | No. No data (`OPEN_QUESTIONS.md:63`) | none | Schneider | Definition and source | G1 |
| Q40 | International assignment derivation | *no answer* (`JAKE!J41`) | PENDING BUSINESS APPROVAL | `transform.py:387-389` (Yes/No derived from home ≠ host country) (D4); `mapping_config.py:138` | Yes (`ASSESS:120`) | `test_conversions_and_rules` (INTERNATIONAL_RELO == "Yes": **⚠ encodes unapproved behaviour**) | Schneider | Approved rule, or approval to leave blank | G1, G10 |
| Q41 | Resource e-mail columns | *no answer* (`JAKE!J42`) | PENDING BUSINESS APPROVAL | No code touches SRC EJ/EL/EN/EP/ER/FE: they end `unmapped` at `transform.py:238-239` | Yes. Populated (`ASSESS:121`) | none | Jake/Schneider | Disposition (lookup key for Q2, or excluded) | G1 |
| Q42 | CF_TEXT024 code vs name | *no answer* (`JAKE!J43`) | PENDING BUSINESS APPROVAL | `mapping_config.py:72`, `:169-177`; `transform.py:328-331` (warning only) | Yes (`ASSESS:122`) | `test_validation_warnings_keep_values` | Alka | Cost-centre lookup | G1, G2 |
| Q43 | Missing-value / placeholder normalisation | *no answer* (`JAKE!J44`) | PENDING BUSINESS APPROVAL | `mapping_config.py:147-152`; `transform.py:44-49`, `:270-274`, `:282-284`, `:369-370` | Yes. SRC N6 holds a glued placeholder (`mapping_config.py:149-150`) | `test_dirty_placeholder_blanked_with_cell_reference`, `test_conversions_and_rules` | Alka | Placeholder pattern list per field | G1 |

Row count: 43 (Q1–Q43).

## 2. Summary by classification
| Classification | Questions |
|---|---|
| CONFIRMED BY JAKE | Q5, Q8, Q10, Q13, Q17, Q18, Q19 |
| PARTIALLY CLARIFIED | Q1, Q11, Q12, Q20 |
| PENDING EQUUS CONFIGURATION | Q2, Q3, Q4, Q6, Q14, Q15, Q24, Q26, Q29, Q35, Q38 |
| PENDING BUSINESS APPROVAL | Q7, Q9, Q16, Q21, Q22, Q23, Q25, Q27, Q28, Q30, Q31, Q32, Q33, Q34, Q36, Q37, Q39, Q40, Q41, Q42, Q43 |
| IMPLEMENTED AND TESTED / UAT VERIFIED / SIGNED OFF | none |

7 + 4 + 11 + 21 = 43.

## 3. Gate status implied by this matrix (snapshot; see ASSESS §5)
| Gate | Status | Driving questions |
|---|---|---|
| G1 | Blocked | Every PENDING row, plus Q1 |
| G2 | Blocked | Q2–Q4, Q6, Q7, Q12, Q14, Q15, Q20, Q29, Q35, Q42 |
| G3 | Blocked | Q1 (D1), Q5 (D2 header), Q11 (D3), Q17, Q18, Q19, Q26 |
| G4 | Pass for the 24 existing tests at base only (per `ASSESS:10`; the runner gate re-executes them) | |
| G5 | Fail | Q27: child rows and non-date values not reconciled (CR-T18, CR-T19); Q25 |
| G6, G7, G8 | NOT EXECUTED | No Equus access |
| G9 | Blocked | Q6, Q9, Q28; CR-R01 (workbooks tracked in git) |
| G10 | Fail | D4 rules active by default (Q1 fallback, Q22, Q23, Q32, Q34, Q35, Q37, Q40) |
| G11, G12 | Not obtained | Q31 |

**Readiness: NOT READY FOR UAT.**
