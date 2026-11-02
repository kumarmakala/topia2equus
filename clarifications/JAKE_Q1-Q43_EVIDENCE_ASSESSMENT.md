# Jake's clarification responses (Q1–Q43): evidence-based assessment

Prepared 2026-11-02 from `clarifications/jake-1.xlsx`
(SHA-256 `30d0666c4f58bd37188831b1d592380a2fd2231c03f356f44d8559e0ed001c09`, 2 sheets:
`datamapping clarrification shef` = the question tracker with column J "Equus Response";
` Mny dependents Ex` = Jake's worked import example, 146 columns, rows 4–7).

Code reviewed at commit `d5b79cf` (branch `migration-hardening`):
`lambda/src/mapping_config.py`, `lambda/src/transform.py`, `lambda/src/handler.py`,
`tests/test_migration.py` (24 tests, all passing on 2026-11-02 with `.venv` Python 3.12.15).

Status vocabulary (never combined): CONFIRMED BY JAKE · PARTIALLY CLARIFIED ·
PENDING BUSINESS APPROVAL · PENDING EQUUS CONFIGURATION · OPEN (no answer) ·
IMPLEMENTED AND TESTED · UAT VERIFIED · SIGNED OFF.
Every question's tracker Status cell (column I) still reads **Open**: nothing below is SIGNED OFF.

## 1. Critical discrepancies found in the evidence

### D1 — Identifier: Jake's text and Jake's example disagree (Q1)
- Text (J2): "Topia ID should be entered for the Assignment Identifier **and** for
  ASSIGNMENT_CF!CF_TEXT029."
- Example (` Mny dependents Ex` row 4): Employee Number C4 = **60936**, Assignment Identifier
  E4 = **60936** (identical to the Employee Number), CF_TEXT029 "Topia File Number" Y4 = **55555**.
  So the example neither uses one value for both fields nor keeps the Assignment Identifier
  distinct from the Employee Number.
- Source ambiguity: Topia column F "Topia ID" is red (not required) and empty in both data
  rows; column X "Assignment ID" (cyan) holds 19338 / 21100; column E "Employee ID" (Talent
  Link ID) holds 60936 / 10500545. Which column Jake means by "Topia ID" is not stated.
- Code today: `mapping_config.ASSIGNMENT_IDENTIFIER_LABEL = "assignment id"` (col X) feeds both
  Assignment Identifier and CF_TEXT029 (`mapping_config.py:54,68`; `transform.py:362,377,450`).
  When X is blank the code **falls back to the Employee Number** (`transform.py:378-380`),
  which BR-01 forbids.
- Decision owner: Jake (Equus) for the field semantics; Alka/John for which Topia column.
  Until resolved: keep col X as the identifier (unique per assignment, matches Q17/Q18 intent),
  remove the fallback, hold rows with no identifier, and flag the example conflict as a blocker.

### D2 — Column A header: "A" vs "Employee Command"
- Bundled template `lambda/src/templates/equus_eafh_template.xlsx` A3 = `A`; the Equus sample
  `DATA MIGRATION SE EAFH TEST FILE.xlsx` A3 = `A`; Jake's example A3 = **`Employee Command`**.
  All other 145 headers (rows 1–3) are identical across the three files.
- Code writes the header literally from the template and addresses the column as `"A"`
  (`mapping_config.py:39`).
- Q5 (J6): "if there are issues in the header row, the full import fails." A wrong A3 header is
  therefore a whole-file failure risk. Owner: Jake. Evidence needed: the current official
  template version. Until then the output must carry a header-version check and a warning.

### D3 — Dependents have an Equus layout but still no keyed Topia source (Q11)
- Jake confirms dependents go in the same workbook, linked by Assignment Identifier (J12, J19,
  J20) — the INSERT-then-UPDATE pattern the code already produces (`transform.py:438-454`).
- The Topia extract's "Dependent Import" sheet (A1:R4) has Child 1 / Child 2 columns and **no
  Employee or Assignment key**; `transform.py` never reads it (`SOURCE_SHEET` only). Dependents
  from the real extract are therefore not migrated today. Owner: Alka/Topia for a keyed extract.

### D4 — Rules in the code that no one has approved (must not reach a UAT-ready file)
| Rule in code | Where | Question | Status |
|---|---|---|---|
| INTERNATIONAL_RELO = Yes/No derived from home ≠ host country | `transform.py:387-389` | Q40 | OPEN |
| Non-empty Cancellation Reason ⇒ PHASE = "Cancelled" | `mapping_config.py:118-120`, `transform.py:385-386` | Q34 | OPEN (source note from SE, unapproved) |
| Salary / family-size effective dates, address BEGIN_DATE, contact START_DATE default to assignment start | `mapping_config.py:145`, `transform.py:390-396,409,427` | Q22, Q23 | OPEN |
| `employee→Employee`, `employed→Active` value maps | `mapping_config.py:156-159` | Q35 | OPEN (yellow) |
| Country names → ISO-3 from a hand-written table | `mapping_config.py:190-243`, `transform.py:292-300` | Q12, Q3 | PARTIALLY (rule confirmed, list pending) |
| Host Job Level → TO_SALARY_GRADE (overrides source row 2 TO_TITLE) | `mapping_config.py:79` | Q32 | OPEN |
| Assignment Identifier falls back to Employee Number | `transform.py:378-380` | Q1 | forbidden by BR-01 |
| `STRICT_PENDING` defaults to false, so yellow columns with data still produce an import file | `handler.py:158` | §5 | must not yield "UAT-ready" output |

### D5 — Confidential and custom attributes are dropped, not held (Q6, Q7, Q9)
- AG Confidentiality Type and DG Expected Assignment Outcome hold data in both rows and are
  flagged MISSING_TARGET and **not written** (`transform.py:245-246`). ED PII Nationality 1 is
  excluded by `FIELD_OVERRIDES` (`mapping_config.py:98`). BR-06/BR-09 require these to be
  reported as held exceptions with an owner, not silently omitted.

## 2. Q1–Q20: Jake's answers, classification, code impact

| Q | Jake's answer (J column, condensed) | Classification | Code today | Action |
|---|---|---|---|---|
| Q1 | Employee Number = Talent Link ID; Assignment Number = "NEW"; Topia ID → Assignment Identifier and CF_TEXT029; Identifier links rows | CONFIRMED BY JAKE, with conflict D1 | Talent Link ID via `KEY_ALIASES` (`mapping_config.py:61`); NEW (`:49`, `transform.py:446`); identifier from col X with fallback | Remove fallback; hold missing identifiers; keep X; escalate D1 |
| Q2 | "I will provide the values … with the other lookup fields" | PENDING EQUUS CONFIGURATION | `RESOURCE_MAP = {}`; names passed through, warned (`transform.py:316-320`) | Versioned reference-data file; unknown names = exception |
| Q3 | Same as Q2 | PENDING EQUUS CONFIGURATION | Q–T shift warned (`transform.py:161-169`); CF_TEXT016 unmapped | No change; keep warning; block G2 |
| Q4 | Same as Q2 | PENDING EQUUS CONFIGURATION | FROM/TO_CURRENCY blank (no source column) | No derivation; blank + exception |
| Q5 | Duplicates can create duplicates; protections not a fail-safe; audit report; row-level rejection; header issues fail the whole file; test first | CONFIRMED BY JAKE (behaviour); UAT VERIFIED still required | S3 idempotency only (`handler.py:144-170`); no source duplicate detection, header check or row classification | Implement BR-05 controls |
| Q6 | Pending with Equus; column to be added to template | PENDING EQUUS CONFIGURATION | AG dropped (MISSING_TARGET) | Report as HELD; never silently omit |
| Q7 | Pending Schneider; SE-424; lookup values and "differs from reason?" needed | PENDING BUSINESS APPROVAL | DG dropped (MISSING_TARGET) | Report as HELD |
| Q8 | Brown has no meaning | CONFIRMED BY JAKE | brown ⇒ `unknown:FFB45F06` ⇒ UNKNOWN_COLOUR (`transform.py:144-158,247`) | Config: FFB45F06 = "none (Q8)"; spouse columns map by row-2 key as now |
| Q9 | John confirmed with Anand not needed; Schneider must decide | PENDING BUSINESS APPROVAL | Excluded (`mapping_config.py:98`) | Keep excluded but list as pending decision, not permanent |
| Q10 | Spouse/Partner can be imported without names | CONFIRMED BY JAKE | Warned "Equus may reject…" (`transform.py:335-339`) | Reword to data-quality note; never block; "Partner" value still needs Q14 list |
| Q11 | Latest template includes dependents; same import; Assignment Identifier links | PARTIALLY CLARIFIED (D3) | Dependent columns handled if present in main sheet; "Dependent Import" sheet unread | Config for a keyed dependent source; hold when no key |
| Q12 | Country and State must match system values (to be provided); City free text | PARTIALLY CLARIFIED | ISO-3 table (unapproved list); state passed through; addresses free text → LINE_1-3 | Treat ISO-3 table as provisional; validate state/country against supplied lists when present |
| Q13 | No email contact type; USER_ID_EMAIL primary, CF_TEXT008 secondary | CONFIRMED BY JAKE | Source G/H map by row-2 key to exactly those headers; only phone contacts created (`mapping_config.py:113-116`) | Compliant; add test |
| Q14 | Yes, will provide | PENDING EQUUS CONFIGURATION | `equus_lists.json` hook (`mapping_config.py:184`) | Versioned reference data; unknown = exception |
| Q15 | Same as Q2 | PENDING EQUUS CONFIGURATION | Code/name pairs warned (`mapping_config.py:171-177`) | No manufactured names |
| Q16 | Same as Q2 | PENDING BUSINESS APPROVAL (source owner) | 158 kept + warned (`mapping_config.py:164-167`) | Compliant; keep |
| Q17 | NEW + INSERT creates; Employee Number links assignments; Identifier links rows | CONFIRMED BY JAKE | INSERT first / IGNORE repeats per employee (`transform.py:444`) | Compliant; add multi-assignment tests |
| Q18 | Spouse/dependents belong to the Assignment; each row carries Assignment Identifier | CONFIRMED BY JAKE | Every row carries identifier (`transform.py:450`) | Compliant; test cross-assignment isolation |
| Q19 | See ` Mny dependents Ex`; as many dependents as needed with the Identifier | CONFIRMED BY JAKE | `n_rows = max(...)` unbounded (`transform.py:440`) | Validate output against the example (commands, identifier, blank Employee Number on UPDATE rows) |
| Q20 | Not recommended, but one dependent field creates a dependent; lookups to follow | CONFIRMED BY JAKE (technical) / PENDING EQUUS CONFIGURATION (lookups) | Records with any value kept; name warning | Keep; distinguish permissibility from data quality |

## 3. Q21–Q43: unanswered, impact on the UAT dataset

| Q | Field / behaviour | Code location | Affects UAT data? | Unapproved assumption today | Severity | Owner | Closure evidence |
|---|---|---|---|---|---|---|---|
| Q21 | Home/host annualized pay → SALARY_AMOUNT/CURRENCY; host pay → OFFER_SALARY | `mapping_config.py:86-87`; FX/FY unmapped (not in template) | Columns empty in sample; yes for full extract | Home pay mapped via overrides; host pay silently unmapped | High | Jake + Schneider comp | Approved header crosswalk |
| Q22 | Blank salary data | `transform.py:392-393` default date | Yes | Effective date defaulted to start | High | Jake | Written rule |
| Q23 | Start/Actual End Date semantics | AO/AP row-2 keys (yellow) | Yes (both rows) | AP = ACTUAL_END_DATE for every status | High | Alka/John | Planned-vs-actual rule |
| Q24 | Date formats | `transform.py:23,254-267,478-479` | Yes | US format output; multiple parse patterns incl. `%m/%d/%y` | Medium | Jake | Accepted formats per field |
| Q25 | Source inventory / cutoff | n/a | Yes (2-row sample only) | None | High | Alka/Schneider | Final extract, hash, counts |
| Q26 | INSERT/UPDATE/IGNORE semantics | `transform.py:443-449` | Yes | IGNORE for employee on repeat rows | High | Jake | Action matrix; UAT test |
| Q27 | Reconciliation criteria | `transform.py:487-532` | Yes | Count equality on INSERT rows only | High | Schneider + Jake | Agreed criteria |
| Q28 | Confidential visibility | none | Yes (AG populated) | None (dropped) | High | Schneider security | Security-group process |
| Q29 | Cyan custom-field inventory | `FIELD_OVERRIDES` | Yes | CF numbers taken from template labels | Medium | Jake/Alka | Signed-off inventory |
| Q30 | Scope & exclusions | red rule (`transform.py:188-191`) | Yes | Red = excluded (Alka legend) | Medium | Schneider | Written disposition |
| Q31 | Approval authority | n/a | Gate G12 | None | High | Schneider PMO | Named signatory |
| Q32 | Job grades / Host Job Level header | `mapping_config.py:79` | Yes (9 / 8) | Override to TO_SALARY_GRADE | Medium | Jake/Alka | Approved correction |
| Q33 | Host Employment Type → CF_TEXT026 | `mapping_config.py:78` | No data in sample | Target from template label | Low | Renniece/Cathy | Confirmation |
| Q34 | Status / Phase / Cancellation precedence | `transform.py:385-386`; AB unmapped | Yes (both Cancelled) | Cancellation Reason ⇒ Cancelled | High | Jake + Schneider | Precedence rule |
| Q35 | Person Status / Policy crosswalk | `mapping_config.py:156-159`; AA passthrough | Yes | employee→Employee; "One-Way Relocation" sent as-is | High | Jake | Lookup lists |
| Q36 | Pay installments | `mapping_config.py` via row-2 keys (yellow) | Yes (12 / 13.33) | Pass-through | Medium | Schneider payroll | Meaning + values |
| Q37 | Family size | overrides `:93-95`, default date `:394-395` | No data | Effective date default | Medium | Schneider | Counting rule |
| Q38 | Default currency | row-2 key GO | No data | None | Low | Jake | Rule |
| Q39 | RMC | row-2 key GR | No data | None | Low | Schneider | Definition |
| Q40 | International Assignment | `transform.py:387-389` | Yes | Derived from countries | High | Schneider | Approved rule |
| Q41 | Resource e-mails | EJ/EL/EN/EP/ER/FE unmapped | Yes (populated) | Dropped | Medium | Jake/Schneider | Disposition |
| Q42 | CF_TEXT024 code vs name | `CODE_NAME_PAIRS` warning | Yes | Warned only | Low | Alka | Lookup |
| Q43 | Placeholder normalisation | `mapping_config.py:148-152` | Yes (N6) | "Not Provided*" ⇒ blank | Medium | Alka | Pattern list per field |

## 4. What is safely implementable now (confirmed) vs. what must stay gated
Implement from CONFIRMED answers: identifier no-fallback + hold (Q1), duplicate/header/row
controls (Q5), brown = no meaning (Q8), nameless spouse/dependent = data-quality not rejection
(Q10, Q20), dependent row validation against Jake's example (Q11, Q18, Q19), e-mail mapping
test (Q13), multi-assignment tests (Q17).
Gate behind explicit approval status in configuration (default OFF for UAT-ready output):
every row of D4. Reference lists (Q2–Q4, Q12, Q14, Q15, Q20): load from a versioned file with
approval metadata; absence = exception, never a guess.

## 5. UAT readiness gates as of this assessment
G1 Blocked (Q2–Q4, Q6, Q7, Q9, Q14, Q15, Q21–Q43 open) · G2 Blocked (lookups and template
fields pending) · G3 Blocked (D1, D2) · G4 Pass for the existing 24 tests only · G5 Fail (child
rows not reconciled) · G6–G8 Not executed (no Equus access from this environment) · G9 Blocked
(Q6, Q28) · G10 Fail (D4 unapproved rules active) · G11–G12 Not obtained.
**Status: NOT READY FOR UAT.**
