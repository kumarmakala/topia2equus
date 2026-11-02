# Code Review Report: mapping_config.py, transform.py, handler.py

Evidence package item 2 of `clarifications/UAT_MASTER_COMMAND.md` §12; the mandatory review
is defined in MC §5. OCNORA task T-1386 (UAT-01).
- Reviewed at base commit `e61d99f`. This task is a review only: no code or tests were changed.
- Line numbers refer to that commit. Later tasks that change code must re-anchor any finding they touch.

**Readiness: NOT READY FOR UAT.** Equus execution evidence is INCOMPLETE / NOT EXECUTED.

## Approach
- Every line of the three files was read.
- Each finding is tagged with one category:
  - **EB** (existing behaviour, recorded as a baseline)
  - **UA** (unsupported assumption: a business rule without approved evidence)
  - **DEF** (defect: violates a stated requirement or loses or corrupts data)
  - **IV** (incomplete validation)
  - **BLK** (UAT blocker)
- Each finding carries the Q number it depends on and an **Approved?** value. No Jake answer is
  formal approval (every tracker status `jake-1.xlsx!I2:I44` reads Open, per
  `JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md:15`), so **no rule below is approved**. "Jake-confirmed"
  marks rules Jake's text supports that still lack formal approval.
- Recommendations describe work for later UAT tasks. None are implemented here.

Severity scale: **Critical** (identity, data loss or confidentiality; blocks UAT) · **High** ·
**Medium** · **Low**.

---

## 1. `lambda/src/mapping_config.py`

| ID | Location | Cat. | Finding | Q | Approved? | Sev. | Recommendation |
|---|---|---|---|---|---|---|---|
| CR-M01 | `mapping_config.py:14-19` | EB / UA | The source sheet name and row layout (key row 2, UI row 3, Topia row 4, data from row 5, column A skipped) are hard-coded from one extract. The "Dependent Import" sheet is never read (D3). | Q25, Q11 | No | Medium | Version the source layout; check it against the final extract inventory (Q25) |
| CR-M02 | `mapping_config.py:21-31` | UA | Colour semantics: red = not required, yellow = pending, cyan = custom field, peach = standard. This follows Alka's legend, which has no formal approval. Brown `FFB45F06` is missing, so it is reported as unknown although Jake said brown has no meaning (Q8). | Q30, Q8 | No (Q8 Jake-confirmed) | Medium | Add brown as "no meaning (Q8)" once approved; record the approval reference per colour |
| CR-M03 | `mapping_config.py:39` | BLK | The column A header is assumed to be `"A"`, matching TPL A3 and the EQ sample. Jake's example `' Mny dependents Ex'!A3` reads "Employee Command" (D2). Jake (J6): header issues fail the whole import. | Q5 | No | Critical | Header-version check plus escalation of D2; do not change the header until Jake confirms the official template |
| CR-M04 | `mapping_config.py:47-48` | UA | INSERT for an employee's first assignment in a file and IGNORE for later ones. INSERT is Jake-confirmed (Q17); IGNORE semantics are unanswered (Q26). Employees already in Equus from earlier batches are not considered. | Q17, Q26 | No | High | Action matrix from Jake (Q26); test for "existing employee with a new assignment" (MC §7) |
| CR-M05 | `mapping_config.py:49` | EB | Assignment Number = "NEW". | Q1 | No (Jake-confirmed) | Low | Keep; cite Q1 in configuration metadata |
| CR-M06 | `mapping_config.py:51-54` | BLK | The Assignment Identifier comes from Topia label "assignment id" (SRC column X). Jake's text says "Topia ID"; SRC F "Topia ID" is red and empty; Jake's example disagrees with his text (D1). | Q1 | No | Critical | Keep X (unique per assignment) only as a documented, held decision; escalate D1 |
| CR-M07 | `mapping_config.py:58-62` | UA | Key aliases rename source row-2 keys: `*_COMPANY_LEVELn → *_COMPANY_SEGMENT_LEVEL_n`, `GROUP → GROUP_NAME`, `EMPLOYEE!SYSTEM_NUMBER → Employee Number`. These are developer inferences. Only the Employee Number alias is backed by Jake (Talent Link ID, Q1). | Q29, Q1 | No | Medium | Inventory approval per alias (Q29) |
| CR-M08 | `mapping_config.py:68` | BLK | The same source column X feeds both CF_TEXT029 ("Topia File Number") and the Assignment Identifier. The example has different values in E4 and Y4 (D1). | Q1 | No | Critical | Resolve D1 before any import-ready output |
| CR-M09 | `mapping_config.py:69` | UA | Nationality 3 → `EMPLOYEE_CF!CF_TEXT013`, taken from a template label. | Q3, Q29 | No | Medium | Custom-field inventory (Q29) |
| CR-M10 | `mapping_config.py:70` | UA | Status Category → `ASSIGNMENT!PHASE`. The column is yellow, and the dropdown values are unconfirmed. | Q34 | No | High | Precedence and value list (Q34, Q14) |
| CR-M11 | `mapping_config.py:71-76` | UA | Cost-centre name and code columns → CF_TEXT023/024/027/028/032/033, taken from template labels. | Q15, Q29, Q42 | No | Medium | Inventory plus cost-centre list |
| CR-M12 | `mapping_config.py:77-78` | UA | Host Position → CF_TEXT025; Host Employment Type → CF_TEXT026. | Q29, Q33 | No | Low | Confirmation (Q33 owner) |
| CR-M13 | `mapping_config.py:79` | UA | Host Job Level → `TO_SALARY_GRADE`, overriding the source row-2 key `TO_TITLE` (D4). | Q32 | No | High | Gate behind an approval switch defaulting OFF; written correction from Jake/Alka |
| CR-M14 | `mapping_config.py:80` | UA | Home Scheduled Weekly Hours → CF_NUMBER004. | Q16, Q29 | No | Low | Inventory |
| CR-M15 | `mapping_config.py:81-83` | UA | End of Assignment Reason → CF_TEXT030; Global ID → CF_TEXT014; Pre-Hire ID → CF_TEXT015. | Q29, Q34 | No | Medium | Inventory |
| CR-M16 | `mapping_config.py:84-85` | UA | Business-justification text → `UT_GOVERNANCE_QUESTIONNAIRE2!BUSINESS_JUSTIFICATION`. The source row 2 asks whether Expected Assignment Outcome belongs here (Q7, SE-424); the code keeps them separate. | Q7 | No | Medium | Wait for SE-424 |
| CR-M17 | `mapping_config.py:86-92` | UA | Compensation overrides: home base pay amount and currency, profit-sharing frequencies, benefits deduction, family allowance currency. Host pay (FX/FY) has no target. | Q21, Q4 | No | High | Approved compensation crosswalk (Q21) |
| CR-M18 | `mapping_config.py:93-95` | UA | Host and home family size → `ASSIGNMENT_FAMILY_SIZE!*`. The source note asks whether to calculate it. | Q37 | No | Medium | Counting rule (Q37) |
| CR-M19 | `mapping_config.py:96-98` | DEF | PII Nationality 1 is permanently excluded (`None`). BR-09: do not exclude without an authorised decision; it must be reported as HELD (D5). | Q9 | No | High | Report as a held exception with owner Schneider |
| CR-M20 | `mapping_config.py:101-108` | UA | Address column → LOCATION_TYPE constants (Departure/Destination Residence/Office). | Q12, Q14 | No | Medium | Confirm the values against the Equus list |
| CR-M21 | `mapping_config.py:109-110` | UA | "Location Type - …" columns are ignored as descriptions of constants. | Q30 | No | Low | Disposition (Q30) |
| CR-M22 | `mapping_config.py:112-116` | UA / EB | Phone contacts only: "Employee Preferred"/"Personal Mobile", type "Telephone". No e-mail contact (consistent with Jake's Q13 answer). The location and contact type values are unconfirmed. | Q13, Q14 | No (Q13 Jake-confirmed) | Low | Q14 list; test that no e-mail contact rows are created (Q13) |
| CR-M23 | `mapping_config.py:118-120` | UA | A non-empty Cancellation Reason forces PHASE = "Cancelled" (D4). | Q34 | No | High | Approval switch defaulting OFF |
| CR-M24 | `mapping_config.py:122-132` | UA | Fixed set of ISO-3 country headers, which works around the template Q–T shift. | Q3, Q12 | No | Medium | Keep the warning; confirm after the template fix (Q3) |
| CR-M25 | `mapping_config.py:134-141` | DEF | `RULE_TARGETS` makes reconciliation treat values set by unapproved rules (PHASE, INTERNATIONAL_RELO, salary and family-size effective dates) as "explained", so unapproved rules pass reconciliation. | Q27, Q22, Q34, Q37, Q40 | No | High | Count rule-set values as explained only when the rule is approved |
| CR-M26 | `mapping_config.py:143-145` | UA | `DEFAULT_DATES_FROM_START = True`: salary and family-size effective dates and address/contact begin dates default to the assignment start (D4). MC §13 forbids this without approval. | Q22, Q23, Q37 | No | Critical | Default OFF for import-ready output |
| CR-M27 | `mapping_config.py:147-152` | UA | Placeholder tokens `not provided, n/a, na, null, -` are blanked in **every** field, including codes and identifiers; the glued-prefix rule for "not provided" also applies. | Q43 | No | Medium | Per-field pattern list (Q43); report blanked cells |
| CR-M28 | `mapping_config.py:154-159` | UA | Value maps `employee→Employee`, `user→User`, `employed→Active`. No evidence for any of them, including "user". | Q35 | No | High | Approval switch defaulting OFF; Equus lists |
| CR-M29 | `mapping_config.py:161-167` | EB | Weekly hours plausibility range 0–80, used as a review threshold only (values kept and warned). | Q16 | No (threshold, not a rule) | Low | Keep |
| CR-M30 | `mapping_config.py:169-177` | UA | Code/name pairs, and the code pattern `^[\w-]+_[A-Z]VC$` derived from sample values. | Q15, Q42 | No | Low | Replace with the approved list |
| CR-M31 | `mapping_config.py:179-180` | DEF | The comment "Family records need a name to be imported as a person" contradicts Jake's Q10/Q20 answers. | Q10, Q20 | No | Low | Reword when Q10/Q20 are approved |
| CR-M32 | `mapping_config.py:182-184` | IV | `equus_lists.json` is optional; a missing file silently means "no list check". BR-14: unknown values must be reported. | Q14 | No | High | A missing list for a dropdown header becomes an explicit exception |
| CR-M33 | `mapping_config.py:186-188` | EB | `RESOURCE_MAP = {}`: no usernames are invented (compliant with BR-02), but display names are sent as is. | Q2 | No | High | Hold unmapped resources instead of passing them through |
| CR-M34 | `mapping_config.py:190-243` | UA | A hand-written country-name → ISO-3 table. It includes the non-ISO code `XKX` (Kosovo), abbreviations (`UK`, `US`, `UAE`) and former names. Jake (J13): values must match Equus system values, and the list is pending. | Q12, Q3 | No | High | Treat as provisional; validate against the Equus list once supplied |
| CR-M35 | whole file | BLK | No configuration element carries a version, approval status (Q number and owner) or effective version, and nothing is validated at load time, as the MC architecture requirement demands. | all | n/a | Critical | Versioned JSON configuration with approval metadata, rejected when invalid |

## 2. `lambda/src/transform.py`

| ID | Location | Cat. | Finding | Q | Approved? | Sev. | Recommendation |
|---|---|---|---|---|---|---|---|
| CR-T01 | `transform.py:22-23`, `:262-266` | DEF | Six parse patterns are tried in order; `%m/%d/%Y` comes first, so an ambiguous `03/04/2020` is silently read as US (MC §13: never convert ambiguous dates). `%m/%d/%y` guesses the century. Output is US `MM/DD/YYYY`. | Q24 | No | High | Report ambiguous text dates; approved formats per field |
| CR-T02 | `transform.py:44-49` | UA | `is_blank` applies BLANK_TOKENS to every field (see CR-M27). | Q43 | No | Medium | As CR-M27 |
| CR-T03 | `transform.py:83-84` | IV | A field counts as a date if its rule says "date format" **or** its header ends with `_DATE`. This is a heuristic. | Q24 | No | Low | Explicit date-field list in configuration |
| CR-T04 | `transform.py:89-90` | UA | Any header containing "RESOURCE" is treated as a username field. | Q2 | No | Low | Explicit list |
| CR-T05 | `transform.py:144-158` | UA | Colour precedence red > yellow > cyan over rows 1–4. Any unknown fill (brown, Q8) is flagged UNKNOWN_COLOUR only when no known colour is present. | Q30, Q8 | No | Medium | Q8 configuration entry |
| CR-T06 | `transform.py:188-191` | EB | Red columns are excluded before every other rule (Alka legend). | Q30 | No | Medium | Disposition per column (Q30) |
| CR-T07 | `transform.py:241-248` | BLK | Yellow (pending) columns are still mapped and written: they are only flagged PENDING (D4, MC §5). Cyan columns with no target (AG, DG) are flagged MISSING_TARGET and **not written** (D5). | Q6, Q7, Q23, Q34–Q36 | No | Critical | Pending and held values become exceptions with an owner; import-ready output excludes them |
| CR-T08 | `transform.py:254-267` | DEF | Any int or float in a date column is read as an Excel serial without a plausibility check (an int like 20200101 gives a far-future date). `datetime` values with a time part are returned as datetime, so types are inconsistent. | Q24 | No | Medium | Validate the serial range; report non-date types |
| CR-T09 | `transform.py:376-382` | DEF / BLK | A missing Assignment Identifier **falls back to the Employee Number** (warning only), which BR-01 and MC §13 forbid. A missing Employee Number is only warned, and the row is still written as INSERT. | Q1 | No | Critical | Hold the row (no output) with an exception; never fall back |
| CR-T10 | `transform.py:367-456` | BLK | No source-level duplicate detection: two source rows with the same Assignment ID produce two INSERT rows with one identifier. Jake (J6): duplicates can create duplicates in Equus. | Q5, Q17 | No | Critical | Duplicate detection on employee, identifier and full row; exception report |
| CR-T11 | `transform.py:365`, `:444`, `:455` | DEF | `seen_employees` also records `None`, so a second row with a missing Employee Number gets `IGNORE`, as if it belonged to the first. | Q1, Q26 | No | High | Never put a missing key into identity sets (follows from CR-T09) |
| CR-T12 | `transform.py:385-386` | UA | The cancellation rule **overwrites** whatever Status Category mapped into PHASE (D4). | Q34 | No | High | Switch defaulting OFF; precedence rule (Q34) |
| CR-T13 | `transform.py:387-389` | UA | INTERNATIONAL_RELO is derived as home ≠ host (D4). MC §13 forbids this without a rule. | Q40 | No | High | Switch defaulting OFF |
| CR-T14 | `transform.py:335-339` | DEF | The warning text says "Equus may reject or create an unnamed person". That is an Equus behaviour claim, and it contradicts Jake's Q10/Q20 answers. | Q10, Q20 | No | Low | Data-quality note wording |
| CR-T15 | `transform.py:390-396` | UA | Default effective and begin dates (see CR-M26) (D4). | Q22, Q23, Q37 | No | Critical | Default OFF |
| CR-T16 | `transform.py:130-141`, `:161-169`, `:616-617` | BLK | Template checks: duplicate headers are rejected and country-rule shifts are warned, but there is no check of the header version or of A3 against the official Equus template (D2). | Q5 | No | Critical | Header fingerprint check; fail fast with a report |
| CR-T17 | `transform.py:400-429` | UA | Every address gets `PRIMARY_ADDRESS = "Yes"` and every contact `PRIMARY_CONTACT = "Yes"`, so an assignment can have several "primary" records. Address text is split into LINE_1–3 with no city, postal or country split. Numeric phone values go through `str(int())`, which keeps no leading zero or "+". **No Q covers the primary flags: new clarification PC-J-15.** | Q12, Q13, Q14 (new) | No | Medium | Raise PC-J-15; report rather than assume |
| CR-T18 | `transform.py:493-499`, `:508` | BLK | Reconciliation compares INSERT rows only. Address, contact, spouse and dependent values on UPDATE rows are **not reconciled** (known concern, MC §5). | Q27 | No | Critical | Reconcile child records and expected expansion counts |
| CR-T19 | `transform.py:508-528` | DEF | For non-date values reconciliation checks presence only, never value equality. Value maps, rule overwrites (CR-T12) or corruption all pass as long as a value is present. | Q27 | No | Critical | Compare normalised source and target values under the approved conversion |
| CR-T20 | `transform.py:292-300` | IV | An unknown country name is passed through raw into an ISO-3 column (warning only). Any code in the table is accepted, which MC §13 forbids ("never assume syntactically valid values are accepted Equus values"). | Q12 | No | High | Hold the value with an exception |
| CR-T21 | `transform.py:599-601` | IV | `load_allowed_values` reads JSON with no schema, version or approval validation. A wrong shape is ignored or crashes later. | Q14 | No | Medium | Validate at load time and reject explicitly |
| CR-T22 | `transform.py:314-315`, `:121-125` | IV | Values missing from VALUE_MAPS pass through unchanged and are reported as "NO" only when a list exists (see CR-M32). | Q14, Q35 | No | High | As CR-M32 |
| CR-T23 | `transform.py:344-353`, `:431-436` | EB | Dependent groups are split by column order (a repeated target starts a new child). A record is kept if any value is present, and the count is unbounded. This is consistent with Jake's Q19/Q20 answers but only tested with 2 dependents. | Q19, Q20 | No (Jake-confirmed) | Medium | Test with 3 or more; validate against EX rows 4–7 |
| CR-T24 | `transform.py:441-449` | IV | UPDATE rows carry `A = IGNORE`, Assignment Command `UPDATE` and the identifier, with no Employee Number. This has not been compared with Jake's example rows. | Q19, Q26 | No | Medium | Comparison test against EX rows 4–7 |
| CR-T25 | `transform.py:473-476` | IV | A value with no Equus column is dropped with a warning only, and reconciliation does not see child-row drops (CR-T18). | Q5, Q27 | No | Medium | Count as an exception |
| CR-T26 | `transform.py:604`, `:625-628` | BLK | `strict_pending` defaults to False. Strict mode considers only PENDING and ignores MISSING_TARGET, UNKNOWN_COLOUR and D4 rule usage. When strict mode raises, **no report is built**, so diagnostics are lost (MC §5). | §5 | n/a | Critical | Classify import-ready vs diagnostic output; always build the report |
| CR-T27 | `transform.py:367-370` | IV | Rows whose cells are all blank or placeholder tokens are skipped without a warning, and `reconcile` (`:496-497`) excludes them the same way, so they never appear in counts. | Q43, Q25 | No | Low | Report skipped rows |
| CR-T28 | `transform.py:304-307` | IV | Floats are turned into ints and CF_TEXT values stringified. Employee Number stays numeric. Identifiers get no leading-zero or type check. | Q1 | No | Medium | Identifier format validation |
| CR-T29 | `transform.py:362` | IV | The identifier column is found by label regardless of its governance status, so an excluded or red column would still be read. | Q1 | No | Low | Respect status or fail explicitly |
| CR-T30 | `transform.py:1-2`, `:19` | EB | Pure functions (bytes in, bytes out). Business constants come from `mapping_config`, but rule logic (cancellation, international, default dates, PRIMARY flags) is coded in `build_records` rather than configured. | architecture | n/a | Medium | Move rule switches into configuration |

## 3. `lambda/src/handler.py`

| ID | Location | Cat. | Finding | Q | Sev. | Recommendation |
|---|---|---|---|---|---|---|
| CR-H01 | `handler.py:36`, `:113-128` | IV | The 50 MB size limit applies only when the S3 event carries a size, so direct invokes bypass it. Prefix, suffix, lock-file and own-output guards work (tested). | Q5 | Low | Check the size after the read as well |
| CR-H02 | `handler.py:58-64` | DEF | The code version hashes the template, `transform.py` and `mapping_config.py`, but not `equus_lists.json` (or any future JSON configuration). After a lookup change the output key stays the same, so the duplicate check (`:152`) skips reprocessing. | Q14, architecture | High | Include every configuration file in the version hash |
| CR-H03 | `handler.py:93-98` | IV | The template from `TEMPLATE_S3_URI` is not checked against the expected header version (D2). | Q5 | High | Header fingerprint check (CR-T16) |
| CR-H04 | `handler.py:101-110`, `:144` | IV | A direct invoke without `source_bucket`/`source_key` raises KeyError. Objects are read without the event's VersionId. | Q5 | Low | Explicit validation and error |
| CR-H05 | `handler.py:157-159` | BLK | `STRICT_PENDING` defaults to "false" (D4, MC §5). `PendingFieldsError` or `ValueError` from `transform` propagates without writing **any** report, and there is no file-level vs row-level error classification. | §5, Q5 | Critical | Always write the diagnostic report; mark output status explicitly |
| CR-H06 | `handler.py:160-166` | EB | A reconciliation failure keeps the evidence, withholds the import file and raises, so the event goes to the failure queue. The `_put_once` result is ignored. | Q27 | Low | Keep |
| CR-H07 | `handler.py:167-171` | DEF | The output is written before the report. If the report write fails, a retry finds the output (`:152`) and skips, so the report is never produced. | Q5 | Medium | Write the report first, or check both keys |
| CR-H08 | `handler.py:149-151` | BLK | The output is always named `_EQUUS_`. Metadata carries only source bucket, key and ETag: no mapping version, reconciliation result, pending or held counts, or import-ready status. Diagnostic output cannot be told apart from import-ready output. | §5, G10 | Critical | Status marker in the key and in metadata |
| CR-H09 | `handler.py:173-182`, `:171` | EB / IV | The CloudWatch summary logs only keys and counts, which is compliant. The report workbook's Warnings sheet holds source values quoted in messages (for example `transform.py:283`, `:331`), so the report object holds personal data and relies on bucket access control. | Q28, G9 | Medium | Document access and retention; restrict the report prefix |
| CR-H10 | whole file | EB | The handler never calls Equus: transformation output is separate from the Equus import (MC §5, §9). | — | — | Keep |
| CR-H11 | `handler.py:137-183` | EB | Records in an event are processed one after another. One failure aborts the rest; outputs already written persist, and a retry relies on idempotency (`:152`). | Q5 | Low | Per-record result reporting |

## 4. Repository and data-handling finding (outside the three files; blocks G9)

| ID | Location | Cat. | Finding | Owner | Sev. |
|---|---|---|---|---|---|
| CR-R01 | git index (added in commit `ade394f`) | BLK | Files named like the real extract and Equus sample (`FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx`, `DATA MIGRATION SE EAFH TEST FILE.xlsx`), plus `backup.xlsx`, `Topia_Equus_Field_Mapping.xlsx` and two `local_s3/dest_s3/equus/**` outputs, are **tracked in git** although `.gitignore` excludes `*.xlsx` and `local_s3/`, and `README.md:91` says real workbooks are git-ignored. Their contents were not opened in this job. If they hold personal data, removing them needs a history rewrite, which requires explicit authorisation and was **not** done. | Repository owner + Pratyush (data storage) | Critical |

## 5. AC-4 register: every hard-coded business assumption, with Q number and approval

Status values:
- **No**: no approval, and Jake has given no supporting answer.
- **No: Jake-confirmed**: Jake's text supports the rule, but it is not formally approved (tracker reads Open).
- Rows marked **(D4)** are active by default today.

| # | Assumption | Location | Q | Approved? |
|---|---|---|---|---|
| 1 | Source sheet name and row layout | `mapping_config.py:14-19` | Q25 | No |
| 2 | Colour meanings (red, yellow, cyan, peach); brown unlisted | `mapping_config.py:26-31`; `transform.py:144-158` | Q30, Q8 | No (Q8: No, Jake-confirmed) |
| 3 | Red = excluded | `transform.py:188-191` | Q30 | No |
| 4 | Column A header literal "A" | `mapping_config.py:39` | Q5 (D2) | No |
| 5 | Employee action INSERT first, IGNORE repeat | `mapping_config.py:47-48`; `transform.py:444`, `:448` | Q17, Q26 | No (Q17: No, Jake-confirmed) |
| 6 | Assignment Number "NEW" | `mapping_config.py:49`; `transform.py:446` | Q1 | No (Jake-confirmed) |
| 7 | Assignment Identifier = Topia "Assignment ID" (column X) | `mapping_config.py:54`; `transform.py:362` | Q1 (D1) | No |
| 8 | Identifier falls back to Employee Number (D4) | `transform.py:378-380` | Q1 | No (forbidden by BR-01) |
| 9 | Key aliases (segment level, group name, system number → Employee Number) | `mapping_config.py:58-62` | Q29, Q1 | No |
| 10 | Assignment ID → CF_TEXT029 | `mapping_config.py:68` | Q1 (D1) | No |
| 11 | Nationality 3 → CF_TEXT013 | `mapping_config.py:69` | Q3, Q29 | No |
| 12 | Status Category → PHASE | `mapping_config.py:70` | Q34 | No |
| 13 | Cost-centre fields → CF_TEXT023/024/027/028/032/033 | `mapping_config.py:71-76` | Q15, Q29, Q42 | No |
| 14 | Host position / employment type → CF_TEXT025/026 | `mapping_config.py:77-78` | Q29, Q33 | No |
| 15 | Host Job Level → TO_SALARY_GRADE (D4) | `mapping_config.py:79` | Q32 | No |
| 16 | Home weekly hours → CF_NUMBER004 | `mapping_config.py:80` | Q16, Q29 | No |
| 17 | End reason / Global ID / Pre-hire ID → CF_TEXT030/014/015 | `mapping_config.py:81-83` | Q29, Q34 | No |
| 18 | Business justification target | `mapping_config.py:84-85` | Q7 | No |
| 19 | Compensation overrides | `mapping_config.py:86-92` | Q21, Q4 | No |
| 20 | Family-size targets | `mapping_config.py:93-95` | Q37 | No |
| 21 | PII Nationality 1 excluded | `mapping_config.py:98` | Q9 | No |
| 22 | Address location-type constants | `mapping_config.py:103-108` | Q12, Q14 | No |
| 23 | "Location Type - …" columns ignored | `mapping_config.py:110` | Q30 | No |
| 24 | Phone contact location/contact types; no e-mail contact | `mapping_config.py:113-116` | Q13, Q14 | No (Q13: No, Jake-confirmed) |
| 25 | Cancellation Reason ⇒ PHASE "Cancelled" (D4) | `mapping_config.py:119-120`; `transform.py:385-386` | Q34 | No |
| 26 | ISO-3 country headers | `mapping_config.py:125-132` | Q3, Q12 | No |
| 27 | Rule-set values count as explained in reconciliation | `mapping_config.py:136-141`; `transform.py:521` | Q27 | No |
| 28 | Effective and begin dates default to assignment start (D4) | `mapping_config.py:145`; `transform.py:390-396`, `:409`, `:427` | Q22, Q23, Q37 | No |
| 29 | Placeholder tokens blanked in all fields | `mapping_config.py:148-152`; `transform.py:44-49`, `:270-284` | Q43 | No |
| 30 | Value maps employee/user/employed (D4) | `mapping_config.py:156-159`; `transform.py:314-315` | Q35 | No |
| 31 | Weekly-hours review range 0–80 | `mapping_config.py:164-167`; `transform.py:309-312` | Q16 | No (warning only) |
| 32 | Cost-centre code pattern and pairs | `mapping_config.py:171-177`; `transform.py:328-331` | Q15, Q42 | No |
| 33 | Family records "need a name" (warning) | `mapping_config.py:179-180`; `transform.py:335-339` | Q10, Q20 | No (Jake's answer contradicts it) |
| 34 | Missing dropdown list file = no check | `mapping_config.py:184`; `transform.py:599-601` | Q14 | No |
| 35 | Resource names passed through when unmapped | `mapping_config.py:188`; `transform.py:316-320` | Q2 | No |
| 36 | Country-name → ISO-3 table (incl. XKX, abbreviations) | `mapping_config.py:191-243`; `transform.py:292-300` | Q12, Q3 | No |
| 37 | Date parse patterns (US first) and US output format | `transform.py:22-23`, `:254-267`, `:477-479` | Q24 | No |
| 38 | Numerics in date columns = Excel serials | `transform.py:259-260` | Q24 | No |
| 39 | `_DATE` suffix ⇒ date field | `transform.py:84` | Q24 | No |
| 40 | "RESOURCE" in header ⇒ username field | `transform.py:90` | Q2 | No |
| 41 | INTERNATIONAL_RELO = home ≠ host (D4) | `transform.py:387-389` | Q40 | No |
| 42 | PRIMARY_ADDRESS / PRIMARY_CONTACT = "Yes" for every record | `transform.py:407`, `:428` | none (new: PC-J-15) | No |
| 43 | Free-text address split into LINE_1–3 | `transform.py:405-413` | Q12 | No |
| 44 | Dependent child grouping by repeated target | `transform.py:344-353` | Q19 | No (Jake-confirmed pattern) |
| 45 | UPDATE row shape (A = IGNORE, no Employee Number) | `transform.py:447-449` | Q19, Q26 | No |
| 46 | Yellow columns still written (strict mode off by default) | `transform.py:242-244`, `:604`; `handler.py:158` | §5, Q23, Q34–Q36 | No |

**Result: 46 hard-coded business assumptions. None is formally approved.**
