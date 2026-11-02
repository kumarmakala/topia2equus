# Pending Clarification Register

Evidence package item 10 of `clarifications/UAT_MASTER_COMMAND.md` §12. OCNORA task T-1386 (UAT-01).
Prepared 2026-11-02 against base commit `e61d99f`.

Companion documents:
- `docs/uat/RTM_Q1-Q43.md` (RTM): one row per question.
- `docs/uat/CODE_REVIEW_REPORT.md` (CR-xx): code findings.

Cell references follow the conventions in RTM §0:
- `JAKE!Jn` = `clarifications/jake-1.xlsx`, sheet `datamapping clarrification shef`, column J.
- `EX!` = sheet ` Mny dependents Ex`.
- `SRC` = the real Topia extract. Only column letters are cited, never values.
- `EQ` = the Equus sample.
- `TPL` = `lambda/src/templates/equus_eafh_template.xlsx`.

Every item is **Open**. A Jake answer is evidence, not approval. Until an item closes:
- The rule it governs stays OFF for import-ready output.
- The affected field is reported as a held or pending exception carrying its Q number and owner.

**Readiness: NOT READY FOR UAT.** Equus execution evidence is INCOMPLETE / NOT EXECUTED.

---

## 1. Escalations: discrepancies D1–D5

| ID | Discrepancy | Evidence cells / lines | Impact | Owner | Severity | Gate | Needed to close |
|---|---|---|---|---|---|---|---|
| **D1** | **Identifier: Jake's text and his example disagree.** The text says Topia ID goes to both the Assignment Identifier and CF_TEXT029. The example puts the Employee Number value in the Assignment Identifier (E4 = C4) and a different value in CF_TEXT029 "Topia File Number" (Y4 ≠ E4). Values redacted in UAT-08, because the assessment records E4 as a real Employee ID; cell references only. SRC F "Topia ID" is red and empty. SRC X "Assignment ID" (cyan) is populated. SRC E is the Talent Link ID. | `JAKE!J2`; `EX!C4`, `EX!E4`, `EX!Y4` (values as recorded in `JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md:22-23`); SRC columns E, F, X (`…ASSESSMENT.md:26-28`); `mapping_config.py:54`, `:68`; `transform.py:362`, `:377-380` | Assignment identity and child-row linkage for every row (Q17, Q18). If the identifier equals the Employee Number, an employee's assignments collide and their UPDATE rows can attach to the wrong assignment. CF_TEXT029 may hold the wrong legacy id. Code today falls back to the Employee Number (CR-T09), which BR-01 forbids. | Jake/Equus (field semantics) + Alka/Topia (which source column) | Critical | G3, G1, G10 | Written answer: which Topia column feeds Assignment Identifier, and which feeds CF_TEXT029. Corrected example or explicit confirmation |
| **D2** | **Column A header: "A" vs "Employee Command".** TPL A3 and EQ A3 read `A`; Jake's example A3 reads `Employee Command`. The other 145 headers in rows 1–3 match. | TPL `Sheet1!A3`; EQ `A3`; `EX!A3` (`…ASSESSMENT.md:38-40`); `JAKE!J6` ("if there are issues in the header row, the full import fails"); `mapping_config.py:39`; `transform.py:130-141` | A wrong header can make Equus reject the whole file (Q5). No header-version check exists (CR-T16, CR-H03). | Jake/Equus | Critical | G3, G2, G8 | The current official template version, with its A3 header |
| **D3** | **Dependents have an Equus layout but no keyed Topia source.** The extract's "Dependent Import" sheet has Child 1/Child 2 columns and no Employee or Assignment key. The code never reads that sheet. | SRC sheet "Dependent Import" `A1:R4` (`…ASSESSMENT.md:50-52`); `JAKE!J12`, `J19`, `J20`; `mapping_config.py:14`; `transform.py:612-614` | Dependents are **not migrated** from the real extract. Spouse data comes only from main-sheet columns. Q11, Q18–Q20 cannot be shown on real data. | Alka/Topia (keyed extract); Jake/Equus (layout, already given in EX rows 4–7) | Critical | G1, G3, G5 | A dependent extract keyed by employee and assignment, with its column map |
| **D4** | **Unapproved rules are active by default.** | `transform.py:387-389` (Q40); `mapping_config.py:118-120` + `transform.py:385-386` (Q34); `mapping_config.py:145` + `transform.py:390-396`, `:409`, `:427` (Q22, Q23, Q37); `mapping_config.py:156-159` (Q35); `mapping_config.py:190-243` (Q12, Q3); `mapping_config.py:79` (Q32); `transform.py:378-380` (Q1); `handler.py:158` (§5) | Import files carry values no one has approved: phase, international flag, effective and begin dates, status values, salary grade, identifier. Reconciliation counts them as "explained" (CR-M25). | Schneider business (Q22, Q34, Q37, Q40), Jake/Equus (Q35, Q12, Q32), Alka/John (Q23) | Critical | G10, G1 | An approval per rule. Until then each rule is switched OFF for import-ready output and reported |
| **D5** | **Confidential and custom attributes are dropped, not held.** AG Confidentiality Type and DG Expected Assignment Outcome hold data and are flagged MISSING_TARGET but not written. ED PII Nationality 1 is excluded by configuration. | SRC AG, DG, ED; `JAKE!J7` (Q6), `JAKE!J8` (Q7), `JAKE!J10` (Q9); `transform.py:245-246`; `mapping_config.py:98` (`…ASSESSMENT.md:66-70`) | Confidential and business data is silently left out of the migration. This violates BR-06 and BR-09, and the confidentiality gate cannot pass. | Jake/Equus (Q6 template column); Schneider business (Q7 SE-424, Q9, Q28) | Critical | G9, G1, G2 | A target field or an authorised exclusion per column. Meanwhile each column is reported as HELD with its owner |

---

## 2. Jake / Equus

| ID | Q | What exactly is needed | Field / cell it unblocks | Gate |
|---|---|---|---|---|
| PC-J-01 | Q1 (D1) | Which value goes to the Assignment Identifier and which to `ASSIGNMENT_CF!CF_TEXT029`. Reconcile `JAKE!J2` with `EX!E4`/`EX!Y4` | Equus `Assignment Identifier`, `ASSIGNMENT_CF!CF_TEXT029`; SRC X/F/E; `mapping_config.py:54`, `:68` | G3 |
| PC-J-02 | Q5 (D2) | The official current template: is A3 `A` or `Employee Command`? | TPL `Sheet1!A3`; `mapping_config.py:39` | G3 |
| PC-J-03 | Q2 | Resource username list (display name or e-mail → username) | SRC EI, EK, EM, EO, EQ, FD; EQ CC/CH; `RESOURCE_MAP` (`mapping_config.py:188`) | G2 |
| PC-J-04 | Q3 | The meaning of `EMPLOYEE_CF!CF_TEXT016`; corrected labels for template Q–T; approved citizenship and birth-country values | TPL `Q2:T3`; `EMPLOYEE!PRIMARY_CITIZENSHIP`, `SECONDARY_CITIZENSHIP`, `BIRTH_COUNTRY` | G2 |
| PC-J-05 | Q4 | Currency list and the approved source for home, host, salary and default currency | EQ AK/AO (`FROM_CURRENCY`/`TO_CURRENCY`); `ANNUAL_BASE_PAY_CURRENCY` | G2 |
| PC-J-06 | Q6 | A template version with the confidentiality column, and its header | SRC AG → (no Equus column yet) | G2, G9 |
| PC-J-07 | Q12 | Equus country and state lists | `ASSIGNMENT!FROM_COUNTRY`/`TO_COUNTRY`, `ASSIGNMENT_MAILING_ADDRESS!COUNTRY_CODE`, state fields; SRC FJ/FL/FN/FP | G2 |
| PC-J-08 | Q14 | Every dropdown list, as a versioned file. Includes whether "Partner" is valid for the spouse relationship | All row-1 "dropdown" headers; SRC GW; `equus_lists.json` | G2 |
| PC-J-09 | Q15 | Cost-centre code → name list | `ASSIGNMENT_CF!CF_TEXT005/012/023/024/027/028/032/033`; SRC BL | G2 |
| PC-J-10 | Q20 | Dependent lookup values | `ASSIGNMENT_DEPENDENT!*` dropdown fields | G2 |
| PC-J-11 | Q24 | Accepted date formats per field, and a rule for ambiguous dates | Every `*_DATE` header; `transform.py:22-23` | G1 |
| PC-J-12 | Q26 | Action matrix: Employee Command and Assignment Command per row type (new employee, existing employee, repeat assignment, child UPDATE row) | Equus column A, `Assignment Command`; `transform.py:443-449` | G3, G8 |
| PC-J-13 | Q29 | Signed-off custom-field inventory, cyan column → CF field (with Alka) | Every `*_CF!*` target in `mapping_config.py:67-95` | G1, G2 |
| PC-J-14 | Q35 | Person Status / Employment Status / Relocation Policy crosswalk and lists | SRC S, AA; `EMPLOYEE!TYPE`, `EMPLOYEE!EMPLOYMENT_STATUS`, `ASSIGNMENT!RELOCATION_POLICY_ID` | G2, G10 |
| PC-J-15 | **new** (no Q) | Whether every address and contact record may be flagged primary, or which one should be (code sets "Yes" on all: CR-T17) | `ASSIGNMENT_MAILING_ADDRESS!PRIMARY_ADDRESS`, `ASSIGNMENT_EMPLOYEE_CONTACT!PRIMARY_CONTACT` | G1 |
| PC-J-16 | Q38 | Rule and source for the default currency | SRC GO → `ASSIGNMENT!DEFAULT_CURRENCY` | G1 |
| PC-J-17 | Q5, Q19 | A controlled trial import in an authorised Equus UAT tenant: row-level vs file-level rejection, duplicate behaviour, and the INSERT/UPDATE dependent rows. Needs access this environment does not have | G6–G8 evidence | G6, G7, G8 |
| PC-J-18 | Q19 | Confirm the UPDATE-row shape: column A value, blank Employee Number, identifier only. Compare with `EX` rows 5–7 | `transform.py:447-449` | G3 |
| PC-J-19 | Q32 | Written correction that Host Job Level maps to `TO_SALARY_GRADE`, not to the row-2 `TO_TITLE` (with Alka) | SRC Host Job Level column row 2; `mapping_config.py:79` | G1, G10 |

## 3. Schneider business

| ID | Q | What exactly is needed | Field / cell it unblocks | Gate |
|---|---|---|---|---|
| PC-S-01 | Q7 | SE-424 decision: target field, lookup values, and whether it differs from reason or business justification | SRC DG → (no Equus column yet) | G1, G2 |
| PC-S-02 | Q9 | Decision to include, exclude or remap PII Nationality 1 (Jake reports John and Anand said it is not needed; Schneider must decide) | SRC ED; `mapping_config.py:98` | G1, G9 |
| PC-S-03 | Q16 | Source-owner confirmation or correction of host weekly hours = 158 | SRC CU6 → `ASSIGNMENT_CF!CF_NUMBER003` | G1 |
| PC-S-04 | Q21 | Compensation header crosswalk, including host pay (no target today) (with Jake) | SRC FX/FY; `ASSIGNMENT_COMPENSATION_SALARY!*`; `mapping_config.py:86-92` | G1, G10 |
| PC-S-05 | Q22 | Salary effective-date rule when the source has none | `ASSIGNMENT_COMPENSATION_SALARY!EFFECTIVE_DATE` | G10 |
| PC-S-06 | Q25 | The final extract, identified by hash and cutoff date, with expected employee and assignment counts (with Alka) | Whole source; Reconciliation sheet counts | G5 |
| PC-S-07 | Q27 | Reconciliation acceptance criteria covering child rows and value comparison (with Jake) | `transform.py:487-532` | G5 |
| PC-S-08 | Q28 | Visibility process for confidential records (security groups) | SRC AG; Equus security configuration | G9 |
| PC-S-09 | Q30 | Written disposition of every red (excluded) column | All red SRC columns (74 in the fixture layout) | G1 |
| PC-S-10 | Q31 | Named UAT signatory and the approval process | Gates G11, G12 | G11, G12 |
| PC-S-11 | Q33 | Confirm Host Employment Type → CF_TEXT026 (Renniece/Cathy per ASSESS) | `ASSIGNMENT_CF!CF_TEXT026` | G1 |
| PC-S-12 | Q34 | Precedence between Status Category, Assignment Status and Cancellation Reason for PHASE (with Jake) | SRC AC, AB and the Cancellation Reason column; `ASSIGNMENT!PHASE` | G1, G10 |
| PC-S-13 | Q36 | Meaning and allowed values of pay installments (payroll) | SRC AK/AL → `ASSIGNMENT_CF!CF_NUMBER001/002` | G1 |
| PC-S-14 | Q37 | Family-size counting rule and effective-date rule | SRC GP/GQ → `ASSIGNMENT_FAMILY_SIZE!*` | G1, G10 |
| PC-S-15 | Q39 | Definition and source of "RMC they use" | SRC GR → `UT_CALCULATION_INPUT!RMC` | G1 |
| PC-S-16 | Q40 | International-assignment rule, or approval to leave it blank | `ASSIGNMENT!INTERNATIONAL_RELO` | G1, G10 |
| PC-S-17 | Q41 | Disposition of the resource e-mail columns (with Jake) | SRC EJ, EL, EN, EP, ER, FE | G1 |
| PC-S-18 | Q5, Q8, Q10, Q13, Q17, Q18, Q19 | Formal approval of the answers Jake confirmed. The tracker Status (`JAKE!I2:I44`) still reads Open | Rules recorded as CONFIRMED BY JAKE in the RTM | G1, G11 |

## 4. Alka / Topia

| ID | Q | What exactly is needed | Field / cell it unblocks | Gate |
|---|---|---|---|---|
| PC-A-01 | Q11 (D3) | Dependent extract keyed by employee and assignment, with its column layout | SRC "Dependent Import" `A1:R4`; `ASSIGNMENT_DEPENDENT!*` | G1, G3, G5 |
| PC-A-02 | Q1 (D1) | Which Topia column is "Topia ID" in Jake's answer: E (Talent Link / Employee ID), F (Topia ID, red, empty) or X (Assignment ID, cyan) | SRC E, F, X | G3 |
| PC-A-03 | Q23 | Planned vs actual end-date semantics for historical records; begin-date rule for addresses and contacts (with John) | SRC AO/AP → `ASSIGNMENT!START_DATE`/`ACTUAL_END_DATE`; `*!BEGIN_DATE`, `*!START_DATE` | G1, G10 |
| PC-A-04 | Q42 | Lookup that turns the host cost-centre code into its name | SRC BL → `ASSIGNMENT_CF!CF_TEXT024` | G2 |
| PC-A-05 | Q43 | Placeholder (missing-value) pattern list, per field | SRC N (glued placeholder in N6) and all fields; `mapping_config.py:148-152` | G1 |
| PC-A-06 | Q25 | Source inventory and cutoff (shared with PC-S-06) | Whole source | G5 |

## 5. Administrative and repository items

| ID | Item | Owner | Needed |
|---|---|---|---|
| PC-ADM-01 | The column-J texts in the RTM are condensed transcriptions from `JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md`. The runner sandbox blocked reading `jake-1.xlsx` in this job. **UAT-08 update:** `J2:J44` and `I2:I44` were read with openpyxl and added verbatim in `docs/uat/RTM_Q1-Q43.md` §4. The row rule was confirmed from column A. The matrix rows keep their condensed text, and §4 is authoritative. | Next UAT task with workbook access, or the reviewer | Read `JAKE!J2:J44` and `I2:I44` with openpyxl; replace each "[condensed per ASSESS]" quote with the exact text; confirm the row rule Qn → row n+1 |
| PC-ADM-03 | **Added in UAT-08 (SEC-01):** the tracked root file `FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx` is byte-identical (SHA-256 `006ba3b5…ac18b8`) to the real extract, and `DATA MIGRATION SE EAFH TEST FILE.xlsx` is byte-identical to the Equus sample (`docs/uat/DEFECT_EXCEPTION_REGISTER.md` §2) | Repository owner + Pratyush (data storage) | Authorised removal from the index and the git history |
| PC-ADM-02 | CR-R01: workbooks named like the real extract and sample, plus `local_s3` outputs, are tracked in git (commit `ade394f`) | Repository owner + Pratyush (data storage) | Check their contents. If they hold personal data, an authorised history clean-up |

## 6. Counts
- Escalations: 5 (D1–D5).
- Clarifications: 43 in total.
  - Jake/Equus: 19.
  - Schneider business: 18.
  - Alka/Topia: 6.
- Administrative items: 3. PC-ADM-03 was added in UAT-08.
- All are Open. None is approved.
