# Open questions: Topia → Equus EAFH mapping

For Alka (business mapping) and John (Equus). Each question cites the cell
it comes from, what the code does today, and the setting that changes it.
Cell references are to:

- **SRC**: `FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx`, sheet "Full Historical Past Assignment"
- **EQ**: `DATA MIGRATION SE EAFH TEST FILE.xlsx`, the Equus sample

Every run's mapping report lists the same items in its Governance and
Warnings sheets.

**Business rules and readiness (UAT-02):** the value-changing rules below
(Q34 cancellation phase, Q40 international relocation, Q22/Q23 defaulted
dates, Q35 value maps, Q32 host job level, Q12/Q3 country table) are now
declared in `lambda/src/business_rules.json` with their Q number, status
and owner, and are **not applied** to the import file until
`approved_for_import` is set to `true` by the owner. Reference lists
(Q2 resource usernames, Q15 cost centres, Q12 country/state) live in
`lambda/src/reference_data/` as empty placeholders. Each run's Rules and
Readiness sheets list every gated rule, provisional conversion and pending
or held item; the import file is named `*_EQUUS_DRAFT_*` while readiness is
`NOT_READY`. Where a section below says "Code today" for one of these
rules, read it as "behaviour once approved". See
`docs/uat/UAT-02_business_rules_configuration.md`.

## Identity

### 1. Which value is the Assignment Identifier?
- **Evidence:**
  - EQ E2 labels the column "Topia Id", but EQ E4 holds the same value as the Employee Number (EQ C4 = SRC E5).
  - SRC X (Assignment ID, cyan) holds a different, per-assignment number. Its row-1 note says "keep this (Topia ID) as a legacy/Data Migration field".
  - SRC F (Topia ID) is red, "not required".
  - Jake said verbally: Topia ID → Assignment Identifier and Assignment Number. This hasn't been confirmed in writing.
- **Options:** Employee Number (matches the EQ sample); Topia Assignment ID (unique per assignment, so one employee can have several historical assignments).
- **Code today:** Topia Assignment ID (SRC X), falling back to the Employee Number when it's blank. The same value also goes to `ASSIGNMENT_CF!CF_TEXT029` (Topia File Number).
- **Setting:** `ASSIGNMENT_IDENTIFIER_LABEL` in `lambda/src/mapping_config.py`.
- **Why it matters:** the Employee Number repeats across one person's assignments. If it's used as the identifier, their UPDATE rows can attach to the wrong assignment.

### 2. Resource usernames
- **Evidence:** SRC EI, EK, EM, EO, EQ and FD hold display names in "Last, First" form. EQ CC4/CH4 hold usernames (`Hugh.HRBP`, `JJones`, the sample's test values). Row 1 says "Match Resource dropdown".
- **Code today:** names are passed through unchanged, with one warning per distinct name.
- **Setting:** `RESOURCE_MAP` (name or e-mail → Equus username). SRC EJ, EL, EN, EP, ER and FE hold e-mail addresses, which are probably the safer lookup key.
- **Needed:** Equus's resource username list.

## Template (Equus side)

### 3. Template columns Q–T: labels shifted against field keys
- **Evidence:** EQ row 2 reads Primary Citizenship | Secondary Citizenship | Country of Birth | PII Nationality, over the row-3 keys `EMPLOYEE_CF!CF_TEXT016` | `PRIMARY_CITIZENSHIP` | `SECONDARY_CITIZENSHIP` | `BIRTH_COUNTRY`. The row-1 "Match Country Code" rules follow row 2, not row 3.
- **Code today:** country fields are listed explicitly (`COUNTRY_HEADERS`), so Country of Birth is converted to ISO-3. Each run warns about the shift. Nothing is mapped to `CF_TEXT016`.
- **Question for John:** what is `EMPLOYEE_CF!CF_TEXT016`? Is it the PII Nationality field (see Q8)? Should the template labels be corrected?

### 4. Home/Host currency
- **Evidence:** EQ AK/AO (`FROM_CURRENCY`/`TO_CURRENCY`, "Match Currency") hold EGP/INR in the sample. No SRC column carries a currency.
- **Code today:** left blank.
- **Options:** derive it from the home/host country (needs an agreed country → currency table); take it from another export; or leave it blank.

### 5. Does Equus reject a whole file or single rows?
- **Evidence:** the Equus help centre is login-only. Its public summary says imports do not delete and many cannot be undone.
- **Question for John:** what does the import do when a row fails? Can a trial load run in a non-production tenant first?

## Colour-coded items

### 6. Yellow: pending clarification
All of these are mapped today and flagged PENDING. Set `STRICT_PENDING=true` to block output until they're resolved.

| SRC col | Topia field | Mapped to | Data present? | Open point |
|---|---|---|---|---|
| S | Person Status | `EMPLOYEE!TYPE` | yes | Is the value mapping `employee → Employee` correct? |
| AA | Assignment Type Category | `ASSIGNMENT!RELOCATION_POLICY_ID` | yes | SRC says "One-Way Relocation", EQ sample says "Localization". Which dropdown value? |
| AB | Assignment Status | *not mapped* | yes | Row 2: "Calculated Value"; row 1: "Calculated field based on Date Entered". Confirm Equus calculates it. |
| AC | Status Category | `ASSIGNMENT!PHASE` | yes | Confirm the mapping and its dropdown values. |
| AK / AL | Home/Host Pay Installments | `ASSIGNMENT_CF!CF_NUMBER001/002` | yes | Confirm. One source value is 13.33. |
| AO / AP | Assignment Start/End Date | `START_DATE` / `ACTUAL_END_DATE` | yes | Confirm that ACTUAL (not estimated) end date is right for historical records. |
| GO | Default Currency | `ASSIGNMENT!DEFAULT_CURRENCY` | no | Source of the data? |
| GP / GQ | Host/Home Family Size | `ASSIGNMENT_FAMILY_SIZE!*` | no | Row 3 asks whether to calculate it from spouse/dependents at host. The effective date defaults to the start date. |
| GR | RMC they use | `UT_CALCULATION_INPUT!RMC` | no | Source of the data? |

### 7. Cyan custom fields with no Equus column
- **AG Confidentiality Type:** row 2 says AG and AH are combined into one text field (default "Non-Confidential"), and row 3 mentions a separate security-group import. The EQ template has no such column.
- **DG Expected Assignment Outcome:** row 2 asks whether this is "reason" or belongs on the business justification tab. The EQ template has no such column.
- **Code today:** both are flagged MISSING_TARGET and not migrated, although both columns hold data.
- **Needed:** the Equus custom field and header, or a decision to drop them.

### 8. PII Nationality 1 (SRC ED, cyan)
- **Evidence:** the row-2 note says "Need to confirm with Anand".
- **Code today:** excluded (`FIELD_OVERRIDES["pii nationality 1"] = None`).

### 9. Brown fill on spouse columns (SRC GS–GX, row 4, `FFB45F06`)
- **Evidence:** the colour isn't covered by the agreed colour rules.
- **Answered (Q8, Jake):** brown has no business meaning.
- **Code today:** mapped by field key; colour reported as `none (Q8: no meaning)`, no longer UNKNOWN_COLOUR.

## Data shape

### 10. Spouse and dependents
- **Spouse:** SRC GW5/GW6 hold a relationship but no spouse name. **Answered (Q10, Jake):** spouse/partner records can be imported without names. **Code today:** written as sent with a DATA_QUALITY note ("importable per Q10/Q20, business review recommended"), never held. Still open: is "Partner" a valid Equus value? (Q14 list from Jake; reported "not in list" until then.)
- **Dependents:** **Answered (Q11/Q18/Q19/Q20, Jake):** same import; every dependent row carries the Assignment Identifier; any number of dependents on UPDATE rows (example sheet ` Mny dependents Ex`); one populated field creates a record. **Code today:** the "Dependent Import" sheet has no Employee/Assignment key, so it is reported HELD_DEPENDENTS_UNKEYED with its row counts and nothing is attached. A keyed sheet can be configured in `business_rules.json` `$.dependent_source` (joined by Assignment ID only). Still open: a keyed dependent extract from Topia (Alka, D3). See docs/uat/UAT-04_family_rules.md.

### 11. Addresses and contacts
- **Addresses:** SRC FJ/FL/FN/FP are free text. The code puts lines into LINE_1–3 and warns, but Equus wants CITY, COUNTRY_CODE and POSTAL_CODE separately (EQ CU–CX). What format does the export use?
- **Contacts:** EQ CZ4 shows a contact row of type "Email". The code only writes phone contacts (SRC FR/FS). Should an e-mail contact row be created? From which column?

### 12. Values to confirm with the business
- **CU6** (host weekly hours) = 158. It's kept and warned (review threshold 80, set in `NUMERIC_RANGES`).
- **BL5/BL6** (host cost-centre name) holds the code `1000-US018907_NVC`. Row 3 says names should come from a new cost-centre list, so the cost-centre code → name list is needed.
- **Dropdown lists:** drop Equus's dropdown lists in as `lambda/src/equus_lists.json` (`{"EQUUS!HEADER": ["Value", ...]}`). The report's "Allowed?" column will then check every value sent.
