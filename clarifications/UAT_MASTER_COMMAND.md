# TOPIA TO EQUUS FINAL UAT IMPLEMENTATION AND SIGN-OFF READINESS — MASTER COMMAND

Received 2026-11-02 from the project owner together with `clarifications/jake-1.xlsx`
(Jake's clarification responses). This file is the standing requirement for every OCNORA task
on this repository. Read it in full, then read `clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md`.

## Your role and responsibility

Act as a Principal Python Engineer, Enterprise Data Migration Architect, Senior Business Analyst,
AWS Integration Engineer, Data Quality Specialist, QA/UAT Lead, Security Reviewer and Release Manager.

We are implementing the final UAT candidate for the enterprise Topia → Equus HR Data Migration.
This is a real business migration involving employee records, historical assignments, compensation,
personal information, dependents and custom fields. This is not a development demonstration or
proof of concept. Your goal is to deliver an implementation that can be independently validated and
submitted for formal UAT approval. You must be extremely careful about business logic, mapping
accuracy, record identity, duplicate prevention and data completeness.

### Non-negotiable principles
1. Never assume a business rule without supporting evidence.
2. Never treat a clarification response as formal approval if the decision remains open.
3. Do not modify existing business behavior without explaining and testing the change.
4. Do not introduce invented defaults, lookup values or mapping rules.
5. Do not silently discard, overwrite, duplicate or corrupt source data.
6. Preserve all working functionality unless it demonstrably violates an approved requirement.
7. Treat data integrity, employee identity, assignment linkage, compensation and confidentiality as critical.
8. Do not declare UAT ready merely because Python executes successfully.
9. Do not declare UAT successful without actual Equus import results and reconciliation.
10. Final UAT sign-off must come from the authorized business approver, not from the developer or AI.

## 1. Project context
| File | Responsibility |
|---|---|
| `lambda/src/handler.py` | AWS Lambda and S3 event handling |
| `lambda/src/transform.py` | Source-to-target Excel transformation |
| `lambda/src/mapping_config.py` | Mapping and business configurations |
| Topia source workbook | Employee and historical assignment data (`FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx`, git-ignored, in the main checkout) |
| Equus target workbook | Official import format (`lambda/src/templates/equus_eafh_template.xlsx`; sample `DATA MIGRATION SE EAFH TEST FILE.xlsx`) |
| `clarifications/jake-1.xlsx` | Jake's latest business clarification responses and dependent example |

Architecture: Topia source → AWS S3 → Lambda → Python transformation → Equus import workbook →
Validation and reconciliation → Controlled Equus UAT import → Equus audit and business sign-off.
The current implementation can generate transformed output, but a complete source-to-target
comparison and an accepted Equus UAT result have not yet been established. Do not claim otherwise.

## 2. Authoritative requirements and approval status
Read `jake-1.xlsx` carefully, including every worksheet, field, answer, annotation, example and
relevant cell. Jake provided responses for Q1–Q20. Q21–Q43 are unanswered. The tracker marks all
questions as Open, so a response may provide enough information to implement a rule without
establishing final business approval. Classification (do not combine): CONFIRMED BY JAKE ·
PARTIALLY CLARIFIED · PENDING BUSINESS APPROVAL · PENDING EQUUS CONFIGURATION · IMPLEMENTED AND
TESTED · UAT VERIFIED · SIGNED OFF. When a rule conflicts with Jake's example, escalate the
conflict rather than choosing whichever value appears convenient.

## 3. Business requirements — Jake Q1 to Q20
- **BR-01 Identity (Q1):** Employee Number = Topia Talent Link ID; Assignment Number = "NEW";
  Assignment Identifier = Topia ID; `ASSIGNMENT_CF!CF_TEXT029` retains the legacy Topia ID;
  the Assignment Identifier links additional rows to the same assignment. Critical discrepancy:
  Jake's dependent example has different values for Assignment Identifier and CF_TEXT029. Do not
  assume they are the same source field until resolved. Verify labels, uniqueness, repeated
  identifiers, leading zeros and multiple historical assignments. Missing identifiers must not
  fall back automatically to Employee Number.
- **BR-02 Resource usernames (Q2):** pending. Approved reference mappings only; never derive
  usernames from display names or e-mails without a verified rule.
- **BR-03 Citizenship / birth country (Q3):** pending. Verify Q–T header alignment; never
  interchange citizenship, birth country and nationality.
- **BR-04 Currency (Q4):** pending. Keep home, host, salary and default currency separate; no
  derivation or defaulting without approval.
- **BR-05 Duplicates and errors (Q5):** implement source-level duplicate detection, assignment
  identity validation, header/schema validation, file-level and row-level error classification,
  rejection and correction reporting, controlled reprocessing without duplicating accepted rows,
  post-import duplicate audit. Actual Equus behaviour must be demonstrated during UAT.
- **BR-06 Confidentiality (Q6):** pending Equus configuration. Do not invent custom fields or
  silently omit confidential attributes; confidentiality is a mapping and an access-control requirement.
- **BR-07 Expected Assignment Outcome (Q7):** pending Schneider and Equus ticket SE-424. Do not
  assume it equals reason or justification.
- **BR-08 Brown columns (Q8):** brown has no business meaning; never use it to decide anything.
- **BR-09 PII Nationality 1 (Q9):** decision belongs to Schneider; do not include, permanently
  exclude or remap automatically.
- **BR-10 Spouse/partner without names (Q10):** importable; never invent names or reject solely
  for absent names; validate remaining attributes against confirmed requirements.
- **BR-11 Dependents in the same workbook (Q11):** assignment-linked rows within the approved
  structure; verify the Topia extract actually contains the required dependent data.
- **BR-12 Country, state, city (Q12):** country and state must match approved Equus values;
  city is free text; never assume every ISO code is accepted.
- **BR-13 E-mails (Q13):** primary `EMPLOYEE!USER_ID_EMAIL`, secondary `EMPLOYEE_CF!CF_TEXT008`;
  no unnecessary e-mail contact records; keep phone/contact mappings separate.
- **BR-14 Dropdowns (Q14):** lists to come from Jake; unknown values reported, never converted.
- **BR-15 Cost centres (Q15):** pending; keep code/name relationships; never manufacture a name.
- **BR-16 Host weekly hours (Q16):** 158 needs clarification; warn, never cap or alter.
- **BR-17 Multiple historical assignments (Q17):** support several assignments per employee;
  never merge on Employee Number alone; verify dates, overlaps, cancelled records, re-import matching.
- **BR-18 Dependent linkage (Q18):** every related row carries the correct Assignment
  Identifier; prevent cross-assignment linkage.
- **BR-19 Multiple dependents (Q19):** follow the ` Mny dependents Ex` sheet: INSERT then UPDATE
  rows with the same Assignment Identifier; any number of dependents; validate generated rows
  against the example; do not assume capacity limits are proven.
- **BR-20 Minimum dependent data (Q20):** one populated field is technically enough; no invented
  mandatory-name validation; no entirely empty dependent records; distinguish technical
  permissibility from business data quality; lookups pending.

## 4. Unanswered requirements — Q21 to Q43
For each: identify the field/behaviour, the code and template location, whether it affects the
UAT dataset, the unapproved assumption, impact and severity, decision owner, closure evidence;
block UAT sign-off when unresolved items could materially affect correctness, completeness or
confidentiality. Do not treat an unanswered question as out of scope, and do not block an
unrelated scenario without explaining the dependency. Focus areas: Q21–Q22 compensation and
effective dates; Q23–Q24 dates and formats; Q25 source inventory and cutoff; Q26 INSERT/UPDATE/
IGNORE; Q27 reconciliation criteria; Q28 confidential visibility; Q29–Q31 custom fields, scope,
approval authority; Q32–Q35 grades, employment types, status and policy; Q36–Q39 installments,
family size, default currency, RMC; Q40 international derivation; Q41–Q43 resource e-mails,
custom-field interpretation, missing-value normalisation.

## 5. Mandatory code review
Review `mapping_config.py` (mappings, identifiers, commands, template headers, e-mails,
exclusions, lookups, country/currency, compensation, date defaults, numeric warnings, family
fields, custom fields; identify every hardcoded unapproved assumption), `transform.py` (reading,
placeholder normalisation, grouping, identity/uniqueness, historical preservation, dependents,
UPDATE rows, addresses/contacts, dates, salary/currency, preservation, warnings, reconciliation —
known concern: reconciliation compares only INSERT rows; dependent, address and contact rows
must also be reconciled) and `handler.py` (S3 events, validation and limits, retries, output
versioning, duplicate events, failure handling, audit metadata, import-ready status, secure
logging, separation of transformation output from actual Equus import — known concern:
`STRICT_PENDING` defaults to false and must not let an unresolved critical mapping produce an
approved UAT-ready artifact; do not simply enable stricter behaviour without checking that it
classifies pending fields correctly and preserves diagnostics).

## 6. Implementation safety
Baseline commit and branch first; confirm authoritative files; record hashes; run baseline tests;
record existing behaviour; identify exact changes; keep the original reproducible. Smallest safe
changes; no unrelated refactoring; separate business rules from generic logic; make unapproved
rules explicit and keep them out of UAT-ready output; validation must not destroy diagnostic
reports; never silently change identifiers, compensation, dates, statuses or confidential
attributes; reversible changes with a rollback path. Unresolved requirement = validation and
exception reporting where safe, never an invented rule.

## 7. Final UAT test requirements
Employee/assignment: single employee one assignment; one employee multiple historical
assignments; multiple employees; duplicate source rows; missing Employee Number; missing
Assignment Identifier; conflicting identifiers; previously imported assignment; repeated runs;
overlapping or missing dates; existing employee with a new assignment.
Spouse/dependent: none; one; two; three or more; spouse with name; spouse without name;
dependent with one populated field; multiple assignments each with different dependents;
repeated dependent records; missing or incorrect linkage.
Mapping/data quality: approved country/state values; unknown reference values; currency;
invalid and ambiguous dates; salary and effective-date validation; optional vs mandatory
fields; approved custom fields; red and yellow flags; e-mail mappings; placeholders;
confidential records; header mismatch; invalid Equus values.
AWS/import: valid S3 trigger; duplicate event; Lambda retry; output failure; invalid source;
invalid template; partial validation failure; row-level and file-level Equus rejection;
correction and resubmission; re-import without duplicates; post-import audit.
Test positive and negative paths; never assume the current output is right; separate simulated
behaviour from observed Equus behaviour.

## 8. Source-to-target reconciliation (mandatory for sign-off)
Cover employees, assignments, identifiers, dependents, spouses, compensation, dates, addresses,
contacts, custom fields, exclusions, reference fields, duplicates, errors. For each mapped field
and record: source key, target key, source value, transformed value, comparison under the
approved rule, mismatches, authorised conversions, blanked/excluded/missing data, exceptions
with an owner. Compare employee/assignment counts and child-record counts; one input row is not
one output row when child rows are created; separate expected expansion rows from unexpected
discrepancies. Acceptance: no unexplained loss, duplication or corruption of critical in-scope
data; noncritical exceptions explicitly approved and documented.

## 9. Actual Equus UAT execution
Pre-import validation → controlled import in the authorised UAT environment (hash, batch id,
timestamp, job results, accepted/rejected rows) → post-import verification → error correction →
business validation. Never perform the import automatically without explicit authorisation and
access; never mark an unexecuted scenario as passed.

## 10. Sign-off gates
G1 mappings approved/dispositioned · G2 Equus custom fields and lookups available · G3 identity
and linkage conflicts resolved · G4 automated tests pass with recorded results · G5
reconciliation passes · G6 controlled Equus UAT import completed · G7 Equus records independently
validated · G8 duplicate/error/retry scenarios verified · G9 security/confidentiality checks pass
· G10 no unresolved critical defects or unapproved critical exceptions · G11 business reviewers
approve the evidence · G12 authorised signatory signs off.
Final status is exactly one of: NOT READY FOR UAT · READY FOR UAT EXECUTION · UAT EXECUTED —
SIGN-OFF PENDING · UAT SIGNED OFF. Never assign UAT SIGNED OFF yourself. A deadline never waives a gate.

## 11. Production safety
UAT approval is not production authorisation. Prepare: extract version/cutoff, output
versioning, deployment package, execution sequence, pre-import inventory, recovery/rollback,
partial-import handling, duplicate correction, UAT-to-production configuration differences,
data protection/access/retention, go/no-go dependencies. Never delete records, run destructive
recovery, deploy to production or import production data without explicit authorisation.

## 12. Required evidence package
1 Requirements Traceability Matrix (all 43 questions) · 2 Code Review Report · 3 Change Log ·
4 Updated code · 5 Automated test results (commands, counts, logs) · 6 Field Mapping Validation ·
7 Source-to-Output Reconciliation · 8 Equus UAT Execution Evidence · 9 Defect and Exception
Register · 10 Pending Clarification Register · 11 UAT Readiness Assessment (each gate Pass /
Fail / Blocked with evidence) · 12 Business Sign-off Record. Anything that cannot be completed
for lack of access, inputs or approvals is marked incomplete. Never fabricate evidence.

## 13. Critical execution restrictions
Never: guess the Assignment Identifier ↔ legacy Topia ID relationship; assume one assignment per
employee; assume Equus prevents duplicates; invent dropdown/reference mappings; convert
ambiguous dates; default salary or family effective dates without approval; derive
international status without a confirmed rule; infer cancellation from a cancellation reason
without approval; exclude sensitive fields without an authorised decision; fabricate dependent
names; assume syntactically valid values are accepted Equus values; silently allow missing or
invalid identifiers; declare unresolved critical mappings acceptable; claim Equus UAT passed from
local tests; mark UAT signed off without authorised evidence. Incomplete requirements are
project dependencies, not an invitation to manufacture business logic.

## 14. Execution order
Phase 1 evidence baseline → Phase 2 code impact assessment → Phase 3 safe implementation →
Phase 4 automated testing → Phase 5 reconciliation → Phase 6 controlled Equus UAT preparation →
Phase 7 Equus result verification (only with authorised access) → Phase 8 readiness and approval
(SIGN-OFF PENDING where approval is missing, never signed off).

## 15. Final response format
A Executive summary · B Business requirement status Q1–Q43 (Jake's answers vs formally approved
rules) · C Actual code changes (file, function, before, after, reason, test evidence) · D Test
execution results · E Source-to-target reconciliation · F Critical UAT blockers (priority,
impact, missing decision, owner, evidence) · G Gates G1–G12 Pass/Fail/Blocked · H Final
readiness recommendation (one of the four statuses) with supporting evidence.

## Critical architecture requirement — future business clarification support
The implementation must be configuration-driven and extensible. Keep generic transformation
logic independent of configurable business mappings and reference values. Centralise business
configuration in `mapping_config.py` and, where appropriate, version-controlled external JSON
(the Lambda runtime has no YAML library; prefer JSON). Support configurable source-to-target
mappings, dropdown lookups, custom fields, date formats, default rules, field exclusions and
validation requirements. Design pending rules so approved clarifications are usually
configuration updates, not engine rewrites. Never activate unapproved assumptions automatically.
Include configuration versions, approval status, effective versions and audit traceability.
Validate every configuration change against the actual Equus template, approved test cases and
affected records. Reject invalid or unsupported configuration explicitly. Preserve compatibility
with the existing AWS Lambda, S3 integration and reporting. When a clarification needs new
structural logic, implement a focused, tested Python change instead of forcing it into
configuration. Maintain regression tests and reconciliation for every mapping version. Do not
declare UAT ready while critical in-scope mappings remain unresolved.
