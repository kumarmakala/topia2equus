# Pre-import checklist: what the operator looks at before any Equus import

Task: OCNORA T-1392 / UAT-07 (handler hardening). Sources: `clarifications/UAT_MASTER_COMMAND.md`
§6 (lines 128-135), §9 (lines 165-169), §10 (lines 171-179), §14 (lines 207-211);
`clarifications/JAKE_Q1-Q43_EVIDENCE_ASSESSMENT.md` §1 (D1-D5) and §5.
Every tracker Status (`clarifications/jake-1.xlsx` sheet `datamapping clarrification shef`
column I) is still **Open**. Equus UAT execution, import results and post-import audit are
**INCOMPLETE / NOT EXECUTED** from this environment (no Equus access). The best status this
tool can compute is `READY_FOR_UAT_EXECUTION`; it never signs anything off. Expected status for
the current configuration: **NOT READY FOR UAT** (assessment §5).

Assumption (recorded, not asked): "the master command's step 1 items" are read as the first step
of the Equus UAT execution in §9 ("Pre-import validation"), together with the evidence-baseline
items that §6 and §14 Phase 1 put first (authoritative files, hashes, baseline tests). All are
listed below in section 2.

Jake, Q5 (`datamapping clarrification shef`!J6, verbatim): "Duplicate records in the import can
result in duplicate records being created in the system. There are some protections in place to
prevent duplicate employees from being created, but I would not rely on this as a fail-safe. If
duplicates are created, an audit report can be run to identify them, then they can be deleted. In
general, the import will process as many rows as it can and reject (error) rows that contain
invalid data. If this occurs, a subsequent import can be ran to correct rows that were rejected.
Note that if there are issues in the header row, the full import fails. We always recomend running
test a validating data to prevent these scenarios." (I6: `Open`.)
This is why every item below is checked **before** an import and not repaired afterwards.

## 1. The three artefacts of one run (all keyed by the same `<stamp>`)

`<stamp>` = `<source ETag[:12]>_<code version>` (`lambda/src/handler.py:365-369`).

| Object | Meaning | May be imported? |
|---|---|---|
| `equus/<stem>_EQUUS_<stamp>.xlsx` | written only when readiness is `READY_FOR_UAT_EXECUTION` (`handler.py:432`) | Only into the authorised **UAT** environment, after every box in section 2 is ticked |
| `equus/<stem>_EQUUS_DRAFT_<stamp>.xlsx` | transformation artefact while readiness is `NOT_READY` | **No.** Review only. Manifest `output.kind = DRAFT_NOT_IMPORT_READY`, `output.import_ready = false` |
| `equus/reports/<stem>_MAPPING_REPORT_<stamp>.xlsx` | diagnostic mapping report (Column Mapping, Governance, Equus Coverage, Dropdown Values, Reconciliation, Rules, Family, Readiness, Warnings) | No (evidence) |
| `equus/reports/<stem>_RECONCILIATION_FAILED_<stamp>.xlsx` | report of a run that did not reconcile; no import file exists | No |
| `equus/reports/<stem>_FILE_REJECTED_<stamp>.xlsx` | report of a run with a FILE-level finding (UAT-03: `HEADER_MISMATCH`, `TEMPLATE_HASH_DIFFERS` at FILE level, `ROW_STRUCTURE_INVALID`); no import file exists (`handler.py:443-452`) | No |
| `equus/reports/<stem>_RUN_<stamp>.json` | run manifest (audit trail), written on success **and** failure (`handler.py:410,468`) | No (evidence) |

The manifest is the entry point: open it first. Fields (built in `handler.py:377-393`,
`_configuration_section` `:263`, `_counts` `:285`, `_output_section` `:317`, `_readiness_section` `:332`):
`source` (bucket, key, etag, sha256, size_bytes) · `template` (source, sha256, version) ·
`configuration.business_rules` (file, sha256, version, effective_version, generated_at, rules[id,
q, status, approved_for_import, owner], dependent_source) · `configuration.reference_data` ·
`code` (version, sha256) · `strict_pending` · `idempotency` (stamp, import/draft keys) ·
`readiness` (status, reasons with Q and owner, `equus_uat_execution: INCOMPLETE …`) · `output`
(import_file, kind, import_ready, statement, mapping_report) · `counts` (source_rows, accepted,
held_rows, held{assignment_rows, dependent_rows, dependent_records, columns_with_data},
rows_written, warnings, unmapped_source_columns, exceptions_by_code) · `run` (status, started_at,
finished_at, aws_request_id, function_version) · `failure` (code, message) or null.
Evidence: `tests/test_handler_audit.py::test_run_writes_manifest_with_every_listed_field`.

## 2. Step 1 - pre-import validation (tick every box; any "no" = do not import)

### 2a. Evidence baseline (master command §6, §14 Phase 1)
- [ ] The source in the manifest is the authorised extract: `source.bucket`/`source.key` as
      agreed, `source.sha256` recorded in the import log and equal to the hash of the file the
      business approved (Q25 cutoff/inventory is still **Open**, owner Alka/Schneider).
- [ ] `template.sha256` equals the SHA-256 of the current official Equus template. Bundled
      template at this commit: `c90e9fc057e9fbcc5b1b38e8b4821dd266b92bd7557915f48f7925fc3f445be7`.
      Equus has issued no template version id, so `template.version` is `bundled:<sha256[:12]>` or
      the S3 VersionId/ETag of `TEMPLATE_S3_URI`. **D2 open:** template A3 = `A`, Jake's example A3
      = `Employee Command`; owner Jake. Q5: a header-row issue fails the whole import.
- [ ] `configuration.business_rules.version` / `effective_version` / `sha256` are the reviewed
      configuration (shipped `lambda/src/business_rules.json` at this commit: version `1.2.0`,
      SHA-256 `41cf2580c754dd133ac014165f4e3ca1f2c35d7f3bf7317afe7e23a1388b0dd9`) and every rule
      with `approved_for_import: true` has a written approval reference (none at this commit).
- [ ] `code.version` / `code.sha256` belong to a commit whose full test suite passed (record the
      command and counts: `./.venv/bin/python -m pytest -q tests`).
- [ ] The S3 object metadata of the file to be imported agrees with the manifest:
      `readiness`, `import-ready`, `business-rules-version`, `business-rules-effective-version`,
      `code-version`, `template-sha256`, `run-manifest`, `rows-held`, `file-status` (= `ACCEPTED`).

### 2b. Readiness and import status
- [ ] `run.status = SUCCEEDED` and `failure = null`. On failure see section 3; nothing is imported.
- [ ] `readiness.status = READY_FOR_UAT_EXECUTION` **and** `output.kind = IMPORT_CANDIDATE`
      **and** `output.import_ready = true` **and** the file name has no `_DRAFT_`.
      With today's configuration every run is `NOT_READY` / `DRAFT_NOT_IMPORT_READY` - stop here.
- [ ] `readiness.reasons` is empty. Otherwise each reason names its Q number and owner; it is
      resolved by that owner (approval or Equus configuration), not by the operator.
- [ ] `readiness.equus_uat_execution` still reads `INCOMPLETE`: the file has not been tested in
      Equus; this checklist is the gate before that test, not after it.

### 2c. Counts and exceptions (reconciliation, held data)
- [ ] `counts.source_rows = counts.accepted` (each source assignment row produced one INSERT row);
      `counts.held.assignment_rows = 0` (source rows held by UAT-03 ROW_HELD codes are listed in the
      mapping report's Exceptions and Held Rows sheets; reprocess per `docs/uat/REPROCESSING_PROCEDURE.md`).
- [ ] `counts.rows_written - counts.accepted` (UPDATE rows for addresses, contacts, dependents)
      matches the child records expected from the source (master command §8: one input row is not
      one output row; dependent, address and contact rows are not yet reconciled by the tool -
      assessment §5 G5).
- [ ] `counts.held_rows = 0`, `counts.held.dependent_records = 0`,
      `counts.held.columns_with_data = 0`; otherwise held data is **not** in the file and needs a
      disposition from the owner in the mapping report (Family and Readiness sheets).
- [ ] `counts.exceptions_by_code` reviewed code by code against the mapping report's Readiness
      sheet; `counts.warnings` reviewed in the Warnings sheet (warnings never change values).
- [ ] Mapping report, Reconciliation sheet: `Result = OK`, every `Unexplained` = 0.

### 2d. Identity and duplicates (Q1, Q5, Q17, Q18)
- [ ] No identifier exception in `readiness.reasons` (category `identifier`): no row without an
      Employee Number and no Assignment Identifier taken from the Employee Number. **D1 open**
      (Jake's text vs his example for Assignment Identifier / CF_TEXT029), owner Jake + Alka/John.
- [ ] The same source (same `source.sha256`) has not been imported before: check the import log
      for an earlier manifest with the same hash. The function itself skips a duplicate S3 event
      and writes nothing new (`tests/test_handler_audit.py::test_duplicate_event_writes_nothing_new`),
      but Equus does not reliably prevent duplicates (Q5 above).

### 2e. Controlled import record (master command §9; to be filled during the import, not here)
- [ ] Authorised UAT environment and authorised person named; never production.
- [ ] Import batch id, timestamp, file SHA-256 (= the S3 object's hash), Equus job result,
      accepted/rejected row counts, rejected-row file - all attached to the manifest key.
- [ ] Post-import verification and duplicate audit (Q5) performed and recorded.
These records do not exist yet: **Equus UAT execution evidence INCOMPLETE**.

## 3. Failed runs: what the code means and who acts

The raised message starts with `[CODE]` and the manifest holds `failure.code` / `failure.message`
(`handler.py:68-77`, `classify_failure` `:229`). No import file is written for any failure.

| Code | Cause | Who acts | Test |
|---|---|---|---|
| `CONFIG` | `business_rules.json` or `reference_data/*.json` rejected at load (`ConfigurationError`, message names file and JSON path) | configuration maintainer; fix and re-run (new code version = new keys) | `tests/test_handler_audit.py::test_broken_business_rules_is_config_failure_without_import_file` |
| `TEMPLATE` | template unreadable, sheet `Sheet1` missing or a duplicated header (`transform.py:75`, `check_template_structure` `:275`), or a UAT-03 FILE-level header finding against the approved template (`HEADER_MISMATCH`, `TEMPLATE_HASH_DIFFERS`; `check_template_headers` `:774`, `handler.py:77,443-452`); `FILE_REJECTED` report kept | Jake (official template, D2) | `::test_template_header_mismatch_is_template_failure`, `::test_file_level_header_finding_is_template_failure`, `::test_unreadable_template_is_template_failure`, `tests/test_identity_controls.py::test_handler_withholds_import_on_file_level_failure` |
| `RECONCILIATION` | source-to-output comparison did not balance; `RECONCILIATION_FAILED` report kept | developer + Schneider/Jake (Q27) | `::test_reconciliation_failure_is_classified_and_keeps_evidence` |
| `PENDING` | `STRICT_PENDING=true` and yellow (pending) columns hold data | Alka/John (mapping review) | `::test_strict_pending_is_pending_failure` |
| `SOURCE` | source not a readable workbook or no `Full Historical Past Assignment` sheet (`transform.py:81`) | source owner (Alka / Topia) | `::test_unreadable_source_is_source_failure` |
| `INTERNAL` | a UAT-03 FILE-level `ROW_STRUCTURE_INVALID` (engine defect; message lists codes only), or anything else; message withheld (could quote data), log has `file:line:function` frames only | developer | `::test_unexpected_error_message_is_withheld` |

## 4. Logs

CloudWatch lines carry S3 keys, counts, versions, readiness and failure codes only - never cell
values, names or e-mail addresses. Evidence:
`tests/test_handler_audit.py::test_logs_contain_no_fixture_names_or_emails` and
`::test_failure_logs_and_messages_contain_no_fixture_personal_values` (fixture
`tests/fixtures/topia_sample.xlsx`, synthetic markers). Personal data lives only in the source,
the import/draft workbook and the mapping report (Warnings, Dropdown Values, Family sheets),
which must be handled under the project's data-protection rules. The S3 key itself is logged:
source file names must not contain personal names.

## 5. Not covered here / open

- Equus UAT import, Equus results, post-import audit: **NOT EXECUTED** (no access).
- A local run of the real extract with this change was **not executed** in this task (the
  runner's tool policy refused the command); the automated suite uses the synthetic fixture only.
- Readiness gates (master command §10): G1, G2, G3 Blocked; G5 Fail (child rows not reconciled);
  G6-G8 Not executed; G9 Blocked (Q6, Q28); G10 Fail while D4 rules are gated; G11-G12 Not obtained.
  Overall: **NOT READY FOR UAT**.
