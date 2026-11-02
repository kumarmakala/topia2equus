# Topia → Equus migration

AWS Lambda that converts a Topia **Full Historical Past Assignment** export
(`.xlsx`) dropped in a source S3 bucket into the Equus **EAFH** import
workbook, written to a destination S3 bucket together with a mapping report.

## How it works

1. An `.xlsx` uploaded under `topia/incoming/` in the source bucket triggers
   the function (S3 event filter: prefix + `.xlsx` suffix).
2. Each source column is mapped by its row-2 Equus technical field to the
   Equus row-3 import header; exceptions live in `lambda/src/mapping_config.py`.
3. Values are converted from the Equus row-1 rules (US dates, ISO-3 country
   codes, dropdown maps; "Not Provided" → blank).
4. Output: one `INSERT` row per assignment, plus `UPDATE` rows for extra
   addresses, contacts and dependents, written into the bundled Equus template.
5. Written to the destination bucket:
   - `equus/<name>_EQUUS_DRAFT_<etag>_<version>.xlsx` - Equus import file
     while readiness is `NOT_READY` (always, until every open item is
     resolved); `equus/<name>_EQUUS_<etag>_<version>.xlsx` only when
     readiness is `READY_FOR_UAT_EXECUTION`
   - `equus/reports/<name>_MAPPING_REPORT_<etag>_<version>.xlsx` - column
     mapping, governance, Equus coverage, dropdown values sent,
     reconciliation, rules, readiness, warnings
   - `equus/reports/<name>_RUN_<etag>_<version>.json` - run manifest (audit
     trail, also written when the run fails); see **Run manifest** below
6. Before any Equus import, the operator works through
   [docs/uat/PRE_IMPORT_CHECKLIST.md](docs/uat/PRE_IMPORT_CHECKLIST.md).

## Layout

| Path | Purpose |
|---|---|
| `lambda/src/handler.py` | Lambda entry point, S3 I/O, trigger guard |
| `lambda/src/transform.py` | Pure transformation (bytes in, bytes out) |
| `lambda/src/mapping_config.py` | Field overrides, country table, configuration loaders |
| `lambda/src/business_rules.json` | Versioned business rules with Q number, status, owner, `approved_for_import` |
| `lambda/src/reference_data/` | Versioned reference lists (resource usernames, cost centres, country/state); empty placeholders |
| `lambda/src/templates/` | Equus EAFH template (header rows only) |
| `lambda/template.yaml` | SAM stack: function, least-privilege IAM, failure queue |
| `infra/deploy.sh` | Deploy + merge filtered S3 trigger into the source bucket |
| `infra/source-bucket-policy.json` | Optional: TLS-only, `.xlsx`-only uploads |
| `local/run_local.py` | Run the handler with local folders as buckets |
| `tests/` | pytest suite with a synthetic (fake data) fixture |

## Run locally

Requires Python 3.9+ for local runs and git. SAM builds require Python 3.13,
matching the Lambda runtime in `lambda/template.yaml`.

1. Clone and set up a virtual environment:

   ```bash
   git clone git@github.com:kumarmakala/topia2equus.git
   cd topia2equus
   python3 -m venv .venv
   source .venv/bin/activate                    # Windows: .venv\Scripts\activate
   pip install openpyxl==3.1.5 pytest boto3 "moto[s3]"   # moto: idempotency test
   ```

2. Run the tests:

   ```bash
   python3 -m pytest -q tests
   ```

3. Run the migration, with folders under `local_s3/` standing in for the
   source and destination buckets:

   ```bash
   mkdir -p local_s3/source_s3/topia/incoming
   cp <topia-export>.xlsx local_s3/source_s3/topia/incoming/
   python3 local/run_local.py                   # every .xlsx under topia/incoming/
   python3 local/run_local.py "topia/incoming/<file>.xlsx"   # or a single file
   ```

   Outputs land in `local_s3/dest_s3/equus/` (import file) and
   `local_s3/dest_s3/equus/reports/` (mapping report).

4. Optional: end-to-end run through real boto3 against moto-mocked S3,
   including the trigger guards:

   ```bash
   pip install boto3 "moto[s3]"
   python3 local/simulate_s3.py <topia-export>.xlsx <download-dir>
   ```

`*.xlsx` and `local_s3/` are git-ignored, so Topia exports must be copied to
each machine by hand.

## Deploy

Requires AWS credentials and the SAM CLI; buckets must already exist.

```bash
infra/deploy.sh <source-bucket> <dest-bucket> [source-prefix] [dest-prefix] [stack-name]
```

## Data handling

Real Topia/Equus workbooks contain personal data and are git-ignored. The
committed template has header rows only, and the test fixture uses fake data.

## Controls

- **Colour governance:** the fills of source rows 1-4 are read as rules.
  Red columns are excluded. Yellow columns are flagged `PENDING` (set
  `STRICT_PENDING=true` to refuse output while they hold data). Cyan columns
  without an Equus column are flagged `MISSING_TARGET`. Brown (`FFB45F06`) is
  `none (Q8: no meaning)`; any other fill is flagged `UNKNOWN_COLOUR`. See
  the report's Governance sheet.
- **Business rules (configuration-driven):** every rule that changes values
  is declared in `lambda/src/business_rules.json` with `id`, `Q`, `title`,
  `status` (`confirmed_by_jake`, `partially_clarified`,
  `pending_business_approval`, `pending_equus_configuration`, `open`),
  `approved_for_import`, `owner` and `parameters`. A rule changes the import
  file only while `approved_for_import` is `true`; a Jake answer is evidence,
  not approval, so every shipped rule is `false`:

  | Rule | Q | Behaviour while not approved |
  |---|---|---|
  | `international_relo_derivation` | Q40 | `INTERNATIONAL_RELO` not derived |
  | `cancellation_reason_sets_phase` | Q34 | `PHASE` not set; Cancellation Reason column HELD |
  | `default_effective_dates_from_start` | Q22/Q23 | salary/family-size effective dates, address `BEGIN_DATE`, contact `START_DATE` not defaulted |
  | `value_maps_employee_type_and_status` | Q35 | values sent untranslated |
  | `host_job_level_to_salary_grade` | Q32 | Host Job Level column HELD (not mapped) |
  | `country_name_to_iso3` | Q12/Q3 | conversion performed for diagnostics; each converted value listed as a provisional conversion |

  The loader rejects unknown keys, unknown statuses, missing fields,
  duplicate or unsupported rule ids, wrong parameters, duplicate JSON keys
  and invalid JSON with `ConfigurationError` naming the file and JSON path
  (e.g. `$.rules[2].owner`); `transform()` fails before reading any workbook.
  The report's **Rules** sheet lists version, every rule, whether it was
  applied and how many rows it affected. Rollback of a gated behaviour is a
  configuration change: set `approved_for_import: true` (with an
  `approval_reference`) for that rule; the output keys change with it.
- **Reference data:** `lambda/src/reference_data/{resource_usernames,
  cost_centres,country_state}.json` carry version/approval metadata and
  empty `values`. An absent, empty or unapproved list means every value of
  the covered headers is an exception, never a guess.
- **Import readiness:** every run computes `NOT_READY` or
  `READY_FOR_UAT_EXECUTION` (Readiness sheet, S3 metadata `readiness`, log
  summary). It is `NOT_READY` while any gated rule affected a row, any
  provisional conversion, held/pending/missing-target column with data,
  unvalidated reference or dropdown value, identifier fallback, configuration
  mismatch or reconciliation failure remains; each reason carries its Q
  number and owner. Equus UAT execution is not performed by this tool, so
  `READY_FOR_UAT_EXECUTION` is the best status it can report. Readiness is
  independent of `STRICT_PENDING`.
- **Validation:** warns on out-of-range numbers, cost-centre names that hold
  codes, spouse/dependent records without a name, resource names missing from
  `RESOURCE_MAP` (derived from the approved resource list), and dropdown
  values outside `lambda/src/equus_lists.json` (optional). Warnings never
  change values.
- **Reconciliation:** after writing, every mapped source cell is compared
  with the written workbook. Unexplained differences fail the run: only a
  `RECONCILIATION_FAILED` report and the run manifest are written (no import
  file), and the event goes to the failure queue.
- **Run manifest (UAT-07):** `equus/reports/<name>_RUN_<etag>_<version>.json`
  records source bucket/key/ETag/SHA-256/size, template SHA-256 + version
  (`bundled:<sha256[:12]>` or the S3 VersionId/ETag of `TEMPLATE_S3_URI`:
  Equus has issued no template version, D2), business-rules version,
  effective version, file SHA-256 and the rule set (id, Q, status,
  `approved_for_import`, owner), reference lists, code version + code
  SHA-256, `strict_pending`, readiness status and reasons, the output kind
  (`IMPORT_CANDIDATE`, `DRAFT_NOT_IMPORT_READY` or `NONE`) with an explicit
  import-ready statement, counts (source rows, accepted, held rows and held
  detail, rows written, warnings, unmapped columns, exceptions by code),
  timestamps and the failure (if any). The import file, report and manifest
  carry S3 metadata `readiness`, `import-ready`, `business-rules-version`,
  `business-rules-effective-version`, `code-version`, `template-sha256` and
  `run-manifest`.
- **Failure classification:** a failed run raises `MigrationError` whose
  message starts with `[CODE]`, and the manifest's `failure.code` holds the
  same code: `CONFIG` (`ConfigurationError` from `business_rules.json` /
  `reference_data/`), `TEMPLATE` (template unreadable, sheet missing,
  duplicated header, or a UAT-03 FILE-level header finding against the
  approved template - `HEADER_MISMATCH` / `TEMPLATE_HASH_DIFFERS`, with the
  `FILE_REJECTED` report kept - Q5: header issues fail the whole import),
  `RECONCILIATION`, `PENDING` (`STRICT_PENDING=true` with yellow data),
  `SOURCE` (unreadable source or no source sheet), `INTERNAL` (a FILE-level
  `ROW_STRUCTURE_INVALID` engine defect, or anything else; an unexpected
  error's message is withheld because it could quote data). No import file is
  written for any failure. File and manifest metadata also carry `rows-held`
  and `file-status`.
- **Secure logging:** log lines carry S3 keys, counts, versions and codes only
  - never cell values, names or e-mail addresses
  (`tests/test_handler_audit.py::test_logs_contain_no_fixture_names_or_emails`).
  Unexpected errors are logged as `file:line:function` frames without
  messages.
- **Idempotency:** output keys are `<name>_EQUUS[_DRAFT]_<source ETag>_<code
  version>.xlsx`; the code version hashes the code, `business_rules.json`,
  `reference_data/` and the template. A duplicate S3 delivery is skipped
  (under either name) and writes nothing new; a mapping or approval change
  produces new keys.

## Open mapping questions

See [OPEN_QUESTIONS.md](OPEN_QUESTIONS.md), which gives the cell evidence,
current behaviour and config setting for each one. In short:

- Assignment Identifier: Topia Assignment ID (current) vs Employee Number
  (Equus sample).
- Resource fields need a name → Equus username list (`RESOURCE_MAP`).
- Home/Host currency has no source column.
- Equus template columns Q-T: labels are shifted against field keys.
- Yellow (pending), cyan-without-target (AG, DG) and brown (spouse) columns.
- Address text format; source of dependent data (the "Dependent Import"
  sheet has no employee key).
