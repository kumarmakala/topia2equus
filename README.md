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
   - `equus/<name>_EQUUS_<timestamp>.xlsx` - Equus import file
   - `equus/reports/<name>_MAPPING_REPORT_<timestamp>.xlsx` - column mapping,
     Equus coverage, dropdown values sent, warnings

## Layout

| Path | Purpose |
|---|---|
| `lambda/src/handler.py` | Lambda entry point, S3 I/O, trigger guard |
| `lambda/src/transform.py` | Pure transformation (bytes in, bytes out) |
| `lambda/src/mapping_config.py` | Field overrides, value maps, country codes |
| `lambda/src/templates/` | Equus EAFH template (header rows only) |
| `lambda/template.yaml` | SAM stack: function, least-privilege IAM, failure queue |
| `infra/deploy.sh` | Deploy + merge filtered S3 trigger into the source bucket |
| `infra/source-bucket-policy.json` | Optional: TLS-only, `.xlsx`-only uploads |
| `local/run_local.py` | Run the handler with local folders as buckets |
| `tests/` | pytest suite with a synthetic (fake data) fixture |

## Run locally

Requires Python 3.9+ (3.12 recommended) and git.

1. Clone and set up a virtual environment:

   ```bash
   git clone git@github.com:kumarmakala/topia2equus.git
   cd topia2equus
   python3 -m venv .venv
   source .venv/bin/activate                    # Windows: .venv\Scripts\activate
   pip install openpyxl==3.1.5 pytest
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

## Open mapping questions

- Source colour coding (red / yellow / light blue) was not present in the
  files received; mapping is header-based and needs review via the report.
- Assignment Identifier: Topia Assignment ID (current) vs Employee Number.
- Resource fields need a name → Equus username list (`RESOURCE_MAP`).
- Address text format, to split city / postal code / country.
- Source of dependent data (the "Dependent Import" sheet has no employee key).
- Defaults to confirm: effective dates from start date, International
  Assignment derived from countries, PII Nationality excluded.
