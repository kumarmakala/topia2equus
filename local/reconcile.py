"""Run the transformation on a Topia workbook and write the source-to-target
reconciliation (UAT-05) next to the local S3 outputs.

  python3 local/reconcile.py                          # the real extract in local_s3/source_s3/topia/incoming/
  python3 local/reconcile.py path/to/workbook.xlsx    # any workbook (e.g. a synthetic fixture)
  python3 local/reconcile.py WORKBOOK --out-dir DIR --business-rules FILE --json

Writes, under local_s3/dest_s3/equus/reconciliation/ (git-ignored, never
committed - real runs contain personal data):
  <stem>_RECONCILIATION_<sha8>.xlsx   the full report workbook (all sheets)
  <stem>_RECONCILIATION_<sha8>.json   summary + per-column counts (no cell values)

Prints only counts, column letters, Equus headers and exception codes: never
a source value, an identifier or a name. Exit code 0 = reconciliation OK,
1 = reconciliation FAILED (the report is written either way). This is local
evidence only: no Equus import is performed (master command section 9)."""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lambda" / "src"))

import transform as tmod  # noqa: E402

DEFAULT_SOURCE = ROOT / "local_s3" / "source_s3" / "topia" / "incoming" / "FULL HISTORICAL PAST ASSIGNMENT ALL FIELDS.xlsx"
DEFAULT_TEMPLATE = ROOT / "lambda" / "src" / "templates" / "equus_eafh_template.xlsx"
DEFAULT_OUT_DIR = ROOT / "local_s3" / "dest_s3" / "equus" / "reconciliation"

# Per-column keys that never carry a source value (safe to print / store).
SAFE_FIELD_KEYS = [k for k in tmod.RECONCILIATION_KEYS if k not in ("source_date_total", "written_date_total")]
# Per-record keys without identifiers or employee numbers.
SAFE_RECORD_KEYS = [k for k in tmod.RECORDS_KEYS if k not in ("identifier", "employee")]


def run(source: Path, template: Path = DEFAULT_TEMPLATE, business_rules: Path | None = None) -> dict:
    """Transform `source` and return the value-free reconciliation evidence."""
    _out, report_bytes, report = tmod.transform(source.read_bytes(), template.read_bytes(),
                                                business_rules=business_rules)
    return {
        "report": report,
        "report_bytes": report_bytes,
        "evidence": evidence(report),
    }


def evidence(report) -> dict:
    """Counts, column letters, headers and codes only (no source values)."""
    return {
        "source_sha256": report.source_sha256,
        "business_rules_version": report.business_rules_version,
        "reconciliation_ok": report.reconciliation_ok,
        "summary": report.reconciliation_summary,
        "fields": [{k: r.get(k) for k in SAFE_FIELD_KEYS} for r in report.reconciliation],
        "records": [{k: r.get(k) for k in SAFE_RECORD_KEYS} for r in report.reconciliation_records],
        "exception_codes": dict(sorted(Counter(e["code"] for e in report.exceptions).items())),
        "held_dependents": [{k: e[k] for k in ("code", "sheet", "rows", "groups", "q", "owner")}
                            for e in report.dependent_exceptions],
        "readiness": report.readiness.status if report.readiness else None,
    }


def write(result: dict, source: Path, out_dir: Path) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = f"{source.stem}_RECONCILIATION_{result['report'].source_sha256[:8]}"
    xlsx, js = out_dir / f"{stem}.xlsx", out_dir / f"{stem}.json"
    xlsx.write_bytes(result["report_bytes"])
    js.write_text(json.dumps(result["evidence"], indent=2, default=str), encoding="utf-8")
    return xlsx, js


def format_summary(ev: dict) -> str:
    s = ev["summary"]
    lines = [f"Reconciliation: {s.get('result')}  (source SHA-256 {ev['source_sha256'][:16]}..., "
             f"business rules {ev['business_rules_version']})"]
    for key, value in s.items():
        if key not in ("result", "reasons"):
            lines.append(f"  {key}: {value}")
    for reason in s.get("reasons", []):
        lines.append(f"  REASON: {reason}")
    lines.append("Columns not (fully) written, by category (column letter, Equus header, counts):")
    explained = [c for c in tmod.RECON_CATEGORIES if c != "written"]
    for f in ev["fields"]:
        parts = [f"{c}={f[c]}" for c in explained if f.get(c)]
        if parts:
            lines.append(f"  {f['col']:>4} {f['group']:<10} {str(f['header'] or '-'):<48} source={f['source']} "
                         + " ".join(parts) + (f"  [{f['q']}]" if f.get("q") else ""))
    for h in ev["held_dependents"]:
        lines.append(f"Held dependents: {h['code']} sheet {h['sheet']!r}: {h['rows']} row(s), {h['groups']} "
                     f"record(s) ({h['q']}, owner {h['owner']})")
    lines.append(f"Exception codes: {ev['exception_codes']}")
    lines.append(f"Readiness: {ev['readiness']} (Equus UAT execution: NOT EXECUTED from this environment)")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("workbook", nargs="?", type=Path, default=DEFAULT_SOURCE)
    ap.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    ap.add_argument("--business-rules", type=Path, default=None)
    ap.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    ap.add_argument("--json", action="store_true", help="print the value-free evidence as JSON")
    args = ap.parse_args(argv)
    if not args.workbook.exists():
        print(f"workbook not found: {args.workbook}", file=sys.stderr)
        return 2
    result = run(args.workbook, args.template, args.business_rules)
    xlsx, js = write(result, args.workbook, args.out_dir)
    print(json.dumps(result["evidence"], indent=2, default=str) if args.json else format_summary(result["evidence"]))
    print(f"Report written: {xlsx.relative_to(ROOT) if xlsx.is_relative_to(ROOT) else xlsx}")
    print(f"Evidence written: {js.relative_to(ROOT) if js.is_relative_to(ROOT) else js}")
    return 0 if result["report"].reconciliation_ok else 1


if __name__ == "__main__":
    sys.exit(main())
