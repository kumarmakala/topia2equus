"""Run the migration Lambda locally, with two folders standing in for S3.

  local_s3/source_s3/  -> source bucket "source_s3"
  local_s3/dest_s3/    -> destination bucket "dest_s3"

Usage:
  python3 local/run_local.py                      # process every .xlsx under source_s3/topia/incoming/
  python3 local/run_local.py "topia/incoming/x.xlsx"
"""

import io
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BUCKETS = ROOT / "local_s3"
sys.path.insert(0, str(ROOT / "lambda" / "src"))

os.environ.setdefault("DEST_BUCKET", "dest_s3")
os.environ.setdefault("DEST_PREFIX", "equus/")
os.environ.setdefault("SOURCE_PREFIX", "topia/incoming/")
os.environ.setdefault("SOURCE_SUFFIX", ".xlsx")

import handler  # noqa: E402


class FolderS3:
    """Implements the two S3 calls the handler makes, against local folders."""

    def get_object(self, Bucket, Key):
        return {"Body": io.BytesIO((BUCKETS / Bucket / Key).read_bytes())}

    def put_object(self, Bucket, Key, Body, **kwargs):
        path = BUCKETS / Bucket / Key
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(Body)


def s3_event(bucket, key):
    path = BUCKETS / bucket / key
    return {"Records": [{
        "eventName": "ObjectCreated:Put",
        "s3": {"bucket": {"name": bucket},
               "object": {"key": key.replace(" ", "+"), "size": path.stat().st_size}},
    }]}


def main():
    handler._s3 = FolderS3()
    src = BUCKETS / "source_s3"
    keys = sys.argv[1:] or sorted(
        str(p.relative_to(src)) for p in src.rglob("*") if p.is_file())
    for key in keys:
        print(f"\n>>> S3 event: s3://source_s3/{key}")
        result = handler.lambda_handler(s3_event("source_s3", key), None)
        for r in result["results"]:
            for k, v in r.items():
                print(f"    {k}: {v}")


if __name__ == "__main__":
    main()
