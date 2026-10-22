"""End-to-end simulation: real boto3 + handler against moto-mocked S3.

Usage: python3 local/simulate_s3.py <topia-export.xlsx> <download-dir>

Mirrors the deployed setup: source bucket with a topia/incoming/ trigger,
destination bucket for equus/ outputs. Also exercises the trigger guards.
"""
import json
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
OUT = Path(sys.argv[2])
SOURCE_FILE = Path(sys.argv[1])
SRC, DST = "topia-source-sim", "equus-dest-sim"

os.environ.update(AWS_DEFAULT_REGION="us-east-1", AWS_ACCESS_KEY_ID="test",
                  AWS_SECRET_ACCESS_KEY="test", DEST_BUCKET=DST, DEST_PREFIX="equus/",
                  SOURCE_PREFIX="topia/incoming/", SOURCE_SUFFIX=".xlsx")
sys.path.insert(0, str(REPO / "lambda" / "src"))

import boto3
from moto import mock_aws


def event(bucket, key, size):
    return {"Records": [{"eventName": "ObjectCreated:Put",
                         "s3": {"bucket": {"name": bucket},
                                "object": {"key": key.replace(" ", "+"), "size": size}}}]}


with mock_aws():
    import handler
    s3 = boto3.client("s3")
    s3.create_bucket(Bucket=SRC)
    s3.create_bucket(Bucket=DST)

    body = SOURCE_FILE.read_bytes()
    key = f"topia/incoming/{SOURCE_FILE.name}"
    s3.put_object(Bucket=SRC, Key=key, Body=body)
    print(f"Uploaded s3://{SRC}/{key} ({len(body)} bytes)")

    print("\n=== Migration run ===")
    res = handler.lambda_handler(event(SRC, key, len(body)), None)
    print(json.dumps(res, indent=2))

    print("\n=== Trigger guard checks (all should be skipped) ===")
    s3.put_object(Bucket=SRC, Key="other/x.xlsx", Body=body)
    s3.put_object(Bucket=SRC, Key="topia/incoming/notes.csv", Body=b"a,b")
    s3.put_object(Bucket=SRC, Key="topia/incoming/~$lock.xlsx", Body=b"x")
    for k, size in [("other/x.xlsx", len(body)), ("topia/incoming/notes.csv", 3),
                    ("topia/incoming/~$lock.xlsx", 1), ("topia/incoming/huge.xlsx", 60 * 1024 * 1024)]:
        r = handler.lambda_handler(event(SRC, k, size), None)["results"][0]
        print(f"  {k}: {r.get('skipped', 'NOT SKIPPED!')}")

    print(f"\n=== Destination bucket s3://{DST} ===")
    OUT.mkdir(parents=True, exist_ok=True)
    for obj in s3.list_objects_v2(Bucket=DST).get("Contents", []):
        head = s3.head_object(Bucket=DST, Key=obj["Key"])
        print(f"  {obj['Key']}  {obj['Size']} bytes  {head['ContentType'][:40]}...  meta={head['Metadata']}")
        (OUT / Path(obj["Key"]).name).write_bytes(s3.get_object(Bucket=DST, Key=obj["Key"])["Body"].read())
    print(f"\nDownloaded outputs to {OUT}")
