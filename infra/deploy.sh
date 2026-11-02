#!/usr/bin/env bash
# Deploy the migration Lambda and attach the filtered S3 trigger to an
# existing source bucket without overwriting its other notifications.
#
# Usage: infra/deploy.sh <source-bucket> <dest-bucket> [source-prefix] [dest-prefix] [stack-name]
set -euo pipefail

SRC_BUCKET="$1"
DEST_BUCKET="$2"
SRC_PREFIX="${3:-topia/incoming/}"
DEST_PREFIX="${4:-equus/}"
STACK="${5:-topia-equus-migration}"
HERE="$(cd "$(dirname "$0")" && pwd)"

if [[ "$SRC_BUCKET" == "$DEST_BUCKET" && "$DEST_PREFIX" == "$SRC_PREFIX"* ]]; then
  echo "DEST_PREFIX must not sit under SOURCE_PREFIX in the same bucket (would retrigger)" >&2
  exit 1
fi

# Git Bash does not inherit interactive aliases such as sam='sam.cmd'.
if command -v sam >/dev/null 2>&1; then
  SAM_CMD=sam
elif command -v sam.cmd >/dev/null 2>&1; then
  SAM_CMD=sam.cmd
else
  echo "AWS SAM CLI was not found. Add sam or sam.cmd to PATH." >&2
  exit 1
fi

# Validate Python before deploying; Windows may expose a non-working python3
# Microsoft Store shortcut even when python is installed.
PYTHON_CMD=""
for candidate in python python3; do
  if "$candidate" -c 'import sys; sys.exit(0 if sys.version_info.major == 3 else 1)' >/dev/null 2>&1; then
    PYTHON_CMD="$candidate"
    break
  fi
done
if [[ -z "$PYTHON_CMD" ]]; then
  echo "Python 3 was not found. Add a working python or python3 to PATH." >&2
  exit 1
fi

cd "$HERE/../lambda"
"$SAM_CMD" build
"$SAM_CMD" deploy --stack-name "$STACK" --capabilities CAPABILITY_IAM --resolve-s3 \
  --no-fail-on-empty-changeset \
  --parameter-overrides SourceBucketName="$SRC_BUCKET" DestBucketName="$DEST_BUCKET" \
                        SourcePrefix="$SRC_PREFIX" DestPrefix="$DEST_PREFIX"

FN_ARN=$(aws cloudformation describe-stacks --stack-name "$STACK" \
  --query "Stacks[0].Outputs[?OutputKey=='FunctionArn'].OutputValue" --output text)

CURRENT=$(aws s3api get-bucket-notification-configuration --bucket "$SRC_BUCKET")
# A relative path also works with native Windows Python under Git Bash.
MERGED=$(CURRENT="$CURRENT" FN_ARN="$FN_ARN" SRC_PREFIX="$SRC_PREFIX" FILTER="../infra/s3-notification-filter.json" "$PYTHON_CMD" - <<'PY'
import json, os
cfg = json.loads(os.environ["CURRENT"] or "{}")
new = json.load(open(os.environ["FILTER"]))
new["LambdaFunctionArn"] = os.environ["FN_ARN"]
for rule in new["Filter"]["Key"]["FilterRules"]:
    if rule["Name"] == "prefix":
        rule["Value"] = os.environ["SRC_PREFIX"]
lambdas = [c for c in cfg.get("LambdaFunctionConfigurations", []) if c.get("Id") != new["Id"]]
cfg["LambdaFunctionConfigurations"] = lambdas + [new]
print(json.dumps(cfg))
PY
)
aws s3api put-bucket-notification-configuration --bucket "$SRC_BUCKET" \
  --notification-configuration "$MERGED"
echo "Trigger attached: s3://$SRC_BUCKET/$SRC_PREFIX*.xlsx -> $FN_ARN"
