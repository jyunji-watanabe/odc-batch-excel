#!/usr/bin/env bash
# Builds the ODC-ready zip: the publish output (.dll + .deps.json + dependencies) at the zip root.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PUBLISH_DIR="$ROOT/artifacts/publish"
ZIP_PATH="$ROOT/artifacts/OdcBatchExcel.zip"

rm -rf "$PUBLISH_DIR" "$ZIP_PATH"
dotnet publish "$ROOT/src/OdcBatchExcel/OdcBatchExcel.csproj" -c Release -o "$PUBLISH_DIR"

(cd "$PUBLISH_DIR" && zip -q -r "$ZIP_PATH" . -x '*.pdb')

ENTRIES="$(unzip -Z1 "$ZIP_PATH")"
for required in OdcBatchExcel.dll OdcBatchExcel.deps.json; do
  if ! grep -qx "$required" <<< "$ENTRIES"; then
    echo "error: $required is missing from the zip root" >&2
    exit 1
  fi
done

SIZE=$(wc -c < "$ZIP_PATH" | tr -d ' ')
if [ "$SIZE" -gt $((50 * 1024 * 1024)) ]; then
  echo "error: zip is $SIZE bytes; ODC's limit is 50 MB" >&2
  exit 1
fi

unzip -l "$ZIP_PATH"
echo "Package ready: $ZIP_PATH"
