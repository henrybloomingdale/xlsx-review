#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
WORKBOOK_ROOT="${ROOT_DIR}/testdata/public-xlsx-corpus/files"
BINARY_PATH="${BINARY:-${ROOT_DIR}/build/xlsx-review}"
SUITE_PATH="${ROOT_DIR}/testdata/public-xlsx-corpus/suites/read-feature-smoke.tsv"

usage() {
  cat <<'EOF'
Usage: ./scripts/run_feature_smoke.sh [options]

Options:
  --binary <path>   Binary to run. Defaults to build/xlsx-review
  --root <path>     Root directory used to resolve relative workbook paths in the suite
  --suite <path>    TSV file with feature assertions
  -h, --help        Show help
EOF
}

while (($# > 0)); do
  case "$1" in
    --binary)
      BINARY_PATH="$2"
      shift 2
      ;;
    --root)
      WORKBOOK_ROOT="$2"
      shift 2
      ;;
    --suite)
      SUITE_PATH="$2"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage >&2
      exit 1
      ;;
  esac
done

if [[ ! -x "${BINARY_PATH}" ]]; then
  echo "Binary not found or not executable: ${BINARY_PATH}" >&2
  exit 1
fi

if [[ ! -f "${SUITE_PATH}" ]]; then
  echo "Suite file not found: ${SUITE_PATH}" >&2
  exit 1
fi

if [[ ! -d "${WORKBOOK_ROOT}" ]]; then
  echo "Workbook root not found: ${WORKBOOK_ROOT}" >&2
  exit 1
fi

python3 - "${BINARY_PATH}" "${WORKBOOK_ROOT}" "${SUITE_PATH}" <<'PY'
import json
import subprocess
import sys
from collections import OrderedDict
from pathlib import Path


def parse_expected(raw: str):
    for op in (">=", "<=", ">", "<", "="):
        if raw.startswith(op):
            return op, coerce_scalar(raw[len(op):])
    if raw.startswith("contains:"):
        return "contains", raw[len("contains:"):]
    return "=", coerce_scalar(raw)


def coerce_scalar(raw: str):
    text = raw.strip()
    lowered = text.lower()
    if lowered == "null":
        return None
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    try:
        return int(text)
    except ValueError:
        return text


def lookup_value(payload, field: str):
    if field == "warning_count":
        return len(payload.get("warnings", []))

    if field.startswith("sheet:"):
        sheet_key = field[len("sheet:"):]
        sheet_name, remainder = sheet_key.split(".", 1)
        for sheet in payload.get("sheets", []):
            if sheet.get("name") == sheet_name:
                return walk(sheet, remainder)
        raise KeyError(f"sheet '{sheet_name}' not found")

    return walk(payload, field)


def walk(value, path: str):
    current = value
    for segment in path.split("."):
        if isinstance(current, list):
            current = current[int(segment)]
            continue
        current = current[segment]
    return current


def compare(actual, op: str, expected):
    if op == "=":
        return actual == expected
    if op == ">=":
        return actual >= expected
    if op == "<=":
        return actual <= expected
    if op == ">":
        return actual > expected
    if op == "<":
        return actual < expected
    if op == "contains":
        return str(expected) in str(actual)
    raise ValueError(f"unsupported operator: {op}")


binary_path = Path(sys.argv[1])
corpus_dir = Path(sys.argv[2])
suite_path = Path(sys.argv[3])

assertions = OrderedDict()
with suite_path.open("r", encoding="utf-8") as handle:
    for lineno, raw_line in enumerate(handle, start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        parts = raw_line.rstrip("\n").split("\t")
        if len(parts) != 3:
            raise SystemExit(f"Invalid suite line {lineno}: expected 3 tab-separated fields")

        relative_path, field, expected = parts
        assertions.setdefault(relative_path, []).append((field, expected))

failures = []

for relative_path, checks in assertions.items():
    file_path = corpus_dir / relative_path
    completed = subprocess.run(
        [str(binary_path), str(file_path), "--read", "--json"],
        capture_output=True,
        text=True,
        check=False,
    )

    if completed.returncode != 0:
        failures.append(f"{relative_path}: read failed: {completed.stderr.strip()}")
        continue

    payload = json.loads(completed.stdout)
    for field, expected_raw in checks:
        try:
            actual = lookup_value(payload, field)
        except Exception as exc:  # noqa: BLE001
            failures.append(f"{relative_path}: could not resolve '{field}': {exc}")
            continue

        op, expected = parse_expected(expected_raw)
        if not compare(actual, op, expected):
            failures.append(
                f"{relative_path}: assertion failed for '{field}': actual={actual!r}, expected {op} {expected!r}"
            )

if failures:
    print("Feature smoke failed:", file=sys.stderr)
    for failure in failures:
        print(f"  - {failure}", file=sys.stderr)
    raise SystemExit(1)

print(f"Feature smoke passed ({len(assertions)} workbooks, {sum(len(v) for v in assertions.values())} assertions).")
PY
