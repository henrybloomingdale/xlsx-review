#!/usr/bin/env bash

set -uo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CORPUS_DIR="${ROOT_DIR}/testdata/public-xlsx-corpus/files"
REPORT_DIR="${ROOT_DIR}/testdata/public-xlsx-corpus/reports"
BINARY_PATH="${BINARY:-}"
MODE="read"
LIMIT=""
SOURCE_FILTER=""
SUITE_PATH=""
REPORT_PREFIX=""
STRICT_MODE=0
TIMEOUT_SECONDS=30

usage() {
  cat <<'EOF'
Usage: ./scripts/run_public_corpus_check.sh [options]

Options:
  --binary <path>     Binary to run. Defaults to build/xlsx-review or a bin/Release executable.
  --mode <mode>       Check mode: read or textconv. Default: read
  --limit <n>         Only check the first n files after filtering
  --source <name>     Restrict to one source subtree (apache-poi, closedxml, openxml-sdk)
  --suite <path>      Text file with corpus-relative workbook paths to check
  --report-prefix <x> Prefix for report filenames. Defaults to the mode name.
  --report-dir <dir>  Output directory for TSV reports
  --timeout <sec>     Per-workbook timeout in seconds. Default: 30
  --strict            Exit non-zero if any files fail
  -h, --help          Show help
EOF
}

while (($# > 0)); do
  case "$1" in
    --binary)
      BINARY_PATH="$2"
      shift 2
      ;;
    --mode)
      MODE="$2"
      shift 2
      ;;
    --limit)
      LIMIT="$2"
      shift 2
      ;;
    --source)
      SOURCE_FILTER="$2"
      shift 2
      ;;
    --suite)
      SUITE_PATH="$2"
      shift 2
      ;;
    --report-prefix)
      REPORT_PREFIX="$2"
      shift 2
      ;;
    --report-dir)
      REPORT_DIR="$2"
      shift 2
      ;;
    --timeout)
      TIMEOUT_SECONDS="$2"
      shift 2
      ;;
    --strict)
      STRICT_MODE=1
      shift
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

if [[ -z "${BINARY_PATH}" ]]; then
  if [[ -x "${ROOT_DIR}/build/xlsx-review" ]]; then
    BINARY_PATH="${ROOT_DIR}/build/xlsx-review"
  else
    BINARY_PATH="$(find "${ROOT_DIR}/bin/Release" -type f -name xlsx-review -perm -111 2>/dev/null | sort | head -n 1)"
  fi
fi

if [[ -z "${BINARY_PATH}" || ! -x "${BINARY_PATH}" ]]; then
  echo "No executable xlsx-review binary found. Build first or pass --binary." >&2
  exit 1
fi

if [[ ! -d "${CORPUS_DIR}" ]]; then
  echo "Corpus directory not found: ${CORPUS_DIR}" >&2
  echo "Run ./scripts/download_public_corpus.sh first." >&2
  exit 1
fi

if [[ -n "${SUITE_PATH}" && ! -f "${SUITE_PATH}" ]]; then
  echo "Suite file not found: ${SUITE_PATH}" >&2
  exit 1
fi

case "${MODE}" in
  read)
    ;;
  textconv)
    ;;
  *)
    echo "Unsupported mode: ${MODE}" >&2
    exit 1
    ;;
esac

mkdir -p "${REPORT_DIR}"

if [[ -z "${REPORT_PREFIX}" ]]; then
  REPORT_PREFIX="${MODE}"
fi

STATUS_PATH="${REPORT_DIR}/${REPORT_PREFIX}_status.tsv"
FAILURES_PATH="${REPORT_DIR}/${REPORT_PREFIX}_failures.tsv"
SUMMARY_PATH="${REPORT_DIR}/${REPORT_PREFIX}_summary.tsv"
REASONS_PATH="${REPORT_DIR}/${REPORT_PREFIX}_reasons.tsv"

printf 'source\tpath\tstatus\terror\n' > "${STATUS_PATH}"

tmp_err="$(mktemp)"
tmp_files="$(mktemp)"
cleanup() {
  rm -f "${tmp_err}" "${tmp_files}"
}
trap cleanup EXIT

run_checked_command() {
  python3 - "$TIMEOUT_SECONDS" "${tmp_err}" "$@" <<'PY'
import subprocess
import sys

timeout = float(sys.argv[1])
stderr_path = sys.argv[2]
command = sys.argv[3:]

with open(stderr_path, "w", encoding="utf-8") as stderr:
    try:
        completed = subprocess.run(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=stderr,
            timeout=timeout,
            check=False,
            text=True,
        )
    except subprocess.TimeoutExpired:
        stderr.write(f"Timed out after {timeout:g}s\n")
        raise SystemExit(124)

raise SystemExit(completed.returncode)
PY
}

if [[ -n "${SUITE_PATH}" ]]; then
  awk '!/^[[:space:]]*(#|$)/ { print $0 }' "${SUITE_PATH}" | while IFS= read -r relative_path; do
    printf '%s/%s\n' "${CORPUS_DIR}" "${relative_path}"
  done > "${tmp_files}"
else
  find "${CORPUS_DIR}" -type f -iname '*.xlsx' | sort > "${tmp_files}"
fi

if [[ -n "${SOURCE_FILTER}" ]]; then
  awk -v source="${SOURCE_FILTER}" 'index($0, "/" source "/") > 0' "${tmp_files}" > "${tmp_files}.filtered"
  mv "${tmp_files}.filtered" "${tmp_files}"
fi

if [[ -n "${LIMIT}" ]]; then
  awk -v limit="${LIMIT}" 'NR <= limit' "${tmp_files}" > "${tmp_files}.limited"
  mv "${tmp_files}.limited" "${tmp_files}"
fi

while IFS= read -r file_path; do
  [[ -n "${file_path}" ]] || continue

  relative_path="${file_path#${CORPUS_DIR}/}"
  source_name="${relative_path%%/*}"

  if [[ ! -f "${file_path}" ]]; then
    printf '%s\t%s\tfail\tMissing file\n' "${source_name}" "${relative_path}" >> "${STATUS_PATH}"
    continue
  fi

  if [[ "${MODE}" == "read" ]]; then
    run_checked_command "${BINARY_PATH}" "${file_path}" --read --json
    command_status=$?

    if [[ ${command_status} -eq 0 ]]; then
      printf '%s\t%s\tok\t\n' "${source_name}" "${relative_path}" >> "${STATUS_PATH}"
    else
      error_message="$(tr '\n' ' ' < "${tmp_err}" | tr '\t' ' ' | sed 's/[[:space:]]\+/ /g; s/^ *//; s/ *$//')"
      if [[ -z "${error_message}" ]]; then
        error_message="Process exited with status ${command_status}"
      fi
      printf '%s\t%s\tfail\t%s\n' "${source_name}" "${relative_path}" "${error_message}" >> "${STATUS_PATH}"
    fi
  else
    run_checked_command "${BINARY_PATH}" --textconv "${file_path}"
    command_status=$?

    if [[ ${command_status} -eq 0 ]]; then
      printf '%s\t%s\tok\t\n' "${source_name}" "${relative_path}" >> "${STATUS_PATH}"
    else
      error_message="$(tr '\n' ' ' < "${tmp_err}" | tr '\t' ' ' | sed 's/[[:space:]]\+/ /g; s/^ *//; s/ *$//')"
      if [[ -z "${error_message}" ]]; then
        error_message="Process exited with status ${command_status}"
      fi
      printf '%s\t%s\tfail\t%s\n' "${source_name}" "${relative_path}" "${error_message}" >> "${STATUS_PATH}"
    fi
  fi
done < "${tmp_files}"

awk -F '\t' '
BEGIN {
  OFS = "\t"
}
NR == 1 {
  header = $0
  next
}
{
  rows[$2] = $0
}
END {
  print header
  for (path in rows)
    print rows[path]
}
' "${STATUS_PATH}" > "${STATUS_PATH}.dedup"
{
  head -n 1 "${STATUS_PATH}.dedup"
  tail -n +2 "${STATUS_PATH}.dedup" | sort -t "$(printf '\t')" -k1,1 -k2,2
} > "${STATUS_PATH}.tmp"
mv "${STATUS_PATH}.tmp" "${STATUS_PATH}"
rm -f "${STATUS_PATH}.dedup"

printf 'scope\tname\ttotal\tok\tfail\n' > "${SUMMARY_PATH}"
awk -F '\t' '
NR > 1 {
  total_all++
  total[$1]++
  if ($3 == "ok") {
    ok_all++
    ok[$1]++
  } else {
    fail_all++
    fail[$1]++
  }
}
END {
  printf "all\tall\t%d\t%d\t%d\n", total_all, ok_all, fail_all
  for (source in total)
    printf "source\t%s\t%d\t%d\t%d\n", source, total[source], ok[source] + 0, fail[source] + 0
}
' "${STATUS_PATH}" | sort >> "${SUMMARY_PATH}"

awk -F '\t' 'NR == 1 || $3 == "fail"' "${STATUS_PATH}" > "${FAILURES_PATH}"

printf 'count\terror\n' > "${REASONS_PATH}"
awk -F '\t' 'NR > 1 && $3 == "fail" { count[$4]++ } END { for (reason in count) printf "%d\t%s\n", count[reason], reason }' "${STATUS_PATH}" | sort -nr >> "${REASONS_PATH}"

failure_count="$(awk -F '\t' 'NR > 1 && $3 == "fail" { count++ } END { print count + 0 }' "${STATUS_PATH}")"

echo "Binary: ${BINARY_PATH}"
echo "Mode: ${MODE}"
echo "Status report: ${STATUS_PATH}"
echo "Failures: ${FAILURES_PATH}"
echo "Summary: ${SUMMARY_PATH}"
echo "Reasons: ${REASONS_PATH}"
echo ""
cat "${SUMMARY_PATH}"

if [[ ${STRICT_MODE} -eq 1 && "${failure_count}" -gt 0 ]]; then
  exit 1
fi
