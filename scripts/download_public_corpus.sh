#!/usr/bin/env bash

set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CORPUS_DIR="${ROOT_DIR}/testdata/public-xlsx-corpus"
FILES_DIR="${CORPUS_DIR}/files"
TMP_DIR="$(mktemp -d)"

cleanup() {
  rm -rf "${TMP_DIR}"
}
trap cleanup EXIT

mkdir -p "${FILES_DIR}"

MANIFEST_PATH="${CORPUS_DIR}/manifest.tsv"
cat > "${MANIFEST_PATH}" <<'EOF'
source	repo	commit	license	upstream_path	stored_path	bytes	sha256
EOF

copy_repo_files() {
  local source_name="$1"
  local repo_url="$2"
  local sparse_path="$3"
  local license_id="$4"
  local repo_dir="${TMP_DIR}/${source_name}"

  echo "==> Cloning ${repo_url}"
  git clone --depth 1 --filter=blob:none --sparse "${repo_url}" "${repo_dir}" >/dev/null
  (
    cd "${repo_dir}"
    git sparse-checkout set "${sparse_path}" >/dev/null
    local commit
    commit="$(git rev-parse HEAD)"

    find "${sparse_path}" -type f -iname '*.xlsx' -print0 | while IFS= read -r -d '' file; do
      local relative_path="${file#${sparse_path}/}"
      local stored_path="files/${source_name}/${relative_path}"
      local destination="${CORPUS_DIR}/${stored_path}"
      local bytes
      local sha256

      mkdir -p "$(dirname "${destination}")"
      cp "${file}" "${destination}"

      bytes="$(wc -c < "${destination}" | tr -d '[:space:]')"
      sha256="$(shasum -a 256 "${destination}" | awk '{print $1}')"

      printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
        "${source_name}" \
        "${repo_url}" \
        "${commit}" \
        "${license_id}" \
        "${file}" \
        "${stored_path}" \
        "${bytes}" \
        "${sha256}" >> "${MANIFEST_PATH}"
    done
  )
}

copy_repo_files "apache-poi" "https://github.com/apache/poi" "test-data/spreadsheet" "Apache-2.0"
copy_repo_files "closedxml" "https://github.com/ClosedXML/ClosedXML" "ClosedXML.Tests/Resource/TryToLoad" "MIT"
copy_repo_files "openxml-sdk" "https://github.com/dotnet/Open-XML-SDK" "test/DocumentFormat.OpenXml.Tests.Assets" "MIT"

file_count="$(tail -n +2 "${MANIFEST_PATH}" | wc -l | tr -d '[:space:]')"
total_bytes="$(awk -F '\t' 'NR > 1 { sum += $7 } END { print sum + 0 }' "${MANIFEST_PATH}")"

echo ""
echo "Downloaded ${file_count} spreadsheets into ${FILES_DIR}"
echo "Manifest written to ${MANIFEST_PATH}"
echo "Total bytes: ${total_bytes}"
