# Public XLSX Corpus

This directory holds a reproducible public spreadsheet corpus for regression
testing `xlsx-review`.

Run:

```bash
make corpus-download
make corpus-smoke
make corpus-feature-smoke
make corpus-check
```

`make corpus-smoke` runs against the packaged single-file binary.
`make corpus-feature-smoke` and `make corpus-check` use the local Release
configuration via `dotnet run --no-build`, which is more stable for longer or
more assertion-heavy corpus execution.

Outputs:

- `files/`: downloaded `.xlsx` spreadsheets, grouped by source repository
- `manifest.tsv`: provenance for each spreadsheet, including repo URL, commit,
  upstream path, file size, and SHA-256
- `reports/`: generated TSV reports for corpus runs
- `suites/`: curated subsets for smoke tests and targeted debugging

Suite intent:

- `suites/read-smoke.txt`: must-pass read-mode gate used in CI.
- `suites/read-feature-smoke.tsv`: must-pass metadata assertions for workbook
  and worksheet features like chart sheets, dialog sheets, hidden sheets,
  shared formulas, validations, protection, pivots, comments, and external
  links.
- `make corpus-check`: tolerant full-corpus baseline for compatibility and
  hostile-input tracking. This is for measuring behavior, not requiring every
  workbook to load successfully.

Current sources:

- `apache-poi`: `test-data/spreadsheet` from `apache/poi` (Apache-2.0)
- `closedxml`: `ClosedXML.Tests/Resource/TryToLoad` from `ClosedXML/ClosedXML` (MIT)
- `openxml-sdk`: `test/DocumentFormat.OpenXml.Tests.Assets` from `dotnet/Open-XML-SDK` (MIT)

Notes:

- The corpus intentionally includes edge cases from test suites, including
  malformed and robustness-focused `.xlsx` files.
- The feature smoke runner executes `--read --json` and asserts specific JSON
  fields, so it catches metadata regressions that basic read pass/fail checks
  would miss.
- `.xlsx` files are ignored by git in this repo, so the manifest and downloader
  stay versioned while the downloaded corpus remains local.
