# Public XLSX Corpus

This directory holds a reproducible public spreadsheet corpus for regression
testing `xlsx-review`.

Run:

```bash
make corpus-download
make corpus-smoke
make corpus-check
```

Outputs:

- `files/`: downloaded `.xlsx` spreadsheets, grouped by source repository
- `manifest.tsv`: provenance for each spreadsheet, including repo URL, commit,
  upstream path, file size, and SHA-256
- `reports/`: generated TSV reports for corpus runs
- `suites/`: curated subsets for smoke tests and targeted debugging

Current sources:

- `apache-poi`: `test-data/spreadsheet` from `apache/poi` (Apache-2.0)
- `closedxml`: `ClosedXML.Tests/Resource/TryToLoad` from `ClosedXML/ClosedXML` (MIT)
- `openxml-sdk`: `test/DocumentFormat.OpenXml.Tests.Assets` from `dotnet/Open-XML-SDK` (MIT)

Notes:

- The corpus intentionally includes edge cases from test suites, including
  malformed and robustness-focused `.xlsx` files.
- `.xlsx` files are ignored by git in this repo, so the manifest and downloader
  stay versioned while the downloaded corpus remains local.
