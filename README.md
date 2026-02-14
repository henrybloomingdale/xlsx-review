# xlsx-review

A CLI tool for **programmatic Excel (.xlsx) editing** using Microsoft's [Open XML SDK](https://github.com/dotnet/Open-XML-SDK). Takes an `.xlsx` file and a JSON edit manifest, produces a modified spreadsheet with highlighted changes and comments — no macros, no compatibility issues.

**Ships as a single ~12MB native binary.** No runtime, no Docker required.

## Why Open XML SDK?

We evaluated three approaches for programmatic spreadsheet editing:

| Approach | Cell Editing | Comments | Formatting |
|----------|:-:|:-:|:-:|
| **Open XML SDK (.NET)** | ✅ 100% | ✅ 100% | ✅ Preserved |
| openpyxl (Python) | ✅ 100% | ⚠️ ~90% | ⚠️ Partial |
| csv manipulation | ❌ No formulas | ❌ None | ❌ Lost |

Open XML SDK is the gold standard — it's Microsoft's own library for manipulating Office documents. Cell values, formulas, formatting, and comments all work correctly.

## Quick Start

### Option 1: Native Binary (recommended)

```bash
git clone https://github.com/henrybloomingdale/xlsx-review.git
cd xlsx-review
make install    # Builds + installs to /usr/local/bin
```

Requires [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) for building (`brew install dotnet@8`). The resulting binary is self-contained — no .NET runtime needed to run it.

### Option 2: Docker

```bash
make docker     # Builds Docker image
docker run --rm -v "$(pwd):/work" -w /work xlsx-review input.xlsx edits.json -o edited.xlsx
```

### Usage

```bash
# Basic usage — edit a spreadsheet
xlsx-review input.xlsx edits.json -o edited.xlsx

# Pipe JSON from stdin
cat edits.json | xlsx-review input.xlsx -o edited.xlsx

# Read spreadsheet contents as JSON
xlsx-review input.xlsx --read --json

# Read spreadsheet (human-readable)
xlsx-review input.xlsx --read

# Custom author name for comments
xlsx-review input.xlsx edits.json -o edited.xlsx --author "Dr. Smith"

# Dry run (validate without modifying)
xlsx-review input.xlsx edits.json --dry-run

# JSON output for pipelines
xlsx-review input.xlsx edits.json -o edited.xlsx --json
```

## JSON Manifest Format

```json
{
  "author": "Reviewer Name",
  "changes": [
    { "type": "set_cell", "sheet": "Sheet1", "cell": "A1", "value": "New Value" },
    { "type": "set_cell", "sheet": "Sheet1", "cell": "B2", "value": "42", "format": "number" },
    { "type": "set_formula", "sheet": "Sheet1", "cell": "C1", "formula": "=SUM(A1:B1)" },
    { "type": "insert_row", "sheet": "Sheet1", "after": 5 },
    { "type": "delete_row", "sheet": "Sheet1", "row": 10 },
    { "type": "insert_column", "sheet": "Sheet1", "after": "C" },
    { "type": "delete_column", "sheet": "Sheet1", "column": "D" },
    { "type": "add_sheet", "name": "Summary" },
    { "type": "rename_sheet", "from": "Sheet1", "to": "Data" },
    { "type": "delete_sheet", "name": "Old Sheet" }
  ],
  "comments": [
    { "sheet": "Sheet1", "cell": "A1", "text": "This value was updated" }
  ]
}
```

### Change Types

| Type | Required Fields | Description |
|------|----------------|-------------|
| `set_cell` | `sheet`, `cell`, `value` | Set cell value. Modified cells are highlighted yellow (#FFFF00). Optional `format`: `"number"` for numeric values. |
| `set_formula` | `sheet`, `cell`, `formula` | Set a cell formula (e.g., `=SUM(A1:B1)`) |
| `insert_row` | `sheet`, `after` | Insert blank row after specified row number |
| `delete_row` | `sheet`, `row` | Delete specified row (shifts rows up) |
| `insert_column` | `sheet`, `after` | Insert blank column after specified column letter |
| `delete_column` | `sheet`, `column` | Delete specified column (shifts columns left) |
| `add_sheet` | `name` | Add a new worksheet |
| `rename_sheet` | `from`, `to` | Rename a worksheet |
| `delete_sheet` | `name` | Delete a worksheet |

### Comment Format

Each comment needs:
- `sheet` — worksheet name (case-sensitive)
- `cell` — cell reference in A1 notation
- `text` — the comment content

Comments are added as legacy Notes for maximum compatibility across Excel versions.

## CLI Flags

| Flag | Description |
|------|-------------|
| `-o`, `--output <path>` | Output file path (default: `<input>_edited.xlsx`) |
| `--author <name>` | Author name for comments (overrides manifest `author`) |
| `--json` | Output results as JSON (for scripting/pipelines) |
| `--dry-run` | Validate the manifest without modifying the spreadsheet |
| `--read` | Read spreadsheet contents (no manifest needed) |
| `-v`, `--version` | Show version |
| `-h`, `--help` | Show help |

## Build Targets

```
make              # Build native binary for current platform (~12MB, self-contained)
make install      # Build + install to /usr/local/bin
make all          # Cross-compile for macOS ARM64, macOS x64, Linux x64, Linux ARM64
make docker       # Build Docker image
make test         # Run test (requires TEST_DOC=path/to/spreadsheet.xlsx)
make clean        # Remove build artifacts
make help         # Show all targets
```

## Exit Codes

- `0` — All changes and comments applied successfully
- `1` — One or more edits failed, or invalid input

## JSON Output Mode

With `--json`, the tool outputs structured results:

```json
{
  "input": "data.xlsx",
  "output": "data_edited.xlsx",
  "author": "Dr. Smith",
  "changes_attempted": 5,
  "changes_succeeded": 5,
  "comments_attempted": 2,
  "comments_succeeded": 2,
  "success": true,
  "results": [
    { "index": 0, "type": "set_cell", "success": true, "message": "Set Sheet1!A1 = \"New Value\"" },
    { "index": 0, "type": "comment", "success": true, "message": "Comment added on Sheet1!A1" }
  ]
}
```

## Read Mode

Extract spreadsheet contents without a manifest:

```bash
xlsx-review input.xlsx --read --json
```

```json
{
  "sheets": [
    {
      "name": "Sheet1",
      "rows": [
        { "row": 1, "cells": [{ "cell": "A1", "value": "Name" }, { "cell": "B1", "value": "Age" }] },
        { "row": 2, "cells": [{ "cell": "A2", "value": "Alice" }, { "cell": "B2", "value": "30" }] }
      ]
    }
  ]
}
```

## How It Works

1. Copies the input `.xlsx` to the output path
2. Opens the spreadsheet using Open XML SDK
3. Applies changes in manifest order (set_cell, insert/delete rows/columns, sheet operations)
4. Highlights modified cells with yellow fill (#FFFF00)
5. Adds legacy comments (Notes) for maximum compatibility
6. Saves and reports results

## Development

```bash
# Build native binary (requires .NET 8 SDK)
make build

# Build and run locally
dotnet run -- input.xlsx edits.json -o edited.xlsx

# Read mode
dotnet run -- input.xlsx --read --json

# Cross-compile all platforms
make all
# → build/osx-arm64/xlsx-review  (macOS Apple Silicon)
# → build/osx-x64/xlsx-review    (macOS Intel)
# → build/linux-x64/xlsx-review  (Linux x64)
# → build/linux-arm64/xlsx-review (Linux ARM64)
```

## License

MIT — see [LICENSE](LICENSE).

---

*Built by [CinciNeuro](https://github.com/henrybloomingdale) for AI-assisted data review workflows.*
