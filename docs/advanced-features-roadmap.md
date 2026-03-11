# Advanced Feature Roadmap

`xlsx-review` should distinguish between four support levels for Excel features:

- `read`: inspect and report the feature in `--read --json`
- `edit`: modify the feature in place on an existing workbook
- `create`: generate the feature when using `--create`
- `diff`: detect semantic changes to the feature across workbooks

## Current Baseline

Already supported well:

- cells and basic formulas
- row and column insertion/deletion
- worksheet add/rename/delete
- legacy comments
- workbook creation from blank or template

Already readable, but not fully editable/diffable across the board:

- hidden and very hidden sheets
- chart sheets and dialog sheets
- defined names
- workbook protection
- sheet protection
- external links
- shared, array, and data-table formula counts
- tables
- data validation
- conditional formatting
- pivot-table counts
- threaded comment counts

## Phase Order

### Phase 1: Workbook Metadata Control

Status: shipped

Deliverables:

- `set_sheet_visibility`
- `set_defined_name`
- `delete_defined_name`
- `set_workbook_protection`
- `set_sheet_protection`
- local read-back smoke coverage for edit and template-create flows

Reason:

These features already exist in read mode, are high-value in real workbooks, and do not require complex drawing, pivot, or style graph manipulation.

### Phase 2: Metadata Diff Parity

Status: shipped

Deliverables:

- diff sheet visibility changes
- diff workbook protection changes
- diff sheet protection changes
- diff defined-name additions, deletions, and modifications

Reason:

Once the tool can edit workbook metadata, `--diff` needs to stop pretending those changes do not exist.

### Phase 3: Worksheet UX Features

Status: in progress

Deliverables:

- merged cells
- hyperlinks
- freeze panes and selections
- auto-filter and sort state
- print area and page setup

Reason:

These are common workbook behaviors that users perceive as part of the document, not as secondary metadata.

### Phase 4: Tables and Rules

Deliverables:

- table create/update/delete
- data validation create/update/delete
- conditional formatting create/update/delete

Reason:

This is the first tranche where manifests need richer payloads and stronger validation, so it should come after the metadata foundation is stable.

### Phase 5: Formula and Calculation Fidelity

Deliverables:

- shared formula authoring
- array formula authoring
- data-table formula authoring
- calculation chain handling
- recalc flags and workbook calculation properties

Reason:

Formula correctness matters, but it is safer to add after table and validation support because those features often interact with references and sheet structure.

### Phase 6: Rich Objects

Deliverables:

- charts and chart sheets
- drawings, images, and shapes
- pivot tables and pivot caches
- slicers and timelines

Reason:

These features span multiple interrelated package parts and need stronger fixture coverage than earlier phases.

### Phase 7: Connected and Enterprise Features

Deliverables:

- external data connections and query tables
- macro-aware workflows and `.xlsm` handling rules
- workbook revisions and collaboration metadata

Reason:

These are high-risk features with a larger preservation burden and should only land after the core workbook-editing model is trustworthy.

## Test Strategy

- Keep bundled examples for deterministic edit/create smoke tests.
- Keep corpus feature smoke for read-mode detection across public files.
- Add one local read-back feature suite per tranche under `testdata/`.
- Only extend the default `make smoke` target with features that have stable local fixtures.
