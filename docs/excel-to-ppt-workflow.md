# Excel to editable PowerPoint

This change addresses the November 2026 Excel-to-PPT failure report without
assuming that every empty tool transcript was an execution failure.

## Changes

- Compressed tool receipts explicitly say `compacted` and `payloadOmitted`,
  retain error diagnostics, and preserve a bounded preview for plain-text reads.
  The newest result survives removal of its large call arguments when it fits.
- Excel/data requests select the full editing route instead of text/discovery.
- Shared and PowerPoint tool wrappers validate schemas before executing.
- `xlsx-to-json input.xlsx output.jsonl [sheet] [A1:Q70]` preserves cell
  addresses, cached values, formulas, number formats and stored fills.
  Each invocation covers at most 10,000 cells. Read JSONL with offset/limit.
- Chart edits validate changed series indices/orders, data-point ordering,
  cache indices/values and chart content types before PowerPoint import.
  Existing embedded-workbook references may not silently disappear.
  `require_editable_data=true` requires embedded workbook references for all
  charts in the exported slide package.
- Chart scripts can use `loadZip(bytes)` to patch embedded XLSX, and
  `sharedAxisBounds(values, step?, includeZero?)` /
  `centeredAxisBounds(values, minimumSpan?)` for consistent axes.

## Suggested request

Use the first two slides as references and fill the remaining slides from the
uploaded Excel file. Preserve editable native charts and embedded workbooks.
Map orange stored fills to red text and green stored fills to green text.
Center each mini line chart and give paired bar charts identical axis bounds.
Inspect the template, save the source-to-target mapping, edit one slide and
verify it, then apply the mapping to the remaining slides.

## Validation boundaries

These are targeted package checks, not a complete OOXML schema validator.
Formula caches are not recalculated. Conditional-format colors are not
evaluated; unresolved theme/index colors must not be guessed.
Charts still require visual checks and desktop PowerPoint/WPS Edit Data
verification. Workbook/chart formula-value consistency is a workflow
responsibility; the package validator checks references and archive structure.
No private Excel data or customer presentation is included in tests.

GitHub code changes do not automatically update a previously installed local
add-in. Build this branch using the repository's pnpm commands and load its
development manifest to test. The inherited hosted manifest still points to
the upstream deployment; do not use it to verify these fork changes.
