---
name: ag0-billing
description: "Process AG0 billing Excel workbooks for Argentina month-end billing runs. Use when the user provides a source billing workbook plus the AG0 working workbook and wants Codex to verify the monthly partition across Invoice Data, Manual, Credit Notes Data, and HOLD, preserve or validate Org Codes control values, refresh Excel queries in the working file, and export import/review CSVs from the CSV sheet."
---

# AG0 Billing

Use this skill for the AG0 Argentina billing workflow that turns a prepared monthly source workbook and the AG0 working workbook into refreshed CSV exports.

## Requirements

- Run on Windows with Excel Desktop available through COM automation.
- Prefer the bundled PowerShell script for deterministic runs.
- Keep the original workbooks untouched by default; work on `-codex-run` copies.

## Workflow

1. Confirm the two workbook paths and the source month sheet name.
2. Run `scripts/export_ag0_billing.ps1` with safe-copy defaults.
3. Preserve current `Org Codes` values unless the user explicitly asks to change them.
4. Treat `Manual`, `Credit Notes Data`, and `HOLD` as exception tabs outside the generated CSV path unless the user explicitly asks for a different policy.
5. Return the generated CSV paths plus any validation caveats from the `CSV` sheet.

## Script

Run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/export_ag0_billing.ps1 `
  -SourceWorkbookPath "C:\path\to\source.xlsx" `
  -WorkingWorkbookPath "C:\path\to\working.xlsx" `
  -SourceSheetName "Mar 26"
```

Useful optional parameters:

- `-OutputDirectory` to place safe copies and CSVs in a different folder
- `-OutputPrefix` to control the CSV filenames
- `-ExpectedPostDate`, `-ExpectedInvoiceDate`, `-ExpectedBranch`, `-ExpectedDept`, `-ExpectedExchangeRate` to validate `Org Codes`
- `-ExpectedInvoiceCount` to assert `CSV!B6`
- `-AllowHoldRows` only if the user explicitly wants to proceed with rows in `HOLD`

## Validation Rules

Require these checks after refresh:

- `CSV!B2 = 0`
- `CSV!D2 = Good`
- `CSV!D4 = Good`
- `CSV!D5 = Good - excluded $0 and CN`

If `-ExpectedInvoiceCount` is provided, also require `CSV!B6` to match it.

Do not fail the run solely because the `ARS` summary row shows a small rounding-difference check in row 3. Surface that as a caveat in the final response instead.

## Exports

- Import CSV: `CSV!A8:Q(last non-empty row in column A)`
- Review CSV: `CSV!A1:Q(last non-empty row in column A)`

The bundled script exports displayed sheet text, not raw numeric values, so fields like `0350.210` stay formatted exactly as shown in Excel.
