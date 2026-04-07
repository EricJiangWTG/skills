param(
    [Parameter(Mandatory = $true)]
    [string]$SourceWorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$WorkingWorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$SourceSheetName,

    [string]$OutputDirectory,
    [string]$OutputPrefix,
    [string]$ExpectedPostDate,
    [string]$ExpectedInvoiceDate,
    [string]$ExpectedBranch,
    [string]$ExpectedDept,
    [string]$ExpectedExchangeRate,
    [int]$ExpectedInvoiceCount = 0,
    [switch]$AllowHoldRows
)

$ErrorActionPreference = 'Stop'

function Release-ComObject($obj) {
    if ($null -ne $obj) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
        } catch {
        }
    }
}

function Get-LastNonEmptyRowInColumnA($ws) {
    $lastRow = 0
    $usedRows = $ws.UsedRange.Rows.Count
    for ($r = 1; $r -le $usedRows; $r++) {
        $text = [string]$ws.Cells.Item($r, 1).Text
        if ($text.Trim() -ne '') {
            $lastRow = $r
        }
    }
    return $lastRow
}

function Get-RowSignatures($ws, [int]$startRow, [int]$endRow, [int]$startCol, [int]$endCol) {
    $items = New-Object System.Collections.Generic.List[string]
    for ($r = $startRow; $r -le $endRow; $r++) {
        $cells = @()
        $hasValue = $false
        for ($c = $startCol; $c -le $endCol; $c++) {
            $v = $ws.Cells.Item($r, $c).Value2
            $text = if ($null -eq $v) { '' } else { [string]$v }
            if ($text -ne '') {
                $hasValue = $true
            }
            $cells += $text.Trim()
        }
        if ($hasValue) {
            $items.Add(($cells -join '|'))
        }
    }
    return $items
}

function Compare-Multiset($left, $right) {
    $l = @{}
    foreach ($item in $left) {
        if ($l.ContainsKey($item)) {
            $l[$item]++
        } else {
            $l[$item] = 1
        }
    }

    $r = @{}
    foreach ($item in $right) {
        if ($r.ContainsKey($item)) {
            $r[$item]++
        } else {
            $r[$item] = 1
        }
    }

    if ($l.Count -ne $r.Count) {
        return $false
    }

    foreach ($key in $l.Keys) {
        if (-not $r.ContainsKey($key)) {
            return $false
        }
        if ($r[$key] -ne $l[$key]) {
            return $false
        }
    }

    return $true
}

function Get-CellText($ws, $address) {
    return [string]$ws.Range($address).Text
}

function ConvertTo-CsvField([string]$value) {
    if ($null -eq $value) {
        return ''
    }
    $escaped = $value.Replace('"', '""')
    if ($escaped.Contains(',') -or $escaped.Contains('"') -or $escaped.Contains("`r") -or $escaped.Contains("`n")) {
        return '"' + $escaped + '"'
    }
    return $escaped
}

function Export-RangeTextToUtf8Csv($ws, [int]$startRow, [int]$endRow, [int]$startCol, [int]$endCol, [string]$csvPath) {
    $lines = New-Object System.Collections.Generic.List[string]
    for ($r = $startRow; $r -le $endRow; $r++) {
        $fields = @()
        for ($c = $startCol; $c -le $endCol; $c++) {
            $fields += ConvertTo-CsvField ([string]$ws.Cells.Item($r, $c).Text)
        }
        $lines.Add(($fields -join ','))
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllLines($csvPath, $lines, $utf8NoBom)
}

if (-not (Test-Path -LiteralPath $SourceWorkbookPath)) {
    throw "Source workbook not found: $SourceWorkbookPath"
}
if (-not (Test-Path -LiteralPath $WorkingWorkbookPath)) {
    throw "Working workbook not found: $WorkingWorkbookPath"
}

$sourceItem = Get-Item -LiteralPath $SourceWorkbookPath
$workingItem = Get-Item -LiteralPath $WorkingWorkbookPath

if (-not $OutputDirectory) {
    $OutputDirectory = Split-Path -Parent $WorkingWorkbookPath
}
if (-not $OutputPrefix) {
    $OutputPrefix = [System.IO.Path]::GetFileNameWithoutExtension($SourceWorkbookPath)
}

$sourceCopy = Join-Path $OutputDirectory ($sourceItem.BaseName + '-codex-run' + $sourceItem.Extension)
$workingCopy = Join-Path $OutputDirectory ($workingItem.BaseName + '-codex-run' + $workingItem.Extension)
$importCsv = Join-Path $OutputDirectory ($OutputPrefix + '_import.csv')
$reviewCsv = Join-Path $OutputDirectory ($OutputPrefix + '_review.csv')

$beforeHash = Get-FileHash -LiteralPath $SourceWorkbookPath, $WorkingWorkbookPath -Algorithm SHA256 | Select-Object Path, Hash

foreach ($path in @($sourceCopy, $workingCopy, $importCsv, $reviewCsv)) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

Copy-Item -LiteralPath $SourceWorkbookPath -Destination $sourceCopy
Copy-Item -LiteralPath $WorkingWorkbookPath -Destination $workingCopy

$excel = $null
$sourceWb = $null
$workingWb = $null
$srcWs = $null
$invWs = $null
$manualWs = $null
$cnWs = $null
$holdWs = $null
$orgWs = $null
$csvWs = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3

    $sourceWb = $excel.Workbooks.Open($sourceCopy, 0, $true)
    $workingWb = $excel.Workbooks.Open($workingCopy, 0, $false)

    $srcWs = $sourceWb.Worksheets.Item($SourceSheetName)
    $invWs = $workingWb.Worksheets.Item('Invoice Data')
    $manualWs = $workingWb.Worksheets.Item('Manual')
    $cnWs = $workingWb.Worksheets.Item('Credit Notes Data')
    $holdWs = $workingWb.Worksheets.Item('HOLD')
    $orgWs = $workingWb.Worksheets.Item('Org Codes')
    $csvWs = $workingWb.Worksheets.Item('CSV')

    $sourceRows = Get-RowSignatures $srcWs 2 (Get-LastNonEmptyRowInColumnA $srcWs) 1 13
    $workingRows = New-Object System.Collections.Generic.List[string]
    foreach ($item in (Get-RowSignatures $invWs 2 (Get-LastNonEmptyRowInColumnA $invWs) 1 13)) {
        $workingRows.Add($item)
    }
    foreach ($item in (Get-RowSignatures $manualWs 2 $manualWs.UsedRange.Rows.Count 2 14)) {
        $workingRows.Add($item)
    }
    foreach ($item in (Get-RowSignatures $cnWs 2 $cnWs.UsedRange.Rows.Count 2 14)) {
        $workingRows.Add($item)
    }
    $holdRows = Get-RowSignatures $holdWs 2 $holdWs.UsedRange.Rows.Count 2 14

    if (-not $AllowHoldRows -and $holdRows.Count -ne 0) {
        throw 'HOLD sheet is not empty; stopping rather than guessing how to classify rows.'
    }
    if ($workingRows.Count -ne $sourceRows.Count) {
        throw "Prepared working partition count ($($workingRows.Count)) does not match source count ($($sourceRows.Count))."
    }
    if (-not (Compare-Multiset $sourceRows $workingRows)) {
        throw 'Prepared working rows do not match the source rows after combining Invoice Data, Manual, and Credit Notes Data.'
    }

    $expectedOrg = [ordered]@{}
    if ($ExpectedPostDate) { $expectedOrg['C3'] = $ExpectedPostDate }
    if ($ExpectedInvoiceDate) { $expectedOrg['D3'] = $ExpectedInvoiceDate }
    if ($ExpectedBranch) { $expectedOrg['C7'] = $ExpectedBranch }
    if ($ExpectedDept) { $expectedOrg['D7'] = $ExpectedDept }
    if ($ExpectedExchangeRate) { $expectedOrg['E7'] = $ExpectedExchangeRate }

    foreach ($addr in $expectedOrg.Keys) {
        $actual = Get-CellText $orgWs $addr
        if ($actual -ne $expectedOrg[$addr]) {
            throw "Org Codes $addr expected '$($expectedOrg[$addr])' but found '$actual'."
        }
    }

    $workingWb.RefreshAll()
    try {
        $excel.CalculateUntilAsyncQueriesDone()
    } catch {
    }
    $excel.CalculateFullRebuild()
    Start-Sleep -Seconds 5

    $lastCsvRow = Get-LastNonEmptyRowInColumnA $csvWs
    if ($lastCsvRow -lt 10) {
        throw "CSV payload boundary is unexpectedly small: row $lastCsvRow."
    }

    $validations = [ordered]@{
        'B2' = '0'
        'D2' = 'Good'
        'D4' = 'Good'
        'D5' = 'Good - excluded $0 and CN'
    }
    foreach ($addr in $validations.Keys) {
        $actual = Get-CellText $csvWs $addr
        if ($actual -ne $validations[$addr]) {
            throw "Validation failed at CSV!$addr. Expected '$($validations[$addr])' but found '$actual'."
        }
    }
    if ($ExpectedInvoiceCount -gt 0) {
        $actualCount = Get-CellText $csvWs 'B6'
        if ($actualCount -ne [string]$ExpectedInvoiceCount) {
            throw "Validation failed at CSV!B6. Expected '$ExpectedInvoiceCount' but found '$actualCount'."
        }
    }

    $csvWs.Columns('A:Q').AutoFit() | Out-Null
    Export-RangeTextToUtf8Csv $csvWs 8 $lastCsvRow 1 17 $importCsv
    Export-RangeTextToUtf8Csv $csvWs 1 $lastCsvRow 1 17 $reviewCsv
    $roundingCaveat = Get-CellText $csvWs 'D3'

    $workingWb.Save()
    $workingWb.Close($true)
    $workingWb = $null
    $sourceWb.Close($false)
    $sourceWb = $null

    $afterHash = Get-FileHash -LiteralPath $SourceWorkbookPath, $WorkingWorkbookPath -Algorithm SHA256 | Select-Object Path, Hash
    $originalsUnchanged = $true
    for ($i = 0; $i -lt $beforeHash.Count; $i++) {
        if ($beforeHash[$i].Hash -ne $afterHash[$i].Hash) {
            $originalsUnchanged = $false
        }
    }

    [pscustomobject]@{
        source_copy = $sourceCopy
        working_copy = $workingCopy
        import_csv = $importCsv
        review_csv = $reviewCsv
        export_encoding = 'UTF8'
        csv_last_row = $lastCsvRow
        import_line_count = (Get-Content -LiteralPath $importCsv).Count
        review_line_count = (Get-Content -LiteralPath $reviewCsv).Count
        originals_unchanged = $originalsUnchanged
        rounding_caveat = $roundingCaveat
    } | ConvertTo-Json -Depth 4
} finally {
    if ($null -ne $sourceWb) {
        try {
            $sourceWb.Close($false)
        } catch {
        }
    }
    if ($null -ne $workingWb) {
        try {
            $workingWb.Close($false)
        } catch {
        }
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        } catch {
        }
    }

    Release-ComObject $csvWs
    Release-ComObject $orgWs
    Release-ComObject $holdWs
    Release-ComObject $cnWs
    Release-ComObject $manualWs
    Release-ComObject $invWs
    Release-ComObject $srcWs
    Release-ComObject $sourceWb
    Release-ComObject $workingWb
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
