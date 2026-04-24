param(
    [string]$FolderPath
)

# ---------------------------------------------------------------------------
# Folder
# ---------------------------------------------------------------------------
if (-not $FolderPath) {
    $FolderPath = (Read-Host "`nEnter folder path").Trim('"').Trim("'")
}
$FolderPath = $FolderPath.Trim('"').Trim("'")
if (-not (Test-Path $FolderPath -PathType Container)) {
    Write-Host "Error: '$FolderPath' is not a valid directory." -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# Find Excel files
# ---------------------------------------------------------------------------
$files = @(Get-ChildItem -Path $FolderPath -File |
    Where-Object {
        ($_.Extension -ieq '.xlsx' -or $_.Extension -ieq '.xls') -and
        $_.Name -notmatch '^consolidated_\d{8}_\d{6}\.xlsx$'
    } | Sort-Object Name)

if ($files.Count -eq 0) {
    Write-Host "No .xlsx or .xls files found in the folder." -ForegroundColor Yellow
    exit 0
}

Write-Host "`nFound $($files.Count) Excel file(s):"
foreach ($f in $files) {
    Write-Host ("  - {0}  ({1:F1} MB)" -f $f.Name, ($f.Length / 1MB))
}

# ---------------------------------------------------------------------------
# Start Excel
# ---------------------------------------------------------------------------
try {
    $xl = New-Object -ComObject Excel.Application
} catch {
    Write-Host "`nERROR: Could not start Excel. Is Microsoft Excel installed?" -ForegroundColor Red
    exit 1
}
$xl.Visible        = $false
$xl.DisplayAlerts  = $false
$xl.ScreenUpdating = $false
$xl.EnableEvents   = $false
$xl.Calculation    = -4135   # xlCalculationManual

function ReleaseCom($o) {
    if ($null -ne $o) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch {}
    }
}

# Whole-number floats become Long to prevent invoice-number sig-fig loss
function CoerceVal($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [double] -and [math]::Truncate($v) -eq $v -and
        [math]::Abs($v) -le [long]::MaxValue) {
        return [long]$v
    }
    return $v
}

# ---------------------------------------------------------------------------
# Collect all column headers (one read-only pass)
# ---------------------------------------------------------------------------
Write-Host "`nScanning column headers across all files..."
$seenLower = [System.Collections.Specialized.OrderedDictionary]::new()

foreach ($file in $files) {
    $wb = $null
    try {
        $wb = $xl.Workbooks.Open($file.FullName, 0, $true)
        $sheets = $wb.Worksheets
        foreach ($ws in $sheets) {
            $used = $ws.UsedRange
            if ($null -eq $used) { ReleaseCom $ws; continue }
            $firstRow = $used.Row
            $ncols    = $used.Columns.Count
            $firstCol = $used.Column
            for ($c = 0; $c -lt $ncols; $c++) {
                $v = $ws.Cells($firstRow, $firstCol + $c).Value2
                if ($null -ne $v) {
                    $name = $v.ToString().Trim()
                    $key  = $name.ToLower()
                    if ($name -ne '' -and -not $seenLower.Contains($key)) {
                        $seenLower[$key] = $name
                    }
                }
            }
            ReleaseCom $used
            ReleaseCom $ws
        }
        ReleaseCom $sheets
    } catch {
        Write-Host "  Warning: could not read headers from '$($file.Name)': $_" -ForegroundColor Yellow
    } finally {
        if ($null -ne $wb) { $wb.Close($false); ReleaseCom $wb; $wb = $null }
    }
}

$allCols = @($seenLower.Values)
if ($allCols.Count -eq 0) {
    Write-Host "No column headers found in any file." -ForegroundColor Red
    $xl.Quit(); ReleaseCom $xl; exit 1
}

# ---------------------------------------------------------------------------
# Column selection
# ---------------------------------------------------------------------------
Write-Host "`nUnique columns found across all files:"
for ($i = 0; $i -lt $allCols.Count; $i++) {
    Write-Host ("  {0,4}. {1}" -f ($i + 1), $allCols[$i])
}
$raw = Read-Host "`nEnter column numbers to EXCLUDE (comma-separated), or press Enter to keep all"

$included = [System.Collections.Generic.List[string]]::new()
if ([string]::IsNullOrWhiteSpace($raw)) {
    foreach ($c in $allCols) { $included.Add($c) }
} else {
    $xnums = @{}
    foreach ($tok in $raw.Split(',')) {
        $n = 0
        if ([int]::TryParse($tok.Trim(), [ref]$n)) { $xnums[$n] = $true }
    }
    for ($i = 0; $i -lt $allCols.Count; $i++) {
        if (-not $xnums.ContainsKey($i + 1)) { $included.Add($allCols[$i]) }
    }
    $xd = for ($i = 0; $i -lt $allCols.Count; $i++) {
        if ($xnums.ContainsKey($i + 1)) { $allCols[$i] }
    }
    if ($xd) { Write-Host "Excluding: $($xd -join ', ')" }
}
if ($included.Count -eq 0) {
    Write-Host "No columns remaining after exclusion." -ForegroundColor Yellow
    $xl.Quit(); ReleaseCom $xl; exit 0
}
Write-Host "`nRetaining $($included.Count) column(s)."

$colMap = @{}
for ($i = 0; $i -lt $included.Count; $i++) { $colMap[$included[$i].ToLower()] = $i }

# ---------------------------------------------------------------------------
# Filters
# ---------------------------------------------------------------------------
$filters = [System.Collections.Generic.List[hashtable]]::new()
while ($true) {
    if ((Read-Host "`nFilter by a column value? (y/n)").Trim().ToLower() -ne 'y') { break }

    Write-Host "`nIncluded columns:"
    for ($i = 0; $i -lt $included.Count; $i++) {
        Write-Host ("  {0,4}. {1}" -f ($i + 1), $included[$i])
    }
    $cn = 0
    $cr = Read-Host "Enter column number to filter on"
    if (-not [int]::TryParse($cr.Trim(), [ref]$cn) -or $cn -lt 1 -or $cn -gt $included.Count) {
        Write-Host "Invalid - skipping." -ForegroundColor Yellow; continue
    }
    $ci    = $cn - 1
    $cname = $included[$ci]
    $vr    = Read-Host "Enter value(s) to keep in column '$cname' (comma-separated, partial match OK)"
    if ([string]::IsNullOrWhiteSpace($vr)) {
        Write-Host "No values entered - skipping." -ForegroundColor Yellow; continue
    }
    $fv = @($vr.Split(',') | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ -ne '' })
    $filters.Add(@{ Idx = $ci; Vals = $fv })
    Write-Host "  Filter added: '$cname' contains any of [$($fv -join ', ')]"
    if ((Read-Host "Add another filter? (y/n)").Trim().ToLower() -ne 'y') { break }
}

# ---------------------------------------------------------------------------
# Append RYAN SOURCE FILE column
# ---------------------------------------------------------------------------
$included.Add("RYAN SOURCE FILE")
$nOut   = $included.Count
$srcIdx = $nOut - 1

# ---------------------------------------------------------------------------
# Create output workbook
# ---------------------------------------------------------------------------
$ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath = Join-Path $FolderPath "consolidated_$ts.xlsx"
$outWb   = $xl.Workbooks.Add()

while ($outWb.Worksheets.Count -gt 1) {
    $outWb.Worksheets($outWb.Worksheets.Count).Delete()
}
$outWs      = $outWb.Worksheets(1)
$outWs.Name = "Consolidated"

for ($c = 0; $c -lt $nOut; $c++) { $outWs.Cells(1, $c + 1).Value2 = $included[$c] }
$outWs.Rows(1).Font.Bold = $true

# ---------------------------------------------------------------------------
# Process files
# ---------------------------------------------------------------------------
Write-Host "`nMerging into: $outPath`n"

$allRows      = [System.Collections.Generic.List[object[]]]::new()
$totalOk      = 0
$totalSkipped = 0

foreach ($file in $files) {
    $fname           = $file.Name
    $fileRows        = 0
    $sheetsProcessed = 0
    $wb              = $null
    Write-Host "  [ ] $fname"

    try {
        $wb     = $xl.Workbooks.Open($file.FullName, 0, $true)
        $sheets = $wb.Worksheets

        foreach ($ws in $sheets) {
            $sname = $ws.Name
            $used  = $ws.UsedRange

            if ($null -eq $used -or $used.Rows.Count -lt 2) {
                ReleaseCom $used; ReleaseCom $ws; continue
            }

            $nrows = $used.Rows.Count
            $ncols = $used.Columns.Count

            Write-Host "      '$sname'..." -NoNewline

            $data  = $used.Value2
            $isArr = $data -is [System.Array]

            $shMap = @{}
            for ($c = 1; $c -le $ncols; $c++) {
                $h = if ($isArr) { $data[1, $c] } else { $data }
                if ($null -ne $h) {
                    $key = $h.ToString().Trim().ToLower()
                    if ($colMap.ContainsKey($key)) { $shMap[$c] = $colMap[$key] }
                }
            }

            $sheetRows = 0
            for ($r = 2; $r -le $nrows; $r++) {
                $row   = New-Object object[] $nOut
                $empty = $true

                foreach ($fc in $shMap.Keys) {
                    $v = if ($isArr) { $data[$r, $fc] } else { $null }
                    $v = CoerceVal $v
                    $row[$shMap[$fc]] = $v
                    if ($null -ne $v -and "$v" -ne '') { $empty = $false }
                }
                if ($empty) { continue }

                $pass = $true
                foreach ($f in $filters) {
                    $s   = if ($null -ne $row[$f.Idx]) { $row[$f.Idx].ToString().ToLower() } else { '' }
                    $hit = $false
                    foreach ($v in $f.Vals) { if ($s.Contains($v)) { $hit = $true; break } }
                    if (-not $hit) { $pass = $false; break }
                }
                if (-not $pass) { continue }

                $row[$srcIdx] = $fname
                $allRows.Add($row)
                $sheetRows++
                $fileRows++

                if ($sheetRows % 5000 -eq 0) {
                    Write-Host ("`r      '$sname': {0:N0} rows so far..." -f $sheetRows) -NoNewline
                }
            }

            Write-Host ("`r      '$sname': {0:N0} rows          " -f $sheetRows)
            $sheetsProcessed++
            ReleaseCom $used
            ReleaseCom $ws
        }

        ReleaseCom $sheets
        $sw = if ($sheetsProcessed -eq 1) { "sheet" } else { "sheets" }
        Write-Host ("  [+] {0} - {1} {2}, {3:N0} rows total" -f $fname, $sheetsProcessed, $sw, $fileRows)
        $totalOk++

    } catch {
        Write-Host ("`r  [!] {0} - skipped: {1}" -f $fname, $_) -ForegroundColor Yellow
        $totalSkipped++
    } finally {
        if ($null -ne $wb) { $wb.Close($false); ReleaseCom $wb; $wb = $null }
    }
}

# ---------------------------------------------------------------------------
# Write all rows to output in one bulk operation
# ---------------------------------------------------------------------------
$totalRows = $allRows.Count
Write-Host ("`nWriting {0:N0} rows to output file..." -f $totalRows) -NoNewline

if ($totalRows -gt 0) {
    $arr = [System.Array]::CreateInstance([object], @($totalRows, $nOut), @(1, 1))
    for ($r = 0; $r -lt $totalRows; $r++) {
        for ($c = 0; $c -lt $nOut; $c++) {
            $arr[$r + 1, $c + 1] = $allRows[$r][$c]
        }
    }
    $range = $outWs.Range($outWs.Cells(2, 1), $outWs.Cells($totalRows + 1, $nOut))
    $range.Value2 = $arr
    ReleaseCom $range
}
Write-Host " done."

Write-Host "Saving..." -NoNewline
$outWb.SaveAs($outPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
$outWb.Close($false)
ReleaseCom $outWs
ReleaseCom $outWb
Write-Host " done."

$xl.Quit()
ReleaseCom $xl
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$sep = "=" * 56
Write-Host "`n$sep`n  SUMMARY`n$sep"
Write-Host "  Files processed  : $totalOk"
if ($totalSkipped -gt 0) {
    Write-Host "  Files skipped    : $totalSkipped" -ForegroundColor Yellow
}
Write-Host "  Total rows merged: $("{0:N0}" -f $totalRows)"
Write-Host "  Output file      : $outPath"
Write-Host $sep
