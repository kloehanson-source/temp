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
try { $xl.Calculation = -4135 } catch {}   # xlCalculationManual - optional optimisation

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

function ColumnLetter([int]$n) {
    $s = ''
    while ($n -gt 0) { $n--; $s = [char](65 + ($n % 26)) + $s; $n = [int]($n / 26) }
    $s
}

function XmlEsc([string]$s) {
    $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;')
}

$InvCulture = [System.Globalization.CultureInfo]::InvariantCulture

# ---------------------------------------------------------------------------
# Collect sheet names AND column headers in one read-only pass
# ---------------------------------------------------------------------------
Write-Host "`nScanning worksheets and column headers across all files..."
$seenSheets  = [System.Collections.Specialized.OrderedDictionary]::new()  # lower -> display
$seenLower   = [System.Collections.Specialized.OrderedDictionary]::new()  # lower col -> display col
$colToSheets = @{}   # lower col -> hashtable of lower sheet names it appeared in

foreach ($file in $files) {
    $wb = $null
    try {
        $wb = $xl.Workbooks.Open($file.FullName, 0, $true)
        $sheets = $wb.Worksheets
        foreach ($ws in $sheets) {
            $shName = $ws.Name
            $shKey  = $shName.ToLower()
            if (-not $seenSheets.Contains($shKey)) { $seenSheets[$shKey] = $shName }

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
                    if ($name -ne '') {
                        if (-not $seenLower.Contains($key)) {
                            $seenLower[$key] = $name
                            $colToSheets[$key] = @{}
                        }
                        $colToSheets[$key][$shKey] = $true
                    }
                }
            }
            ReleaseCom $used
            ReleaseCom $ws
        }
        ReleaseCom $sheets
    } catch {
        Write-Host "  Warning: could not read '$($file.Name)': $_" -ForegroundColor Yellow
    } finally {
        if ($null -ne $wb) { $wb.Close($false); ReleaseCom $wb; $wb = $null }
    }
}

if ($seenSheets.Count -eq 0) {
    Write-Host "No worksheets found in any file." -ForegroundColor Red
    $xl.Quit(); ReleaseCom $xl; exit 1
}

# ---------------------------------------------------------------------------
# Sheet selection
# ---------------------------------------------------------------------------
$allSheetNames = @($seenSheets.Values)
Write-Host "`nWorksheets found across all files:"
for ($i = 0; $i -lt $allSheetNames.Count; $i++) {
    Write-Host ("  {0,4}. {1}" -f ($i + 1), $allSheetNames[$i])
}
$rawSheets = Read-Host "`nEnter sheet numbers to EXCLUDE (comma-separated), or press Enter to keep all"

$sheetFilter = @{}   # lowercase name -> true for included sheets
if ([string]::IsNullOrWhiteSpace($rawSheets)) {
    foreach ($n in $allSheetNames) { $sheetFilter[$n.ToLower()] = $true }
} else {
    $xnumsS = @{}
    foreach ($tok in $rawSheets.Split(',')) {
        $n = 0
        if ([int]::TryParse($tok.Trim(), [ref]$n)) { $xnumsS[$n] = $true }
    }
    for ($i = 0; $i -lt $allSheetNames.Count; $i++) {
        if (-not $xnumsS.ContainsKey($i + 1)) { $sheetFilter[$allSheetNames[$i].ToLower()] = $true }
    }
    $xdS = for ($i = 0; $i -lt $allSheetNames.Count; $i++) {
        if ($xnumsS.ContainsKey($i + 1)) { $allSheetNames[$i] }
    }
    if ($xdS) { Write-Host "Excluding sheets: $($xdS -join ', ')" }
}
if ($sheetFilter.Count -eq 0) {
    Write-Host "No sheets remaining after exclusion." -ForegroundColor Yellow
    $xl.Quit(); ReleaseCom $xl; exit 0
}
Write-Host "Including $($sheetFilter.Count) sheet(s)."

# Filter the column list to only columns that appeared in at least one selected sheet
$filteredCols = [System.Collections.Specialized.OrderedDictionary]::new()
foreach ($key in $seenLower.Keys) {
    foreach ($sk in $sheetFilter.Keys) {
        if ($colToSheets[$key].ContainsKey($sk)) {
            $filteredCols[$key] = $seenLower[$key]
            break
        }
    }
}

$allCols = @($filteredCols.Values)
if ($allCols.Count -eq 0) {
    Write-Host "No column headers found in the selected sheets." -ForegroundColor Red
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
    $vr = Read-Host "Enter value(s) to keep in column '$cname' (comma-separated; wrap in `"quotes`" for exact match)"
    if ([string]::IsNullOrWhiteSpace($vr)) {
        Write-Host "No values entered - skipping." -ForegroundColor Yellow; continue
    }
    $fv = @($vr.Split(',') | ForEach-Object {
        $tok = $_.Trim()
        if ($tok.Length -ge 2 -and $tok[0] -eq '"' -and $tok[-1] -eq '"') {
            @{ V = $tok.Substring(1, $tok.Length - 2).ToLower(); E = $true }
        } else {
            @{ V = $tok.ToLower(); E = $false }
        }
    } | Where-Object { $_.V -ne '' })
    $filters.Add(@{ Idx = $ci; Vals = $fv })
    $fvDesc = ($fv | ForEach-Object { if ($_.E) { '="' + $_.V + '"' } else { '~' + $_.V } }) -join ', '
    Write-Host "  Filter added: '$cname' matches any of [$fvDesc]  (= exact, ~ partial)"
    if ($filters.Count -ge 2) {
        Write-Host ("  NOTE: all {0} filters must be true on the SAME ROW (AND logic)." -f $filters.Count) `
            -ForegroundColor Cyan
    }
    if ((Read-Host "Add another filter? (y/n)").Trim().ToLower() -ne 'y') { break }
}

# ---------------------------------------------------------------------------
# Append RYAN SOURCE FILE column
# ---------------------------------------------------------------------------
$included.Add("RYAN SOURCE FILE")
$nOut   = $included.Count
$srcIdx = $nOut - 1

# ---------------------------------------------------------------------------
# Create output XLSX via Open XML (bypasses 32-bit COM numeric write issues)
# ---------------------------------------------------------------------------
$ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
$outPath = Join-Path $FolderPath "consolidated_$ts.xlsx"
$noBom   = New-Object System.Text.UTF8Encoding $false

$tmpDir = Join-Path $env:TEMP ("xlmerge_" + [System.Guid]::NewGuid().ToString("N"))
[System.IO.Directory]::CreateDirectory("$tmpDir\_rels")         | Out-Null
[System.IO.Directory]::CreateDirectory("$tmpDir\xl\_rels")      | Out-Null
[System.IO.Directory]::CreateDirectory("$tmpDir\xl\worksheets") | Out-Null

$MAX_DATA_ROWS = 1048575   # rows 2..1048576 per sheet (row 1 = header)

$sheetNames  = [System.Collections.Generic.List[string]]::new()
$curSheetNum = 1
$curDataRows = 0

function Open-SheetWriter([int]$num) {
    $p = Join-Path $tmpDir "xl\worksheets\sheet$num.xml"
    $w = [System.IO.StreamWriter]::new($p, $false, $noBom)
    $w.WriteLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
    $w.Write('<row r="1">')
    for ($hc = 0; $hc -lt $nOut; $hc++) {
        $hRef = (ColumnLetter ($hc + 1)) + '1'
        $w.Write('<c r="' + $hRef + '" s="1" t="inlineStr"><is><t>' + (XmlEsc $included[$hc]) + '</t></is></c>')
    }
    $w.WriteLine('</row>')
    return $w
}

$curSW = Open-SheetWriter 1
$sheetNames.Add("Consolidated")

# ---------------------------------------------------------------------------
# Process files
# ---------------------------------------------------------------------------
Write-Host "`nMerging into: $outPath`n"

$outRow       = 2   # next row to write in output worksheet (row 1 = header)
$totalRows    = 0
$totalOk      = 0
$totalSkipped = 0

# Per-filter hit counters - incremented independently so we can warn if a
# filter value never matched any row regardless of other filters.
$filterHits = New-Object int[] $filters.Count

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

            if (-not $sheetFilter.ContainsKey($sname.ToLower())) {
                ReleaseCom $ws; continue
            }

            $used  = $ws.UsedRange
            if ($null -eq $used -or $used.Rows.Count -lt 2) {
                ReleaseCom $used; ReleaseCom $ws; continue
            }

            $nrows = $used.Rows.Count
            $ncols = $used.Columns.Count

            Write-Host "      '$sname'..." -NoNewline

            $data  = $used.Value2
            $isArr = $data -is [System.Array]

            # Auto-detect header row — handles files with a title row above the real headers
            $headerRow   = 1
            $bestMatches = 0
            $scanTo      = [Math]::Min(5, $nrows)
            for ($tr = 1; $tr -le $scanTo; $tr++) {
                $m = 0
                for ($c = 1; $c -le $ncols; $c++) {
                    $h = if ($isArr) { $data[$tr, $c] } else { $null }
                    if ($null -ne $h) {
                        if ($colMap.ContainsKey($h.ToString().Trim().ToLower())) { $m++ }
                    }
                }
                if ($m -gt $bestMatches) { $bestMatches = $m; $headerRow = $tr }
            }

            $shMap = @{}
            for ($c = 1; $c -le $ncols; $c++) {
                $h = if ($isArr) { $data[$headerRow, $c] } else { $data }
                if ($null -ne $h) {
                    $key = $h.ToString().Trim().ToLower()
                    if ($colMap.ContainsKey($key)) { $shMap[$c] = $colMap[$key] }
                }
            }

            $sheetRows = 0
            for ($r = ($headerRow + 1); $r -le $nrows; $r++) {
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
                for ($fi = 0; $fi -lt $filters.Count; $fi++) {
                    $f   = $filters[$fi]
                    $s   = if ($null -ne $row[$f.Idx]) { $row[$f.Idx].ToString().ToLower() } else { '' }
                    $hit = $false
                    foreach ($v in $f.Vals) {
                        if (($v.E -and $s -eq $v.V) -or (-not $v.E -and $s.Contains($v.V))) {
                            $hit = $true; break
                        }
                    }
                    if ($hit)      { $filterHits[$fi]++ }
                    if (-not $hit) { $pass = $false }
                }
                if (-not $pass) { continue }

                $row[$srcIdx] = $fname

                # Roll over to a new sheet if Excel row limit reached
                if ($curDataRows -ge $MAX_DATA_ROWS) {
                    $curSW.Write('</sheetData></worksheet>')
                    $curSW.Flush(); $curSW.Close(); $curSW.Dispose()
                    $curSheetNum++
                    $curSW = Open-SheetWriter $curSheetNum
                    $sheetNames.Add("Consolidated ($curSheetNum)")
                    $outRow = 2
                    $curDataRows = 0
                }

                # Stream row to Open XML sheet
                $curSW.Write('<row r="' + $outRow + '">')
                for ($c = 0; $c -lt $nOut; $c++) {
                    $v = $row[$c]
                    if ($null -eq $v) { continue }
                    $ref = (ColumnLetter ($c + 1)) + $outRow
                    if ($v -is [string]) {
                        $curSW.Write('<c r="' + $ref + '" t="inlineStr"><is><t>' + (XmlEsc $v) + '</t></is></c>')
                    } elseif ($v -is [bool]) {
                        $curSW.Write('<c r="' + $ref + '" t="b"><v>' + (if ($v) {'1'} else {'0'}) + '</v></c>')
                    } elseif ($v -is [long]) {
                        if ([math]::Abs($v) -ge 1000000000) {
                            # 10+ digit integer: store as text to prevent scientific notation
                            $curSW.Write('<c r="' + $ref + '" t="inlineStr"><is><t>' + $v.ToString() + '</t></is></c>')
                        } else {
                            $curSW.Write('<c r="' + $ref + '" t="n"><v>' + $v.ToString() + '</v></c>')
                        }
                    } else {
                        # Double, Int32, etc. - numeric cell, invariant decimal format
                        $curSW.Write('<c r="' + $ref + '" t="n"><v>' + $v.ToString('G15', $InvCulture) + '</v></c>')
                    }
                }
                $curSW.WriteLine('</row>')
                $curDataRows++
                $outRow++
                $totalRows++
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
        $shLabel = if ($sheetsProcessed -eq 1) { "sheet" } else { "sheets" }
        Write-Host ("  [+] {0} - {1} {2}, {3:N0} rows total" -f $fname, $sheetsProcessed, $shLabel, $fileRows)
        $totalOk++

    } catch {
        Write-Host ("`r  [!] {0} - skipped: {1}" -f $fname, $_) -ForegroundColor Yellow
        $totalSkipped++
    } finally {
        if ($null -ne $wb) { $wb.Close($false); ReleaseCom $wb; $wb = $null }
    }
}

# ---------------------------------------------------------------------------
# Warn about any filter that never matched a single row
# ---------------------------------------------------------------------------
for ($fi = 0; $fi -lt $filters.Count; $fi++) {
    if ($filterHits[$fi] -eq 0) {
        $f = $filters[$fi]
        $fDesc = ($f.Vals | ForEach-Object { if ($_.E) { '="' + $_.V + '"' } else { '~' + $_.V } }) -join ', '
        Write-Host ("WARNING: filter on '{0}' for value(s) [$fDesc] matched zero rows across all files." `
            -f $included[$f.Idx]) -ForegroundColor Yellow
    }
}

if ($totalRows -eq 0 -and $filters.Count -gt 1) {
    $allHadHits = $true
    for ($fi = 0; $fi -lt $filters.Count; $fi++) {
        if ($filterHits[$fi] -eq 0) { $allHadHits = $false; break }
    }
    if ($allHadHits) {
        Write-Host ("WARNING: no rows matched all $($filters.Count) filters on the same row.") `
            -ForegroundColor Yellow
        Write-Host "  Each filter's individual match count:" -ForegroundColor Yellow
        for ($fi = 0; $fi -lt $filters.Count; $fi++) {
            $f = $filters[$fi]
            $fDesc = ($f.Vals | ForEach-Object { if ($_.E) { '="' + $_.V + '"' } else { '~' + $_.V } }) -join ', '
            Write-Host ("    Filter $($fi+1) -- '{0}' matches [$fDesc]: {1:N0} rows" `
                -f $included[$f.Idx], $filterHits[$fi]) -ForegroundColor Yellow
        }
        Write-Host "  If both counts look right, those values may not co-exist on any single row." `
            -ForegroundColor Yellow
    }
}

$xl.Quit()
ReleaseCom $xl
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ("`n{0:N0} rows written." -f $totalRows)
Write-Host "Building output file..." -NoNewline

# Close the last (or only) sheet writer
$curSW.Write('</sheetData></worksheet>')
$curSW.Flush(); $curSW.Close(); $curSW.Dispose()

# OOXML namespace constants
$nsWs  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
$nsSty = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
$nsOD  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
$ctWs  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
$ctWb  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
$ctSty = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
$ctRel = 'application/vnd.openxmlformats-package.relationships+xml'
$nsPkg = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Build per-sheet XML fragments
$ctOverrides = ''
$sheetElems  = ''
$wsRels      = ''
for ($i = 1; $i -le $sheetNames.Count; $i++) {
    $sn = $sheetNames[$i - 1]
    $ctOverrides += '  <Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="' + $ctWs + '"/>' + "`n"
    $sheetElems  += '    <sheet name="' + (XmlEsc $sn) + '" sheetId="' + $i + '" r:id="rId' + $i + '"/>' + "`n"
    $wsRels      += '  <Relationship Id="rId' + $i + '" Type="' + $nsWs + '" Target="worksheets/sheet' + $i + '.xml"/>' + "`n"
}
$styleRId = 'rId' + ($sheetNames.Count + 1)
$wsRels  += '  <Relationship Id="' + $styleRId + '" Type="' + $nsSty + '" Target="styles.xml"/>' + "`n"

$xmlFiles = @{
    '_rels\.rels' = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
        '<Relationships xmlns="' + $nsPkg + '">' + "`n" +
        '  <Relationship Id="rId1" Type="' + $nsOD + '" Target="xl/workbook.xml"/>' + "`n" +
        '</Relationships>')

    'xl\workbook.xml' = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' + "`n" +
        '          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + "`n" +
        '  <sheets>' + "`n" + $sheetElems + '  </sheets>' + "`n" +
        '</workbook>')

    'xl\_rels\workbook.xml.rels' = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
        '<Relationships xmlns="' + $nsPkg + '">' + "`n" +
        $wsRels +
        '</Relationships>')

    'xl\styles.xml' = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + "`n" +
        '  <fonts count="2">' + "`n" +
        '    <font><sz val="11"/><name val="Calibri"/></font>' + "`n" +
        '    <font><b/><sz val="11"/><name val="Calibri"/></font>' + "`n" +
        '  </fonts>' + "`n" +
        '  <fills count="2">' + "`n" +
        '    <fill><patternFill patternType="none"/></fill>' + "`n" +
        '    <fill><patternFill patternType="gray125"/></fill>' + "`n" +
        '  </fills>' + "`n" +
        '  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' + "`n" +
        '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' + "`n" +
        '  <cellXfs count="2">' + "`n" +
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + "`n" +
        '    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' + "`n" +
        '  </cellXfs>' + "`n" +
        '</styleSheet>')
}

foreach ($rel in $xmlFiles.Keys) {
    [System.IO.File]::WriteAllText((Join-Path $tmpDir $rel), $xmlFiles[$rel], $noBom)
}

$ctXml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + "`n" +
    '  <Default Extension="rels" ContentType="' + $ctRel + '"/>' + "`n" +
    '  <Default Extension="xml" ContentType="application/xml"/>' + "`n" +
    '  <Override PartName="/xl/workbook.xml" ContentType="' + $ctWb + '"/>' + "`n" +
    $ctOverrides +
    '  <Override PartName="/xl/styles.xml" ContentType="' + $ctSty + '"/>' + "`n" +
    '</Types>')
# Use IO.Path.Combine so brackets in filename are treated as literals
[System.IO.File]::WriteAllText(
    [System.IO.Path]::Combine($tmpDir, '[Content_Types].xml'), $ctXml, $noBom)

Add-Type -AssemblyName System.IO.Compression.FileSystem
if (Test-Path $outPath) { Remove-Item $outPath -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmpDir, $outPath)
[System.IO.Directory]::Delete($tmpDir, $true)

Write-Host " done."

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
