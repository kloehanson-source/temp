# ============================================================
#  Flatten & Merge PDFs - PowerShell + iTextSharp
#  Drop run.bat and this .ps1 into a folder with your PDFs,
#  double-click run.bat, and get merged.pdf out.
#  No installs. Downloads required DLLs on first run (~3 MB).
# ============================================================

$ErrorActionPreference = "Stop"
trap {
    Write-Host ""
    Write-Host "UNEXPECTED ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "   PDF Flatten & Merge" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$outputName = "merged.pdf"

# --- Step 1: Gather PDFs ---
$pdfFiles = Get-ChildItem -Path $scriptDir -Filter "*.pdf" |
    Where-Object { $_.Name -ne $outputName } |
    Sort-Object Name

if ($pdfFiles.Count -eq 0) {
    Write-Host "No PDF files found in this folder." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

Write-Host "Found $($pdfFiles.Count) PDF(s) to process (alphabetical order):"
$i = 1
foreach ($f in $pdfFiles) {
    Write-Host "  $i. $($f.Name)" -ForegroundColor White
    $i++
}
Write-Host ""

# --- Step 2: Download iTextSharp + BouncyCastle if needed ---
# iTextSharp 5.5.13.3 requires BouncyCastle.Crypto.dll as a separate
# dependency. Both must be present or iTextSharp fails to load.

$libDir    = Join-Path $scriptDir ".pdftools"
$dllPath   = Join-Path $libDir "itextsharp.dll"
$bcDllPath = Join-Path $libDir "BouncyCastle.Crypto.dll"

$needsDownload = (-not (Test-Path $dllPath)) -or (-not (Test-Path $bcDllPath))

if ($needsDownload) {
    Write-Host "First run: downloading PDF libraries (~3 MB, one-time)..." -ForegroundColor Gray
    New-Item -ItemType Directory -Path $libDir -Force | Out-Null
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # -- iTextSharp --
    Write-Host "  Fetching iTextSharp..." -NoNewline -ForegroundColor Gray
    try {
        $zipPath    = Join-Path $env:TEMP "itextsharp_pkg.zip"
        $extractDir = Join-Path $env:TEMP "itextsharp_extract"

        Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3" `
            -OutFile $zipPath -UseBasicParsing
        if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
        Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

        $found = Get-ChildItem -Path $extractDir -Filter "itextsharp.dll" -Recurse |
            Where-Object { $_.FullName -notmatch "netstandard|netcore|net5|net6|net7|net8|net9" } |
            Where-Object { $_.FullName -match "net4" } |
            Select-Object -First 1

        if (-not $found) {
            $found = Get-ChildItem -Path $extractDir -Filter "itextsharp.dll" -Recurse |
                Where-Object { $_.FullName -notmatch "netstandard|netcore|net5|net6|net7|net8|net9" } |
                Select-Object -First 1
        }

        if (-not $found) { throw "itextsharp.dll not found in package" }
        Copy-Item $found.FullName $dllPath -Force
        Remove-Item $zipPath    -Force -ErrorAction SilentlyContinue
        Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host " OK" -ForegroundColor Green
    }
    catch {
        Write-Host " FAILED: $_" -ForegroundColor Red
        Remove-Item $libDir -Recurse -Force -ErrorAction SilentlyContinue
        Read-Host "Press Enter to exit"
        exit 1
    }

    # -- BouncyCastle (required by iTextSharp) --
    Write-Host "  Fetching BouncyCastle..." -NoNewline -ForegroundColor Gray
    try {
        $zipPath    = Join-Path $env:TEMP "bc_pkg.zip"
        $extractDir = Join-Path $env:TEMP "bc_extract"

        Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/BouncyCastle.Crypto/1.9.0" `
            -OutFile $zipPath -UseBasicParsing
        if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
        Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

        $found = Get-ChildItem -Path $extractDir -Filter "BouncyCastle.Crypto.dll" -Recurse |
            Where-Object { $_.FullName -notmatch "netstandard|netcore|net5|net6|net7|net8|net9" } |
            Where-Object { $_.FullName -match "net4" } |
            Select-Object -First 1

        if (-not $found) {
            $found = Get-ChildItem -Path $extractDir -Filter "BouncyCastle.Crypto.dll" -Recurse |
                Where-Object { $_.FullName -notmatch "netstandard|netcore|net5|net6|net7|net8|net9" } |
                Select-Object -First 1
        }

        if (-not $found) { throw "BouncyCastle.Crypto.dll not found in package" }
        Copy-Item $found.FullName $bcDllPath -Force
        Remove-Item $zipPath    -Force -ErrorAction SilentlyContinue
        Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host " OK" -ForegroundColor Green
    }
    catch {
        Write-Host " FAILED: $_" -ForegroundColor Red
        Remove-Item $libDir -Recurse -Force -ErrorAction SilentlyContinue
        Read-Host "Press Enter to exit"
        exit 1
    }

    Write-Host "Libraries ready." -ForegroundColor Green
    Write-Host ""
}

# BouncyCastle must be loaded before iTextSharp
try {
    Add-Type -Path $bcDllPath
    Add-Type -Path $dllPath
}
catch {
    Write-Host "Failed to load PDF libraries: $_" -ForegroundColor Red
    Write-Host "Deleting cached files - re-run to download again." -ForegroundColor Yellow
    Remove-Item $libDir -Recurse -Force -ErrorAction SilentlyContinue
    Read-Host "Press Enter to exit"
    exit 1
}

# --- Step 3: Flatten each PDF into a temp file ---
# PdfStamper.FormFlattening renders field values as static page content
# so no interactive fields survive to conflict on merge.

$tempDir = Join-Path $scriptDir ".pdftemp"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
$tempFiles = [System.Collections.Generic.List[string]]::new()

foreach ($pdfFile in $pdfFiles) {
    Write-Host "Flattening: $($pdfFile.Name) ..." -NoNewline

    $tempPath = Join-Path $tempDir $pdfFile.Name

    try {
        $reader    = New-Object iTextSharp.text.pdf.PdfReader($pdfFile.FullName)
        $outStream = [System.IO.File]::Create($tempPath)
        $stamper   = New-Object iTextSharp.text.pdf.PdfStamper($reader, $outStream)

        $stamper.FormFlattening     = $true
        $stamper.FreeTextFlattening = $true

        $pageCount = $reader.NumberOfPages
        $stamper.Close()
        $reader.Close()

        $tempFiles.Add($tempPath)
        Write-Host " OK ($pageCount pages)" -ForegroundColor Green
    }
    catch {
        Write-Host " ERROR: $_" -ForegroundColor Red
    }
}

if ($tempFiles.Count -eq 0) {
    Write-Host "No files were successfully flattened. Exiting." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# --- Step 4: Merge the flattened temp files ---
Write-Host ""
Write-Host "Merging $($tempFiles.Count) flattened PDF(s) ..." -NoNewline

$outputPath = Join-Path $scriptDir $outputName
$document   = New-Object iTextSharp.text.Document
$outStream  = [System.IO.File]::Create($outputPath)
$copy       = New-Object iTextSharp.text.pdf.PdfCopy($document, $outStream)
$document.Open()

$totalPages = 0
foreach ($tempFile in $tempFiles) {
    try {
        $reader = New-Object iTextSharp.text.pdf.PdfReader($tempFile)
        $pages  = $reader.NumberOfPages
        for ($p = 1; $p -le $pages; $p++) {
            $copy.AddPage($copy.GetImportedPage($reader, $p))
        }
        $totalPages += $pages
        $reader.Close()
    }
    catch {
        Write-Host ""
        Write-Host "  Error merging $([System.IO.Path]::GetFileName($tempFile)): $_" -ForegroundColor Red
    }
}

$document.Close()

# --- Step 5: Cleanup ---
Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host " Done!" -ForegroundColor Green
Write-Host ""
Write-Host "merged.pdf created - $totalPages total pages, no form fields." -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
