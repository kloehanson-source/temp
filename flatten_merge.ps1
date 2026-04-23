# ============================================================
#  Flatten & Merge PDFs - PowerShell + iTextSharp
#  Drop run.bat and this .ps1 into a folder with your PDFs,
#  double-click run.bat, and get merged.pdf out.
#  No installs. Downloads iTextSharp DLL on first run (~2 MB).
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

# --- Step 2: Get iTextSharp if needed ---
$libDir  = Join-Path $scriptDir ".pdftools"
$dllPath = Join-Path $libDir "itextsharp.dll"

if (-not (Test-Path $dllPath)) {
    Write-Host "First run: downloading iTextSharp (~2 MB, one-time)..." -ForegroundColor Gray

    $nugetUrl   = "https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3"
    $zipPath    = Join-Path $env:TEMP "itextsharp_pkg.zip"
    $extractDir = Join-Path $env:TEMP "itextsharp_extract"

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $nugetUrl -OutFile $zipPath -UseBasicParsing

        if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
        Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

        # Windows PowerShell needs the net40 build - netstandard/netcore will not load
        $sourceDll = Get-ChildItem -Path $extractDir -Filter "itextsharp.dll" -Recurse |
            Where-Object { $_.FullName -notmatch "netstandard|netcore|net5|net6|net7|net8|net9" } |
            Where-Object { $_.FullName -match "net4" } |
            Select-Object -First 1

        if (-not $sourceDll) {
            # Fallback: anything that is not netstandard/netcore
            $sourceDll = Get-ChildItem -Path $extractDir -Filter "itextsharp.dll" -Recurse |
                Where-Object { $_.FullName -notmatch "netstandard|netcore|net5|net6|net7|net8|net9" } |
                Select-Object -First 1
        }

        if (-not $sourceDll) { throw "itextsharp.dll not found in NuGet package" }

        Write-Host "  Selected: $($sourceDll.FullName)" -ForegroundColor Gray

        New-Item -ItemType Directory -Path $libDir -Force | Out-Null
        Copy-Item $sourceDll.FullName $dllPath -Force

        Remove-Item $zipPath    -Force -ErrorAction SilentlyContinue
        Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue

        Write-Host "iTextSharp ready." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR downloading iTextSharp: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "Manual fix: download itextsharp.dll and place it in:" -ForegroundColor Yellow
        Write-Host "  $libDir" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit
    }
}

# If the cached DLL is the wrong build it will fail here - delete and re-run to fix
try {
    Add-Type -Path $dllPath
} catch {
    Write-Host "Failed to load iTextSharp DLL (wrong build cached). Deleting and re-run to fix." -ForegroundColor Red
    Remove-Item $libDir -Recurse -Force -ErrorAction SilentlyContinue
    Read-Host "Press Enter to exit"
    exit 1
}

# --- Step 3: Flatten each PDF into a temp file ---
# iTextSharp's PdfStamper.FormFlattening renders field values as static
# page content, so no interactive fields survive to conflict on merge.

$tempDir = Join-Path $scriptDir ".pdftemp"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
$tempFiles = [System.Collections.Generic.List[string]]::new()

foreach ($pdfFile in $pdfFiles) {
    Write-Host "Flattening: $($pdfFile.Name) ..." -NoNewline

    $tempPath = Join-Path $tempDir $pdfFile.Name

    try {
        $reader  = New-Object iTextSharp.text.pdf.PdfReader($pdfFile.FullName)
        $outStream = [System.IO.File]::Create($tempPath)
        $stamper = New-Object iTextSharp.text.pdf.PdfStamper($reader, $outStream)

        $stamper.FormFlattening     = $true   # bakes AcroForm field values into page
        $stamper.FreeTextFlattening = $true   # bakes free-text annotations too

        $pageCount = $reader.NumberOfPages    # read before closing
        $stamper.Close()   # also closes $outStream
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
        $pages  = $reader.NumberOfPages       # read before closing
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
