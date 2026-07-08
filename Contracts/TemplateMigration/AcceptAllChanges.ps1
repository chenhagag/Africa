# ============================================================
# AcceptAllChanges.ps1
# Opens all .docx files in Old Contracts subfolders,
# accepts all tracked changes, stops tracking, saves, and closes.
#
# Usage:
#   .\AcceptAllChanges.ps1
#   .\AcceptAllChanges.ps1 -WhatIf      (preview only)
#   .\AcceptAllChanges.ps1 -Folder "ד"  (specific subfolder only)
# ============================================================

param(
    [string]$BasePath = "$PSScriptRoot\Old Contracts",
    [string]$Folder = "",
    [switch]$WhatIf
)

if (-not (Test-Path $BasePath)) {
    Write-Error "Folder not found: $BasePath"
    exit 1
}

# Get all .docx files (exclude temp files starting with ~)
$searchPath = if ($Folder) { Join-Path $BasePath $Folder } else { $BasePath }
$files = Get-ChildItem -Path $searchPath -Filter "*.docx" -Recurse | Where-Object { $_.Name -notlike '~*' }

if ($files.Count -eq 0) {
    Write-Host "No .docx files found in $searchPath" -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($files.Count) files." -ForegroundColor Cyan

if ($WhatIf) {
    foreach ($f in $files) {
        Write-Host "  [WhatIf] Would process: $($f.FullName)" -ForegroundColor Gray
    }
    exit 0
}

# Start Word
Write-Host "Starting Word..." -ForegroundColor Cyan
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone

$succeeded = 0
$failed = 0

foreach ($f in $files) {
    Write-Host "Processing: $($f.Name)" -ForegroundColor Yellow -NoNewline

    try {
        $doc = $word.Documents.Open($f.FullName, $false, $false, $false, "", "", $false, "", "", 0)

        # Accept all revisions
        $doc.AcceptAllRevisions()

        # Stop tracking changes
        $doc.TrackRevisions = $false

        # Save and close
        $doc.Save()
        $doc.Close(0)  # wdDoNotSaveChanges (already saved)

        Write-Host " - OK" -ForegroundColor Green
        $succeeded++
    } catch {
        Write-Host " - FAILED: $_" -ForegroundColor Red
        $failed++
        try { $doc.Close(0) } catch {}
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
}

$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Done! Succeeded: $succeeded, Failed: $failed" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
