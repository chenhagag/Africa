# Import apartments from CSV to SharePoint list "andreius3"
# IMPORTANT: This script targets ONLY the "andreius3" list. Do NOT change the list name.
# Site: https://africaisrael.sharepoint.com/mesirot/
#
# Prerequisites: PnP.PowerShell module
#   Install-Module PnP.PowerShell -Scope CurrentUser
#
# Usage:
#   Connect-PnPOnline -Url "https://africaisrael.sharepoint.com/mesirot/" -Interactive
#   .\import_to_sp.ps1

$SiteUrl = "https://africaisrael.sharepoint.com/mesirot/"
$ListName = "andreius3"
$CsvPath = Join-Path $PSScriptRoot "andreius3_import.csv"

# Verify connection
try {
    $ctx = Get-PnPContext
    if ($null -eq $ctx) { throw "Not connected" }
    Write-Host "Connected to: $($ctx.Url)" -ForegroundColor Green
}
catch {
    Write-Host "Not connected to SharePoint. Run first:" -ForegroundColor Red
    Write-Host "  Connect-PnPOnline -Url '$SiteUrl' -Interactive" -ForegroundColor Yellow
    exit 1
}

# Verify we're on the right site
$currentUrl = (Get-PnPWeb).Url
if ($currentUrl -notlike "*mesirot*") {
    Write-Host "ERROR: Connected to '$currentUrl' - expected mesirot site!" -ForegroundColor Red
    exit 1
}

# Verify list exists
try {
    $list = Get-PnPList -Identity $ListName -ErrorAction Stop
    Write-Host "Target list: $($list.Title) (Items: $($list.ItemCount))" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: List '$ListName' not found!" -ForegroundColor Red
    exit 1
}

# Confirm before proceeding
if ($list.ItemCount -gt 0) {
    Write-Host "WARNING: List already has $($list.ItemCount) items!" -ForegroundColor Yellow
    $confirm = Read-Host "Continue? (y/n)"
    if ($confirm -ne 'y') { exit 0 }
}

# Import CSV
$data = Import-Csv -Path $CsvPath -Encoding UTF8
Write-Host "Loaded $($data.Count) rows from CSV" -ForegroundColor Cyan

$count = 0
$errors = 0

foreach ($row in $data) {
    try {
        $values = @{
            "Title"           = $row.Title
            "Tower"           = $row.Tower
            "Floor"           = $row.Floor
            "msrAppNum"       = $row.msrAppNum
            "ShadowNumber"    = $row.ShadowNumber
            "ApartmentType"   = $row.ApartmentType
            "Rooms"           = $row.Rooms
            "Standard"        = $row.Standard
            "msrAppStatus"    = $row.msrAppStatus
            "HandoverStatus"  = $row.HandoverStatus
            "FinishStatus"    = $row.FinishStatus
        }

        Add-PnPListItem -List $ListName -Values $values | Out-Null
        $count++

        if ($count % 50 -eq 0) {
            Write-Host "  Imported $count / $($data.Count)..." -ForegroundColor Gray
        }
    }
    catch {
        $errors++
        Write-Host "Error on row $($row.msrAppNum): $_" -ForegroundColor Red
    }
}

Write-Host "`nDone! Imported: $count | Errors: $errors" -ForegroundColor Green
