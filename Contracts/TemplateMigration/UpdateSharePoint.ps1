# ============================================================
# UpdateSharePoint.ps1
# Reads ExportedData.csv and updates SharePoint Online library columns
#
# Prerequisites:
#   Install-Module PnP.PowerShell -Scope CurrentUser
#
# Usage:
#   .\UpdateSharePoint.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/yoursite" -LibraryName "Contracts"
#
# The script will:
# 1. Connect to SharePoint Online (browser login)
# 2. Read the CSV exported by ExportContractData.bas
# 3. Match each row to a document in the library by filename
# 4. Update the SP columns with values from the CSV
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$true)]
    [string]$LibraryName,

    [string]$CsvPath = "$PSScriptRoot\ExportedData.csv",

    [switch]$WhatIf
)

# --- Connect to SharePoint ---
Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -UseWebLogin
Write-Host "Connected." -ForegroundColor Green

# --- Read CSV ---
if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    Write-Host "Run ExportAllContracts macro in Word first to generate the CSV."
    exit 1
}

$rows = Import-Csv -Path $CsvPath -Encoding UTF8
Write-Host "Loaded $($rows.Count) rows from CSV." -ForegroundColor Cyan

# --- SP column names that we update (from CSV headers, excluding FileName) ---
# These match the column internal names used by saveToHelper in the Add-in
$spColumns = @(
    "ContractNumber",
    "contractVersion",
    "ContractTemplate",
    "project",
    "SiteName",
    "Municipality",
    "WorkDescription",
    "signDate",
    "StartDate",
    "DurationMonths",
    "ExpectedEndDate",
    "status",
    "recipient",
    "otherSides",
    "partyAName",
    "PartyAContactNamePercent",
    "supplierName",
    "CostCompMethod",
    "CostContractScope",
    "CostCurrency",
    "CostIndexType",
    "CostBaseIndexDate",
    "CostIndexMode",
    "CostIndexPoints",
    "CostPaymentTerms",
    "customField1",
    "customField2",
    "customField3",
    "customField4",
    "customField5",
    "customField6",
    "customField7",
    "customField8"
)

# --- Discover which columns actually exist in the library ---
Write-Host "Checking library columns..." -ForegroundColor Cyan
$listFields = Get-PnPField -List $LibraryName | Select-Object -ExpandProperty InternalName
$validColumns = $spColumns | Where-Object { $_ -in $listFields }
$invalidColumns = $spColumns | Where-Object { $_ -notin $listFields }

if ($invalidColumns.Count -gt 0) {
    Write-Host "Columns NOT found in library (will be skipped):" -ForegroundColor Yellow
    foreach ($col in $invalidColumns) {
        Write-Host "  - $col" -ForegroundColor Yellow
    }
}
Write-Host "Valid columns: $($validColumns.Count)/$($spColumns.Count)" -ForegroundColor Green

# --- Process each row ---
$succeeded = 0
$failed = 0
$skipped = 0

foreach ($row in $rows) {
    $fileName = $row.FileName
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        Write-Warning "Empty filename in CSV row, skipping."
        $skipped++
        continue
    }

    Write-Host "`nProcessing: $fileName" -ForegroundColor Yellow

    # Find the file in the library — get all items and filter locally for Hebrew filename support
    try {
        $allItems = Get-PnPListItem -List $LibraryName -Fields "FileLeafRef","ID" -PageSize 500
    } catch {
        Write-Warning "  Error querying library: $_"
        $failed++
        continue
    }

    $item = $null
    foreach ($li in $allItems) {
        if ($li.FieldValues["FileLeafRef"] -eq $fileName) {
            $item = $li
            break
        }
    }

    if ($null -eq $item) {
        Write-Warning "  File not found in library: $fileName"
        $skipped++
        continue
    }

    $itemId = $item.FieldValues["ID"]
    Write-Host "  Found item ID: $itemId" -ForegroundColor Green

    # Build field values hashtable (only non-empty, non-placeholder values)
    $fieldValues = @{}
    foreach ($col in $validColumns) {
        $val = $row.$col
        if ([string]::IsNullOrWhiteSpace($val)) { continue }

        # Skip placeholder text
        if ($val -match "(?i)click or tap here" -or
            $val -match "(?i)enter text" -or
            $val -match "^\[.*\]$") {
            Write-Host "    SKIP $col = placeholder" -ForegroundColor DarkYellow
            continue
        }

        # Fix status: "3;#Draft" -> "Draft"
        if ($col -eq "status" -and $val -match "^\d+;#(.+)$") {
            $val = $Matches[1]
        }

        # Fix ContractNumber: "CONT-5-668CONT-10-2195" -> take the last CONT-xx-xxxx
        if ($col -eq "ContractNumber" -and $val -match "(CONT-\d+-\d+)$") {
            $val = $Matches[1]
        }

        $fieldValues[$col] = $val
    }

    if ($fieldValues.Count -eq 0) {
        Write-Host "  No data to update for $fileName" -ForegroundColor Gray
        $skipped++
        continue
    }

    # Update the list item
    if ($WhatIf) {
        Write-Host "  [WhatIf] Would update item #$itemId with $($fieldValues.Count) fields:" -ForegroundColor Magenta
        foreach ($key in $fieldValues.Keys) {
            $displayVal = if ($fieldValues[$key].Length -gt 50) { $fieldValues[$key].Substring(0, 50) + "..." } else { $fieldValues[$key] }
            Write-Host "    $key = $displayVal" -ForegroundColor Gray
        }
    } else {
        try {
            Set-PnPListItem -List $LibraryName -Identity $itemId -Values $fieldValues
            Write-Host "  Updated item #$itemId with $($fieldValues.Count) fields." -ForegroundColor Green
            $succeeded++
        } catch {
            Write-Warning "  Error updating item #$itemId : $_"
            $failed++
        }
    }
}

# --- Summary ---
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Done!" -ForegroundColor Green
Write-Host "  Succeeded: $succeeded"
Write-Host "  Failed:    $failed"
Write-Host "  Skipped:   $skipped"
Write-Host "========================================" -ForegroundColor Cyan

Disconnect-PnPOnline
