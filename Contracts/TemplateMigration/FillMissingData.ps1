# ============================================================
# FillMissingData.ps1
# Reads oldSiteData_en.csv (old SP data with English headers) and fills
# missing column values in the new SP library.
#
# Matching: by contract number (CONT-xx-xxxx).
# Only fills empty fields — does NOT overwrite existing data.
#
# Usage:
#   .\FillMissingData.ps1 -SiteUrl "https://africaisrael.sharepoint.com/ContractsNEW" -LibraryName "ContractsDocs" -WhatIf
#   .\FillMissingData.ps1 -SiteUrl "https://africaisrael.sharepoint.com/ContractsNEW" -LibraryName "ContractsDocs"
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$true)]
    [string]$LibraryName,

    [string]$CsvPath = "$PSScriptRoot\oldSiteData_en.csv",

    [switch]$WhatIf
)

# --- CSV column (English) -> SP column internal name ---
# Only columns we want to fill. CSV header = SP column in most cases.
$columnsToFill = @(
    "docCreator",
    "CostCompMethod",
    "CostContractScope",
    "CostCurrency",
    "CostIndexType",
    "CostBaseIndexDate",
    "CostIndexMode",
    "CostIndexPoints",
    "CostPaymentTerms",
    "ContractTemplate",
    "project",
    "SiteName",
    "Municipality",
    "WorkDescription",
    "signDate",
    "StartDate",
    "ExpectedEndDate",
    "DurationMonths",
    "status",
    "supplierName",
    "recipient",
    "otherSides",
    "partyAName",
    "PartyAContactNamePercent",
    "contractVersion",
    "customField1",
    "customField2",
    "customField3",
    "customField4",
    "customField5",
    "customField6",
    "customField7",
    "customField8"
)

# --- Connect to SharePoint ---
Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -UseWebLogin
Write-Host "Connected." -ForegroundColor Green

# --- Read old data CSV ---
if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

$oldRows = Import-Csv -Path $CsvPath -Encoding UTF8
Write-Host "Loaded $($oldRows.Count) rows from old data." -ForegroundColor Cyan

# --- Discover which columns actually exist in the new library ---
Write-Host "Checking library columns..." -ForegroundColor Cyan
$listFields = Get-PnPField -List $LibraryName | Select-Object -ExpandProperty InternalName
$validColumns = $columnsToFill | Where-Object { $_ -in $listFields }
$invalidColumns = $columnsToFill | Where-Object { $_ -notin $listFields }

if ($invalidColumns.Count -gt 0) {
    Write-Host "Columns NOT found in library (will be skipped):" -ForegroundColor Yellow
    foreach ($col in $invalidColumns) {
        Write-Host "  - $col" -ForegroundColor Yellow
    }
}
Write-Host "Valid columns: $($validColumns.Count)/$($columnsToFill.Count)" -ForegroundColor Green

# --- Load all items from new library (once) ---
Write-Host "Loading all items from new library..." -ForegroundColor Cyan
$newItemFields = @("FileLeafRef", "ID", "ContractNumber") + $validColumns
$allNewItems = Get-PnPListItem -List $LibraryName -Fields $newItemFields -PageSize 500
Write-Host "Found $($allNewItems.Count) items in new library." -ForegroundColor Green

# --- Build lookup: CONT number -> new library item(s) ---
# Extract ALL CONT-xx-xxxx patterns from filename + ContractNumber column
$contLookup = @{}
foreach ($item in $allNewItems) {
    $fileName = $item.FieldValues["FileLeafRef"]
    $contCol = $item.FieldValues["ContractNumber"]

    $allConts = @()

    # Get all CONT numbers from filename (handles CONT-5-565CONT-10-1927)
    $matchResult = [regex]::Matches($fileName, "CONT-\d+-\d+")
    foreach ($m in $matchResult) {
        $allConts += $m.Value
    }

    # Also from ContractNumber column
    if (-not [string]::IsNullOrWhiteSpace($contCol)) {
        $contColMatch = [regex]::Matches("$contCol", "CONT-\d+-\d+")
        foreach ($m in $contColMatch) {
            $allConts += $m.Value
        }
    }

    foreach ($contNum in ($allConts | Sort-Object -Unique)) {
        if (-not $contLookup.ContainsKey($contNum)) {
            $contLookup[$contNum] = @()
        }
        $contLookup[$contNum] += $item
    }
}
Write-Host "Indexed $($contLookup.Count) unique CONT numbers in new library." -ForegroundColor Cyan

# --- Build site+name lookup: "siteName|baseName" -> new library item(s) ---
# For fallback matching when CONT number doesn't match
$siteNameLookup = @{}
foreach ($item in $allNewItems) {
    $fileName = $item.FieldValues["FileLeafRef"]
    $siteName = "$($item.FieldValues['SiteName'])".Trim().ToLower()
    # Extract base name: remove extension and " - CONT-xx-xxxx" suffix
    $baseName = $fileName -replace "\.docx$", "" -replace "\s*-\s*CONT-\d+-\d+.*$", ""
    $baseName = $baseName.Trim().ToLower()
    if ($baseName -ne "" -and $siteName -ne "") {
        $key = "$siteName|$baseName"
        if (-not $siteNameLookup.ContainsKey($key)) {
            $siteNameLookup[$key] = @()
        }
        $siteNameLookup[$key] += $item
    }
}
Write-Host "Indexed $($siteNameLookup.Count) unique site+name combinations in new library." -ForegroundColor Cyan

# --- Build site lookup: siteName -> all items in that site ---
$siteLookup = @{}
foreach ($item in $allNewItems) {
    $siteName = "$($item.FieldValues['SiteName'])".Trim().ToLower()
    if ($siteName -ne "") {
        if (-not $siteLookup.ContainsKey($siteName)) {
            $siteLookup[$siteName] = @()
        }
        $siteLookup[$siteName] += $item
    }
}
Write-Host "Indexed $($siteLookup.Count) unique site names in new library." -ForegroundColor Cyan

# --- Process each old row ---
$succeeded = 0
$failed = 0
$skipped = 0
$noMatch = 0
$noMatchList = @()
$matchedByCont = 0
$matchedByName = 0
$matchedBySiteCont = 0

foreach ($oldRow in $oldRows) {
    $contNum = $oldRow.OldContractNumber.Trim()

    # Fix duplicated CONT numbers — take the last one
    if ($contNum -match "(CONT-\d+-\d+)$") {
        $contNum = $Matches[1]
    }

    # Strategy 1: match by CONT number
    $matchingItems = $null
    if ($contLookup.ContainsKey($contNum)) {
        $matchingItems = $contLookup[$contNum]
        $matchedByCont++
    }

    # Strategy 2: match by folder name (from path) + filename
    if ($null -eq $matchingItems -or $matchingItems.Count -eq 0) {
        $oldFileName = $oldRow.OldFileName -replace "\.docx$", ""
        $oldFileNameLower = $oldFileName.Trim().ToLower()
        $oldFolderName = "$($oldRow.OldFolderName)".Trim().ToLower()
        if ($oldFileNameLower -ne "" -and $oldFolderName -ne "") {
            $key = "$oldFolderName|$oldFileNameLower"
            if ($siteNameLookup.ContainsKey($key)) {
                $matchingItems = $siteNameLookup[$key]
                $matchedByName++
            }
        }
    }

    # Strategy 3: match by site (folder name, with fuzzy) + CONT in filename or ContractNumber
    if ($null -eq $matchingItems -or $matchingItems.Count -eq 0) {
        $oldFolderName = "$($oldRow.OldFolderName)".Trim().ToLower()
        if ($oldFolderName -ne "") {
            # Find matching site: exact match first, then partial (old contained in new or new contained in old)
            $siteItems = $null
            if ($siteLookup.ContainsKey($oldFolderName)) {
                $siteItems = $siteLookup[$oldFolderName]
            } else {
                # Normalize: remove dashes, extra spaces for comparison
                $oldNorm = ($oldFolderName -replace "[-\s]+", " ").Trim()
                foreach ($newSite in $siteLookup.Keys) {
                    $newNorm = ($newSite -replace "[-\s]+", " ").Trim()
                    if ($newNorm.Contains($oldNorm) -or $oldNorm.Contains($newNorm)) {
                        $siteItems = $siteLookup[$newSite]
                        break
                    }
                }
            }
            if ($null -ne $siteItems) {
                foreach ($si in $siteItems) {
                    $siFN = "$($si.FieldValues['FileLeafRef'])"
                    $siCN = "$($si.FieldValues['ContractNumber'])"
                    if ($siFN -match [regex]::Escape($contNum) -or $siCN -match [regex]::Escape($contNum)) {
                        $matchingItems = @($si)
                        $matchedBySiteCont++
                        break
                    }
                }
            }
        }
    }

    if ($null -eq $matchingItems -or $matchingItems.Count -eq 0) {
        $noMatch++
        $oldFN = $oldRow.OldFileName
        $oldFolder = $oldRow.OldFolderName
        $noMatchList += [PSCustomObject]@{
            ContractNumber = $contNum
            FileName = $oldFN
            FolderName = $oldFolder
        }
        continue
    }

    foreach ($newItem in $matchingItems) {
        $itemId = $newItem.FieldValues["ID"]
        $fileName = $newItem.FieldValues["FileLeafRef"]

        # Build update hashtable: only fill fields that are EMPTY in new library
        $updates = @{}
        foreach ($col in $validColumns) {
            $oldVal = $oldRow.$col

            # Skip empty old values
            if ([string]::IsNullOrWhiteSpace($oldVal)) { continue }

            # Check if new library field is empty
            $newVal = $newItem.FieldValues[$col]
            if (-not [string]::IsNullOrWhiteSpace($newVal)) {
                continue
            }

            $cleanVal = $oldVal.Trim()

            # Skip placeholder text
            if ($cleanVal -match "(?i)click or tap here" -or $cleanVal -match "^\[.*\]$") {
                continue
            }

            $updates[$col] = $cleanVal
        }

        if ($updates.Count -eq 0) {
            $skipped++
            continue
        }

        if ($WhatIf) {
            Write-Host "`n[WhatIf] $contNum -> $fileName (item #$itemId) - $($updates.Count) fields:" -ForegroundColor Magenta
            foreach ($key in $updates.Keys) {
                $displayVal = if ($updates[$key].Length -gt 60) { $updates[$key].Substring(0, 60) + "..." } else { $updates[$key] }
                Write-Host "    $key = $displayVal" -ForegroundColor Gray
            }
        } else {
            try {
                Set-PnPListItem -List $LibraryName -Identity $itemId -Values $updates
                Write-Host "  Updated $contNum -> $fileName (#$itemId) - $($updates.Count) fields" -ForegroundColor Green
                $succeeded++
            } catch {
                Write-Warning "  Error updating $contNum (#$itemId): $_"
                $failed++
            }
        }
    }
}

# --- Summary ---
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Done!" -ForegroundColor Green
Write-Host "  Matched by CONT:     $matchedByCont"
Write-Host "  Matched by site+CONT: $matchedBySiteCont"
Write-Host "  Matched by name:     $matchedByName"
Write-Host "  Succeeded:           $succeeded"
Write-Host "  Failed:              $failed"
Write-Host "  Skipped:             $skipped (no missing data)"
Write-Host "  No match:            $noMatch"
if ($noMatchList.Count -gt 0) {
    $unmatchedPath = "$PSScriptRoot\unmatched_contracts.csv"
    $noMatchList | Export-Csv -Path $unmatchedPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n  Unmatched list saved to: $unmatchedPath" -ForegroundColor Yellow
}
Write-Host "========================================" -ForegroundColor Cyan

Disconnect-PnPOnline
