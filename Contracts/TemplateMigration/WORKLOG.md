# Template Migration — Work Log

## Goal
Convert old Word templates with SharePoint-bound content controls to clean controls that the Word Add-in can manage.

## Key Discovery
The Word JS API (`document.body.contentControls`) **cannot see** old SharePoint-bound controls at all (returns 0).
VBA COM (`doc.ContentControls`) sees all of them (48 in test file).
This means the Add-in cannot migrate them — only VBA can.

## Approaches Tried

### 1. Python scripts (May 2026, FAILED)
- `migrate_templates.py` — regex cleanup of XML. Output in `ContractTemplates/Fixed Templates/`. Didn't work.
- `fix_templates.py` — rebuild SDT elements from scratch. Output in `ContractTemplates/NewFixedTemplates/`. Didn't work.
- `accept_changes.py` — accept Track Changes before migration.

### 2. Add-in migration button (May 2026, FAILED)
- Added `migrateContentControls()` in `taskpane.ts` + button in HTML.
- Word JS API finds 0 controls in old templates. Dead end.

### 3. VBA Macro (May 18 2026, IN PROGRESS)
Files: `DiagnoseControls.bas`, `MigrateTemplates.bas`

#### What we tried:
- **DiagnoseControls** — confirmed VBA sees 42 body + 6 footer controls. WORKS.
- **Clean in place** (remove binding, update tag/title without deleting control) — 48/48 cleaned, BUT Word JS API still doesn't see them. Old SP controls cannot be reused.
- **Delete + recreate using position offsets** — 26/48 succeeded. 22 failed because positions shift after each deletion.
- **Delete + recreate using bookmarks** — FAILED, error #4605 (can't bookmark inside CC).
- **Delete + recreate using Range variable** — WORKS! 48/48 on אדריכל ראשי.docx. Fallback to PlainText for nested controls.

#### Bugs fixed:
- `Chr$()` doesn't support Hebrew (values >255) — changed to `ChrW$()`
- `step` is a VBA reserved word — renamed to `stp`

## Bookmark Approach — ABANDONED (May 25 2026)
Bookmark approach failed — error #4605: can't add bookmark to a range inside a content control.

## Range Variable Approach — WORKING (May 25 2026)
Replaced bookmarks with saving `cc.Range` to a VBA Range variable before deletion.
The Range object stays valid after `cc.Delete False` since text is kept.
- Added fallback: if RichText fails, tries PlainText (fixes cntTzadA/cntTzadB).
- **Result: 48/48 controls converted successfully on אדריכל ראשי.docx**
- Add-in confirmed seeing all controls.

## Batch Processing — Documents.Open Doesn't Work (May 26 2026)
VBA `Documents.Open` + process + `SaveAs2` hangs on Save when opening files programmatically.
`MigrateSingleFile`/`MigrateSingleContract` work fine because they use `ActiveDocument`.
**Solution:** Changed `MigrateAllFiles` and `MigrateAllContracts` to work on already-open documents instead of opening files from a folder. User opens all files first, then runs the macro.

## Nested Content Controls (May 26 2026)
Old SP templates have nested CCs (same tag inside parent with same tag).
After deleting inner CC, can't create new one because parent CC blocks it.
**Solution:** In `MigrateContracts.bas`, if creating new CC fails, delete the parent CC first, then retry.

## Current Status (May 26 2026)
Both migration macros are **DONE and WORKING**:
- `MigrateTemplates.bas` — template migration (empty placeholders) — **TESTED OK**
- `MigrateContracts.bas` — contract migration (preserves values) — **TESTED OK** on 2 contracts

### How to use:
1. Open all files to migrate in Word (make sure they're not in Protected View)
2. Alt+F8 → run `MigrateAllFiles` (templates) or `MigrateAllContracts` (contracts)
3. Macro lists all open documents, asks for confirmation, processes each one
4. Output saved to `New Templates/` or `New Contracts/`, then document is closed

For a single file: open it, Alt+F8 → `MigrateSingleFile` or `MigrateSingleContract`

## Files
- `MigrateTemplates.bas` — for blank templates: replaces content with `[placeholder]` in blue. Output: `New Templates/`
- `MigrateContracts.bas` — for filled-in contracts: preserves existing values. Output: `New Contracts/`
- `DiagnoseControls.bas` — diagnostic macro to list all content controls

## Tag Mappings (old -> new)
| Old Tag | New Tag |
|---------|---------|
| `cntTzadB_x002C__x0020_cntTzadC_x002C__x0020_cntTzadD` | `cntTzadB` |
| `cmtTzadAName` | `cntPartyAName` |
| `cntLocalAuth` | `cntMunicipality` |
| `cntContractType` | `cntTemplateName` |
| `_dlc_DocId` | `cntContractNumber` |
| `cntJobDesc` | `cntWorkDescription` |

## Folder Structure
- `New Templates/` — output for blank template migration
- `New Contracts/` — output for filled contract migration
- `Old Templates/` — old templates (for reference)
- `Sourc of truth/` — 3 manually edited reference templates

## Data Sync: Document → SharePoint Columns (June 21 2026)

After migration, the data exists in the document content controls but NOT in the SP library columns.
Two approaches built for syncing data back:

### Approach 1: VBA Export + PowerShell PnP (for bulk — 400 docs)
- `ExportContractData.bas` — VBA macro that reads all CC values from open documents, exports to `ExportedData.csv`
- `UpdateSharePoint.ps1` — PnP PowerShell script that reads the CSV and updates SP library columns
- Matching is by filename (FileLeafRef)
- Usage: open docs → run `ExportAllContracts` → run `.\UpdateSharePoint.ps1 -SiteUrl "..." -LibraryName "..."`
- Supports `-WhatIf` flag for dry run

### Approach 2: Add-in "Load from Document" button (for one-by-one)
- New button "טען נתונים מהמסמך" added to the Add-in panel
- Reads all content control values into the UI panel
- User can review, then click "שמור שינויים במערכת" to push to SP columns
- Uses existing FieldsUpdateHelper → Power Automate flow

### Tag → SP Column mapping (used by both)
| CC Tag | SP Column | CC Tag | SP Column |
|--------|-----------|--------|-----------|
| cntContractNumber | ContractNumber | cntTzadA | recipient |
| cntContractVersion | contractVersion | cntTzadB | otherSides |
| cntTemplateName | ContractTemplate | cntPartyAName | partyAName |
| cntProjectName | project | cmtTzadAPercent | PartyAContactNamePercent |
| cntSite | siteName | cntCostCompMethod | CostCompMethod |
| cntMunicipality | Municipality | cntCostContractScope | CostContractScope |
| cntWorkDescription | WorkDescription | cntCostCurrency | CostCurrency |
| cntSignDate | signDate | cntCostIndexType | CostIndexType |
| cntStartDate | StartDate | cntCostBaseIndexDate | CostBaseIndexDate |
| cntDurationMonths | DurationMonths | cntCostIndexMode | CostIndexMode |
| cntExpectedEndDate | ExpectedEndDate | cntCostIndexPoints | CostIndexPoints |
| cntStatus | status | cntCostPaymentTerms | CostPaymentTerms |
| cntSupllierName | supplierName | cntCustomField1-8 | customField1-8 |

## CSV Export Integrated into MigrateAllContracts (June 22 2026)
Merged `ExportContractData.bas` functionality directly into `MigrateContracts.bas`:
- `MigrateAllContracts` now exports CSV automatically during migration (no separate step)
- CSV includes fallback lookups for old tags (e.g. `_dlc_DocId` → `cntContractNumber`)
- `FixContractNumber` handles duplicated CONT numbers (e.g. "CONT-5-668CONT-10-2195" → "CONT-10-2195")
- `SiteName` extracted from parent folder path (more reliable than CC value)
- Output filename includes contract number: `baseName - CONT-xx-xxxx.docx`

## Error Handling Fix (June 23 2026)
**Problem:** If `ProcessContract` threw an error on one document, the entire `MigrateAllContracts` loop crashed — no remaining documents were processed.
**Also:** `newFileName` variable was never assigned a value (bug), causing `SaveAs2` to fail.

**Fix:**
- Wrapped `ProcessContract` call in `On Error Resume Next` — if a doc fails, it's logged and skipped
- Added `newFileName` construction: `baseName & " - " & contractNum & ".docx"`
- Failed docs are closed gracefully, loop continues to next document
- Summary shows both successes and failures at the end

## Migration Progress (June 24 2026)

### Completed folders (Old Contracts):
- **א-ג (A-G)** — DONE, but SiteName was not extracted from folder (needs CSV fix)
- **נתניה - מגרש 1409** — DONE (with error handling fix applied mid-session)

### Known issues in migrated batches:
- SiteName missing/wrong for folders א-ג (early batches before folder extraction was added)
- Some deeply nested CCs fail to convert (FAIL at 6c-retry-9) — data still captured via old tag fallback in CSV
- Some documents have duplicate CONT numbers — handled by FixContractNumber

### Remaining folders (Old Contracts):
All other folders from ד onwards (except נתניה - מגרש 1409 which is done).
Check `Old Contracts/` for the full list of remaining subfolders.

### Workflow reminder:
1. Open all .docx files from a subfolder in Word
2. Make sure none are in Protected View
3. Alt+F8 → `MigrateAllContracts`
4. Output goes to `New Contracts/` + `ExportedData.csv` is written
5. Run `UpdateSharePoint.ps1` to push CSV data to SP columns

## Contract Migration — COMPLETE (June 24-29 2026)
All Old Contracts folders migrated. All documents uploaded to new SP library.

## Data Sync from Old SP (June 29-30 2026)
**Problem:** Many SP columns were empty after migration — docCreator, cost fields, etc.
**Solution:** `FillMissingData.ps1` script:
- Reads `oldSiteData_en.csv` (exported from old SP library, converted to English headers)
- Matches old→new by: (1) CONT number, (2) folder name + filename
- Fills only empty fields — does NOT overwrite existing data
- **Result:** 196 contracts updated successfully, 109 not found (not uploaded to new library)
- Unmatched list saved in `unmatched_contracts.csv`

### Key files for data sync:
- `oldSiteData.xlsx` — raw export from old SP library
- `oldSiteData_en.csv` — same data with English column headers (generated by Python)
- `FillMissingData.ps1` — PnP script to fill missing data
- `new_library_items.csv` — snapshot of all items in new SP library (for analysis)

## Panel Fixes (June 30 2026)
Fixed issues with add-in panel not displaying loaded data:

1. **Template field** — was never loaded from SP. Added `pickField(fields, "ContractTemplate")` to `applyLoadedFieldsToUI`.
2. **Version field** — was always showing SP file version instead of contract version. Fixed to show the higher of the two.
3. **setSelectValue silent failure** — if a value from SP didn't exist in the dropdown options, it was silently ignored. Fixed to auto-add missing values as new options. This fixed template, site, municipality, and status fields.
4. **"Load from Document" button** — hidden with `display:none` (code preserved for future use).
5. **Extra fields (custom 1-8)** — for migrated documents without template mapping in extraFields list, now shows fields read-only with yellow notice: "מסמך זה עבר הסבה..."

## Template Migration — COMPLETE (June 30 2026)
All 22 templates migrated locally (19 from Old Templates + 3 new).
13 uploaded to SP so far, 9 remaining to upload.

## Extra Fields Fix (July 6 2026)
1. **List name bug** — `EXTRA_FIELDS_LIST_DISPLAY_NAME` was `"extraFields"`, should be `"ניהול שדות נוספים"`. Fixed.
2. **Template matching** — Template field in the list returns as string (not lookup array). Changed `loadAndApplyExtraFields` to match by template name from `templateSelect` instead of LookupId from filename.
3. **Migrated docs** — Fields beyond the label range (orphaned) shown read-only with yellow notice. Fields within label range shown editable with labels from management list.
4. **Removed `loadFromDocumentIntoUI` button** — already hidden with `display:none`, code preserved.

## Server Migration (July 6 2026)
Moving add-in hosting from `https://knowedge.co.il/matrix/downloads/` to `https://m.res.afi-g.com/contracts/`.
- Updated `manifest.xml` — all URLs
- Updated `taskpane.ts` — MSAL redirectUri
- **Pending:** upload dist/ files, update Azure AD redirect URI, deploy manifest via M365 Admin Center on customer tenant

## Notes
- `ContractTemplates/New Templates/` is the source of truth — do NOT delete
- Add-in being migrated to `https://m.res.afi-g.com/contracts/` — upload BOTH `dist/taskpane.js` AND `dist/taskpane.html` after `npm run build`
