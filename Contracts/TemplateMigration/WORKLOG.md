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

## Notes
- `ContractTemplates/New Templates/` is the source of truth — do NOT delete
- Add-in deployed to `https://knowedge.co.il/matrix/downloads/` — upload `dist/` files after `npm run build`
