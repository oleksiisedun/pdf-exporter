# pdf-exporter

A standalone Google Apps Script **library** project (not container-bound to any spreadsheet). It's added as a dependency by other Apps Script projects that need to export a Google Sheets spreadsheet to PDF — see README.md for the consumer-facing API and usage examples.

## Invariant: never mutate the source spreadsheet

Every export duplicates the source spreadsheet to a temporary Drive file first (`duplicateSpreadsheetFile` in `SpreadsheetDuplicator.js`), and all changes — hiding excluded sheets, flattening IMPORTRANGE cells, anything run via the `beforeExport` hook — happen on that copy. `Main.js`'s `exportSpreadsheetToPdfBlob` always deletes the copy in a `finally`, even on error. Any change to this library must preserve that: never call a mutating method (`hideSheet`, `setValue`, etc.) on the spreadsheet resolved from `options.spreadsheetId`/the active spreadsheet — only on the duplicate. Values read *from* the source for flattening (IMPORTRANGE cells) are fine — only writes are forbidden.

`options.direct: true` is the one sanctioned exception: it skips duplication entirely and reads+exports the source directly, with no mutation involved at all (no copy, no hiding, no flattening). Because it never mutates anything, it doesn't violate the invariant — but it also can't support `includeSheets`/`excludeSheets`/`hideColumns`/`beforeExport` (all of which rely on a copy to act on), so `exportSpreadsheetToPdfBlob` throws if `direct` is combined with any of them. Keep that guard if refactoring this code path.

## Why IMPORTRANGE gets special handling

A Drive copy is a new file ID, so it never inherits the source's `IMPORTRANGE` "Allow access" grant — there's no API to grant it programmatically. `ImportRangeFlattener.js` detects `IMPORTRANGE` cells (and their best-effort spill extent) on the source before duplicating, waits for any still-resolving ones, then after duplication overwrites those cells on the copy with the source's already-authorized values. Don't remove this thinking it's dead weight — without it, any spreadsheet using `IMPORTRANGE` exports a PDF full of access-denied errors.

## File layout

Only `src/` is deployed (`.clasp.json` has `"rootDir": "src"`); tooling and docs stay at the repo root. Keep `src/` flat — clasp flattens subfolders into script file names.

- `src/Main.js` — public API (`exportSpreadsheetToPdfBlob`, `exportSpreadsheetToPdfFile`) and sheet include/exclude resolution.
- `src/SpreadsheetDuplicator.js` — Drive-copy creation and hiding sheets/columns on the copy. The copy is always placed explicitly in the source file's own parent folder (never Drive's default root) — `DriveUtils.js`'s orphan sweep only searches the source's parent folder(s), so a copy left in root would never get cleaned up.
- `src/ImportRangeFlattener.js` — IMPORTRANGE detection, the "Loading..." wait, and flattening cells on the copy (see above). Detection of "still loading" relies on matching the literal string `'Loading...'` (`IMPORTRANGE_LOADING_PLACEHOLDER`) — there's no Apps Script API for calculation status. This is unverified against locale/character variants; treat it as fragile if touching this file.
- `src/PdfFetch.js` — builds the `export?format=pdf` query string from `PdfExportOptions` and fetches the PDF bytes.
- `src/DriveUtils.js` — saving to a Drive folder, retrying deletes, sweeping orphaned temp copies left behind by hard-killed executions, and `getParentFoldersOrRoot` (shared by the sweep and by `SpreadsheetDuplicator.js`'s copy placement).

## Commands

```bash
npm run check      # Lint + type check (run after every edit; fast, offline)
npm run lint       # ESLint only
npm run typecheck  # tsc checkJs over src/ — resolves cross-file globals against @types/google-apps-script
clasp open   # Open this project in the Apps Script editor
clasp push   # Push local changes to Apps Script (requires confirmation — see Deploying)
clasp pull   # Pull changes made in the Apps Script editor back to local files
```

Cutting a new library deployment (**Deploy > New deployment** in the Apps Script editor) is a manual UI step with no `clasp` equivalent — see README.md's "Deploying the library" section.

## Testing

Run `npm run check` after every edit for static checks (`no-undef` is off in ESLint because all files share one global scope — `typecheck` is what catches undefined cross-file references). No automated test framework in Apps Script. Test manually from the Apps Script editor — see README.md's "Testing" section for the checklist.

## Deploying

`clasp push` and cutting a new library deployment both require explicit user confirmation before running — never do either without being asked in that moment (per the user's standing instruction not to auto-push Apps Script code).
