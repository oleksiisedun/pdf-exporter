/**
 * Name prefix for temporary Drive copies made during export. Shared between
 * the copy step below and DriveUtils.js's orphan sweep, which identifies
 * leftover copies from previous runs by this prefix.
 * @type {string}
 */
const PDF_EXPORT_TEMP_FILE_PREFIX = '__pdf_export_tmp__';

/**
 * Prefix for every error thrown by the public API, so a caller (or an error
 * log) can always tell which library function raised it. Shared across
 * source files — e.g. SpreadsheetDuplicator.js's hideColumnsBySheetName
 * throws on the caller's behalf for an invalid hideColumns entry — so it's
 * declared once here rather than duplicated as a string literal per throw
 * site.
 * @type {string}
 */
const EXPORT_ERROR_PREFIX = 'exportSpreadsheetToPdfBlob: ';

/**
 * @typedef {Object} ExportPdfOptions
 * @property {string} [spreadsheetId] - ID of the source Google Sheets spreadsheet. Never mutated. Defaults to the active spreadsheet (SpreadsheetApp.getActiveSpreadsheet()) if omitted — only resolvable when the library is called from a bound script context (e.g. a container-bound script or a simple/installable trigger), not from a standalone script or webapp with no active spreadsheet.
 * @property {string[]} [includeSheets] - Sheet names to include. Mutually exclusive with excludeSheets.
 * @property {string[]} [excludeSheets] - Sheet names to exclude. Mutually exclusive with includeSheets.
 * @property {Object<string, Array<number|string>>} [hideColumns] - Columns to hide per sheet, keyed by sheet name: `{ 'Report': [3, 'E:G'] }`. Each column is a 1-based index, a column letter (`'C'`), or a letter range (`'E:G'`). Applied on the temporary copy after IMPORTRANGE flattening and before `beforeExport`. Columns are hidden, not deleted, so formulas referencing them keep working. Not allowed together with `direct: true` — there is no copy to hide them on.
 * @property {string} [fileName] - Base file name (no extension) for the export; defaults to the source spreadsheet's name. The current date/time is always appended.
 * @property {PdfExportOptions} [pdfOptions] - PDF rendering options (paper size, margins, ...). See PdfFetch.js.
 * @property {(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) => void} [beforeExport] - Optional hook invoked on the temporary Drive copy, after excluded sheets are hidden and IMPORTRANGE cells are flattened, but before the PDF is fetched. Use it for any prep beyond whole-sheet include/exclude — collapsing row/column groups, toggling a display flag, etc. Runs against the copy, never the source. Not allowed together with `direct: true` — there is no copy for it to run against.
 * @property {number} [importRangeWaitTimeoutMs] - Max time (ms) to wait for pending IMPORTRANGE calculations on the source to settle before exporting. Defaults to 60000 (1 min). Pass 0 to skip the wait entirely.
 * @property {number} [importRangeWaitPollIntervalMs] - Delay (ms) between IMPORTRANGE status re-checks while waiting. Defaults to 2000 (2s).
 * @property {boolean} [direct] - Skip the Drive-copy step entirely and export the source spreadsheet as-is: no duplicate file, no sheet hiding, no IMPORTRANGE flattening, no `beforeExport`. Faster for simple cases that don't need include/exclude or pre-export prep. Since there's no copy to hide sheets on, this is mutually exclusive with `includeSheets`, `excludeSheets`, `hideColumns`, and `beforeExport`. IMPORTRANGE cells work fine here since the source already holds its own access grant — only the still-loading wait (`importRangeWaitTimeoutMs`) still applies, to avoid capturing a "Loading..." placeholder in the PDF.
 */

/**
 * Exports a Google Sheets spreadsheet to a PDF Blob.
 *
 * By default (`direct` not set), the source spreadsheet is only ever read,
 * never mutated — all changes (hiding excluded sheets, flattening IMPORTRANGE
 * cells, anything done in `beforeExport`) happen on a temporary Drive copy,
 * which is deleted afterward even if an error occurs.
 *
 * IMPORTRANGE cells are flattened to a static value read from the source
 * because IMPORTRANGE's cross-spreadsheet access grant is tied to the
 * specific destination file ID: the duplicate is a brand-new file that was
 * never granted access, so its own live evaluation would render as an
 * access-denied error instead of data. See ImportRangeFlattener.js.
 *
 * Pass `direct: true` to skip the copy entirely and export the source
 * spreadsheet as-is, for simple cases that don't need sheet include/exclude
 * or a `beforeExport` hook.
 * @param {ExportPdfOptions} options
 * @returns {GoogleAppsScript.Base.Blob}
 */
function exportSpreadsheetToPdfBlob(options) {
  const {
    spreadsheetId,
    includeSheets,
    excludeSheets,
    hideColumns,
    fileName,
    pdfOptions,
    beforeExport,
    direct,
    importRangeWaitTimeoutMs,
    importRangeWaitPollIntervalMs,
  } = options || {};

  const sourceSs = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!sourceSs) {
    throw new Error(`${EXPORT_ERROR_PREFIX}options.spreadsheetId was not provided and there is no active spreadsheet.`);
  }

  const hasHideColumns = !!hideColumns && Object.keys(hideColumns).length > 0;

  if (direct) {
    if ((includeSheets && includeSheets.length) || (excludeSheets && excludeSheets.length) || hasHideColumns || beforeExport) {
      throw new Error(
        `${EXPORT_ERROR_PREFIX}direct: true cannot be combined with includeSheets, excludeSheets, hideColumns, or beforeExport — there is no Drive copy for these to act on.`,
      );
    }
    const allSheetNames = getAllSheetNames(sourceSs);
    waitForImportRangesToSettle(sourceSs, allSheetNames, importRangeWaitTimeoutMs, importRangeWaitPollIntervalMs);
    const baseFileName = fileName || sourceSs.getName();
    const timestampedFileName = buildTimestampedFileName(new Date(), baseFileName, sourceSs.getSpreadsheetTimeZone());
    return fetchPdfBlob(sourceSs.getId(), pdfOptions).setName(`${timestampedFileName}.pdf`);
  }

  const resolvedSpreadsheetId = sourceSs.getId();
  const allSheetNames = getAllSheetNames(sourceSs);
  const includedSheetNames = resolveIncludedSheetNames(allSheetNames, includeSheets, excludeSheets);
  const excludedSheetNames = allSheetNames.filter((n) => !includedSheetNames.includes(n));
  if (hasHideColumns) assertSheetNamesExist(Object.keys(hideColumns), new Set(allSheetNames), 'hideColumns');

  waitForImportRangesToSettle(sourceSs, includedSheetNames, importRangeWaitTimeoutMs, importRangeWaitPollIntervalMs);

  const baseFileName = fileName || sourceSs.getName();
  const timestampedFileName = buildTimestampedFileName(new Date(), baseFileName, sourceSs.getSpreadsheetTimeZone());

  cleanUpOrphanedExportTempFiles(resolvedSpreadsheetId, PDF_EXPORT_TEMP_FILE_PREFIX);
  const copiedFile = duplicateSpreadsheetFile(resolvedSpreadsheetId, `${PDF_EXPORT_TEMP_FILE_PREFIX}${baseFileName}__${Date.now()}`);

  try {
    const dupSs = SpreadsheetApp.openById(copiedFile.getId());
    hideSheetsByName(dupSs, excludedSheetNames);
    flattenImportRangeCells(sourceSs, dupSs, includedSheetNames);
    if (hasHideColumns) hideColumnsBySheetName(dupSs, hideColumns);
    if (beforeExport) beforeExport(dupSs);
    SpreadsheetApp.flush();
    return fetchPdfBlob(copiedFile.getId(), pdfOptions).setName(`${timestampedFileName}.pdf`);
  } finally {
    deleteFileWithRetry(copiedFile.getId());
  }
}

/**
 * Convenience helper: exports a spreadsheet to PDF and saves it into a Drive folder.
 * @param {ExportPdfOptions} options
 * @param {string} folderId - Destination Drive folder ID.
 * @param {string} [fileName] - Overrides the exported file's name (including the date/time suffix that would otherwise be appended).
 * @returns {GoogleAppsScript.Drive.File}
 */
function exportSpreadsheetToPdfFile(options, folderId, fileName) {
  const blob = exportSpreadsheetToPdfBlob(options);
  return saveBlobToDriveFolder(blob, folderId, fileName);
}

/**
 * Formats a date as "dd.MM.yyyy HH:mm" in the given IANA time zone, using
 * `Intl.DateTimeFormat` rather than `Utilities.formatDate` — `Intl` is
 * available in both the Apps Script V8 runtime and plain Node, so this stays
 * pure logic that can be unit tested with a fixed date, instead of depending
 * on an Apps Script service.
 * @param {Date} date
 * @param {string} timeZone - IANA time zone, e.g. from Spreadsheet.getSpreadsheetTimeZone().
 * @returns {string}
 */
function formatDateDdMmYyyyHhMm(date, timeZone) {
  const partValues = Object.fromEntries(
    new Intl.DateTimeFormat('en-CA', {
      timeZone,
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hourCycle: 'h23',
    })
      .formatToParts(date)
      .map((part) => [part.type, part.value]),
  );
  return `${partValues.day}.${partValues.month}.${partValues.year} ${partValues.hour}:${partValues.minute}`;
}

/**
 * Appends a date/time to a base file name, in DD.MM.YYYY HH:MM format. Takes
 * `date` as a parameter (rather than reading `new Date()` internally) so this
 * pure formatting logic can be unit tested with a fixed date.
 * @param {Date} date
 * @param {string} baseFileName
 * @param {string} timeZone - IANA time zone, e.g. from Spreadsheet.getSpreadsheetTimeZone().
 * @returns {string}
 */
function buildTimestampedFileName(date, baseFileName, timeZone) {
  return `${baseFileName} ${formatDateDdMmYyyyHhMm(date, timeZone)}`;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @returns {string[]} Sheet names in the spreadsheet's own tab order.
 */
function getAllSheetNames(spreadsheet) {
  return spreadsheet.getSheets().map((s) => s.getName());
}

/**
 * Resolves the final ordered list of sheet names to include, preserving the
 * source spreadsheet's original tab order. If neither includeSheets nor
 * excludeSheets is supplied, all sheets are included.
 * @param {string[]} allSheetNames
 * @param {string[]|undefined} includeSheets
 * @param {string[]|undefined} excludeSheets
 * @returns {string[]}
 */
function resolveIncludedSheetNames(allSheetNames, includeSheets, excludeSheets) {
  const hasInclude = Array.isArray(includeSheets) && includeSheets.length > 0;
  const hasExclude = Array.isArray(excludeSheets) && excludeSheets.length > 0;
  if (hasInclude && hasExclude) {
    throw new Error(`${EXPORT_ERROR_PREFIX}pass either includeSheets or excludeSheets, not both.`);
  }

  const allSet = new Set(allSheetNames);
  if (hasInclude) {
    assertSheetNamesExist(includeSheets, allSet, 'includeSheets');
    return allSheetNames.filter((name) => includeSheets.includes(name));
  }

  if (hasExclude) assertSheetNamesExist(excludeSheets, allSet, 'excludeSheets');
  const excludeSet = new Set(excludeSheets || []);
  const included = allSheetNames.filter((name) => !excludeSet.has(name));
  if (included.length === 0) {
    throw new Error(`${EXPORT_ERROR_PREFIX}resulting included-sheet set is empty.`);
  }
  return included;
}

/**
 * @param {string[]} names
 * @param {Set<string>} allSheetNamesSet
 * @param {string} optionLabel
 * @returns {void}
 */
function assertSheetNamesExist(names, allSheetNamesSet, optionLabel) {
  const missing = names.filter((name) => !allSheetNamesSet.has(name));
  if (missing.length > 0) {
    throw new Error(`${EXPORT_ERROR_PREFIX}${optionLabel} contains unknown sheet name(s): ${missing.join(', ')}`);
  }
}
