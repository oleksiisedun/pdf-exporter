/**
 * Name prefix for temporary Drive copies made during export. Shared between
 * the copy step below and DriveUtils.js's orphan sweep, which identifies
 * leftover copies from previous runs by this prefix.
 * @type {string}
 */
const PDF_EXPORT_TEMP_FILE_PREFIX = '__pdf_export_tmp__';

/**
 * @typedef {Object} ExportPdfOptions
 * @property {string} [spreadsheetId] - ID of the source Google Sheets spreadsheet. Never mutated. Defaults to the active spreadsheet (SpreadsheetApp.getActiveSpreadsheet()) if omitted — only resolvable when the library is called from a bound script context (e.g. a container-bound script or a simple/installable trigger), not from a standalone script or webapp with no active spreadsheet.
 * @property {string[]} [includeSheets] - Sheet names to include. Mutually exclusive with excludeSheets.
 * @property {string[]} [excludeSheets] - Sheet names to exclude. Mutually exclusive with includeSheets.
 * @property {string} [fileName] - Base file name (no extension) for the export; defaults to the source spreadsheet's name. The current date/time is always appended.
 * @property {PdfExportOptions} [pdfOptions] - PDF rendering options (paper size, margins, ...). See PdfFetch.js.
 * @property {function(GoogleAppsScript.Spreadsheet.Spreadsheet): void} [beforeExport] - Optional hook invoked on the temporary Drive copy, after excluded sheets are hidden and IMPORTRANGE cells are flattened, but before the PDF is fetched. Use it for any prep beyond whole-sheet include/exclude — collapsing row/column groups, toggling a display flag, etc. Runs against the copy, never the source.
 * @property {number} [importRangeWaitTimeoutMs] - Max time (ms) to wait for pending IMPORTRANGE calculations on the source to settle before exporting. Defaults to 60000 (1 min). Pass 0 to skip the wait entirely.
 * @property {number} [importRangeWaitPollIntervalMs] - Delay (ms) between IMPORTRANGE status re-checks while waiting. Defaults to 2000 (2s).
 */

/**
 * Exports a Google Sheets spreadsheet to a PDF Blob. The source spreadsheet is
 * only ever read, never mutated — all changes (hiding excluded sheets,
 * flattening IMPORTRANGE cells, anything done in `beforeExport`) happen on a
 * temporary Drive copy, which is deleted afterward even if an error occurs.
 *
 * IMPORTRANGE cells are flattened to a static value read from the source
 * because IMPORTRANGE's cross-spreadsheet access grant is tied to the
 * specific destination file ID: the duplicate is a brand-new file that was
 * never granted access, so its own live evaluation would render as an
 * access-denied error instead of data. See ImportRangeFlattener.js.
 * @param {ExportPdfOptions} options
 * @returns {GoogleAppsScript.Base.Blob}
 */
function exportSpreadsheetToPdfBlob(options) {
  const {
    spreadsheetId, includeSheets, excludeSheets, fileName, pdfOptions, beforeExport,
    importRangeWaitTimeoutMs, importRangeWaitPollIntervalMs,
  } = options || {};

  const sourceSs = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!sourceSs) {
    throw new Error('exportSpreadsheetToPdfBlob: options.spreadsheetId was not provided and there is no active spreadsheet.');
  }
  const resolvedSpreadsheetId = sourceSs.getId();
  const allSheetNames = sourceSs.getSheets().map((s) => s.getName());
  const includedSheetNames = resolveIncludedSheetNames(allSheetNames, includeSheets, excludeSheets);
  const excludedSheetNames = allSheetNames.filter((n) => !includedSheetNames.includes(n));

  waitForImportRangesToSettle(sourceSs, includedSheetNames, importRangeWaitTimeoutMs, importRangeWaitPollIntervalMs);

  const baseFileName = fileName || sourceSs.getName();
  const timestampedFileName = buildTimestampedFileName(baseFileName, sourceSs.getSpreadsheetTimeZone());

  cleanUpOrphanedExportTempFiles(resolvedSpreadsheetId, PDF_EXPORT_TEMP_FILE_PREFIX);
  const copiedFile = duplicateSpreadsheetFile(resolvedSpreadsheetId, `${PDF_EXPORT_TEMP_FILE_PREFIX}${baseFileName}__${Date.now()}`);

  try {
    const dupSs = SpreadsheetApp.openById(copiedFile.getId());
    hideSheetsByName(dupSs, excludedSheetNames);
    flattenImportRangeCells(sourceSs, dupSs, includedSheetNames);
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
 * Appends the current date/time to a base file name, in DD.MM.YYYY HH:MM format.
 * @param {string} baseFileName
 * @param {string} timeZone - IANA time zone, e.g. from Spreadsheet.getSpreadsheetTimeZone().
 * @returns {string}
 */
function buildTimestampedFileName(baseFileName, timeZone) {
  const timestamp = Utilities.formatDate(new Date(), timeZone, 'dd.MM.yyyy HH:mm');
  return `${baseFileName} ${timestamp}`;
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
    throw new Error('exportSpreadsheetToPdfBlob: pass either includeSheets or excludeSheets, not both.');
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
    throw new Error('exportSpreadsheetToPdfBlob: resulting included-sheet set is empty.');
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
    throw new Error(`exportSpreadsheetToPdfBlob: ${optionLabel} contains unknown sheet name(s): ${missing.join(', ')}`);
  }
}
