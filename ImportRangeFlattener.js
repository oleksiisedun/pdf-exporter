/**
 * Literal placeholder text Google Sheets shows while a cross-spreadsheet
 * IMPORTRANGE is still resolving. There is no public Apps Script API that
 * reports calculation status; this string is the only observable signal.
 * Verify this exact string (ellipsis character vs. three periods, locale)
 * against a real IMPORTRANGE cell before relying on it in production.
 * @type {string}
 */
const IMPORTRANGE_LOADING_PLACEHOLDER = 'Loading...';

/**
 * Removes double-quoted string literals from a formula so their contents
 * aren't mistaken for a function call (e.g. a cell whose formula is
 * ="Use IMPORTRANGE()" should not match).
 * @param {string} formula
 * @returns {string}
 */
function stripStringLiterals(formula) {
  return formula.replace(/"(?:[^"\\]|\\.)*"/g, '""');
}

/**
 * @param {string} formula - Formula text as returned by Range.getFormula()/getFormulas(), including the leading '='.
 * @returns {boolean} Whether the formula calls IMPORTRANGE anywhere, including nested inside another function (e.g. SUM(IMPORTRANGE(...))).
 */
function formulaCallsImportRange(formula) {
  if (!formula) return false;
  return /\bIMPORTRANGE\s*\(/i.test(stripStringLiterals(formula));
}

/**
 * Like Spreadsheet.getSheetByName, but throws a descriptive error instead of
 * returning null when the sheet doesn't exist.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheetOrThrow(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`getSheetOrThrow: sheet "${sheetName}" not found in spreadsheet "${spreadsheet.getName()}".`);
  }
  return sheet;
}

/**
 * Bulk-reads a sheet's data range once and locates every cell whose own
 * formula calls IMPORTRANGE, alongside the formulas/values grids so callers
 * don't need to re-fetch them.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{formulas: string[][], values: any[][], anchors: {row:number,col:number}[]}} 0-based anchor positions.
 */
function scanImportRangeAnchors(sheet) {
  const dataRange = sheet.getDataRange();
  const numRows = dataRange.getNumRows();
  const numCols = dataRange.getNumColumns();
  if (numRows === 0 || numCols === 0) return { formulas: [], values: [], anchors: [] };

  const formulas = dataRange.getFormulas();
  const values = dataRange.getValues();
  const anchors = [];
  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      if (formulaCallsImportRange(formulas[r][c])) anchors.push({ row: r, col: c });
    }
  }
  return { formulas, values, anchors };
}

/**
 * Finds every IMPORTRANGE-anchored cell in a sheet's data range and its
 * best-effort spilled output extent. Apps Script's SpreadsheetApp service has
 * no API to query an array formula's result dimensions directly, so this
 * expands right along the anchor's row, then down along the anchor's column,
 * while the next cell has no formula of its own and a non-blank value — the
 * same signal a spilled (not manually typed) cell leaves behind. Uses only
 * the bulk formulas/values grids already read by scanImportRangeAnchors, no
 * additional per-cell API calls.
 *
 * Known limitation: if the imported range itself contains a blank cell (a
 * gap in the source data), this under-detects the spill's true extent past
 * that gap. If unrelated non-blank data sits immediately adjacent to the
 * spill with no gap, this over-detects — harmless, since the value written
 * back for an over-detected cell is simply its own current value (a no-op).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Must be on the SOURCE spreadsheet (read-only).
 * @returns {{row:number,col:number,numRows:number,numCols:number}[]}
 */
function findImportRangeSpillRegions(sheet) {
  const { formulas, values, anchors } = scanImportRangeAnchors(sheet);
  const numDataRows = values.length;
  const numDataCols = numDataRows > 0 ? values[0].length : 0;

  return anchors.map(({ row, col }) => {
    let numCols = 1;
    while (col + numCols < numDataCols && !formulas[row][col + numCols] && values[row][col + numCols] !== '') numCols++;

    let numRows = 1;
    while (row + numRows < numDataRows && !formulas[row + numRows][col] && values[row + numRows][col] !== '') numRows++;

    return { row, col, numRows, numCols };
  });
}

/**
 * Freezes every IMPORTRANGE-anchored cell (and its best-effort detected
 * spill region) in the included sheets to a static value read from the
 * SOURCE spreadsheet, writing the result into the DUPLICATE. This is
 * necessary because IMPORTRANGE's cross-spreadsheet access grant is tied to
 * the specific destination file ID — the duplicate is a brand-new file that
 * was never granted access, so its own live evaluation of the formula would
 * render as an access-denied error instead of data. The source's
 * already-granted access is what makes the value written here correct.
 *
 * Every anchor across all included sheets is cleared on the duplicate first
 * (releasing the still-live-but-about-to-error array formula's ownership of
 * its spilled cells) and flushed, before any value is written — so a
 * partially-cleared spill never blocks a later write on the same sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} duplicateSpreadsheet
 * @param {string[]} includedSheetNames
 * @returns {void}
 */
function flattenImportRangeCells(sourceSpreadsheet, duplicateSpreadsheet, includedSheetNames) {
  const perSheetRegions = includedSheetNames
    .map((sheetName) => ({ sheetName, regions: findImportRangeSpillRegions(getSheetOrThrow(sourceSpreadsheet, sheetName)) }))
    .filter(({ regions }) => regions.length > 0);

  if (perSheetRegions.length === 0) return;

  perSheetRegions.forEach(({ sheetName, regions }) => {
    const dupSheet = getSheetOrThrow(duplicateSpreadsheet, sheetName);
    regions.forEach(({ row, col }) => dupSheet.getRange(row + 1, col + 1).clearContent());
  });
  SpreadsheetApp.flush();

  perSheetRegions.forEach(({ sheetName, regions }) => {
    const sourceSheet = getSheetOrThrow(sourceSpreadsheet, sheetName);
    const dupSheet = getSheetOrThrow(duplicateSpreadsheet, sheetName);
    regions.forEach(({ row, col, numRows, numCols }) => {
      const values = sourceSheet.getRange(row + 1, col + 1, numRows, numCols).getValues();
      dupSheet.getRange(row + 1, col + 1, numRows, numCols).setValues(values);
    });
  });
}

/**
 * Blocks until every IMPORTRANGE anchor cell in the included sheets of the
 * SOURCE spreadsheet has finished resolving (no longer shows the
 * "Loading..." placeholder), or throws once timeoutMs elapses. Must run
 * before duplication: flattenImportRangeCells later reads these same cells'
 * values to freeze into the export, and without this wait a transient
 * "Loading..." could get baked in as a literal string.
 * Pass timeoutMs: 0 to skip this check entirely.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet
 * @param {string[]} includedSheetNames
 * @param {number} [timeoutMs] - Defaults to 60000 (1 min). Uses `??`, not `||`, so 0 is a meaningful opt-out.
 * @param {number} [pollIntervalMs] - Delay between re-checks. Defaults to 2000 (2s); floored at 250ms.
 * @returns {void}
 */
function waitForImportRangesToSettle(sourceSpreadsheet, includedSheetNames, timeoutMs, pollIntervalMs) {
  const timeout = timeoutMs ?? 60000;
  if (timeout === 0) return;
  const interval = Math.max(pollIntervalMs ?? 2000, 250);
  const deadline = Date.now() + timeout;

  while (true) {
    const stillLoading = findImportRangeAnchorStillLoading(sourceSpreadsheet, includedSheetNames);
    if (!stillLoading) return;
    if (Date.now() >= deadline) {
      throw new Error(
        `waitForImportRangesToSettle: "${stillLoading.sheetName}"!${stillLoading.a1Notation} still shows ` +
        `"${IMPORTRANGE_LOADING_PLACEHOLDER}" after waiting ${timeout}ms; aborting export before it bakes that ` +
        'placeholder into a static value. Pass a larger importRangeWaitTimeoutMs if this import genuinely needs ' +
        'longer, or importRangeWaitTimeoutMs: 0 to skip this check.'
      );
    }
    Utilities.sleep(interval);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet
 * @param {string[]} includedSheetNames
 * @returns {{sheetName:string,a1Notation:string}|null}
 */
function findImportRangeAnchorStillLoading(sourceSpreadsheet, includedSheetNames) {
  for (const sheetName of includedSheetNames) {
    const sheet = getSheetOrThrow(sourceSpreadsheet, sheetName);
    const { values, anchors } = scanImportRangeAnchors(sheet);
    for (const { row, col } of anchors) {
      if (values[row][col] === IMPORTRANGE_LOADING_PLACEHOLDER) {
        return { sheetName, a1Notation: sheet.getRange(row + 1, col + 1).getA1Notation() };
      }
    }
  }
  return null;
}
