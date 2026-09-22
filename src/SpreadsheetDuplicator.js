/**
 * Creates a full copy of the source spreadsheet file in Drive, placed
 * explicitly in the source's own parent folder (or Drive's root, if the
 * source has none — see getParentFoldersOrRoot in DriveUtils.js). Without an
 * explicit destination, `File.makeCopy(name)` drops the copy in the current
 * user's Drive root regardless of where the source lives — which would
 * silently break `cleanUpOrphanedExportTempFiles` in DriveUtils.js, since it
 * searches the source's parent folder(s) for leftover copies and would never
 * find any placed in root instead.
 * @param {string} spreadsheetId
 * @param {string} copyName
 * @returns {GoogleAppsScript.Drive.File}
 */
function duplicateSpreadsheetFile(spreadsheetId, copyName) {
  const sourceFile = DriveApp.getFileById(spreadsheetId);
  const [destinationFolder] = getParentFoldersOrRoot(sourceFile);
  return sourceFile.makeCopy(copyName, destinationFolder);
}

/**
 * Hides the given sheets (by name) in a spreadsheet, if present. Hiding
 * rather than deleting keeps any formula on a still-visible sheet that
 * references an "excluded" one intact (no #REF!), and Google's PDF export
 * endpoint already omits hidden sheets from the output.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string[]} sheetNames
 * @returns {void}
 */
function hideSheetsByName(spreadsheet, sheetNames) {
  sheetNames.forEach((name) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) sheet.hideSheet();
  });
}

/**
 * Hides columns on the given sheets. Hiding rather than deleting keeps any
 * formula that references those columns intact, and Google's PDF export
 * endpoint omits hidden columns from the output.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {Object<string, Array<number|string>>} columnsBySheetName - Sheet name → columns to hide. Each column is a 1-based index (`3`), a column letter (`'C'`), or a letter range (`'E:G'`).
 * @returns {void}
 */
function hideColumnsBySheetName(spreadsheet, columnsBySheetName) {
  Object.entries(columnsBySheetName).forEach(([sheetName, columns]) => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;
    columns.forEach((column) => {
      if (typeof column === 'number') {
        sheet.hideColumns(column);
      } else if (typeof column === 'string') {
        sheet.hideColumn(sheet.getRange(column.includes(':') ? column : `${column}:${column}`));
      } else {
        throw new Error(`exportSpreadsheetToPdfBlob: hideColumns["${sheetName}"] contains an invalid column ${JSON.stringify(column)} — use a 1-based index, a column letter, or a letter range like "E:G".`);
      }
    });
  });
}
