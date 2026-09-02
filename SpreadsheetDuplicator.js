/**
 * Creates a full copy of the source spreadsheet file in Drive, placed
 * explicitly in the source's own parent folder (or Drive's root, if the
 * source has none). Without an explicit destination, `File.makeCopy(name)`
 * drops the copy in the current user's Drive root regardless of where the
 * source lives — which would silently break `cleanUpOrphanedExportTempFiles`
 * in DriveUtils.js, since it searches the source's parent folder(s) for
 * leftover copies and would never find any placed in root instead.
 * @param {string} spreadsheetId
 * @param {string} copyName
 * @returns {GoogleAppsScript.Drive.File}
 */
function duplicateSpreadsheetFile(spreadsheetId, copyName) {
  const sourceFile = DriveApp.getFileById(spreadsheetId);
  const parentIterator = sourceFile.getParents();
  const destinationFolder = parentIterator.hasNext() ? parentIterator.next() : DriveApp.getRootFolder();
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
