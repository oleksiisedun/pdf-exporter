/**
 * @typedef {Object} PdfExportOptions
 * @property {boolean} [portrait] - true for portrait, false for landscape. Defaults to false.
 * @property {string} [size] - Paper size: 'letter' | 'legal' | 'a4' | ... Defaults to 'a4'.
 * @property {boolean} [fitw] - Fit sheet width to one page width. Defaults to true.
 * @property {boolean} [gridlines] - Include cell gridlines. Defaults to false.
 * @property {boolean} [printtitle] - Include the spreadsheet's title above the content. Defaults to false.
 * @property {number} [margin] - Shorthand: sets all four margins (inches). Defaults to 0.5.
 * @property {number} [topMargin] - Overrides the top margin (inches).
 * @property {number} [bottomMargin] - Overrides the bottom margin (inches).
 * @property {number} [leftMargin] - Overrides the left margin (inches).
 * @property {number} [rightMargin] - Overrides the right margin (inches).
 * @property {Object<string,string|number|boolean>} [extraExportParams] - Any other
 *   `export?format=pdf` query param not covered above, passed through verbatim
 *   (e.g. { pagenum: 'CENTER', horizontal_alignment: 'CENTER' }).
 */

/** @type {Partial<PdfExportOptions>} */
const PDF_EXPORT_OPTION_DEFAULTS = {
  portrait: false,
  size: 'a4',
  fitw: true,
  gridlines: false,
  printtitle: false,
  margin: 0.5,
};

/**
 * Builds the `export?format=pdf&...` query string from a PdfExportOptions object,
 * layering caller overrides over PDF_EXPORT_OPTION_DEFAULTS.
 * @param {PdfExportOptions} [pdfOptions]
 * @returns {string}
 */
function buildPdfExportOptionsQueryString(pdfOptions) {
  const opts = { ...PDF_EXPORT_OPTION_DEFAULTS, ...(pdfOptions || {}) };
  const params = {
    format: 'pdf',
    portrait: opts.portrait,
    size: opts.size,
    fitw: opts.fitw,
    gridlines: opts.gridlines,
    printtitle: opts.printtitle,
    top_margin: opts.topMargin ?? opts.margin,
    bottom_margin: opts.bottomMargin ?? opts.margin,
    left_margin: opts.leftMargin ?? opts.margin,
    right_margin: opts.rightMargin ?? opts.margin,
    ...(opts.extraExportParams || {}),
  };

  return Object.entries(params)
    .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
    .join('&');
}

/**
 * Fetches the .pdf bytes for a Google Sheets file via the native export endpoint.
 * The URL is always built from spreadsheetFileId internally — a caller can never
 * inject an arbitrary host through pdfOptions, so there's no need for a
 * host-allowlist check here.
 * @param {string} spreadsheetFileId
 * @param {PdfExportOptions} [pdfOptions]
 * @returns {GoogleAppsScript.Base.Blob}
 */
function fetchPdfBlob(spreadsheetFileId, pdfOptions) {
  const query = buildPdfExportOptionsQueryString(pdfOptions);
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetFileId}/export?${query}`;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
  });
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`fetchPdfBlob: export request failed with status ${code}: ${response.getContentText().slice(0, 300)}`);
  }
  return response.getBlob();
}
