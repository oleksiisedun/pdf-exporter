'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadFromSrc } = require('./helpers/load-src');

const { buildPdfExportOptionsQueryString } = loadFromSrc('PdfFetch.js', ['buildPdfExportOptionsQueryString']);

/** @param {string} query @returns {Record<string, string>} */
function parseQuery(query) {
  return Object.fromEntries(query.split('&').map((pair) => pair.split('=').map(decodeURIComponent)));
}

test('applies defaults when no options are given', () => {
  const params = parseQuery(buildPdfExportOptionsQueryString());
  assert.equal(params.format, 'pdf');
  assert.equal(params.size, 'a4');
  assert.equal(params.fitw, 'true');
  assert.equal(params.top_margin, '0.5');
});

test('margin overrides all four individual margins', () => {
  const params = parseQuery(buildPdfExportOptionsQueryString({ margin: 1 }));
  assert.equal(params.top_margin, '1');
  assert.equal(params.bottom_margin, '1');
  assert.equal(params.left_margin, '1');
  assert.equal(params.right_margin, '1');
});

test('an individual margin override takes precedence over margin', () => {
  const params = parseQuery(buildPdfExportOptionsQueryString({ margin: 1, topMargin: 0.25 }));
  assert.equal(params.top_margin, '0.25');
  assert.equal(params.bottom_margin, '1');
});

test('a zero individual margin override is respected, not treated as missing', () => {
  const params = parseQuery(buildPdfExportOptionsQueryString({ margin: 1, topMargin: 0 }));
  assert.equal(params.top_margin, '0');
});

test('extraExportParams are passed through and can override named options', () => {
  const params = parseQuery(buildPdfExportOptionsQueryString({ extraExportParams: { pagenum: 'CENTER', size: 'letter' } }));
  assert.equal(params.pagenum, 'CENTER');
  assert.equal(params.size, 'letter');
});
