'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadFromSrc } = require('./helpers/load-src');

const { formulaCallsImportRange } = loadFromSrc('ImportRangeFlattener.js', ['formulaCallsImportRange']);

test('detects a direct IMPORTRANGE call', () => {
  assert.equal(formulaCallsImportRange('=IMPORTRANGE("id","Sheet1!A1:B2")'), true);
});

test('detects IMPORTRANGE nested inside another function', () => {
  assert.equal(formulaCallsImportRange('=SUM(IMPORTRANGE("id","Sheet1!A1:B2"))'), true);
});

test('is case-insensitive', () => {
  assert.equal(formulaCallsImportRange('=importrange("id","Sheet1!A1:B2")'), true);
});

test('ignores IMPORTRANGE mentioned only inside a string literal', () => {
  assert.equal(formulaCallsImportRange('="Use IMPORTRANGE() here"'), false);
});

test('returns false for an empty or missing formula', () => {
  assert.equal(formulaCallsImportRange(''), false);
  assert.equal(formulaCallsImportRange(undefined), false);
});

test('returns false for an unrelated formula', () => {
  assert.equal(formulaCallsImportRange('=SUM(A1:A10)'), false);
});
