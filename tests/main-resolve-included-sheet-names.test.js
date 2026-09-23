'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadFromSrc } = require('./helpers/load-src');

const { resolveIncludedSheetNames } = loadFromSrc('Main.js', ['resolveIncludedSheetNames']);

const ALL_SHEETS = ['Summary', 'Data', 'Internal Notes'];

test('includes every sheet when neither option is given', () => {
  assert.deepEqual(resolveIncludedSheetNames(ALL_SHEETS, undefined, undefined), ALL_SHEETS);
});

test('includeSheets filters to the given names, preserving source order', () => {
  assert.deepEqual(resolveIncludedSheetNames(ALL_SHEETS, ['Data', 'Summary'], undefined), ['Summary', 'Data']);
});

test('excludeSheets removes the given names, preserving source order', () => {
  assert.deepEqual(resolveIncludedSheetNames(ALL_SHEETS, undefined, ['Internal Notes']), ['Summary', 'Data']);
});

test('throws when both includeSheets and excludeSheets are given', () => {
  assert.throws(() => resolveIncludedSheetNames(ALL_SHEETS, ['Data'], ['Summary']), /includeSheets or excludeSheets, not both/);
});

test('throws when includeSheets names an unknown sheet', () => {
  assert.throws(() => resolveIncludedSheetNames(ALL_SHEETS, ['Nope'], undefined), /unknown sheet name/);
});

test('throws when excludeSheets names an unknown sheet', () => {
  assert.throws(() => resolveIncludedSheetNames(ALL_SHEETS, undefined, ['Nope']), /unknown sheet name/);
});

test('throws when excluding every sheet leaves nothing included', () => {
  assert.throws(() => resolveIncludedSheetNames(ALL_SHEETS, undefined, ALL_SHEETS), /empty/);
});
