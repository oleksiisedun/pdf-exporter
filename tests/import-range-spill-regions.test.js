'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadFromSrc, toPlain } = require('./helpers/load-src');

const { computeImportRangeSpillRegions } = loadFromSrc('ImportRangeFlattener.js', ['computeImportRangeSpillRegions']);

test('a single-cell import (no spill) reports a 1x1 region', () => {
  const formulas = [['=IMPORTRANGE("id","A1")']];
  const values = [['x']];
  const regions = computeImportRangeSpillRegions(formulas, values, [{ row: 0, col: 0 }]);
  assert.deepEqual(toPlain(regions), [{ row: 0, col: 0, numRows: 1, numCols: 1 }]);
});

test('detects a multi-row, multi-column spill', () => {
  const formulas = [
    ['=IMPORTRANGE("id","A1:B2")', '', ''],
    ['', '', ''],
    ['', '', ''],
  ];
  const values = [
    ['a', 'b', ''],
    ['c', 'd', ''],
    ['', '', ''],
  ];
  const regions = computeImportRangeSpillRegions(formulas, values, [{ row: 0, col: 0 }]);
  assert.deepEqual(toPlain(regions), [{ row: 0, col: 0, numRows: 2, numCols: 2 }]);
});

test('known limitation: under-detects past a blank gap inside the spill', () => {
  // Spill is really 3 columns wide, but a blank cell at col 1 stops expansion early.
  const formulas = [['=IMPORTRANGE("id","A1:C1")', '', '']];
  const values = [['a', '', 'c']];
  const regions = computeImportRangeSpillRegions(formulas, values, [{ row: 0, col: 0 }]);
  assert.deepEqual(toPlain(regions), [{ row: 0, col: 0, numRows: 1, numCols: 1 }]);
});

test('known limitation: over-detects into adjacent unrelated non-blank data with no gap', () => {
  const formulas = [['=IMPORTRANGE("id","A1")', '']];
  const values = [['a', 'unrelated']];
  const regions = computeImportRangeSpillRegions(formulas, values, [{ row: 0, col: 0 }]);
  assert.deepEqual(toPlain(regions), [{ row: 0, col: 0, numRows: 1, numCols: 2 }]);
});

test('stops expanding at a cell that has its own formula', () => {
  const formulas = [['=IMPORTRANGE("id","A1")', '=SUM(1,2)']];
  const values = [['a', 3]];
  const regions = computeImportRangeSpillRegions(formulas, values, [{ row: 0, col: 0 }]);
  assert.deepEqual(toPlain(regions), [{ row: 0, col: 0, numRows: 1, numCols: 1 }]);
});

test('handles multiple independent anchors', () => {
  const formulas = [
    ['=IMPORTRANGE("id","A1:A2")', ''],
    ['', ''],
  ];
  const values = [
    ['a', ''],
    ['b', ''],
  ];
  const regions = computeImportRangeSpillRegions(formulas, values, [{ row: 0, col: 0 }]);
  assert.deepEqual(toPlain(regions), [{ row: 0, col: 0, numRows: 2, numCols: 1 }]);
});
