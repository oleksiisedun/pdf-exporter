'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadFromSrc } = require('./helpers/load-src');

const { formatDateDdMmYyyyHhMm, buildTimestampedFileName } = loadFromSrc('Main.js', ['formatDateDdMmYyyyHhMm', 'buildTimestampedFileName']);

test('formats a date as dd.MM.yyyy HH:mm in the given time zone', () => {
  assert.equal(formatDateDdMmYyyyHhMm(new Date('2026-07-28T15:51:00Z'), 'Europe/Kyiv'), '28.07.2026 18:51');
});

test('zero-pads single-digit day, month, hour and minute', () => {
  assert.equal(formatDateDdMmYyyyHhMm(new Date('2026-01-05T03:07:00Z'), 'UTC'), '05.01.2026 03:07');
});

test('formats midnight as 00:mm, not 24:mm', () => {
  assert.equal(formatDateDdMmYyyyHhMm(new Date('2026-01-01T00:05:00Z'), 'UTC'), '01.01.2026 00:05');
});

test('buildTimestampedFileName appends the formatted date to the base name', () => {
  assert.equal(buildTimestampedFileName(new Date('2026-07-28T15:51:00Z'), 'Data Export', 'Europe/Kyiv'), 'Data Export 28.07.2026 18:51');
});
