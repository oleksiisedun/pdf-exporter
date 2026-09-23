'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SRC_DIR = path.join(__dirname, '..', '..', 'src');

/**
 * Loads a src/*.js file — a plain Apps Script script with no imports/exports,
 * sharing one global scope by design — into a fresh vm context, and returns
 * the requested top-level functions. Functions pulled out this way keep
 * their closure over the file's other top-level declarations (consts, other
 * functions), so cross-references within the same file still resolve.
 * @param {string} fileName - e.g. 'PdfFetch.js'
 * @param {string[]} names - top-level function names to pull out
 * @returns {Record<string, Function>}
 */
function loadFromSrc(fileName, names) {
  const code = fs.readFileSync(path.join(SRC_DIR, fileName), 'utf8');
  const context = vm.createContext({});
  vm.runInContext(code, context, { filename: fileName });

  const result = {};
  for (const name of names) result[name] = context[name];
  return result;
}

/**
 * A `vm` context is a separate JS realm with its own `Object`/`Array`
 * constructors, so a plain object or array returned by a function loaded via
 * loadFromSrc is not `instanceof` this file's `Object` — `assert/strict`'s
 * `deepEqual` (aliased to `deepStrictEqual`) then reports "same structure but
 * are not reference-equal" even for identical plain data. Round-tripping
 * through JSON rebuilds the value using this realm's own built-ins, so
 * assertions on plain JSON-serializable data (objects/arrays of strings,
 * numbers, booleans) can compare normally. Only use this for values that are
 * plain data — it would silently drop functions, undefined, etc.
 * @param {unknown} value
 * @returns {unknown}
 */
function toPlain(value) {
  return JSON.parse(JSON.stringify(value));
}

module.exports = { loadFromSrc, toPlain };
