/**
 * Helper to load .gs source files into the test environment.
 *
 * Google Apps Script files define functions and var constants as globals.
 * We preprocess the code to ensure top-level `var` and `function` declarations
 * become properties on the `global` object, then eval in global scope.
 */

const fs = require('fs');
const path = require('path');

/**
 * Loads a .gs source file into the global scope.
 *
 * Rewrites top-level `var X =` to `global.X =` and
 * `function X(` to `global.X = function X(` so that
 * all declarations become globally accessible in Jest's sandbox.
 *
 * @param {string} filename - Filename relative to src/, e.g. '00_Security.gs'
 */
function loadSource(filename) {
  const filePath = path.resolve(__dirname, '..', 'src', filename);
  let code = fs.readFileSync(filePath, 'utf8');

  // Rewrite top-level var declarations to global assignments
  // Match lines starting with `var IDENTIFIER =` (possibly with leading whitespace at indent 0)
  code = code.replace(/^var\s+(\w+)\s*=/gm, 'global.$1 =');

  // Rewrite top-level function declarations to global assignments
  // Match lines starting with `function IDENTIFIER(` at indent 0
  code = code.replace(/^function\s+(\w+)\s*\(/gm, 'global.$1 = function $1(');

  // Rewrite const/let at top level (ES6 in V8 runtime)
  code = code.replace(/^(const|let)\s+(\w+)\s*=/gm, 'global.$2 =');

  // Use indirect eval to run in global scope
  // eslint-disable-next-line no-eval
  (0, eval)(code);
}

/**
 * Loads multiple source files in order (simulates GAS load order).
 * @param {string[]} filenames - Array of filenames to load
 */
function loadSources(filenames) {
  filenames.forEach(loadSource);
}

module.exports = { loadSource, loadSources };
