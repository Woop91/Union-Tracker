/**
 * Helper to load .gs source files into the test environment.
 *
 * Google Apps Script files define functions and var constants as globals.
 * We preprocess the code to ensure top-level `var` and `function` declarations
 * become properties on the `global` object, then eval in global scope.
 *
 * KNOWN LIMITATIONS of regex-based source loading:
 * 1. var/const/let rewriting matches line-starts in multi-line strings (indented code is safe)
 * 2. Brace-counting for function extraction doesn't handle braces in string literals or regex
 * 3. Destructured declarations (const { a, b } = ...) are not rewritten to global
 * 4. Destructured declarations (const { a, b } = ...) are NOT rewritten to global scope.
 *    Verify no .gs source files use top-level destructuring before adding any.
 * These are acceptable for this codebase where top-level declarations are not indented
 * and functions don't contain unmatched braces in strings.
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
 * Google Apps Script hoists function declarations so they are available
 * anywhere in the file, regardless of definition order. To simulate this
 * in Node.js, we extract top-level function definitions, evaluate them
 * first, then evaluate the full code.
 *
 * @param {string} filename - Filename relative to src/, e.g. '00_Security.gs'
 */
function loadSource(filename) {
  const filePath = path.resolve(__dirname, '..', 'src', filename);
  let code = fs.readFileSync(filePath, 'utf8');

  // --- Phase 1: Hoist function declarations ---
  // In GAS, `function foo()` is hoisted to the top of the script scope.
  // Extract top-level function declarations and evaluate them first so
  // they are available when var/const/let assignments run.
  const fnPattern = /^function\s+(\w+)\s*\(/gm;
  let match;
  const functionNames = [];
  while ((match = fnPattern.exec(code)) !== null) {
    functionNames.push(match[1]);
  }

  if (functionNames.length > 0) {
    // Build a hoisted-only version: just the function declarations rewritten as global assignments
    let hoisted = code.replace(/^function\s+(\w+)\s*\(/gm, 'global.$1 = function $1(');
    // Remove everything that isn't a global.FUNCNAME assignment (var, const, let, other statements)
    // by only keeping lines that are part of function expressions
    // Simpler approach: eval the full rewritten code but skip non-function top-level statements
    // Actually simplest: just eval function assignments extracted via a second pass

    // Extract complete function bodies for hoisting
    const fnBodies = [];
    for (const name of functionNames) {
      const fnRegex = new RegExp('^function\\s+' + name + '\\s*\\(', 'm');
      const fnMatch = fnRegex.exec(code);
      if (fnMatch) {
        // Find the matching closing brace by counting braces
        let braceCount = 0;
        let started = false;
        let endIdx = fnMatch.index;
        for (let i = fnMatch.index; i < code.length; i++) {
          if (code[i] === '{') { braceCount++; started = true; }
          if (code[i] === '}') { braceCount--; }
          if (started && braceCount === 0) { endIdx = i + 1; break; }
        }
        const fnCode = code.substring(fnMatch.index, endIdx);
        fnBodies.push('global.' + name + ' = ' + fnCode.replace(/^function/, 'function'));
      }
    }

    if (fnBodies.length > 0) {
      // eslint-disable-next-line no-eval
      (0, eval)(fnBodies.join('\n'));
    }
  }

  // --- Phase 2: Evaluate the full file ---
  // Rewrite top-level var declarations to global assignments
  // Match lines starting with `var IDENTIFIER =` (possibly with leading whitespace at indent 0)
  // NOTE: These regexes use ^...gm which matches start-of-line, not start-of-string.
  // This means nested function/var declarations inside IIFEs or blocks that happen to
  // start at column 0 will ALSO be rewritten. For the current codebase this works
  // because IIFE bodies are indented, but it could break if formatting changes.
  // A proper AST-based approach (e.g., babel transform) would be more robust.
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
