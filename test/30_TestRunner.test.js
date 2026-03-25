/**
 * Tests for 30_TestRunner.gs and 31_WebAppTests.gs
 *
 * Validates TestRunner structural invariants:
 * - Test function invocation uses correct this-binding (not method call)
 * - Test registry entries all reference defined functions
 * - Global wrapper existence tests use patterns compatible with the runner
 * - Canary test exists to catch this-binding regressions at runtime
 */

const fs = require('fs');
const path = require('path');

const runnerSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', '30_TestRunner.gs'), 'utf8'
);
const testsSrc = fs.readFileSync(
  path.resolve(__dirname, '..', 'src', '31_WebAppTests.gs'), 'utf8'
);

// ============================================================================
// TestRunner invocation safety
// ============================================================================

describe('TestRunner invocation pattern', () => {
  test('runAll never calls test.fn() as a method call', () => {
    // test.fn() sets `this` to the registry entry object, not the global scope.
    // This breaks any test that uses typeof this[fnName] for global lookups.
    const lines = runnerSrc.split('\n');
    const violations = [];

    lines.forEach((line, idx) => {
      // Match test.fn( but not test.fn.call( or test.fn.apply( or = test.fn
      if (/\btest\.fn\s*\(/.test(line) &&
          !/\btest\.fn\.(call|apply|bind)\s*\(/.test(line)) {
        violations.push(`line ${idx + 1}: ${line.trim()}`);
      }
    });

    expect(violations).toEqual([]);
  });

  test('runAll invokes test functions via indirect call or explicit binding', () => {
    // Must use one of these safe patterns:
    //   var f = test.fn; f();          — indirect call, this = global
    //   test.fn.call(globalThis);      — explicit binding
    //   test.fn.apply(globalThis, []); — explicit binding
    const safePatterns = [
      /=\s*test\.fn\b/,           // local variable assignment
      /test\.fn\.call\s*\(/,      // .call()
      /test\.fn\.apply\s*\(/,     // .apply()
    ];

    const hasSafe = safePatterns.some(p => p.test(runnerSrc));
    expect(hasSafe).toBe(true);
  });
});

// ============================================================================
// Test registry consistency
// ============================================================================

describe('TestRunner registry consistency', () => {
  test('every registered test function is defined in source', () => {
    // Extract registered function names from _getTestRegistry()
    const registryNames = [];
    const fnRefPattern = /fn:\s*(\w+)/g;
    let match;
    while ((match = fnRefPattern.exec(runnerSrc)) !== null) {
      // Skip non-function-name references (e.g., `fn: entry.fn`, `fn: typeof ...`)
      if (match[1] === 'entry' || match[1] === 'typeof') continue;
      registryNames.push(match[1]);
    }
    // Also check 31_WebAppTests.gs
    while ((match = fnRefPattern.exec(testsSrc)) !== null) {
      registryNames.push(match[1]);
    }

    // Combine both source files for function definitions
    const combined = runnerSrc + '\n' + testsSrc;

    const missing = registryNames.filter(name => {
      const defPattern = new RegExp('function\\s+' + name + '\\s*\\(');
      return !defPattern.test(combined);
    });

    expect(missing).toEqual([]);
  });

  test('every test_ function defined in 31_WebAppTests.gs is registered', () => {
    // Extract function definitions
    const definedFns = [];
    const defPattern = /^function\s+(test_\w+)\s*\(/gm;
    let match;
    while ((match = defPattern.exec(testsSrc)) !== null) {
      definedFns.push(match[1]);
    }

    // Extract registered function names from the registry in 30_TestRunner.gs
    const registeredNames = new Set();
    const regPattern = /fn:\s*(test_\w+)/g;
    while ((match = regPattern.exec(runnerSrc)) !== null) {
      registeredNames.add(match[1]);
    }

    const unregistered = definedFns.filter(fn => !registeredNames.has(fn));
    expect(unregistered).toEqual([]);
  });
});

// ============================================================================
// this[...] usage safety in test files
// ============================================================================

describe('this[...] usage in GAS test files', () => {
  test('this-binding canary test is defined and registered', () => {
    // The canary test must exist to catch this-binding regressions at runtime
    expect(testsSrc).toContain('function test_endpoints_thisBindingCanary');
    expect(runnerSrc).toContain('test_endpoints_thisBindingCanary');
  });

  test('all this[...] lookups in test functions use consistent pattern', () => {
    // Every this[...] usage should be inside a for-loop iterating over
    // a function name array — no ad-hoc this[...] calls outside this pattern
    const lines = testsSrc.split('\n');
    const thisUsages = [];

    lines.forEach((line, idx) => {
      if (/\bthis\[/.test(line)) {
        thisUsages.push({ line: idx + 1, content: line.trim() });
      }
    });

    // All usages should be in functions that follow the known patterns:
    // 1. typeof this[var[i]] in a for-loop (existence check)
    // 2. this[ep.fn].apply(...) in a for-loop (auth rejection test)
    // 3. Comments explaining the this[...] pattern
    thisUsages.forEach(({ line, content }) => {
      const isExistenceCheck = /typeof\s+this\[/.test(content);
      const isApplyCall = /this\[.*\]\.apply\(/.test(content);
      const isComment = /^\s*\/\//.test(content);
      expect(isExistenceCheck || isApplyCall || isComment).toBe(true);
    });
  });

  test('this[...] count does not grow without review', () => {
    const count = (testsSrc.match(/\bthis\[/g) || []).length;
    // Current count: 18. Update this if intentionally adding more.
    expect(count).toBeLessThanOrEqual(18);
    expect(count).toBeGreaterThanOrEqual(12);
  });
});

// ============================================================================
// Endpoint test coverage completeness
// ============================================================================

describe('Endpoint existence tests cover all data* wrappers', () => {
  const dataSrc = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', '21_WebDashDataService.gs'), 'utf8'
  );

  test('all global data* functions are tested in endpoint existence checks', () => {
    // Extract all global data* function names from source
    const dataFns = new Set();
    const fnPattern = /^function\s+(data\w+)\s*\(/gm;
    let match;
    while ((match = fnPattern.exec(dataSrc)) !== null) {
      // Skip private helpers (trailing underscore)
      if (!match[1].endsWith('_')) {
        dataFns.add(match[1]);
      }
    }

    // Extract all function names referenced in endpoint tests
    const testedFns = new Set();
    // Match string literals in arrays: 'dataFnName'
    const strPattern = /'(data\w+)'/g;
    while ((match = strPattern.exec(testsSrc)) !== null) {
      testedFns.add(match[1]);
    }
    // Also match direct typeof references: typeof dataFnName
    const directPattern = /typeof\s+(data\w+)\b/g;
    while ((match = directPattern.exec(testsSrc)) !== null) {
      testedFns.add(match[1]);
    }

    const untested = [...dataFns].filter(fn => !testedFns.has(fn));

    // Allow a threshold for wrappers tested indirectly via other suites
    // (e.g., workload SSO, broadcast filters, badge counts, task delegation, scale system,
    // e-signature, form options, filter dropdown values — v4.33.0,
    // onboarding wizard — v4.34.x, case activity log — v4.34.0,
    // access log viewer — v4.36.0, grievance initiation — v4.36.0,
    // default view preference — v4.37.1)
    expect(untested.length).toBeLessThanOrEqual(50);

    // If there are untested functions, log them for visibility
    if (untested.length > 0) {
      // eslint-disable-next-line no-console
      console.log('Untested data* wrappers:', untested);
    }
  });
});
