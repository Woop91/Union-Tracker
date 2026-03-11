/**
 * Simple Trigger Lint Tests — Prevent GAS Authorization Violations
 *
 * GAS simple triggers (onOpen, onEdit, onChange, onFormSubmit) run with
 * restricted authorization. Calling ScriptApp.getProjectTriggers(),
 * ScriptApp.newTrigger(), MailApp, or GmailApp inside them silently fails —
 * no error is shown, no stack trace is produced.
 *
 * These tests parse the source of each simple trigger function and flag any
 * forbidden API calls before they reach production.
 *
 * ST1: onOpen body must not call ScriptApp methods
 * ST2: onEdit body must not call ScriptApp methods
 * ST3: onOpen must not use finally with trigger cleanup pattern
 * ST4: onOpen must not spawn time-based triggers
 */

const fs   = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '..', 'src');

// Extracts the body of a top-level function by name from source text.
// Returns null if function not found.
function extractFunctionBody(src, fnName) {
  const start = src.indexOf('function ' + fnName + '(');
  if (start === -1) return null;
  let depth = 0;
  let inBody = false;
  let bodyStart = -1;
  for (let i = start; i < src.length; i++) {
    if (src[i] === '{') {
      depth++;
      if (!inBody) { inBody = true; bodyStart = i; }
    } else if (src[i] === '}') {
      depth--;
      if (depth === 0 && inBody) return src.slice(bodyStart, i + 1);
    }
  }
  return null;
}

// Strip single-line and block comments from source
function stripComments(src) {
  return src
    .replace(/\/\/[^\n]*/g, '')
    .replace(/\/\*[\s\S]*?\*\//g, '');
}

const MAIN_SRC = fs.readFileSync(path.join(SRC_DIR, '10_Main.gs'), 'utf8');
const MAIN_STRIPPED = stripComments(MAIN_SRC);

const FORBIDDEN_IN_SIMPLE = [
  'ScriptApp.getProjectTriggers',
  'ScriptApp.newTrigger',
  'ScriptApp.deleteTrigger',
  'MailApp.sendEmail',
  'GmailApp.',
];

describe('ST1 — onOpen must not call forbidden ScriptApp/Mail APIs', () => {
  const body = extractFunctionBody(MAIN_STRIPPED, 'onOpen');

  test('onOpen function is found in 10_Main.gs', () => {
    expect(body).not.toBeNull();
  });

  FORBIDDEN_IN_SIMPLE.forEach(api => {
    test(`onOpen must not call ${api}`, () => {
      expect(body).not.toContain(api);
    });
  });
});

describe('ST2 — onEdit must not call forbidden ScriptApp/Mail APIs', () => {
  const MAINTENANCE_SRC = fs.readFileSync(path.join(SRC_DIR, '06_Maintenance.gs'), 'utf8');
  const MAINTENANCE_STRIPPED = stripComments(MAINTENANCE_SRC);
  const body = extractFunctionBody(MAINTENANCE_STRIPPED, 'onEdit');

  // onEdit may not exist in this file — skip gracefully
  if (!body) {
    test.skip('onEdit not found — skipping', () => {});
    return;
  }

  FORBIDDEN_IN_SIMPLE.forEach(api => {
    test(`onEdit must not call ${api}`, () => {
      expect(body).not.toContain(api);
    });
  });
});

describe('ST3 — onOpen must not use finally with trigger cleanup pattern', () => {
  const body = extractFunctionBody(MAIN_STRIPPED, 'onOpen');

  test('onOpen must not have a finally block', () => {
    expect(body).not.toMatch(/\bfinally\b/);
  });
});

describe('ST4 — onOpen must not spawn time-based triggers', () => {
  const body = extractFunctionBody(MAIN_STRIPPED, 'onOpen');

  test('onOpen must not call .timeBased()', () => {
    expect(body).not.toContain('.timeBased()');
  });

  test('onOpen must not call .after(', () => {
    expect(body).not.toContain('.after(');
  });
});
