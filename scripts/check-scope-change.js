#!/usr/bin/env node
/**
 * check-scope-change.js
 * Detects OAuth scope additions in appsscript.json vs git HEAD~1.
 * Run as pre-push hook or npm run check:scopes.
 *
 * Rationale (v4.25.8 bug class):
 *   Adding a scope to appsscript.json does NOT auto-authorize the deployed
 *   web app. If a new scope is added and the admin doesn't create a new
 *   deployment + re-authorize, GAS services using that scope will throw
 *   auth errors at runtime — caught only in the catch block, returned as
 *   {success:false}. Silent to users; very hard to debug.
 *
 * Exit codes:
 *   0 = no scope additions detected (or no git history to compare against)
 *   1 = scope additions detected — re-auth reminder printed, CI fails
 */

'use strict';
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const MANIFEST = path.join(ROOT, 'appsscript.json');

function getScopes(json) {
  try {
    return JSON.parse(json).oauthScopes || [];
  } catch (_) {
    return [];
  }
}

// Read current manifest
if (!fs.existsSync(MANIFEST)) {
  console.log('[scope-check] appsscript.json not found — skipping');
  process.exit(0);
}
const currentScopes = getScopes(fs.readFileSync(MANIFEST, 'utf8'));

// Try to read the previous version from git
let previousScopes = [];
try {
  const prev = execSync('git show HEAD:appsscript.json', { cwd: ROOT, stdio: ['pipe', 'pipe', 'pipe'] }).toString();
  previousScopes = getScopes(prev);
} catch (_) {
  // No previous commit (initial commit) or not a git repo — skip
  console.log('[scope-check] No previous commit to diff against — skipping scope comparison');
  process.exit(0);
}

const added = currentScopes.filter(s => !previousScopes.includes(s));
const removed = previousScopes.filter(s => !currentScopes.includes(s));

if (removed.length > 0) {
  console.log('[scope-check] ⚠️  Scopes REMOVED (verify nothing breaks):');
  removed.forEach(s => console.log('   - ' + s));
  console.log('');
}

if (added.length === 0) {
  console.log('[scope-check] ✅ No OAuth scopes added — no re-authorization required');
  process.exit(0);
}

// ADDED scopes — must warn and block
console.error('');
console.error('╔══════════════════════════════════════════════════════════════════╗');
console.error('║  ⚠️  OAUTH SCOPE CHANGE DETECTED — ACTION REQUIRED              ║');
console.error('╚══════════════════════════════════════════════════════════════════╝');
console.error('');
console.error('Scopes added to appsscript.json:');
added.forEach(s => console.error('   + ' + s));
console.error('');
console.error('WHY THIS MATTERS:');
console.error('  Google Apps Script web app deployments cache authorized scopes');
console.error('  at deploy time. Adding a scope to appsscript.json does NOT');
console.error('  auto-authorize the live deployment. The live app will throw');
console.error('  auth errors for the new service — caught silently in catch blocks.');
console.error('  This caused the v4.25.8 bug (email login broken for all users).');
console.error('');
console.error('REQUIRED AFTER CLASP PUSH:');
console.error('  1. Apps Script editor → Deploy → Manage Deployments');
console.error('  2. Create a NEW deployment version (not just edit description)');
console.error('  3. Click through the authorization prompt → authorize all scopes');
console.error('  4. Update any saved deployment URLs if the URL changed');
console.error('');
console.error('VERIFY: After re-authorizing, run the emailsend suite in TestRunner');
console.error('        OR run testAuthEmailSend() from the Apps Script editor.');
console.error('');
console.error('To bypass this check (if you know what you are doing):');
console.error('  git push --no-verify');
console.error('');
process.exit(1);
