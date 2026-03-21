#!/usr/bin/env node
/**
 * Automated inline style → CSS utility class replacement.
 * Safely replaces exact-match style objects in el() calls.
 *
 * Usage: node scripts/replace-inline-styles.js [--dry-run]
 */
'use strict';

const fs = require('fs');
const path = require('path');

const DRY_RUN = process.argv.includes('--dry-run');

// Map of exact style object text → CSS class name
// Keys must match the EXACT text between style: { and }
const REPLACEMENTS = [
  // ── Typography combos (exact match) ──
  { style: "fontSize: '11px', color: 'var(--muted)'", cls: 'text-xs-muted' },
  { style: "fontSize: '12px', color: 'var(--muted)'", cls: 'text-sm-muted' },
  { style: "fontSize: '13px', color: 'var(--text)'", cls: 'text-base' },
  { style: "fontSize: '13px', color: 'var(--muted)'", cls: 'text-base-muted' },
  { style: "fontSize: '12px', color: 'var(--text)'", cls: 'text-sm-text' },
  { style: "fontSize: '12px'", cls: 'text-sm' },

  // ── Typography + spacing combos ──
  { style: "fontSize: '11px', color: 'var(--muted)', marginTop: '4px'", cls: 'text-xs-muted-mt-4' },
  { style: "fontSize: '11px', color: 'var(--muted)', marginBottom: '4px'", cls: 'text-xs-muted-mb-4' },
  { style: "fontSize: '11px', color: 'var(--muted)', marginTop: '2px'", cls: 'text-xs-muted', extraStyle: "marginTop: '2px'" },
  { style: "fontSize: '12px', color: 'var(--muted)', marginBottom: '12px'", cls: 'text-sm-muted-mb-12' },

  // ── Layout ──
  { style: "flex: '1'", cls: 'flex-1' },
  { style: "display: 'flex', gap: '8px', flexWrap: 'wrap'", cls: 'flex-wrap-gap' },
  { style: "display: 'flex', alignItems: 'center', gap: '8px'", cls: 'flex-row' },

  // ── Empty states ──
  { style: "textAlign: 'center', padding: '30px 0', color: 'var(--danger)'", cls: 'empty-state-danger' },
  { style: "textAlign: 'center', padding: '30px 0', color: 'var(--muted)'", cls: 'empty-state' },
  { style: "color: 'var(--muted)', padding: '20px 0', textAlign: 'center'", cls: 'empty-state', extraStyle: "padding: '20px 0'" },
  { style: "textAlign: 'center', padding: '40px 0'", cls: 'empty-state-lg' },

  // ── Spacing ──
  { style: "padding: '16px', marginBottom: '10px'", cls: 'p-16-mb-10' },
  { style: "padding: '16px'", cls: 'p-16' },
  { style: "marginBottom: '14px'", cls: 'mb-14' },
  { style: "marginBottom: '8px'", cls: 'mb-8' },
  { style: "marginBottom: '16px'", cls: 'mb-16' },
  { style: "marginBottom: '4px'", cls: 'mb-4' },

  // ── Common combos ──
  { style: "fontSize: '32px', marginBottom: '10px'", cls: 'emoji-lg' },
  { style: "fontSize: '13px', minHeight: '20px'", cls: 'status-msg' },
  { style: "display: 'none'", cls: '', extraStyle: "display: 'none'" }, // skip — display:none is dynamic

  // ── Pass 2: Display font combos ──
  { style: "fontFamily: 'var(--fontDisplay)', fontSize: '13px', fontWeight: '600', color: 'var(--text)'", cls: 'display-section' },
  { style: "fontFamily: 'var(--fontDisplay)', fontSize: '14px', fontWeight: '600', color: 'var(--text)'", cls: 'display-title-sm' },
  { style: "fontSize: '10px', color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px'", cls: 'text-label' },

  // ── Pass 2: Empty state variants ──
  { style: "color: 'var(--muted)', padding: '20px 0', textAlign: 'center'", cls: 'empty-state-sm' },
  { style: "textAlign: 'center', padding: '30px 0', color: 'var(--muted)', fontSize: '13px'", cls: 'empty-state-13' },
  { style: "textAlign: 'center', padding: '40px 0', color: 'var(--muted)', fontSize: '13px'", cls: 'empty-state-lg' },

  // ── Pass 2: Typography + spacing combos ──
  { style: "fontSize: '11px', color: 'var(--muted)', marginTop: '2px'", cls: 'text-xs-muted-mt-2' },
  { style: "fontSize: '11px', color: 'var(--muted)', marginTop: '8px'", cls: 'text-xs-muted-mt-8' },
  { style: "fontSize: '11px', color: 'var(--muted)', marginBottom: '6px'", cls: 'text-xs-muted-mb-6' },
  { style: "fontSize: '11px', color: 'var(--muted)', marginBottom: '8px'", cls: 'text-xs-muted-mb-8' },
  { style: "fontSize: '13px', fontWeight: '600', color: 'var(--text)'", cls: 'display-section' },

  // ── Pass 2: Flex combos ──
  { style: "display: 'flex', justifyContent: 'space-between', marginBottom: '4px'", cls: 'flex-between' },
  { style: "display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px'", cls: 'flex-between-center' },
  { style: "display: 'flex', gap: '6px'", cls: 'flex-gap-6' },
  { style: "display: 'flex', gap: '8px'", cls: 'flex-gap-8' },

  // ── Pass 2: Spacing ──
  { style: "marginBottom: '6px'", cls: 'mb-6' },
  { style: "marginBottom: '12px'", cls: 'mb-12' },
  { style: "marginBottom: '10px'", cls: 'mb-10' },
  { style: "padding: '14px'", cls: 'p-14' },
  { style: "padding: '16px', marginBottom: '12px'", cls: 'p-16-mb-12' },
  { style: "color: 'var(--accent)'", cls: '', extraStyle: true }, // skip — too generic, may conflict
];

// Filter out skipped entries
const activeReplacements = REPLACEMENTS.filter(r => r.cls);

const FILES = [
  'src/member_view.html',
  'src/steward_view.html',
  'src/shared_components.html',
];

let totalReplacements = 0;

FILES.forEach(filePath => {
  if (!fs.existsSync(filePath)) return;
  let src = fs.readFileSync(filePath, 'utf8');
  let fileReplacements = 0;

  activeReplacements.forEach(({ style, cls, extraStyle }) => {
    // Escape special regex chars in the style string
    const escaped = style.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    // Allow flexible whitespace
    const flexStyle = escaped.replace(/, /g, ',\\s*');

    if (extraStyle) {
      // Partial match — replace matched props with className, keep extra as style
      // Pattern: style: { <matched props> }
      // This case is complex; skip automated replacement for partial matches
      return;
    }

    // Case 1: el('tag', { style: { ... } }, ...) — no className
    // Replace style: { ... } with className: 'cls'
    const reNoClass = new RegExp(
      `(\\bstyle:\\s*\\{\\s*)${flexStyle}(\\s*\\})`,
      'g'
    );

    // Count matches first
    const matches1 = (src.match(reNoClass) || []).length;

    if (matches1 > 0) {
      // Check if each match is in a context WITHOUT an existing className
      // We need to be more careful: only replace when className is not present
      // Strategy: match the full { ... style: { ... } ... } attrs object

      // Simple approach: replace style: { <exact> } with className: '<cls>'
      // This works because if className exists, it's a separate key in the object
      src = src.replace(reNoClass, (match) => {
        return `className: '${cls}'`;
      });

      // Now handle cases where the element ALSO had a className before style
      // Pattern: className: 'existing', className: 'cls' → className: 'existing cls'
      // This is a post-processing step
      const doubleCls = /className:\s*'([^']+)',\s*className:\s*'([^']+)'/g;
      src = src.replace(doubleCls, (_, cls1, cls2) => `className: '${cls1} ${cls2}'`);

      // Also handle: className: 'cls', className: 'existing' (reversed order)
      // This shouldn't happen since style: comes after className: typically

      fileReplacements += matches1;
    }
  });

  if (fileReplacements > 0) {
    if (!DRY_RUN) {
      fs.writeFileSync(filePath, src);
    }
    console.log(`${filePath}: ${fileReplacements} replacements${DRY_RUN ? ' (dry run)' : ''}`);
    totalReplacements += fileReplacements;
  }
});

console.log(`\nTotal: ${totalReplacements} inline styles replaced with CSS utility classes`);
if (DRY_RUN) console.log('(Dry run — no files modified)');
