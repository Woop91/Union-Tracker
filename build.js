/**
 * Build Script for Dashboard
 * Copies individual source files into dist/ for multi-file CLASP deployment.
 * GAS V8 loads files in alphabetical filename order — numbered filenames
 * (00_, 01_, …) guarantee correct load order AND give the GAS editor a
 * navigable file sidebar.
 *
 * Usage:
 *   node build.js           - Build (includes all files)
 *   node build.js --prod    - Production build (excludes DevTools.gs)
 *   node build.js --clean   - Clean dist directory
 *
 * Production builds (--prod or --production):
 *   Excludes development/test files that should not be deployed:
 *   - 07_DevTools.gs (contains test data seeding functions like NUKE_SEEDED_DATA)
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC_DIR = path.join(__dirname, 'src');
const DIST_DIR = path.join(__dirname, 'dist');

// .gs files in load order — alphabetical filename = correct GAS load order
const BUILD_ORDER = [
  '00_Security.gs',
  '00_DataAccess.gs',
  '01_Core.gs',
  '02_DataManagers.gs',
  '03_UIComponents.gs',
  '04a_UIMenus.gs',
  '04b_AccessibilityFeatures.gs',
  '04c_InteractiveDashboard.gs',
  '04d_ExecutiveDashboard.gs',
  '04e_PublicDashboard.gs',
  '05_Integrations.gs',
  '06_Maintenance.gs',
  '07_DevTools.gs',
  '08a_SheetSetup.gs',
  '08b_SearchAndCharts.gs',
  '08c_FormsAndNotifications.gs',
  '08d_AuditAndFormulas.gs',
  '08e_SurveyEngine.gs',
  '09_Dashboards.gs',
  '10a_SheetCreation.gs',
  '10b_SurveyDocSheets.gs',
  '10c_FormHandlers.gs',
  '10d_SyncAndMaintenance.gs',
  '10_Main.gs',
  '11_CommandHub.gs',
  '12_Features.gs',
  '13_MemberSelfService.gs',
  '14_MeetingCheckIn.gs',
  '15_EventBus.gs',
  '16_DashboardEnhancements.gs',
  '17_CorrelationEngine.gs',
  // Web-dashboard SPA modules (load after all DDS modules)
  '19_WebDashAuth.gs',
  '20_WebDashConfigReader.gs',
  '21_WebDashDataService.gs',
  '22_WebDashApp.gs',
  '23_PortalSheets.gs',
  '24_WeeklyQuestions.gs',
  '25_WorkloadService.gs',
  '26_QAForum.gs',
  '27_TimelineService.gs',
  '28_FailsafeService.gs',
  '29_Migrations.gs',
  '30_TestRunner.gs',
];

// .html files — copied as actual GAS HTML files (required for HtmlService.createTemplateFromFile)
const HTML_FILES = [
  // Web-dashboard SPA templates
  'index.html',
  'styles.html',
  'auth_view.html',
  'steward_view.html',
  'member_view.html',
  'error_view.html',
  'org_chart.html',
  'poms_reference.html',
];

/**
 * Validates all source files for syntax errors BEFORE copying to dist.
 * Catches: JS syntax errors in .gs files, broken <script> blocks in HTML.
 * Aborts build on first error — no broken code reaches dist/.
 */
function validate(gsFiles, htmlFiles) {
  console.log('Validating source files...\n');
  let errors = 0;

  // 1. Check .gs files
  for (const file of gsFiles) {
    const filePath = path.join(SRC_DIR, file);
    if (!fs.existsSync(filePath)) continue;
    const code = fs.readFileSync(filePath, 'utf8');
    try {
      new vm.Script(code, { filename: file });
    } catch (e) {
      const line = (e.stack || '').match(/:(\d+)/)?.[1] || '?';
      console.error(`  ❌ SYNTAX ERROR in ${file} (line ${line}): ${e.message}`);
      errors++;
    }
  }

  // 2. Check <script> blocks in HTML files
  for (const file of htmlFiles) {
    const filePath = path.join(SRC_DIR, file);
    if (!fs.existsSync(filePath)) continue;
    const content = fs.readFileSync(filePath, 'utf8');
    const scriptRegex = /<script[^>]*>([\s\S]*?)<\/script>/gi;
    let match;
    let blockIndex = 0;
    while ((match = scriptRegex.exec(content)) !== null) {
      const js = match[1];
      if (js.trim().length < 10 || js.includes('<?')) continue; // skip empty/template blocks
      try {
        new vm.Script(js, { filename: `${file}:script[${blockIndex}]` });
      } catch (e) {
        const line = (e.stack || '').match(/:(\d+)/)?.[1] || '?';
        console.error(`  ❌ SYNTAX ERROR in ${file} <script> block ${blockIndex} (line ${line}): ${e.message}`);
        errors++;
      }
      blockIndex++;
    }
  }

  if (errors > 0) {
    console.error(`\n❌ Validation failed: ${errors} syntax error(s). Build aborted.`);
    console.error('   Fix the errors above, then run the build again.\n');
    process.exit(1);
  }

  console.log('  ✓ All files validated — no syntax errors.\n');
}

function build(fileList) {
  const startTime = Date.now();
  console.log('Building dashboard (multi-file mode)...\n');

  // Ensure dist directory exists
  if (!fs.existsSync(DIST_DIR)) {
    fs.mkdirSync(DIST_DIR, { recursive: true });
    console.log('Created dist/ directory');
  }

  let totalLines = 0;
  let copiedGs = 0;
  let copiedHtml = 0;

  // Copy each .gs file to dist/
  for (const file of fileList) {
    const src = path.join(SRC_DIR, file);
    const dest = path.join(DIST_DIR, file);

    if (!fs.existsSync(src)) {
      console.error(`ERROR: Source file not found: ${file}`);
      process.exit(1);
    }

    fs.copyFileSync(src, dest);
    const lineCount = fs.readFileSync(src, 'utf8').split('\n').length;
    totalLines += lineCount;
    copiedGs++;
    console.log(`  Copied: ${file} (${lineCount} lines)`);
  }

  // Copy each .html file to dist/
  for (const file of HTML_FILES) {
    const src = path.join(SRC_DIR, file);
    const dest = path.join(DIST_DIR, file);

    if (!fs.existsSync(src)) {
      console.warn(`WARNING: HTML file not found: ${file} — skipping`);
      continue;
    }

    fs.copyFileSync(src, dest);
    const lineCount = fs.readFileSync(src, 'utf8').split('\n').length;
    totalLines += lineCount;
    copiedHtml++;
    console.log(`  Copied: ${file} (${lineCount} lines)`);
  }

  const elapsed = Date.now() - startTime;
  console.log(`\nBuild complete! (${elapsed}ms)`);
  console.log(`  .gs files:  ${copiedGs}`);
  console.log(`  .html files: ${copiedHtml}`);
  console.log(`  Total lines: ${totalLines}`);
  console.log(`  Output dir:  ${DIST_DIR}`);
}

function clean() {
  console.log('Cleaning dist directory...');
  if (fs.existsSync(DIST_DIR)) {
    // Remove only .gs and .html files; keep appsscript.json
    const files = fs.readdirSync(DIST_DIR);
    let removed = 0;
    for (const f of files) {
      if (f.endsWith('.gs') || f.endsWith('.html')) {
        fs.unlinkSync(path.join(DIST_DIR, f));
        removed++;
      }
    }
    console.log(`Removed ${removed} files from dist/\n`);
  } else {
    console.log('Dist directory does not exist.\n');
  }
}

// Parse command line arguments
const args = process.argv.slice(2);
const shouldClean = args.includes('--clean');
const isProd = args.includes('--prod') || args.includes('--production');

// Files to exclude in production builds
const PROD_EXCLUDE = ['07_DevTools.gs'];

if (shouldClean) {
  clean();
} else {
  const fileList = isProd
    ? BUILD_ORDER.filter(f => !PROD_EXCLUDE.includes(f))
    : BUILD_ORDER;

  if (isProd) {
    console.log('Production build: Excluding DevTools...\n');
  }

  // Auto-clean before build to prevent orphaned files from persisting
  validate(fileList, HTML_FILES);
  clean();
  build(fileList);
}
