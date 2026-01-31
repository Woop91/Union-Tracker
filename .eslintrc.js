/**
 * ESLint Configuration for Google Apps Script
 *
 * This configuration is tailored for the 509 Dashboard project,
 * which uses Google Apps Script (GAS) with its specific globals
 * and ES5-compatible JavaScript.
 */

module.exports = {
  env: {
    browser: false,
    es6: false,
    'googleappsscript/googleappsscript': true
  },
  extends: [
    'eslint:recommended'
  ],
  plugins: [
    'googleappsscript'
  ],
  parserOptions: {
    ecmaVersion: 5
  },
  globals: {
    // Google Apps Script globals
    'SpreadsheetApp': 'readonly',
    'DriveApp': 'readonly',
    'GmailApp': 'readonly',
    'CalendarApp': 'readonly',
    'DocumentApp': 'readonly',
    'FormApp': 'readonly',
    'SlidesApp': 'readonly',
    'Logger': 'readonly',
    'Utilities': 'readonly',
    'UrlFetchApp': 'readonly',
    'HtmlService': 'readonly',
    'PropertiesService': 'readonly',
    'CacheService': 'readonly',
    'LockService': 'readonly',
    'ScriptApp': 'readonly',
    'Session': 'readonly',
    'Browser': 'readonly',
    'ContentService': 'readonly',
    'MailApp': 'readonly',
    'Maps': 'readonly',
    'Jdbc': 'readonly',
    'XmlService': 'readonly',
    'console': 'readonly',

    // Project-specific globals (defined in 01_Constants.gs)
    'SHEETS': 'readonly',
    'COLORS': 'readonly',
    'GRIEVANCE_COLS': 'readonly',
    'MEMBER_COLS': 'readonly',
    'DEADLINE_COLS': 'readonly',
    'STATUS_OPTIONS': 'readonly',
    'ISSUE_CATEGORIES': 'readonly',
    'DEPARTMENTS': 'readonly',
    'STEWARD_OPTIONS': 'readonly',
    'DASHBOARD_VERSION': 'readonly',
    'CACHE_CONFIG': 'readonly',
    'CACHE_KEYS': 'readonly',
    'UNDO_CONFIG': 'readonly'
  },
  rules: {
    // Error prevention
    'no-undef': 'error',
    'no-unused-vars': ['warn', {
      'vars': 'all',
      'args': 'none',
      'varsIgnorePattern': '^(test_|run|show|create|setup|handle|on|get|set|add|remove|update|delete|validate|format|parse|build|init|load|save|clear|reset|toggle|apply|check|is|has|can|should|do|make|find|search|filter|sort|map|reduce|process|execute|perform|trigger|fire|emit|dispatch|notify|log|debug|info|warn|error|assert|expect|describe|it|before|after|beforeEach|afterEach)'
    }],
    'no-redeclare': 'error',
    'no-dupe-keys': 'error',
    'no-duplicate-case': 'error',
    'no-empty': 'warn',
    'no-extra-semi': 'warn',
    'no-unreachable': 'error',
    'valid-typeof': 'error',

    // Best practices
    'eqeqeq': ['warn', 'smart'],
    'no-eval': 'error',
    'no-implied-eval': 'error',
    'no-new-func': 'error',
    'no-with': 'error',
    'no-throw-literal': 'warn',
    'no-return-assign': 'warn',
    'no-self-compare': 'error',
    'no-sequences': 'warn',
    'no-unmodified-loop-condition': 'warn',
    'no-useless-concat': 'warn',
    'no-useless-escape': 'warn',

    // Variables
    'no-shadow-restricted-names': 'error',
    'no-use-before-define': ['error', { 'functions': false, 'classes': true }],

    // Style (warnings only - not blocking)
    'semi': ['warn', 'always'],
    'quotes': ['off'],
    'indent': ['off'],
    'comma-dangle': ['off'],
    'no-trailing-spaces': 'off',
    'eol-last': 'off',
    'max-len': 'off',

    // Disabled rules for GAS compatibility
    'strict': 'off',
    'no-var': 'off',  // GAS uses var
    'prefer-const': 'off',
    'prefer-arrow-callback': 'off',
    'object-shorthand': 'off',
    'prefer-template': 'off'
  },
  overrides: [
    {
      // Test files can have more relaxed rules
      files: ['**/15_TestFramework.gs', '**/test_*.gs', '**/*_test.gs'],
      rules: {
        'no-unused-vars': 'off'
      }
    }
  ]
};
