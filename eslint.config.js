/**
 * ESLint Flat Configuration for Google Apps Script
 *
 * This configuration is compatible with ESLint 9.x flat config format.
 * It's set to be lenient for the existing codebase while catching critical issues.
 *
 * @see https://eslint.org/docs/latest/use/configure/migration-guide
 */

module.exports = [
  {
    // Apply to all .gs files
    files: ['src/**/*.gs'],

    languageOptions: {
      ecmaVersion: 2020,
      sourceType: 'script',
      globals: {
        // Google Apps Script globals
        SpreadsheetApp: 'readonly',
        DriveApp: 'readonly',
        GmailApp: 'readonly',
        CalendarApp: 'readonly',
        DocumentApp: 'readonly',
        FormApp: 'readonly',
        SlidesApp: 'readonly',
        Logger: 'readonly',
        Utilities: 'readonly',
        UrlFetchApp: 'readonly',
        HtmlService: 'readonly',
        PropertiesService: 'readonly',
        CacheService: 'readonly',
        LockService: 'readonly',
        ScriptApp: 'readonly',
        Session: 'readonly',
        Browser: 'readonly',
        ContentService: 'readonly',
        MailApp: 'readonly',
        Maps: 'readonly',
        Jdbc: 'readonly',
        XmlService: 'readonly',
        console: 'readonly',

        // Project-specific globals (constants defined in 01_Core.gs)
        SHEETS: 'readonly',
        COLORS: 'readonly',
        ICONS: 'readonly',
        MEMBER_COLS: 'readonly',
        MEMBER_COLUMNS: 'readonly',
        GRIEVANCE_COLS: 'readonly',
        GRIEVANCE_COLUMNS: 'readonly',
        CONFIG_COLS: 'readonly',
        COMMAND_CONFIG: 'readonly',
        ERROR_LEVEL: 'readonly',
        ERROR_CONFIG: 'readonly',
        API_VERSION: 'readonly',
        TIME_CONSTANTS: 'readonly',
        ACCESS_CONTROL: 'readonly',

        // Data Access Layer
        DataAccess: 'readonly',

        // Security functions (00_Security.gs)
        escapeHtml: 'readonly',
        sanitizeForHtml: 'readonly',
        sanitizeObjectForHtml: 'readonly',
        escapeForFormula: 'readonly',
        safeSheetNameForFormula: 'readonly',
        buildSafeQuery: 'readonly',
        checkWebAppAuthorization: 'readonly',
        validateWebAppRequest: 'readonly',
        getAccessDeniedPage: 'readonly',
        maskEmail: 'readonly',
        maskPhone: 'readonly',
        maskName: 'readonly',
        maskMemberForLog: 'readonly',
        maskGrievanceForLog: 'readonly',
        secureLog: 'readonly',
        isValidSafeString: 'readonly',
        isValidMemberId: 'readonly',
        isValidGrievanceId: 'readonly',
        escapeHtmlContent: 'readonly',
        safeSendEmail_: 'readonly',

        // Time utilities
        calculateDeadline: 'readonly',
        daysBetween: 'readonly',
        getDeadlineUrgency: 'readonly',

        // Error handling (01_Core.gs)
        handleError: 'readonly',
        withErrorHandling: 'readonly',
        successResponse: 'readonly',
        errorResponse: 'readonly',
        isTruthyValue: 'readonly',
        PerformanceTimer: 'readonly',
        sanitizeHtml: 'readonly',
        sanitizeForQuery: 'readonly',
        sanitizeEmail: 'readonly',
        sanitizePhone: 'readonly',
        sanitizeSheetName: 'readonly'
      }
    },

    rules: {
      // ==========================================
      // ENABLED RULES (Important for code quality)
      // ==========================================

      // Catch critical errors
      'no-dupe-keys': 'error',
      'no-duplicate-case': 'error',
      'no-empty-pattern': 'error',
      'no-func-assign': 'error',
      'no-import-assign': 'error',
      'no-self-assign': 'error',
      'no-self-compare': 'error',
      'no-template-curly-in-string': 'warn',
      'no-unreachable': 'warn',
      'no-unsafe-finally': 'error',
      'use-isnan': 'error',
      'valid-typeof': 'warn',

      // Security-related
      'no-eval': 'error',
      'no-implied-eval': 'error',
      'no-new-func': 'error',
      'no-with': 'error',

      // ==========================================
      // DISABLED RULES (For legacy compatibility)
      // ==========================================

      // These are disabled to allow existing code to pass
      // Enable gradually as code is cleaned up
      'no-undef': 'off',
      'no-unused-vars': ['warn', { args: 'none', varsIgnorePattern: '^_', caughtErrorsIgnorePattern: '^_', vars: 'local' }],
      'no-redeclare': 'off',
      'no-empty': 'off',
      'eqeqeq': 'warn',
      'semi': 'off',
      'quotes': 'off',
      'indent': 'off',
      'comma-dangle': 'off',
      'no-trailing-spaces': 'off',
      'eol-last': 'off',
      'max-len': 'off',
      'no-extra-semi': 'off',
      'strict': 'off',
      'no-var': 'off',
      'prefer-const': 'off',
      'no-useless-escape': 'off',
      'no-inner-declarations': 'off',
      'no-case-declarations': 'off',
      'no-prototype-builtins': 'off',
      'getter-return': 'warn',
      'no-constant-condition': 'off',
      'no-control-regex': 'off',
      'no-debugger': 'error',
      'no-dupe-args': 'error',
      'no-dupe-else-if': 'off',
      'no-empty-character-class': 'off',
      'no-ex-assign': 'off',
      'no-extra-boolean-cast': 'off',
      'no-invalid-regexp': 'off',
      'no-irregular-whitespace': 'off',
      'no-loss-of-precision': 'off',
      'no-misleading-character-class': 'off',
      'no-obj-calls': 'off',
      'no-regex-spaces': 'off',
      'no-setter-return': 'off',
      'no-sparse-arrays': 'off',
      'no-unexpected-multiline': 'off',
      'no-unsafe-negation': 'off',
      'no-unsafe-optional-chaining': 'off',
      'require-yield': 'off'
    }
  },

  // Ignore patterns
  {
    ignores: [
      'dist/**',
      'node_modules/**',
      'build.js',
      '*.config.js'
    ]
  }
];
