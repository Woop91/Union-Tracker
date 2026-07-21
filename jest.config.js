module.exports = {
  testMatch: ['**/test/**/*.test.js'],
  testEnvironment: 'node',
  verbose: true,
  collectCoverage: true,
  coverageDirectory: 'coverage',
  collectCoverageFrom: [
    'src/**/*.gs',
    '!src/07_DevTools.gs'
  ],
  coverageReporters: ['text', 'text-summary', 'html', 'lcov'],
  // Real GAS coverage baseline measured 2026-07-21. The prior 70/60 settings
  // silently evaluated 0/0 because eval-loaded .gs files were not instrumented.
  // These floors now fail on an actual regression; raise them as coverage grows.
  coverageThreshold: {
    global: {
      lines: 33,
      branches: 25,
      functions: 39,
      statements: 32
    }
  },
  // Unified error summary reporter — aggregates all failures into one readable block
  reporters: [
    'default',
    './test/webapp-error-reporter.js'
  ],
  watchPathIgnorePatterns: ['dist/', 'coverage/', 'node_modules/'],
  modulePathIgnorePatterns: ['<rootDir>/.worktrees/'],
  testPathIgnorePatterns: ['<rootDir>/.worktrees/']
};
