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
  // Real GAS coverage baseline re-measured and ratcheted 2026-07-24.
  // Decimal floors preserve measurable gains without claiming untested coverage.
  coverageThreshold: {
    global: {
      lines: 33.4,
      branches: 25.3,
      functions: 39.9,
      statements: 32.8
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
