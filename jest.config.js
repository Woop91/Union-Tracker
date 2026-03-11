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
  // TEST-01: Coverage thresholds raised from 50/40 — incrementally moving toward 80% target
  coverageThreshold: {
    global: {
      lines: 70,
      branches: 60,
      functions: 70,
      statements: 70
    }
  },
  watchPathIgnorePatterns: ['dist/', 'coverage/', 'node_modules/']
};
