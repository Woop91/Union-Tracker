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
  // Coverage thresholds set to realistic baselines — goal is to increase over time
  coverageThreshold: {
    global: {
      lines: 50,
      branches: 40,
      functions: 50,
      statements: 50
    }
  },
  watchPathIgnorePatterns: ['dist/', 'coverage/', 'node_modules/']
};
