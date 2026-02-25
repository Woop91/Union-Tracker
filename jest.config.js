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
  coverageThreshold: {
    global: {
      lines: 100,
      branches: 50
    }
  },
  watchPathIgnorePatterns: ['dist/', 'coverage/', 'node_modules/']
};
