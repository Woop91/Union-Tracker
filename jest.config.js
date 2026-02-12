module.exports = {
  testMatch: ['**/test/**/*.test.js'],
  testEnvironment: 'node',
  verbose: true,
  collectCoverage: true,
  coverageDirectory: 'coverage',
  coverageReporters: ['text', 'text-summary'],
  coverageThreshold: {
    global: {
      lines: 30
    }
  }
};
