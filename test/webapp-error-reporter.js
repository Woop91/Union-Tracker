/**
 * webapp-error-reporter.js
 * Custom Jest reporter that aggregates all errors from webapp test suites
 * and prints a unified summary at the end of the test run.
 *
 * Usage:
 *   jest --reporters=default --reporters=./test/webapp-error-reporter.js
 *   or via jest.config.js reporters array
 *
 * Output: A consolidated error log showing every failure across all test files,
 * grouped by file, with the assertion message, expected/received, and location.
 */

class WebappErrorReporter {
  constructor(globalConfig, reporterOptions) {
    this._globalConfig = globalConfig;
    this._options = reporterOptions || {};
  }

  onRunComplete(_testContexts, results) {
    var failedSuites = [];
    var totalErrors = 0;

    for (var i = 0; i < results.testResults.length; i++) {
      var suite = results.testResults[i];
      var suiteFails = [];

      for (var j = 0; j < suite.testResults.length; j++) {
        var test = suite.testResults[j];
        if (test.status === 'failed') {
          totalErrors++;
          suiteFails.push({
            name: test.ancestorTitles.concat(test.title).join(' > '),
            errors: test.failureMessages || [],
            duration: test.duration
          });
        }
      }

      if (suiteFails.length > 0) {
        // Extract just the filename from the full path
        var filePath = suite.testFilePath || '';
        var fileName = filePath.split('/').pop().split('\\').pop();

        failedSuites.push({
          file: fileName,
          path: filePath,
          failures: suiteFails,
          runtime: suite.perfStats
            ? suite.perfStats.end - suite.perfStats.start
            : 0
        });
      }
    }

    if (totalErrors === 0) {
      console.log('\n' + _line('='));
      console.log('  ALL WEBAPP TESTS PASSED (' + results.numPassedTests + ' tests across ' + results.numPassedTestSuites + ' suites)');
      console.log(_line('=') + '\n');
      return;
    }

    // Print unified error report
    console.log('\n' + _line('='));
    console.log('  WEBAPP TEST ERROR SUMMARY');
    console.log('  ' + totalErrors + ' failure(s) across ' + failedSuites.length + ' file(s)');
    console.log(_line('='));

    for (var s = 0; s < failedSuites.length; s++) {
      var entry = failedSuites[s];
      console.log('\n' + _line('-'));
      console.log('  FILE: ' + entry.file);
      console.log(_line('-'));

      for (var f = 0; f < entry.failures.length; f++) {
        var fail = entry.failures[f];
        console.log('\n  [FAIL] ' + fail.name);
        if (fail.duration != null) {
          console.log('         Duration: ' + fail.duration + 'ms');
        }

        for (var e = 0; e < fail.errors.length; e++) {
          // Strip ANSI color codes for cleaner log output
          var cleanMsg = _stripAnsi(fail.errors[e]);
          // Indent each line of the error message
          var lines = cleanMsg.split('\n');
          for (var l = 0; l < lines.length; l++) {
            console.log('         ' + lines[l]);
          }
        }
      }
    }

    console.log('\n' + _line('='));
    console.log('  TOTALS: ' + totalErrors + ' failed, '
      + results.numPassedTests + ' passed, '
      + results.numTotalTests + ' total');
    console.log(_line('=') + '\n');
  }
}

function _line(ch) {
  var s = '';
  for (var i = 0; i < 72; i++) s += ch;
  return s;
}

function _stripAnsi(str) {
  // eslint-disable-next-line no-control-regex
  return str.replace(/\x1b\[[0-9;]*m/g, '');
}

module.exports = WebappErrorReporter;
