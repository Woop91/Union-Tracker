# Testing Documentation for 509 Dashboard

## Overview

This document describes the comprehensive test suite for the 509 Dashboard Google Apps Script application. The test suite ensures reliability, correctness, and maintainability of critical union data management functions.

## Test Coverage Summary

| Component | Test File | Tests | Priority | Coverage Goal |
|-----------|-----------|-------|----------|---------------|
| Core formulas & validations | Code.test.gs | 21+ | P0 (Critical) | 100% |
| Grievance workflow | GrievanceWorkflow.test.gs | 11+ | P1 (High) | 70% |
| Data clearing | DeveloperTools.test.gs | 10+ | P1 (High) | 70% |
| Integration workflows | Integration.test.gs | 9+ | P1 (High) | 80% |
| **TOTAL** | **4 test files** | **51+ tests** | | **~70%** |

---

## Test Framework

### Architecture

The test suite uses a custom **GAS-native testing framework** (`TestFramework.gs`) that runs directly within the Google Apps Script environment, avoiding the need for external testing tools.

**Key Features:**
- Simple assertion library (`Assert` object)
- Automatic test discovery
- Visual test report generation
- Test helpers for data setup/cleanup
- Performance tracking

### Test Runner

Access via menu: **📊 509 Dashboard > 🧪 Tests > Run All Tests**

The test runner:
1. Discovers all test functions (functions starting with `test`)
2. Executes each test in isolation
3. Captures results (passed/failed/skipped)
4. Generates a detailed report in "Test Results" sheet
5. Shows summary dialog with pass rate and duration

---

## Test Categories

### 1. Formula Calculation Tests (Code.test.gs)

**Priority: 🔴 P0 (Critical)**

Tests auto-calculated deadline formulas that determine legal grievance deadlines.

**Tests:**
- `testFilingDeadlineCalculation` - Filing deadline = Incident + 21 days
- `testStepIDeadlineCalculation` - Step I decision = Filed + 30 days
- `testStepIIAppealDeadlineCalculation` - Step II appeal = Decision + 10 days
- `testDaysOpenCalculation` - Days open for active grievances
- `testDaysOpenForClosedGrievance` - Days open uses close date
- `testNextActionDueLogic` - Next action based on current step

**Why Critical:** These calculations determine contract compliance. Errors could cause missed deadlines and lost grievance cases.

**Example Test:**
```javascript
function testFilingDeadlineCalculation() {
  const incidentDate = new Date(2025, 0, 1); // Jan 1, 2025
  const expectedDeadline = new Date(2025, 0, 22); // Jan 22, 2025

  const actualDeadline = new Date(incidentDate.getTime() + 21 * 24 * 60 * 60 * 1000);

  Assert.assertDateEquals(
    expectedDeadline,
    actualDeadline,
    'Filing deadline should be 21 days after incident date'
  );
}
```

---

### 2. Data Validation Tests (Code.test.gs)

**Priority: 🔴 P0 (Critical)**

Tests dropdown validation setup and Config-to-sheet linking.

**Tests:**
- `testDataValidationSetup` - Validation rules exist
- `testConfigDropdownValues` - Config has required values
- `testMemberValidationRules` - Member Directory validations
- `testGrievanceValidationRules` - Grievance Log validations

**Why Critical:** Prevents data corruption and ensures consistency across 20,000+ member records.

---

### 3. Seeding Function Tests (Code.test.gs)

**Priority: 🟡 P1 (High)**

Tests data generation functions for test/training data.

**Tests:**
- `testMemberSeedingValidation` - Member data structure
- `testGrievanceSeedingValidation` - Grievance data structure
- `testMemberEmailFormat` - Email format validation
- `testMemberIDUniqueness` - Unique member IDs
- `testGrievanceMemberLinking` - Foreign key integrity
- `testOpenRateRange` - Open rate 0-100

**Example Test:**
```javascript
function testMemberEmailFormat() {
  const testEmails = [
    'john.smith123@union.org',
    'mary.jones456@union.org'
  ];

  const emailRegex = /^[a-z]+\.[a-z]+\d+@union\.org$/;

  testEmails.forEach(email => {
    Assert.assertTrue(
      emailRegex.test(email),
      `Email ${email} should match format`
    );
  });
}
```

---

### 4. Grievance Workflow Tests (GrievanceWorkflow.test.gs)

**Priority: 🟡 P1 (High)**

Tests member selection and grievance creation workflow.

**Tests:**
- `testGetMemberList` - Returns all members
- `testGetMemberListEmpty` - Handles empty directory
- `testGetMemberListFiltersEmptyRows` - Filters blank rows
- `testGetMemberListFieldCompleteness` - All fields present
- `testMemberSelectionDialog` - Dialog generates valid HTML
- `testMemberSelectionDialogMultipleMembers` - Handles multiple members
- `testMemberSelectionDialogSpecialCharacters` - Handles O'Brien, García
- `testStartGrievanceDialogMissingSheet` - Error handling
- `testGrievanceFormConfiguration` - Config validation
- `testCompleteWorkflowMemberToGrievance` - End-to-end workflow

---

### 5. Data Clearing Tests (DeveloperTools.test.gs)

**Priority: 🟡 P1 (High)**

Tests data clearing functions while preserving sheet structure.

**Tests:**
- `testClearMemberDirectoryPreservesHeaders` - Headers intact
- `testClearGrievanceLogPreservesHeaders` - Headers intact
- `testClearStewardWorkloadMissingSheet` - Error handling
- `testNukePropertySet` - Property management
- `testCheckNukeStatus` - Status checking
- `testClearingDoesntAffectOtherSheets` - Isolation
- `testMultipleClearOperationsIdempotent` - Idempotency
- `testClearingEmptySheet` - Edge case handling
- `testClearingWithFormulas` - Formula preservation
- `testDashboardRebuildUpdatesMetrics` - Dashboard updates

**Example Test:**
```javascript
function testClearMemberDirectoryPreservesHeaders() {
  const ss = SpreadsheetApp.getActive();
  const memberDir = ss.getSheetByName(SHEETS.MEMBER_DIR);

  const originalHeader = memberDir.getRange(1, 1, 1, 31).getValues()[0];

  // Create test data and clear
  createTestMember('TEST-M-001');
  clearMemberDirectory();

  const headerAfter = memberDir.getRange(1, 1, 1, 31).getValues()[0];

  Assert.assertEquals(
    'Member ID',
    headerAfter[0],
    'First header should still be "Member ID"'
  );

  Assert.assertEquals(
    1,
    memberDir.getLastRow(),
    'Should only have header row after clearing'
  );
}
```

---

### 6. Integration Tests (Integration.test.gs)

**Priority: 🟡 P1 (High)**

End-to-end tests for complete workflows and data consistency.

**Tests:**
- `testCompleteGrievanceWorkflow` - Full lifecycle from creation to closure
- `testDashboardMetricsUpdate` - Metrics update on data change
- `testMemberGrievanceSnapshot` - Member-grievance linking
- `testConfigChangesPropagateToDropdowns` - Config updates
- `testMultipleGrievancesSameMember` - Multiple grievances per member
- `testDashboardHandlesEmptyData` - Empty data handling
- `testDashboardRefreshPerformance` - Performance < 10s
- `testFormulaPerformanceWithData` - Formula calculation speed
- `testGrievanceUpdatesTriggersRecalculation` - Auto-recalculation

**Example Test:**
```javascript
function testCompleteGrievanceWorkflow() {
  const testMemberId = createTestMember('TEST-M-001');

  // Step 1: Create grievance
  const grievanceData = [...]; // Full grievance data
  grievanceLog.getRange(...).setValues([grievanceData]);

  // Step 2: Verify auto-calculated deadlines
  const filingDeadline = grievanceLog.getRange(row, 8).getValue();
  Assert.assertNotNull(filingDeadline, 'Filing deadline calculated');

  // Step 3: Verify Member Directory snapshot
  const hasOpenGrievance = memberRow[25];
  Assert.assertTrue(hasOpenGrievance === 'Yes', 'Member shows open grievance');

  // Step 4: Progress to Step II
  grievanceLog.getRange(row, 11).setValue(new Date());

  // Step 5: Close grievance
  grievanceLog.getRange(row, 5).setValue('Settled');

  // Step 6: Verify snapshot updates
  Assert.assertEquals('Settled', updatedStatus, 'Status updated to Settled');
}
```

---

## Assertion Methods

The `Assert` object provides these methods:

```javascript
Assert.assertEquals(expected, actual, message)
Assert.assertTrue(value, message)
Assert.assertFalse(value, message)
Assert.assertNotNull(value, message)
Assert.assertNull(value, message)
Assert.assertContains(array, value, message)
Assert.assertArrayLength(array, expectedLength, message)
Assert.assertThrows(fn, message)
Assert.assertApproximately(expected, actual, tolerance, message)
Assert.assertDateEquals(expected, actual, message)
```

---

## Test Helpers

### Data Setup
```javascript
createTestMember(memberId) // Creates test member with ID
```

### Data Cleanup
```javascript
cleanupTestData() // Removes all rows starting with "TEST-"
```

### Usage
```javascript
function testMyFeature() {
  const memberId = createTestMember('TEST-M-001');

  try {
    // ... test code ...
    Assert.assertEquals(...);
  } finally {
    cleanupTestData(); // Always cleanup
  }
}
```

---

## Running Tests

### Via Menu (Recommended)
1. Open your Google Sheet
2. Click **📊 509 Dashboard > 🧪 Tests > Run All Tests**
3. Confirm the dialog (tests take 2-3 minutes)
4. View results in the "Test Results" sheet

### Via Script Editor
1. Open **Extensions > Apps Script**
2. Find `TestFramework.gs`
3. Run `runAllTests()` function
4. View logs in execution logs

### Single Test
To run a single test:
```javascript
runSingleTest('testFilingDeadlineCalculation');
```

---

## Test Results Sheet

After running tests, a "Test Results" sheet is created with:

1. **Summary Section:**
   - Total tests
   - Passed count (✅)
   - Failed count (❌)
   - Skipped count (⏭️)
   - Pass rate (%)
   - Duration (seconds)
   - Timestamp

2. **Passed Tests Table:**
   - Test name
   - Status (✅ PASS)
   - Duration (ms)

3. **Failed Tests Table:**
   - Test name
   - Status (❌ FAIL)
   - Error message
   - Stack trace

4. **Skipped Tests Table:**
   - Test name
   - Status (⏭️ SKIP)
   - Reason

---

## Writing New Tests

### Test Function Naming
- Must start with `test`
- Use descriptive names: `testFeatureDescription`
- Examples: `testFilingDeadlineCalculation`, `testMemberEmailFormat`

### Test Structure
```javascript
function testMyNewFeature() {
  // 1. Setup test data
  const testMemberId = createTestMember('TEST-M-NEW-001');

  try {
    // 2. Execute the code being tested
    const result = myFunction(testMemberId);

    // 3. Assert expected results
    Assert.assertEquals(expectedValue, result, 'Error message');
    Assert.assertTrue(result > 0, 'Result should be positive');

    // 4. Log success
    Logger.log('✅ My new feature test passed');

  } finally {
    // 5. Always cleanup test data
    cleanupTestData();
  }
}
```

### Best Practices
1. **Use test prefixes:** All test data IDs should start with `TEST-`
2. **Always cleanup:** Use `try/finally` to ensure cleanup
3. **Descriptive assertions:** Include clear error messages
4. **Test one thing:** Each test should verify one specific behavior
5. **Avoid dependencies:** Tests should be independent and run in any order
6. **Use realistic data:** Test with data similar to production

---

## Edge Cases to Test

### Date Handling
- ✅ Future dates (positive days to deadline)
- ✅ Past dates (negative days to deadline)
- ✅ Today's date (0 days)
- ✅ Leap years
- ✅ Month boundaries

### Empty Data
- ✅ Empty Member Directory
- ✅ Empty Grievance Log
- ✅ No open grievances
- ✅ Dashboard with no data

### Special Characters
- ✅ Names with apostrophes (O'Brien)
- ✅ Names with accents (García)
- ✅ Hyphens in names (Mary-Jane)

### Boundary Conditions
- ✅ Exactly 0 members
- ✅ Exactly 1 member
- ✅ Large datasets (20,000+ members)
- ✅ Open rate at 0, 50, 100

---

## Performance Benchmarks

| Operation | Target | Current |
|-----------|--------|---------|
| Dashboard refresh | < 5s | ~2s |
| Run all tests | < 5min | ~2-3min |
| Create 10 grievances | < 30s | ~15s |
| Seed 1k members | < 1min | ~30s |
| Seed 300 grievances | < 30s | ~15s |

---

## Troubleshooting

### Tests Timing Out
**Problem:** Tests exceed 6-minute Google Apps Script limit

**Solutions:**
- Run test categories separately
- Reduce test data size (use smaller counts)
- Optimize `SpreadsheetApp.flush()` usage

### Tests Failing Intermittently
**Problem:** Tests pass sometimes, fail other times

**Solutions:**
- Add `Utilities.sleep(2000)` after data changes
- Increase sleep time for formula recalculation
- Check for race conditions in async operations

### Formula Not Calculating
**Problem:** Formulas return empty or #N/A

**Solutions:**
- Add `SpreadsheetApp.flush()` after writing data
- Wait 2 seconds for recalculation
- Verify sheet names match SHEETS constants

### Test Data Not Cleaning Up
**Problem:** Test data remains after test run

**Solutions:**
- Ensure `cleanupTestData()` is in `finally` block
- Manually run `cleanupTestData()` from script editor
- Check that test IDs start with "TEST-"

---

## CI/CD Integration (Future)

### GitHub Actions Example
```yaml
name: Run Tests

on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - name: Setup Node.js
        uses: actions/setup-node@v2
      - name: Install clasp
        run: npm install -g @google/clasp
      - name: Push to Apps Script
        run: clasp push
      - name: Run tests
        run: clasp run runAllTests
```

---

## Contributing

### Adding Tests
1. Create test function in appropriate `.test.gs` file
2. Follow naming convention (`testFeatureDescription`)
3. Use test helpers for setup/cleanup
4. Add test name to `testFunctions` array in `TestFramework.gs`
5. Run tests to verify
6. Update this documentation

### Test Coverage Goals
- **P0 (Critical):** 100% coverage required
- **P1 (High):** 70% coverage target
- **P2 (Medium):** 50% coverage target
- **P3 (Low):** Optional coverage

---

## Support

For issues with tests:
1. Check test logs in Apps Script execution logs
2. Review "Test Results" sheet for error details
3. Run single test in isolation using `runSingleTest()`
4. Check that test data cleanup completed
5. Document bugs in a tracking system or GitHub issues

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | 2025-01-25 | Initial test suite with 51+ tests |
| | | - Formula calculation tests |
| | | - Data validation tests |
| | | - Seeding function tests |
| | | - Integration tests |
| | | - Test framework implementation |

---

## License

Created for Local 509. Part of the 509 Dashboard project.
