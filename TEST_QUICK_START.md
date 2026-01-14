# Test Suite Quick Start

## 🚀 Running Tests in 30 Seconds

### Step 1: Open Your Sheet
Open the 509 Dashboard Google Sheet

### Step 2: Run Tests
Click: **📊 509 Dashboard > 🧪 Tests > Run All Tests**

### Step 3: View Results
Results appear automatically in "Test Results" sheet

**That's it!** ✅

---

## 📊 Understanding Test Results

### Summary Section
- **Total Tests:** Number of tests run
- **✅ Passed:** Tests that succeeded
- **❌ Failed:** Tests that need attention
- **⏭️ Skipped:** Tests not implemented yet
- **Pass Rate:** Percentage of passing tests
- **Duration:** How long tests took

### Goal: 100% Pass Rate

---

## 🔧 What Tests Cover

### Critical (Must Pass 100%)
- ✅ **Filing Deadlines** - 21 days from incident
- ✅ **Step I Deadlines** - 30 days from filing
- ✅ **Step II Deadlines** - 10 days from Step I decision
- ✅ **Days Open** - Accurate day counts
- ✅ **Data Validation** - Dropdowns work correctly

### Important (70%+ Pass Rate)
- ✅ **Member Creation** - Test member generation
- ✅ **Grievance Linking** - Members link to grievances
- ✅ **Dashboard Updates** - Metrics update automatically
- ✅ **Data Clearing** - Headers preserved when clearing

### Integration (80%+ Pass Rate)
- ✅ **Complete Workflow** - Grievance lifecycle from start to finish
- ✅ **Snapshot Updates** - Member directory syncs with grievances
- ✅ **Performance** - Operations complete in reasonable time

---

## ❌ If Tests Fail

### Common Issues

**1. Formulas Not Calculating**
- **Fix:** Wait 10 seconds and re-run tests
- **Why:** Formulas need time to recalculate

**2. Test Timeout**
- **Fix:** Run smaller test groups
- **Why:** Google limits script execution to 6 minutes

**3. Test Data Not Cleaned**
- **Fix:** Go to Member Directory, delete rows starting with "TEST-"
- **Why:** Previous test run may have crashed

### Get Help
1. Check error message in Test Results sheet
2. Look at "Stack Trace" column for details
3. Review [TESTING.md](TESTING.md) for troubleshooting

---

## 🎯 Test Status Indicators

| Icon | Meaning | Action |
|------|---------|--------|
| ✅ | Test passed | No action needed |
| ❌ | Test failed | Review error, fix code |
| ⏭️ | Test skipped | Feature not implemented yet |

---

## 🔄 When to Run Tests

### Always Run Tests:
- ✅ After modifying formulas in Code.gs
- ✅ After changing data validation rules
- ✅ After updating grievance workflow
- ✅ Before deploying to production
- ✅ Weekly as part of maintenance

### Optional (But Recommended):
- After adding new Config dropdown values
- After major data imports
- After Google Sheets updates

---

## 📈 Performance Expectations

| Operation | Expected Time |
|-----------|---------------|
| Run all tests | 2-3 minutes |
| Dashboard refresh test | < 10 seconds |
| Formula calculation test | < 30 seconds |
| Integration test | < 1 minute |

---

## 💡 Pro Tips

### Tip 1: Run Tests After Setup
After running `CREATE_509_DASHBOARD()`, run tests to verify everything works.

### Tip 2: Check Test Results Regularly
Keep the "Test Results" sheet visible to track quality over time.

### Tip 3: Clean Test Data
If you see rows with IDs starting with "TEST-", it's safe to delete them.

### Tip 4: Use Single Test Function
To debug a specific test:
```javascript
// In Apps Script Editor
runSingleTest('testFilingDeadlineCalculation');
```

### Tip 5: Test Before Seeding
Run tests on empty sheets before seeding to ensure setup is correct.

---

## 🎓 Learn More

- **Full Documentation:** [TESTING.md](TESTING.md)
- **Main README:** [README.md](README.md)
- **Test Framework Code:** `TestFramework.gs`
- **Test Examples:** `Code.test.gs`, `Integration.test.gs`

---

## ✨ Quick Reference

```
Run Tests:       📊 509 Dashboard > 🧪 Tests > Run All Tests
View Results:    📊 509 Dashboard > 🧪 Tests > View Test Results
Expected Time:   2-3 minutes
Goal:            100% pass rate for critical tests
Test Data:       All IDs start with "TEST-"
Cleanup:         Automatic (via finally blocks)
```

---

**Ready to test?** Click **📊 509 Dashboard > 🧪 Tests > Run All Tests** now! 🚀
