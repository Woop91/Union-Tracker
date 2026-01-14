# 🚨 Developer Tools - Exit Demo Mode

## Overview

The **Developer Tools** feature (in `DeveloperTools.gs`) allows you to seed demo data for testing and remove it when transitioning to production mode. This file is designed to be **deleted before production** - once you delete the file from the Apps Script editor, all demo functions are permanently gone.

> **⚠️ IMPORTANT**: The `DeveloperTools.gs` file should be **DELETED FROM THE SCRIPT EDITOR** before going live. This ensures stewards cannot accidentally trigger a data wipe.

---

## ⚠️ What Does "Nuke Seed Data" Do?

When you execute the **Nuke Seed Data** function, the system will:

### Data Removal
1. **Remove ALL Members**: Delete all test members from Member Directory
2. **Remove ALL Grievances**: Delete all test grievances from Grievance Log
3. **Clear Survey Responses**: Clears all survey data from Member Satisfaction (preserves sheet structure)
4. **Delete Feedback Sheet**: Completely removes the Feedback & Development sheet
5. **Delete Menu Checklist**: Completely removes the Menu Checklist sheet
6. **Clear Steward Workload**: Remove all test steward assignments
7. **Clear Config Demo Data**: Remove demo entries from Config tab (if any exist):
   - Job Titles (Column A)
   - Office Locations (Column B)
   - Units (Column C)
   - Supervisors (Column F)
   - Managers (Column G)
   - Stewards (Column H)
   - Grievance Coordinators (Column O)
   - Home Towns (Column AF)
   - Office Addresses (Column AN)

   > **NOTE (v3.11+):** These fields are now LEFT EMPTY during CREATE_509_DASHBOARD. Users populate them with their own data. If no user data was added, there's nothing to clear.

### Demo Mode Disabling
8. **Disable Demo Menu**: Sets a flag (`DEMO_MODE_DISABLED`) to hide the "🎭 Demo Data" submenu on next refresh
9. **Clear Tracking**: Removes the tracked seeded ID lists from Script Properties

### Preserved Items
10. **Preserve Organization Info**: Keep your real organization settings:
   - Organization Name, Local Number, Main Address, Phone
   - Union Parent, State/Region, Website
   - Main Fax, Toll Free numbers
   - All deadline and contract reference columns
11. **Preserve Structure**: Keep all headers, formulas, and sheet structure intact
12. **Preserve Manually Entered Data**: Only rows with seeded ID patterns (M/G + 4 letters + 3 digits) are deleted

> **🔴 IMPORTANT**: The nuke operation only deletes **seeded data rows** (matching the pattern `MJOSM123` or `GJOSM456`), clears **survey responses** from Member Satisfaction, and completely removes the **Feedback & Development** and **Menu Checklist** sheets. Manually entered data with different ID formats is preserved. The Demo menu is hidden but the seed functions remain in the code.

---

## 🎯 When Should You Nuke Seed Data?

**Nuke the seed data when:**
- You're ready to start using the dashboard with real member data
- You've completed all testing and training with the demo data
- You understand how the system works and are ready for production
- You want to start fresh without test data cluttering your sheets

**DON'T nuke if:**
- You're still learning how to use the system
- You want to keep testing features
- You haven't backed up the spreadsheet yet
- You're using this for training or demonstration purposes

---

## 📋 Pre-Nuke Checklist

Before nuking seed data, make sure you:

- [ ] **Understand the system**: You know how to add members and grievances
- [ ] **Have a backup**: Copy the spreadsheet if you want to keep a demo version
- [ ] **Review Config settings**: Ensure dropdown values match your needs
- [ ] **Prepare real data**: Have member and grievance data ready to import
- [ ] **Train your team**: Everyone knows how to use the system
- [ ] **Document processes**: You have procedures for data entry and workflow

---

## 🚀 How to Nuke Seed Data

### Step 1: Access the Nuke Function

**Menu**: `🔧 Admin > 🎭 Demo Data > ☢️ NUKE SEEDED DATA`

*Note: This menu item only appears if seed data hasn't been nuked yet.*

### Step 2: Confirm the Action

You'll see **TWO confirmation dialogs**:

**First Confirmation:**
```
⚠️ WARNING: Remove All Seeded Data & Functions

This will PERMANENTLY remove:
• All test data from Member Directory, Grievance Log, Steward Workload
• Config Tab Demo Entries (Job Titles, Locations, etc.)
• ALL seed functions from the script code
• ALL seed menu items
• THIS NUKE FUNCTION ITSELF (complete self-deletion)

After this operation, there will be NO trace of seed OR nuke functionality.

This action CANNOT be undone!

Are you sure you want to proceed?
```

**Second Confirmation (Final):**
```
🚨 FINAL CONFIRMATION

This is your last chance!

ALL test data will be cleared and the Demo menu will be permanently hidden.

To fully remove demo functionality, delete the DeveloperTools.gs file from the script editor after running this.

Click YES to proceed.
```

### Step 3: Wait for Processing

After confirming:
- You'll see a "⏳ Removing seeded data..." message
- The process takes a few moments
- **Do not close the spreadsheet** while processing

### Step 4: Verify Clean State

After nuking, verify the system is ready:
- **Setup checklist**: Steps to configure your production environment
- **Steward contact info reminder**: Enter your steward details
- **Next actions**: What to do first

---

## 📊 What Happens After Nuking

### Immediate Changes

1. **Empty Sheets** (with headers intact):
   - Member Directory: 0 members
   - Grievance Log: 0 grievances
   - Steward Workload: Empty

2. **Dashboards Reset**:
   - All metrics show zero
   - Charts will be empty
   - No overdue grievances

3. **Menu Changes**:
   - "🎭 Demo Data" submenu hidden from Admin menu
   - Cleaner menu structure
   - Focus on production tools

4. **Menu Changes**:
   - Demo menu is hidden (DEMO_MODE_DISABLED property set)
   - The Demo submenu will not appear in the Admin menu
   - **To fully remove**: Delete `DeveloperTools.gs` from the script editor
   - Once the file is deleted, all seed/nuke functions are permanently gone

### What Remains Intact

✅ All sheet headers and column structure
✅ Config tab dropdown lists
✅ Timeline rules table
✅ Steward contact info section (if already filled)
✅ All analytical tabs and dashboard layouts
✅ Color schemes and formatting
✅ All custom menu items (except seed options)
✅ Triggers and automation

---

## 🎬 Post-Nuke Setup Steps

After nuking, follow these steps to set up for production:

### 1. Configure Steward Contact Info

**Location**: `⚙️ Config > Column U`

Enter:
- **Row 2**: Steward Name
- **Row 3**: Steward Email
- **Row 4**: Steward Phone

Or use: `📋 Grievances > ➕ Start New Grievance` (then configure steward info in Config sheet)

### 2. Review Config Dropdown Lists

**Location**: `⚙️ Config > Columns A-N`

Verify these match your organization:
- Job Titles (Column A)
- Work Locations (Column B)
- Units (Column C)
- Grievance Types (Column G)
- Committee Types (Column L)

Add, remove, or modify as needed.

### 3. Add Real Members

**Location**: `👥 Member Directory`

You can:
- **Manual Entry**: Type directly into the sheet
- **CSV Import**: Use File > Import
- **Copy/Paste**: From another spreadsheet

Required fields:
- Member ID
- First Name
- Last Name
- Is Steward (Yes/No)

### 4. Set Up Triggers

**Menu**: `⚙️ Administrator > 🔧 Setup & Triggers > ⚡ Install Auto-Sync Trigger`

This enables:
- Automatic deadline calculations
- Real-time dashboard updates
- Member snapshot updates

### 5. Customize Custom View Dashboard

**Tab**: Click the `🎯 Custom View` sheet tab, or **Menu**: `📊 509 Dashboard > 📊 Dashboard`

Create custom views for:
- Key metrics you track
- Charts you need most
- Your preferred layout

### 6. Test Grievance Workflow (Optional)

If using the grievance workflow:
1. Set up Google Form (see [Grievance Workflow Guide](GRIEVANCE_WORKFLOW_GUIDE.md))
2. Configure form submission trigger
3. Test with a sample grievance

---

## 🔄 Can I Undo the Nuke?

**No, the data deletion is permanent and irreversible.**

After nuking:
- Seeded data rows are permanently deleted from Member Directory and Grievance Log
- The Demo menu is hidden (based on a Script Property flag)
- Config dropdowns are cleared

**Note**: If you haven't deleted `DeveloperTools.gs` yet, the seed functions still exist. To re-seed:
1. Delete the `DEMO_MODE_DISABLED` Script Property manually (Script Properties → Delete `DEMO_MODE_DISABLED`)
2. Refresh the spreadsheet - the Demo menu will reappear
3. Run seed functions again

**Recovery options:**
- **Re-enable Demo menu**: Script Properties → Delete `DEMO_MODE_DISABLED` (only works if DeveloperTools.gs still exists)
- **Restore from backup**: If you made a copy of the spreadsheet before nuking
- **Fresh deployment**: Deploy a new copy from the original source code
- **Import data**: Add your real data to start fresh (recommended approach)

**Final Cleanup (Recommended):**
After running NUKE and verifying you're ready for production:
1. Open Extensions → Apps Script
2. In the Files panel, right-click on `DeveloperTools.gs`
3. Click "Delete" - this permanently removes all demo functionality

---

## 🆘 Troubleshooting

### Problem: Dashboards Still Show Data

**Solution**:
- Go to `📊 Sheet Manager > 📊 Rebuild Dashboard`
- This recalculates all metrics

### Problem: Demo Menu Still Visible

**Solution**:
- Close and reopen the spreadsheet
- The menu is rebuilt on open
- The 🎭 Demo Data submenu only shows when demo mode is enabled

### Problem: Need to Re-Seed for Training

**Solution**:
Since seed functions are permanently deleted after nuking, you cannot re-seed:
1. Deploy a fresh copy from the original source code
2. Or restore from a backup made BEFORE the nuke
3. The `resetNukeFlag()` function no longer restores seed capabilities

### Problem: Accidentally Nuked Too Soon

**Solution**:
- If you have a backup copy, restore from there
- If no backup, you'll need to import your data manually
- For future: Always make a backup first!

---

## 🔧 Manual File Deletion (Final Cleanup)

The new DeveloperTools.gs approach uses **manual file deletion** instead of automatic code removal. This is simpler and more reliable.

### How to Delete DeveloperTools.gs

After running NUKE and verifying your data is clean:

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. In the left sidebar, find `DeveloperTools.gs`
4. Right-click on the file name
5. Click **"Delete"** or the trash icon
6. Confirm the deletion

### Why Manual Deletion?

The manual deletion approach is better because:
- **No API setup required** - No need to enable Apps Script API or set OAuth scopes
- **Clear user action** - You consciously remove the developer tools
- **Zero residue** - Once deleted, no trace of seed/nuke code remains
- **Simpler recovery** - If you need demo tools again, just re-add the file from source

### What Happens After Deletion?

Once `DeveloperTools.gs` is deleted:
- All `SEED_*` functions no longer exist
- All `NUKE_*` functions no longer exist
- The Demo menu never appears (even if you clear the DEMO_MODE_DISABLED property)
- Stewards cannot accidentally seed or nuke data

---

## 💡 Best Practices

### Before Nuking

1. **Make a backup copy**: File > Make a copy
2. **Document your needs**: List required config changes
3. **Prepare import data**: Have member/grievance CSVs ready
4. **Train your team**: Everyone understands the workflow

### After Nuking

1. **Start small**: Add a few test members first
2. **Verify calculations**: Check that auto-calculations work
3. **Test workflows**: Try the grievance workflow with test data
4. **Gradual rollout**: Add real data in batches
5. **Monitor performance**: Watch how dashboards update

---

## 📚 Related Documentation

- [Main README](README.md) - Complete dashboard documentation
- [Grievance Workflow Guide](GRIEVANCE_WORKFLOW_GUIDE.md) - How to use grievance features
- [Interactive Dashboard Guide](INTERACTIVE_DASHBOARD_GUIDE.md) - Customization options
- [Comfort View Guide](COMFORT_VIEW_GUIDE.md) - Accessibility features

---

## 🎯 Quick Reference

### Menu Location (Before Nuke)
```
🔧 Admin > 🎭 Demo Data > ☢️ NUKE SEEDED DATA
```

### What Gets Deleted
- ❌ All members from Member Directory
- ❌ All grievances from Grievance Log
- ❌ All steward workload data
- ❌ Survey responses from Member Satisfaction (data cleared, sheet preserved)
- ❌ Feedback & Development sheet (entire sheet)
- ❌ Menu Checklist sheet (entire sheet)
- ❌ Config demo data (job titles, locations, units, supervisors, managers, stewards, coordinators, home towns, office addresses)

### What Gets Preserved
- ✅ Headers and structure
- ✅ Organization info (name, local number, address, phone, fax, toll-free, website)
- ✅ Deadline settings and contract references
- ✅ Dashboards and charts
- ✅ All formulas and formatting
- ✅ Menu system

### Post-Nuke Priorities
1. Enter steward contact info
2. Review config settings
3. Add real members
4. Set up triggers
5. Test grievance workflow

---

## 🎉 Welcome to Production!

After nuking seed data, you're ready to use the 509 Dashboard with real member and grievance data.

The system is now configured for:
- **Real member tracking**
- **Actual grievance management**
- **Live deadline monitoring**
- **Authentic reporting and analytics**

Your dashboard is **production-ready**! 🚀

---

## See Also

- **`NUKE_SEEDED_DATA()`** - Smart nuke that removes only seeded data (pattern-matched IDs)
  - Menu: `🔧 Admin > 🎭 Demo Data > ☢️ NUKE SEEDED DATA`
- **`NUKE_CONFIG_DROPDOWNS()`** - Clears only Config dropdown values
  - Menu: `🔧 Admin > 🎭 Demo Data > 🧹 Clear Config Dropdowns Only`

---

**Last Updated**: 2026-01-14
**Version**: 2.0.0

---

## Version 2.0.0 Notes - DeveloperTools.gs Separation

In version 2.0.0, the seed/nuke functionality was moved to a separate file for cleaner production deployment:

- **SeedNuke.gs renamed to DeveloperTools.gs** - Clear naming indicates this is for development only
- **Prominent warning banner** - File header states "DELETE THIS FILE BEFORE PRODUCTION"
- **Conditional Demo menu** - Demo menu only appears if DEMO_MODE_DISABLED property is not set
- **Manual file deletion** - No Apps Script API required; just delete the file from the editor
- **Zero residue after deletion** - Once DeveloperTools.gs is deleted, all demo functions are gone

**Production Deployment Steps:**
1. Run `NUKE_SEEDED_DATA()` to clear test data (hides Demo menu)
2. Delete `DeveloperTools.gs` from the Apps Script editor
3. Stewards cannot accidentally trigger data wipes

---

## Version 1.6.0 Notes

In version 1.6.0, the seed functionality was simplified:

- **SEED_GRIEVANCES was merged into SEED_MEMBERS** - Use `SEED_MEMBERS(count, grievancePercent)` to seed members with optional grievances
- **All seeded grievances are directly linked to members** - No orphaned grievances with missing member info
- **Separate SEED_GRIEVANCES function removed** - Use the merged approach instead

**Example:**
- `SEED_MEMBERS(100)` - Seeds 100 members, ~30 grievances (30% default)
- `SEED_MEMBERS(100, 50)` - Seeds 100 members, ~50 grievances (50%)
- `SEED_MEMBERS(100, 0)` - Seeds 100 members only, no grievances
