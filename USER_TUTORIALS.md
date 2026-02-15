# 509 Dashboard - User Tutorials

**Version:** 4.7.0
**Last Updated:** 2026-02-14

Quick, practical tutorials for common tasks in the 509 Dashboard.

---

## Table of Contents

1. [Getting Started](#1-getting-started)
2. [Adding Your First Member](#2-adding-your-first-member)
3. [Filing a New Grievance](#3-filing-a-new-grievance)
4. [Tracking Grievance Deadlines](#4-tracking-grievance-deadlines)
5. [Using the Main Dashboard](#5-using-the-main-dashboard)
6. [Customizing the Interactive Dashboard](#6-customizing-the-interactive-dashboard)
7. [Managing Steward Workload](#7-managing-steward-workload)
8. [Running System Diagnostics](#8-running-system-diagnostics)
9. [Seeding Demo Data (for Training)](#9-seeding-demo-data-for-training)
10. [Exiting Demo Mode](#10-exiting-demo-mode)
11. [Understanding Hidden Sheet Architecture](#11-understanding-hidden-sheet-architecture)
12. [Setting Up a Meeting](#12-setting-up-a-meeting) ⭐ NEW
13. [Managing Meeting Notes & Agenda](#13-managing-meeting-notes--agenda) ⭐ NEW
14. [Creating a Member Drive Folder](#14-creating-a-member-drive-folder) ⭐ NEW

---

## 1. Getting Started

### First-Time Setup (5 minutes)

**Step 1: Open the Dashboard**
1. Open your Google Sheet with the 509 Dashboard
2. Wait for the menus to load (you should see 6 menus)

**Step 2: Verify Setup**
1. Click **⚙️ Administrator** → **System Health** → **🔧 Diagnose Setup**
2. Review the diagnostic report
3. All items should show ✅

**Step 3: Configure Your Organization**
1. Go to the **Config** sheet tab
2. Add your organization's data:
   - Column A: Job Titles
   - Column B: Office Locations
   - Column C: Units
   - Column F: Supervisors
   - Column G: Managers
   - Column H: Stewards

**You're ready to start!**

---

## 2. Adding Your First Member

### Manual Entry (2 minutes per member)

**Step 1: Open Member Directory**
1. Click the **Member Directory** sheet tab

**Step 2: Find the First Empty Row**
1. Scroll down to find an empty row (or row 2 if empty)

**Step 3: Enter Required Fields**
| Column | Field | Example |
|--------|-------|---------|
| A | Member ID | Auto-generated (e.g., MJASM472) or leave blank for auto-fill |
| B | First Name | Jane |
| C | Last Name | Smith |
| D | Job Title | Case Worker (use dropdown) |
| E | Work Location | Boston Office (use dropdown) |
| F | Unit | Unit 8 (use dropdown) |
| I | Email | jane.smith@email.com |
| J | Phone | (555) 123-4567 |
| O | Is Steward | No (use dropdown) |

**Step 4: Optional Fields**
- Column G: Cubicle / workspace ID
- Column H: Office Days (multi-select)
- Column K: Preferred Communication (multi-select)
- Column L: Best Time to Contact (multi-select)
- Columns M-N: Supervisor, Manager
- Columns P-Q: Committees, Assigned Steward
- Columns R-U: Engagement metrics
- Columns V-X: Interest flags

**Columns AB-AD are auto-calculated - don't edit them!**

---

## 3. Filing a New Grievance

### Quick Grievance Entry (3 minutes)

**Step 1: Open Grievance Log**
1. Click the **Grievance Log** sheet tab

**Step 2: Enter Grievance Details**
| Column | Field | Example |
|--------|-------|---------|
| A | Grievance ID | GJASM456 (G + first 2 chars of first/last name + 3 digits) |
| B | Member ID | MJASM123 (must match Member Directory) |
| C | First Name | Jane |
| D | Last Name | Smith |
| E | Status | Open (use dropdown) |
| F | Current Step | Step I (use dropdown) |
| G | Incident Date | 12/01/2025 |
| I | Date Filed | 12/05/2025 |

**Step 3: Add Classification**
| Column | Field | Example |
|--------|-------|---------|
| V | Articles Violated | Art. 23 - Grievance Procedure |
| W | Issue Category | Discipline |
| X | Description | Brief description of the issue |

**Step 4: Verify Auto-Calculations**
After entering, these columns auto-populate:
- Column H: Filing Deadline (Incident + 21 days)
- Column J: Step I Decision Due (Filed + 30 days)
- Column S: Days Open
- Column T: Next Action Due
- Column U: Days to Deadline

---

## 4. Tracking Grievance Deadlines

### Understanding Deadline Colors

| Color | Meaning | Action |
|-------|---------|--------|
| 🟢 Green | 7+ days remaining | On track |
| 🟡 Yellow | 1-7 days remaining | Due soon - prepare |
| 🔴 Red | Overdue | Act immediately |

### Checking Upcoming Deadlines

**Method 1: Dashboard**
1. Go to **Dashboard** sheet
2. Look at "Upcoming Deadlines" section
3. Grievances sorted by urgency

**Method 2: Filter Grievance Log**
1. Go to **Grievance Log**
2. Click Data → Create a filter
3. Filter Column U (Days to Deadline) → Sort ascending
4. Top rows = most urgent

### Updating Grievance Progress

**When Step I Decision is Received:**
1. Enter date in Column K (Step I Decision Rcvd)
2. Column L auto-calculates Step II Appeal Due
3. If appealing, update Column F to "Step II"
4. Enter Column M (Step II Appeal Filed)

---

## 5. Using the Main Dashboard

### Dashboard Sections

**Row 1-2: Header**
- Title and last refresh timestamp

**Rows 3-6: Member Metrics**
| Metric | What It Shows |
|--------|---------------|
| Total Members | Count of all members |
| Active Stewards | Members with Is Steward = Yes |
| Avg Open Rate | Average email engagement |
| YTD Vol. Hours | Total volunteer hours this year |

**Rows 7-10: Grievance Metrics**
| Metric | What It Shows |
|--------|---------------|
| Open Grievances | Status = Open |
| Pending Info | Status = Pending Info |
| Settled (Month) | Resolved this month |
| Avg Days Open | Average time to resolution |

**Rows 11+: Upcoming Deadlines**
- Top 10 most urgent deadlines
- Color-coded by urgency

### Refreshing the Dashboard

1. Click **📊 Sheet Manager** → **🔄 Refresh All Formulas**
2. Wait for "Complete" toast message
3. All metrics now current

---

## 6. Using the Interactive Dashboard

### Accessing the Dashboard

**Option 1: Modal Popup** (⚠️ PROTECTED)
- Click **⚙️ Admin** → **📊 Show Steward Dashboard**
- Opens a tabbed popup with Overview, My Cases, Members, Grievances, Analytics

**Option 2: Sheet Tab**
- Click the **🎯 Interactive Dashboard** sheet tab directly
- Use dropdowns to customize metrics and charts

### Selecting Metrics

**Step 1: Choose Chart 1 Metric**
1. Find "What to show (Chart 1):" dropdown
2. Select a metric (e.g., "Open Grievances by Status")

**Step 2: Choose Chart Type**
1. Find "How to show it (Chart 1):" dropdown
2. Select chart type: Pie, Bar, Line, Column, etc.

**Step 3: Enable Second Chart (Optional)**
1. Set "Show both charts:" to "Yes"
2. Configure Chart 2 using the same process

### Changing Theme

1. Find "Color Scheme:" dropdown
2. Choose from:
   - Professional Blue
   - Union Green
   - Soft Pastels
   - High Contrast
   - Dark Mode
   - Warm Tones

---

## 7. Managing Steward Workload

### Viewing Steward Assignments

1. Click **👨‍⚖️ Steward Workload** sheet tab
2. Review columns:

| Column | Meaning |
|--------|---------|
| Steward Name | Assigned steward |
| Total Cases | All-time cases |
| Active Cases | Currently open |
| Resolved | Cases closed |
| Win Rate | % won or settled |
| Overdue | Past deadline |
| Due This Week | Upcoming deadlines |
| Capacity Status | Green/Yellow/Red |

### Rebalancing Workload

**If a steward is overloaded:**
1. Go to **Grievance Log**
2. Filter by Assigned Steward (Column AB)
3. Update Assigned Steward for some cases
4. Workload sheet updates automatically

---

## 8. Running System Diagnostics

### Quick Health Check

1. Click **⚙️ Administrator** → **System Health** → **🔧 Diagnose Setup**
2. Review the dialog:

**Good Result:**
```
🔧 DIAGNOSTIC REPORT

📊 Sheets Found: 22 / 22 ✅
📋 COLUMN COUNTS:
   Member Directory: 31 columns ✅
   Grievance Log: 34 columns ✅
   Config: 43 columns ✅

🎉 VERDICT: System is healthy!
```

**If Issues Found:**
- Missing sheets: Run CREATE_509_DASHBOARD()
- Wrong column counts: Check for manual edits
- Errors: Review error messages

---

## 9. Seeding Demo Data (for Training)

### Seed All Sample Data

1. Click **🔧 Admin** → **🎭 Demo Data** → **🚀 Seed All Sample Data**
2. Confirm when prompted
3. Wait for seeding to complete

This seeds:
- 1,000 sample members with complete data
- 300 grievances (randomly distributed - some members may have multiple)
- 50 sample survey responses
- Auto-sync trigger for live updates

### Live Wiring

After seeding, Member Directory columns auto-update when you edit Grievance Log:
- **Has Open Grievance?** - Yes/No based on grievance status
- **Grievance Status** - Open/Pending Info/blank
- **Days to Deadline** - Minimum deadline for open grievances

---

## 10. Exiting Demo Mode

### When to Exit Demo Mode

- You've finished training
- You're ready for real member data
- You want a clean production system

### Nuke Process

1. Click **🔧 Admin** → **🎭 Demo Data** → **☢️ NUKE SEEDED DATA**
2. Read the warning carefully
3. Click "Yes" on first confirmation
4. Click "Yes" on second confirmation
5. Wait for completion

### What Gets Removed

- ❌ All test members (pattern-matched IDs only)
- ❌ All test grievances (pattern-matched IDs only)
- ❌ Survey response data
- ❌ Demo menu (hidden after nuke)

### What's Preserved

- ✅ All sheet headers
- ✅ Config dropdown lists
- ✅ Dashboard layouts
- ✅ All formulas
- ✅ Manually entered data with different ID formats

### Final Cleanup (Recommended)

After nuking, delete `07_DevTools.gs` from the Apps Script editor:
1. Go to **Extensions** → **Apps Script**
2. Right-click `07_DevTools.gs` → **Delete**
3. This permanently removes all demo functions

### After Nuking

1. Add your real organization's Config data
2. Enter real members
3. Start tracking real grievances
4. **You're in production mode!**

---

## 11. Understanding Hidden Sheet Architecture

### What Are Hidden Sheets?

The dashboard uses 6 "hidden" calculation sheets (prefixed with "_") to automatically synchronize data between sheets. You don't see these sheets, but they power the auto-updating columns.

### The 6 Hidden Sheets

| Hidden Sheet | What It Does |
|--------------|--------------|
| `_Grievance_Calc` | Calculates grievance data for Member Directory |
| `_Grievance_Formulas` | Self-healing formulas for Grievance Log timeline columns |
| `_Member_Lookup` | Looks up member data for Grievance Log |
| `_Steward_Contact_Calc` | Tracks steward contact data |
| `_Dashboard_Calc` | Dashboard summary metrics (15 key metrics) |
| `_Steward_Performance_Calc` | Per-steward performance scores with weighted formula |

### Auto-Populated Columns

**Member Directory** (these columns update automatically):
| Columns | Data | Source |
|---------|------|--------|
| Q-T | Engagement (Last Virtual Mtg, Last In-Person, Open Rate, Vol Hours) | Meeting Attendance + Volunteer Hours |
| Y-AA | Steward Contact (Date, Who, Notes) | Communications Log |
| AB-AD | Grievance (Has Open?, Status, Deadline) | Grievance Log |
| AF-AH (v3.45) | Grievance Stats (Total Count, Win Rate, Last Date) | Grievance Log |

**Grievance Log** (these columns update automatically):
| Columns | Data | Source |
|---------|------|--------|
| C-D | Member Name (First, Last) | Member Directory |
| X-AA | Member Details (Email, Unit, Location, Steward) | Member Directory |

### How Auto-Sync Works

1. You edit a source sheet (e.g., change a member's email in Member Directory)
2. An automatic trigger fires (runs in the background)
3. The hidden sheet recalculates the values
4. The new values appear in the destination sheet

**⏱️ Updates happen within 2-3 seconds of your edit.**

### Verifying Hidden Sheets Work

**Step 1: Run Verification**
1. Click **Administrator** → **Setup & Triggers** → **🔍 Verify Hidden Sheets**
2. Review the report

**Good Result:**
```
✅ _Grievance_Calc: EXISTS, hidden, formulas present
✅ _Member_Lookup: EXISTS, hidden, formulas present
✅ _Steward_Contact_Calc: EXISTS, hidden, formulas present
✅ _Engagement_Calc: EXISTS, hidden, formulas present
✅ _Steward_Workload_Calc: EXISTS, hidden, formulas present
✅ All 5 triggers installed
```

### Troubleshooting: Columns Not Updating?

**Problem**: Member Directory columns AB-AD aren't updating when I change Grievance Log

**Solution**:
1. Run **Administrator** → **Setup & Triggers** → **🔍 Verify Hidden Sheets**
2. If it shows issues, run `REPAIR_DASHBOARD()` from Apps Script

**Problem**: Grievance Log columns C-D or X-AA are stale

**Solution**:
1. Run `REPAIR_DASHBOARD()` from Apps Script
2. This recreates all hidden sheets and reinstalls triggers

### Setting Up Engagement Tracking

If you want columns Q-T to auto-populate with engagement data:

**Step 1: Create Source Sheets**
1. Click **Administrator** → **Setup & Triggers** → **📅 Setup Engagement Tracking**
2. This creates:
   - 📅 Meeting Attendance sheet
   - 🤝 Volunteer Hours sheet
   - _Engagement_Calc hidden sheet
   - Auto-sync trigger

**Step 2: Enter Data in Source Sheets**

**Meeting Attendance** (enter each meeting):
| Column | Example |
|--------|---------|
| Date | 12/01/2025 |
| Type | Virtual / In-Person / Hybrid |
| Meeting Name | Monthly Chapter Meeting |
| Member ID | MJASM123 |
| Attended | Yes / No |

**Volunteer Hours** (enter each activity):
| Column | Example |
|--------|---------|
| Date | 12/01/2025 |
| Member ID | MJASM123 |
| Activity | Phone Banking |
| Hours | 3 |

**Step 3: Watch Auto-Population**
- After adding entries, Member Directory columns Q-T auto-update
- Q = Last Virtual Meeting date for that member
- R = Last In-Person Meeting date for that member
- T = Total Volunteer Hours for that member

### Repair Functions

If something goes wrong, use these repair functions:

| Function | What It Fixes |
|----------|---------------|
| `REPAIR_DASHBOARD()` | Recreates ALL hidden sheets + ALL triggers (the nuclear option) |
| `VERIFY_HIDDEN_SHEETS()` | Diagnoses what's broken without changing anything |
| `setupGrievanceCalcSheet()` | Repairs only the grievance → member sync |
| `setupMemberLookupSheet()` | Repairs only the member → grievance sync |
| `setupEngagementTracking()` | Creates engagement source sheets + hidden sheet |

**How to run these:**
1. Open **Extensions** → **Apps Script**
2. In the function dropdown, select the function name
3. Click **Run**

---

## 12. Setting Up a Meeting

### Creating a New Meeting (3 minutes)

**Step 1: Open Meeting Setup**
1. Click **Strategic Ops > Meeting Check-In > Setup Meeting**
2. The setup dialog opens with all options

**Step 2: Fill In Meeting Details**
| Field | Example |
|-------|---------|
| Meeting Name | February General Membership Meeting |
| Meeting Date | 02/15/2026 |
| Meeting Time | 6:00 PM |
| Meeting Type | In-Person (dropdown) |
| Duration | 60 minutes |

**Step 3: Configure Steward Notifications**
1. Enter steward emails in "Email Attendance Report To" (comma-separated)
2. Under "Send Agenda Early To (3 days prior)", check specific stewards who need the agenda early
3. Use **Select All** or **Clear All** to quickly toggle all stewards
4. All stewards will receive the agenda at least 1 day before the meeting regardless

**Step 4: Click Create Meeting**
- A Google Calendar event is created
- A **Meeting Notes** Google Doc is created in the `Meeting Notes/` folder
- A **Meeting Agenda** Google Doc is created in the `Meeting Agenda/` folder
- Links are stored in columns N and O of the Meeting Check-In Log

**You're ready! The system handles notifications automatically.**

---

## 13. Managing Meeting Notes & Agenda

### How Document Automation Works

The system automatically sends meeting documents to stewards on a schedule:

| When | What Happens |
|------|--------------|
| **3 days before meeting** | Agenda link emailed to selected stewards (from setup) |
| **1 day before meeting** | Agenda link emailed to ALL stewards |
| **1 day before meeting** | Notes link emailed to notification stewards |
| **1 day after meeting** | Notes set to view-only and visible to members in dashboard |

### Editing Meeting Notes

1. Open the Meeting Check-In Log sheet
2. Find your meeting row
3. Click the URL in column N (Notes Doc URL) to open the Google Doc
4. Edit the document directly - it's a full Google Doc

### Editing Meeting Agenda

1. Open the Meeting Check-In Log sheet
2. Find your meeting row
3. Click the URL in column O (Agenda Doc URL) to open the Google Doc
4. Add agenda items before the meeting

### Viewing Meeting Notes as a Member

1. Open the **Member Dashboard** (Strategic Ops > Command Center > Member Dashboard)
2. Click the **Meeting Notes** tab
3. Browse meetings chronologically or use the search bar
4. Click **View Notes** to open the read-only Google Doc

**Note:** Members can only view notes after the meeting date. Notes are read-only for members.

---

## 14. Creating a Member Drive Folder

### Quick Action Method (30 seconds)

**Step 1: Open Quick Actions**
1. Select a member row in the Member Directory
2. Click **Strategic Ops > Cases > Member Quick Actions**

**Step 2: Click "Create Member Folder"**
1. Click the **Create Member Folder** button (📁 icon)
2. A Google Drive folder is created for the member
3. If the member already has a grievance folder, that folder is reused

**Result:**
- A Drive folder named after the member is created or found
- You'll see a confirmation with the folder URL
- Use this folder to store member-related documents

---

## Quick Reference Card

### Essential Menu Paths

| Task | Menu Path |
|------|-----------|
| Refresh dashboard | 👤 Dashboard → 🔄 Refresh All |
| View operations monitor | 👤 Dashboard → 📊 Dashboards → 🎯 Unified Operations Monitor |
| Check system health | ⚙️ Administrator → System Health → 🔧 Diagnose Setup |
| Seed test data | 🔧 Admin → 🎭 Demo Data → 🚀 Seed All Sample Data |
| Exit demo mode | 🔧 Admin → 🎭 Demo Data → ☢️ NUKE SEEDED DATA |

### Column Reference

**Member Directory (39 columns A-AM)**
- A: Member ID
- B-C: Name
- D-F: Work info (Job Title, Location, Unit)
- G: Cubicle (hidden)
- H: Office Days
- I-L: Contact (Email, Phone, Preferred Comm, Best Time)
- M-N: Supervisor, Manager
- O: Is Steward
- P-Q: Committees, Assigned Steward
- AB-AD: Auto-calculated (don't edit)

**Grievance Log (34 columns A-AH)**
- A: Grievance ID
- B-D: Member info
- E-F: Status/Step
- G: Incident Date
- I: Date Filed
- H, J, L, N, P, S-U: Auto-calculated

### Deadline Rules (Article 23A)

| Deadline | Days | From |
|----------|------|------|
| Filing | 21 | Incident Date |
| Step I Decision | 30 | Date Filed |
| Step II Appeal | 10 | Step I Decision |
| Step II Decision | 30 | Step II Appeal |
| Step III Appeal | 30 | Step II Decision |

---

## Need More Help?

**In-Spreadsheet Help:**
- **📚 Getting Started** tab - Step-by-step setup instructions
- **❓ FAQ** tab - Common questions and answers
- **📖 Config Guide** tab - User guide for dropdown configuration

**Documentation Files:**
- **AIR.md** - Complete technical reference
- **STEWARD_GUIDE.md** - Guide for stewards
- **COMFORT_VIEW_GUIDE.md** - Accessibility and visual comfort features
- **GitHub Issues** - Report bugs or request features

---

**Version:** 4.7.0
**Last Updated:** 2026-02-14
