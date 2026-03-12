# Quick Deploy - Union Steward Dashboard

**Version:** 4.24.4 | **Architecture:** 48-File Production

Deploy the Union Steward Dashboard to a Google Sheet using the source files in `src/`.

---

## What You Get

- 42 .gs files + 7 .html files (41 .gs in production, after removing DevTools)
- SPA web dashboard with SSO and magic link authentication
- Visible sheets: Config, Member Directory, Grievance Log, Dashboard, Member Satisfaction, Getting Started, FAQ, Config Guide
- 18+ hidden sheets including calculation, workload, weekly questions, contact log, steward tasks, resources, and notifications
- Demo data seeding (1,000 members + 300 grievances)
- 4 menu systems with 100+ functions
- Steward Dashboard, Member Dashboard, and Executive Dashboard
- Meeting management with auto-generated Google Docs
- Mobile-optimized quick actions and web app portal
- Standalone workload tracker portal with PIN authentication
- Comfort View accessibility features

---

## Setup

### Step 1: Get the Files

```bash
git clone https://github.com/Woop91/DDS-Dashboard.git
cd DDS-Dashboard/src
```

### Step 2: Create a New Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com)
2. Click **"+ Blank"** to create a new spreadsheet
3. Name it **"Union Steward Dashboard"**

### Step 3: Open Apps Script Editor

1. In your Google Sheet, click **Extensions > Apps Script**
2. Rename the project to **"Union Dashboard Scripts"**

### Step 4: Create and Paste Each Source File

For each `.gs` file in the `src/` folder, create a new script file in Apps Script:

1. Click **+** next to "Files" to add a new script file
2. Name it exactly as shown (without the `.gs` extension)
3. Paste the contents from the corresponding source file
4. Repeat for all 42 .gs files plus the 7 .html files

**Files to copy (in order):**

| Create File | Paste From |
|-------------|------------|
| `00_Security` | `src/00_Security.gs` |
| `00_DataAccess` | `src/00_DataAccess.gs` |
| `01_Core` | `src/01_Core.gs` |
| `02_DataManagers` | `src/02_DataManagers.gs` |
| `03_UIComponents` | `src/03_UIComponents.gs` |
| `04a_UIMenus` | `src/04a_UIMenus.gs` |
| `04b_AccessibilityFeatures` | `src/04b_AccessibilityFeatures.gs` |
| `04c_InteractiveDashboard` | `src/04c_InteractiveDashboard.gs` |
| `04d_ExecutiveDashboard` | `src/04d_ExecutiveDashboard.gs` |
| `04e_PublicDashboard` | `src/04e_PublicDashboard.gs` |
| `05_Integrations` | `src/05_Integrations.gs` |
| `06_Maintenance` | `src/06_Maintenance.gs` |
| `07_DevTools` | `src/07_DevTools.gs` |
| `08a_SheetSetup` | `src/08a_SheetSetup.gs` |
| `08b_SearchAndCharts` | `src/08b_SearchAndCharts.gs` |
| `08c_FormsAndNotifications` | `src/08c_FormsAndNotifications.gs` |
| `08d_AuditAndFormulas` | `src/08d_AuditAndFormulas.gs` |
| `08e_SurveyEngine` | `src/08e_SurveyEngine.gs` |
| `09_Dashboards` | `src/09_Dashboards.gs` |
| `10a_SheetCreation` | `src/10a_SheetCreation.gs` |
| `10b_SurveyDocSheets` | `src/10b_SurveyDocSheets.gs` |
| `10c_FormHandlers` | `src/10c_FormHandlers.gs` |
| `10d_SyncAndMaintenance` | `src/10d_SyncAndMaintenance.gs` |
| `10_Main` | `src/10_Main.gs` |
| `11_CommandHub` | `src/11_CommandHub.gs` |
| `12_Features` | `src/12_Features.gs` |
| `13_MemberSelfService` | `src/13_MemberSelfService.gs` |
| `14_MeetingCheckIn` | `src/14_MeetingCheckIn.gs` |
| `15_EventBus` | `src/15_EventBus.gs` |
| `16_DashboardEnhancements` | `src/16_DashboardEnhancements.gs` |
| `17_CorrelationEngine` | `src/17_CorrelationEngine.gs` |
| `19_WebDashAuth` | `src/19_WebDashAuth.gs` |
| `20_WebDashConfigReader` | `src/20_WebDashConfigReader.gs` |
| `21_WebDashDataService` | `src/21_WebDashDataService.gs` |
| `22_WebDashApp` | `src/22_WebDashApp.gs` |
| `23_PortalSheets` | `src/23_PortalSheets.gs` |
| `24_WeeklyQuestions` | `src/24_WeeklyQuestions.gs` |
| `25_WorkloadService` | `src/25_WorkloadService.gs` |
| `26_QAForum` | `src/26_QAForum.gs` |
| `27_TimelineService` | `src/27_TimelineService.gs` |
| `28_FailsafeService` | `src/28_FailsafeService.gs` |
| `29_Migrations` | `src/29_Migrations.gs` |
| `index` | `src/index.html` (create as HTML file) |
| `styles` | `src/styles.html` (create as HTML file) |
| `auth_view` | `src/auth_view.html` (create as HTML file) |
| `steward_view` | `src/steward_view.html` (create as HTML file) |
| `member_view` | `src/member_view.html` (create as HTML file) |
| `error_view` | `src/error_view.html` (create as HTML file) |
| `org_chart` | `src/org_chart.html` (create as HTML file) |

**Save** the project (Ctrl+S).

### Step 5: Authorize the Script

1. Select **`CREATE_DASHBOARD`** from the function dropdown
2. Click the **Run** button
3. Click **"Review permissions"** > Choose your account
4. Click **"Advanced"** > **"Go to Union Dashboard Scripts (unsafe)"**
5. Click **"Allow"**
6. The script creates all sheets automatically

### Step 6: Refresh the Page

1. Close the Apps Script editor tab
2. Go back to your Google Sheet
3. **Refresh the page** (F5)
4. You should see **4 custom menus** appear:
   - **Union Hub** -- Main menu for search, grievances, members, calendar, view settings
   - **Admin** -- System administration, data sync, cache, setup
   - **Strategic Ops** -- Command center, analytics, steward management
   - **Field Portal** -- Mobile access, field analytics, web app

### Step 7: Seed Test Data (Optional)

1. Click **Admin > Demo Data > Seed All Sample Data**
2. Wait for seeding to complete
3. This creates 1,000 test members, 300 grievances, and 50 survey responses

---

## You're Done!

Your dashboard is fully operational.

### Visible Sheets Created

- **Config** -- Dropdown values and organization settings
- **Member Directory** -- Union member records (40 columns)
- **Grievance Log** -- Grievance tracking (41 columns)
- **Dashboard** -- Summary statistics and metrics
- **Member Satisfaction** -- Survey response tracking
- **Getting Started** -- Setup instructions
- **FAQ** -- Common questions and answers
- **Config Guide** -- Configuration help

### Hidden Sheets (Auto-Managed)

18+ hidden sheets including calculation, workload, and SPA data sheets with self-healing formulas power the auto-updating columns.

---

## Alternative: Deploy with clasp

If you prefer automated deployment:

```bash
cd DDS-Dashboard
npm install
clasp login
clasp create --type sheets --title "Union Dashboard"
clasp push
```

Then open the sheet, run `CREATE_DASHBOARD`, and refresh.

---

## Troubleshooting

### Menus don't appear?
- Refresh the page (F5)
- Run `onOpen()` manually from the Apps Script editor

### Authorization error?
- Complete Step 5 fully
- Close and reopen the Google Sheet

### Dashboard shows errors?
- Run **Admin > System Diagnostics** to check system health
- Run **Admin > Repair Dashboard** to fix common issues

### Hidden sheets missing?
- Run **Admin > Setup > Setup All Hidden Sheets**

---

## Updating

To update to a new version:

1. Pull the latest code: `git pull`
2. In the Apps Script editor, update only the changed files
3. Save and refresh your Google Sheet

---

## Before Production

1. Run **Admin > Demo Data > NUKE SEEDED DATA** to remove all test data
2. Delete `07_DevTools` from your Apps Script project
3. Configure the Config sheet with your organization's real data

See [SEED_NUKE_GUIDE.md](SEED_NUKE_GUIDE.md) for the full production transition process.

---

**Version:** 4.24.4 (48-File Production Architecture)
**Last Updated:** 2026-03-07
