# Quick Deploy - Union Steward Dashboard

## 10-File Modular Deployment

Deploy the **Union Steward Dashboard** using the **10 source files** in the `src/` folder.

---

## What You Get

- 10 modular source files for easy maintenance
- 11 visible sheets + 6 hidden calculation sheets
- Demo data seeding (members + grievances)
- 12-section Dashboard with real-time metrics
- Comfort View accessibility features
- Strategic Command Center with dual dashboards
- Mobile-optimized quick actions and web app

---

## 5-Minute Setup

### **Step 1: Get the Files**

```bash
git clone https://github.com/Woop91/MULTIPLE-SCRIPS-REPO.git
cd MULTIPLE-SCRIPS-REPO/src
```

### **Step 2: Create a New Google Sheet**

1. Go to [sheets.google.com](https://sheets.google.com)
2. Click **"+ Blank"** to create a new spreadsheet
3. Name it **"Union Steward Dashboard"**

### **Step 3: Open Apps Script Editor**

1. In your Google Sheet, click **Extensions → Apps Script**
2. Rename the project to **"Union Dashboard Scripts"**

### **Step 4: Create and Paste Each Source File**

For each file in the `src/` folder, create a new script file in Apps Script:

1. Click **+** next to "Files" to add a new script file
2. Name it exactly as shown (without .gs extension)
3. Paste the contents from the corresponding source file

| Create File | Paste From |
|-------------|------------|
| `01_Constants` | `src/01_Constants.gs` |
| `02_MemberManager` | `src/02_MemberManager.gs` |
| `03_GrievanceManager` | `src/03_GrievanceManager.gs` |
| `04_UIService` | `src/04_UIService.gs` |
| `05_Integrations` | `src/05_Integrations.gs` |
| `06_Maintenance` | `src/06_Maintenance.gs` |
| `07_DevTools` | `src/07_DevTools.gs` |
| `08_Code` | `src/08_Code.gs` |
| `09_Main` | `src/09_Main.gs` |
| `10_CommandCenter` | `src/10_CommandCenter.gs` |

5. **Save** the project (Ctrl+S)

### **Step 5: Authorize the Script**

1. Select **`CREATE_509_DASHBOARD`** from the function dropdown
2. Click the **Run** button
3. Click **"Review permissions"** → Choose your account
4. Click **"Advanced"** → **"Go to Union Dashboard Scripts (unsafe)"**
5. Click **"Allow"**
6. The script creates all sheets (~10-15 seconds)

### **Step 6: Refresh the Page**

1. Close the Apps Script editor tab
2. Go back to your Google Sheet
3. **Refresh the page** (F5)
4. You should see **2 custom menus** appear:
   - **Dashboard** - Main menu with submenus for Search, Grievances, Members, Calendar, View, and Admin
   - **509 Command** - Strategic Command Center with analytics and automation

### **Step 7: Seed Test Data (Optional)**

1. Click **Dashboard → Admin → Demo Data → Seed All Sample Data**
2. Wait for seeding to complete (~30 seconds)

---

## You're Done!

Your dashboard is fully operational with:

### Visible Sheets
- **Config** - Dropdown values and settings
- **Member Directory** - 32 columns for member records
- **Grievance Log** - 35 columns for grievance tracking
- **Dashboard** - 12 analytics sections
- **Interactive Dashboard** - Real-time metrics with My Cases tab
- **Member Satisfaction** - Survey response tracking
- **Feedback** - Bug/feature tracking
- **Function Checklist** - Function reference
- **Getting Started** - Setup instructions
- **FAQ** - Common questions
- **Config Guide** - Configuration help

### Hidden Sheets (Auto-managed)
- 6 calculation sheets with self-healing formulas

---

## Source Files Overview

| File | Purpose |
|------|---------|
| `01_Constants.gs` | Configuration constants, column mappings, themes |
| `02_MemberManager.gs` | Member operations and directory management |
| `03_GrievanceManager.gs` | Grievance lifecycle and deadline tracking |
| `04_UIService.gs` | UI, Comfort View, mobile, Strategic Command Center |
| `05_Integrations.gs` | Drive, Calendar, WebApp integration |
| `06_Maintenance.gs` | Diagnostics, caching, performance |
| `07_DevTools.gs` | Test data generation (DELETE BEFORE PROD) |
| `08_Code.gs` | Core setup, hidden sheets, dashboard creation |
| `09_Main.gs` | Entry point and triggers |
| `10_CommandCenter.gs` | Strategic Command Center features |

---

## Troubleshooting

### **Menus don't appear?**
- Refresh the page (F5)
- Run `onOpen()` manually from Apps Script editor

### **Authorization error?**
- Complete Step 5 fully
- Close and reopen the Google Sheet

### **Dashboard shows errors?**
- Click **Dashboard → Admin → Repair Dashboard**
- Click **Dashboard → Admin → Diagnose Setup** to check system health

### **Hidden sheets missing?**
- Click **Dashboard → Admin → Setup All Hidden Sheets**

---

## Updating

To update to a new version:

1. Pull the latest code: `git pull`
2. In Apps Script editor, update only the changed files
3. Save and refresh your Google Sheet

**Tip:** Only `01_Constants.gs` and `08_Code.gs` typically need updates for bug fixes.

---

## Before Production

**Remove Demo Tools:**
1. Run **Dashboard → Admin → Demo Data → NUKE SEEDED DATA**
2. Delete `07_DevTools` from your Apps Script project

---

**Version:** 4.0.0 (Unified Master Engine)
**Architecture:** 10-File Modular
**Last Updated:** 2026-01-15
