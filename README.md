# Union Steward Dashboard - 9-File Modular Architecture

A comprehensive Google Sheets-based dashboard for managing union grievances, member records, and deadline tracking. This version implements a **9-file modular architecture** following the Separation of Concerns principle.

## Architecture Overview

The dashboard uses a streamlined 9-file modular structure:

```
src/
├── 01_Constants.gs       # Single source of truth (IDs, column mappings, policy rules)
├── 02_MemberManager.gs   # Member directory operations, form submissions
├── 03_GrievanceManager.gs# Grievance lifecycle, deadlines, filing
├── 04_UIService.gs       # UI components, Comfort View, mobile UI
├── 05_Integrations.gs    # Drive, Calendar, WebApp, notifications
├── 06_Maintenance.gs     # Diagnostics, data integrity, caching, audits
├── 07_DevTools.gs        # Test data & nuke/seed (DELETE BEFORE PRODUCTION)
├── 08_Code.gs            # Core setup, hidden sheets, formulas, dashboard creation
└── 09_Main.gs            # Entry point, onOpen, triggers
```

### Module Descriptions

| # | Module | Purpose | Key Functions |
|---|--------|---------|---------------|
| 01 | **Constants.gs** | Single source of truth - IDs, sheet names, column mappings, 21-day deadline rules | `SHEETS`, `MEMBER_COLS`, `GRIEVANCE_COLS`, `CONFIG_COLS`, `generateNameBasedId` |
| 02 | **MemberManager.gs** | Member directory CRUD, contact form handling, multi-select fields | `addNewMember`, `getMemberById`, `searchMembers`, `startGrievanceForMember` |
| 03 | **GrievanceManager.gs** | Grievance creation, step advancement, deadline calculations | `startNewGrievance`, `advanceGrievanceStep`, `calculateNextStepDeadline`, `recalcAllGrievancesBatched` |
| 04 | **UIService.gs** | Dialogs, sidebars, Comfort View (ADHD accessibility), mobile UI | `showDesktopSearch`, `showMultiSelectDialog`, `applyTheme`, `toggleFocusMode`, `showSmartDashboard` |
| 05 | **Integrations.gs** | External services - Drive folders, Calendar sync, WebApp deployment | `setupDriveFolderForGrievance`, `syncDeadlinesToCalendar`, `doGet`, `showOrphanedFolderCleanupDialog` |
| 06 | **Maintenance.gs** | Admin tools, diagnostics, data integrity, caching, undo/redo | `DIAGNOSE_SETUP`, `batchSetValues`, `getCachedData`, `undoLastAction`, `logAuditEvent` |
| 07 | **DevTools.gs** | Test data seeding and cleanup - **DELETE BEFORE PRODUCTION** | `SEED_SAMPLE_DATA`, `NUKE_SEEDED_DATA`, `runAllTests`, `validateEmailAddress` |
| 08 | **Code.gs** | Core dashboard setup, hidden sheets, formulas | `CREATE_509_DASHBOARD`, `createConfigSheet`, `setupAllHiddenSheets`, `syncAllData` |
| 09 | **Main.gs** | Entry point, triggers, menu building | `onOpen`, `onEdit`, `createDashboardMenu`, `setupTriggers` |

## Benefits of 9-File Architecture

1. **Clear Separation**: Each file has one clear purpose
2. **Easy Navigation**: Numbered prefixes show dependency order on GitHub
3. **Production Ready**: Delete `07_DevTools.gs` before go-live (8 files in production)
4. **Isolation of Failures**: A bug in Calendar sync won't break the Member Directory
5. **Easier Maintenance**: Update union rules in one place (`01_Constants.gs`)
6. **Scalability**: Handles 1,000+ members without performance issues

## Installation

### Prerequisites

- Node.js 14+ (for build tools)
- Google account with Google Sheets
- [clasp](https://github.com/google/clasp) CLI (optional, for deployment)

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/Woop91/MULTIPLE-SCRIPS-REPO.git
   cd MULTIPLE-SCRIPS-REPO
   ```

2. Build the consolidated file:
   ```bash
   node build.js
   ```

3. Deploy to Google Apps Script:
   - **Option A**: Copy `dist/ConsolidatedDashboard.gs` to your Google Apps Script project manually
   - **Option B**: Use clasp for automated deployment:
     ```bash
     clasp login
     clasp create --type sheets --title "Union Dashboard"
     clasp push
     ```

## Development Workflow

### Making Changes

1. Edit files in the `src/` directory (numbered 01-09)
2. Run the build to generate consolidated output:
   ```bash
   node build.js
   ```
3. Copy `dist/ConsolidatedDashboard.gs` to Google Apps Script

### Watch Mode (Development)

```bash
node build.js --watch
```

This automatically rebuilds when you save changes to any `.gs` file.

### Build Commands

| Command | Description |
|---------|-------------|
| `node build.js` | Build consolidated file |
| `node build.js --watch` | Watch mode for development |
| `node build.js --clean` | Remove dist folder |
| `node build.js --manifest` | Generate appsscript.json |

## Going Live (Production)

Before deploying to production:

1. **Run cleanup**: Execute `NUKE_SEEDED_DATA()` to remove all test data
2. **Delete DevTools**: Remove `07_DevTools.gs` from the build order in `build.js`
3. **Rebuild**: Run `node build.js` to create production bundle
4. **Deploy**: Copy the consolidated file to your Apps Script project

The Demo menu will automatically disappear when DevTools.gs is removed.

## File Structure After Production

```
src/
├── 01_Constants.gs       # Configuration
├── 02_MemberManager.gs   # Member operations
├── 03_GrievanceManager.gs# Grievance lifecycle
├── 04_UIService.gs       # UI components
├── 05_Integrations.gs    # External services
├── 06_Maintenance.gs     # Admin tools
├── 08_Code.gs            # Core setup
└── 09_Main.gs            # Entry point
```

**8 files in production** (DevTools removed)

## Key Features

- **Grievance Tracking**: Full lifecycle from filing to resolution
- **Member Directory**: Contact info, job metadata, steward assignments
- **Deadline Management**: Auto-calculated deadlines with Calendar sync
- **Drive Integration**: Auto-created folders for each grievance
- **Comfort View**: ADHD-friendly themes, focus mode, reduced motion
- **Mobile UI**: Quick actions optimized for phone access
- **Data Integrity**: Batch operations, validation, orphan detection
- **Audit Logging**: Track all changes with timestamps

## Version

- **Version**: 2.3.0
- **Build**: v3.50
- **Architecture**: 9-File Modular
- **Codename**: Enhanced Grievance Dashboard

## License

Free for use by non-profit collective bargaining groups and unions.
