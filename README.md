# Union Steward Dashboard - Modular Architecture

A comprehensive Google Sheets-based dashboard for managing union grievances, member records, and deadline tracking. This version implements a **modular multi-file architecture** following the Separation of Concerns principle.

## Architecture Overview

The dashboard has been refactored from a single monolithic file into sixteen specialized modules:

```
src/
├── Constants.gs          # Single source of truth for configuration
├── PerformanceUndo.gs    # Caching system and undo/redo functionality
├── ComfortViewFeatures.gs# ADHD/accessibility, themes, focus mode
├── HiddenSheets.gs       # Hidden sheet management
├── FormulaService.gs     # Hidden sheet and formula logic
├── UIService.gs          # Dialogs, sidebars, and UI components
├── GrievanceManager.gs   # Grievance lifecycle management
├── Integrations.gs       # Drive, Calendar, and email services
├── DataIntegrity.gs      # Batch operations, validation, archiving
├── Maintenance.gs        # Admin tools and diagnostics
├── MobileQuickActions.gs # Quick actions, mobile UI
├── WebApp.gs             # Web application deployment
├── TestingValidation.gs  # Test suites and validation
├── DeveloperTools.gs     # Developer utilities
├── Code.gs               # Core setup, forms, dashboard creation
└── Main.gs               # Entry point and triggers
```

### Module Descriptions

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| **Constants.gs** | Configuration constants, sheet names, column mappings, deadline rules | `SHEET_NAMES`, `DEADLINE_RULES`, `GRIEVANCE_STATUS`, `MEMBER_COLS`, `GRIEVANCE_COLS` |
| **PerformanceUndo.gs** | CacheService performance optimization, undo/redo with 50 action history | `getCachedData`, `setCachedData`, `undoLastAction`, `redoAction` |
| **ComfortViewFeatures.gs** | ADHD/accessibility features, themes, focus mode, pomodoro timer | `setTheme`, `toggleFocusMode`, `startPomodoroTimer`, `enableHighContrast` |
| **HiddenSheets.gs** | Hidden sheet management with complex formulas | `setupHiddenSheet`, `protectHiddenSheets`, `updateHiddenCalculations` |
| **FormulaService.gs** | Hidden calculation sheet management | `setupAllHiddenSheets`, `repairAllHiddenSheets` |
| **UIService.gs** | All user interface components | `showDesktopSearch`, `showMultiSelectDialog`, `showQuickActionsMenu` |
| **GrievanceManager.gs** | Grievance creation, step advancement, deadline calculations | `startNewGrievance`, `advanceGrievanceStep`, `recalcAllGrievancesBatched` |
| **Integrations.gs** | External service connections | `setupDriveFolderForGrievance`, `syncDeadlinesToCalendar`, `clearAllCalendarEvents` |
| **DataIntegrity.gs** | Batch operations, validation, orphan detection, auto-archive | `batchUpdateWithRetry`, `detectOrphanedRecords`, `autoArchiveOldRecords`, `validateDataIntegrity` |
| **Maintenance.gs** | Administrative and diagnostic tools | `DIAGNOSE_SETUP`, `REPAIR_DASHBOARD`, `verifyHiddenSheets` |
| **MobileQuickActions.gs** | Quick action dialogs, mobile-optimized UI | `showQuickActions`, `handleQuickAction`, `getMobileMenu` |
| **WebApp.gs** | Web application deployment and endpoints | `doGet`, `doPost`, `getWebInterface` |
| **TestingValidation.gs** | Automated test suites and validation | `runAllTests`, `validateSheetStructure`, `runPerformanceTests` |
| **DeveloperTools.gs** | Developer utilities and debugging | `logDebug`, `exportConfig`, `showDeveloperPanel` |
| **Code.gs** | Core setup, forms, dashboard creation, multi-select | `CREATE_509_DASHBOARD`, `createConfigSheet`, `createMemberDirectory`, `createGrievanceLog` |
| **Main.gs** | Entry point, triggers, initialization | `onOpen`, `onEdit`, `initializeDashboard` |

## Benefits of Modular Architecture

1. **Isolation of Failures**: A bug in Calendar sync won't break the Member Directory
2. **Easier Maintenance**: Update union rules in one place (`Constants.gs`)
3. **Reduced Risk**: UI changes don't risk breaking core data processing
4. **Better Organization**: Find code faster with logical file separation
5. **Team Development**: Multiple developers can work on different modules
6. **Scalability**: Approach 5,000+ lines without chaos

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

2. Install dependencies:
   ```bash
   npm install
   ```

3. Build the consolidated file:
   ```bash
   npm run build
   ```

4. Deploy to Google Apps Script:
   - **Option A**: Copy `dist/ConsolidatedDashboard.gs` to your Google Apps Script project manually
   - **Option B**: Use clasp for automated deployment:
     ```bash
     clasp login
     clasp create --type sheets --title "Union Dashboard"
     npm run deploy
     ```

## Development Workflow

### Making Changes

1. Edit files in the `src/` directory
2. Run the build to generate consolidated output:
   ```bash
   npm run build
   ```
3. For continuous development, use watch mode:
   ```bash
   npm run watch
   ```

### Build Commands

```bash
npm run build      # Build consolidated file
npm run watch      # Watch mode - rebuild on changes
npm run clean      # Remove dist folder
npm run manifest   # Generate appsscript.json
npm run deploy     # Build + push with clasp
```

### File Structure After Build

```
dist/
├── ConsolidatedDashboard.gs  # Combined file for deployment
└── appsscript.json           # App manifest
```

## Features

### Grievance Management
- Create and track grievances through all steps
- Automatic deadline calculations based on Article 23A
- Bulk status updates
- Progress tracking via Current Step

### Member Directory
- Complete member database (32 tracked columns)
- Department and union status tracking
- Data validation for email and phone
- Member Satisfaction Survey processing (89 columns)

### Calendar Integration
- Sync deadlines to Google Calendar
- Configurable reminder days
- Automatic event updates

### Drive Integration
- Auto-create folder structure for each grievance
- Organized subfolders for each step
- Direct links from grievance records

### Self-Healing Formulas
- Six hidden calculation sheets power dashboard statistics
- Automatic recalculation on data changes
- Repair functions for recovery

### Performance & Caching
- CacheService integration for optimized data retrieval
- Undo/Redo system with 50 action history
- Batch operations with retry logic

### Accessibility (Comfort View)
- ADHD-friendly themes and focus mode
- Pomodoro timer integration
- High contrast and reduced motion options
- Customizable color themes

### Data Integrity
- Orphan record detection
- Steward workload balancing
- Auto-archive for old records
- Comprehensive data validation

### Mobile & Quick Actions
- Mobile-optimized UI components
- Quick action dialogs for common tasks
- Responsive design support

### Web Application
- Standalone web app deployment
- REST API endpoints
- Cross-origin request handling

### Developer Tools
- Debug logging utilities
- Configuration export/import
- Test suite automation
- Performance profiling

## Deadline Rules (Article 23A)

| Step | Action | Days |
|------|--------|------|
| Step 1 | Management response | 7 days |
| Step 2 | Appeal deadline | 7 days |
| Step 2 | Management response | 14 days |
| Step 3 | Appeal deadline | 10 days |
| Step 3 | Management response | 21 days |
| Arbitration | Demand deadline | 30 days |

*Deadlines are calculated in business days (excluding weekends)*

## Sheets Structure

### Visible Sheets
- **Member Directory**: Union member records
- **Grievance Tracker**: Active and resolved grievances
- **Calendar Sync**: Calendar integration status
- **Reports**: Generated reports
- **Dashboard**: Summary statistics
- **Settings**: Configuration options
- **Audit Log**: System activity log

### Hidden Sheets (Auto-managed)
- `_CalcMembers`: Member statistics
- `_CalcGrievances`: Grievance aggregations
- `_CalcDeadlines`: Deadline calculations
- `_CalcStats`: Dashboard statistics
- `_CalcSync`: Cross-sheet synchronization
- `_CalcFormulas`: Named formula references

## Troubleshooting

### Running Diagnostics
1. Open **Union Dashboard** menu
2. Select **Admin Tools > System Diagnostics**
3. Review errors and warnings
4. Click **Run Repair** to fix common issues

### Common Issues

| Issue | Solution |
|-------|----------|
| Missing sheets | Run **Repair Dashboard** |
| Broken formulas | Run **Repair Dashboard** |
| Calendar not syncing | Check Calendar permissions in Settings |
| Slow performance | Reduce data in Member Directory |

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes in `src/` files
4. Test with `npm run build`
5. Submit a pull request

## License

MIT License - see LICENSE file for details.

## Version History

- **2.0.0** - Modular multi-file architecture
- **1.x** - Original monolithic version
