# Union Steward Dashboard - 9-File Modular Architecture

A comprehensive Google Sheets-based dashboard for managing union grievances, member records, and deadline tracking. This version implements a **9-file modular architecture** following the Separation of Concerns principle.

## Architecture Overview

The dashboard uses a streamlined 9-file architecture for clarity and maintainability:

```
src/
├── 01_Constants.gs        # Single source of truth for configuration
├── 02_MemberManager.gs    # Member operations and directory management
├── 03_GrievanceManager.gs # Grievance lifecycle and deadline tracking
├── 04_UIService.gs        # UI, Comfort View, mobile, Strategic Command Center
├── 05_Integrations.gs     # Drive, Calendar, WebApp integration
├── 06_Maintenance.gs      # Diagnostics, caching, performance
├── 07_DevTools.gs         # Test data generation (DELETE BEFORE PROD)
├── 08_Code.gs             # Core setup, hidden sheets, dashboard creation
└── 09_Main.gs             # Entry point and triggers
```

### Module Descriptions

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| **01_Constants.gs** | Configuration constants, sheet names, column mappings, deadline rules | `SHEETS`, `MEMBER_COLS`, `GRIEVANCE_COLS`, `CONFIG_COLS`, `COMMAND_CONFIG` |
| **02_MemberManager.gs** | Member directory operations, steward management, ID generation | `addMember`, `promoteToSteward`, `generateMissingMemberIDs`, `checkDuplicateMemberIDs` |
| **03_GrievanceManager.gs** | Grievance creation, step advancement, deadline calculations | `startNewGrievance`, `advanceGrievanceStep`, `recalcAllGrievancesBatched` |
| **04_UIService.gs** | Dialogs, Comfort View, mobile UI, Strategic Command Center dashboards | `showDesktopSearch`, `rebuildExecutiveDashboard`, `rebuildMemberAnalytics`, `setupMidnightTrigger` |
| **05_Integrations.gs** | External service connections (Drive, Calendar, WebApp) | `setupDriveFolderForGrievance`, `syncDeadlinesToCalendar`, `doGet`, `doPost` |
| **06_Maintenance.gs** | Admin tools, diagnostics, caching, performance optimization | `DIAGNOSE_SETUP`, `REPAIR_DASHBOARD`, `getCachedData`, `warmUpCaches` |
| **07_DevTools.gs** | Test data generation and demo utilities (remove before production) | `seedAllSampleData`, `nukeDemoData`, `showDeveloperPanel` |
| **08_Code.gs** | Core setup, hidden sheet management, dashboard creation | `CREATE_509_DASHBOARD`, `setupAllHiddenSheets`, `createConfigSheet` |
| **09_Main.gs** | Entry point, triggers, edit handlers | `onOpen`, `onEdit`, `initializeDashboard` |

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

## Advanced Features

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
- Focus-friendly themes and distraction-free mode
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

### Strategic Command Center (v3.6.0)

The 509 Strategic Command Center provides executive-level analytics and automation:

#### Dual-Dashboard Architecture
- **Executive Command (PII)**: Internal dashboard with member names, steward workload, grievance insights
- **Member Analytics (No PII)**: PII-safe dashboard with morale gauge, leadership pipeline, heatmaps

#### Strategic Intelligence
- **Unit Hot Zones**: Identifies locations with 3+ active grievances
- **Rising Stars**: Highlights top-performing stewards by score and win rate
- **Management Hostility Funnel**: Analyzes denial rates across grievance steps
- **Bargaining Cheat Sheet**: Strategic data for contract negotiations

#### Automation Engines
- **Midnight Auto-Refresh**: Daily 12AM trigger refreshes dashboards and sends overdue alerts
- **Auto-ID Generator**: Creates member IDs using configurable unit codes from Config sheet
- **Stage-Gate Workflow**: Sends escalation alerts when cases reach Step II/III/Arbitration
- **Duplicate Detection**: Finds and highlights duplicate Member IDs

#### Document Generation
- **PDF Engine**: Creates grievance PDFs with digital signature blocks
- **Email Automation**: Weekly PDF snapshots and escalation notifications

#### Steward Management
- **Promote/Demote**: One-click steward status changes with toolkit emails
- **Workload Tracking**: Visual steward case load distribution

#### Dynamic Configuration (Config Sheet)
All settings are configurable without code changes:

| Column | Setting | Format |
|--------|---------|--------|
| AS | Chief Steward Email | email@example.com |
| AT | Unit Codes | `Main Station:MS,Field Ops:FO` |
| AU | Archive Folder ID | Google Drive folder ID |
| AV | Escalation Statuses | `In Arbitration,Appealed` |
| AW | Escalation Steps | `Step II,Step III,Arbitration` |

#### Status Color Mapping
Automatic status-based coloring for Grievance Log:

| Status | Color | Meaning |
|--------|-------|---------|
| Open | Yellow | Active case |
| Pending Info | Purple | Waiting on info |
| Won | Green | Victory |
| Denied | Red | Loss |
| Settled | Blue | Negotiated resolution |
| In Arbitration | Red | High stakes |
| Closed | Gray | Complete |

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

- **3.6.0** - Strategic Command Center with dual-dashboards, midnight auto-refresh, dynamic config
- **2.3.0** - Enhanced grievance dashboard with 9-file modular architecture
- **2.2.0** - Complete feature parity with 16-module modular architecture
- **2.0.0** - Initial modular multi-file architecture
- **1.x** - Original monolithic version
