# 509 Strategic Command Center - v4.3.7 Complete Menu System

A comprehensive Google Sheets-based dashboard for managing union grievances, member records, and deadline tracking. This version implements an **11-file modular architecture** following the Separation of Concerns principle.

## Architecture Overview

The dashboard uses a streamlined 11-file architecture for clarity and maintainability:

```
src/
├── 01_Constants.gs              # Single source of truth for configuration
├── 02_MemberManager.gs          # Member operations and directory management
├── 03_GrievanceManager.gs       # Grievance lifecycle and deadline tracking
├── 04_UIService.gs              # UI, Comfort View, mobile, Strategic Command Center
├── 05_Integrations.gs           # Drive, Calendar, WebApp integration
├── 06_Maintenance.gs            # Diagnostics, caching, performance
├── 07_DevTools.gs               # Test data generation (DELETE BEFORE PROD)
├── 08_Code.gs                   # Core setup, hidden sheets, dashboard creation
├── 09_Main.gs                   # Entry point and triggers
├── 10_CommandCenter.gs          # 509 Strategic Command Center Unified Master Engine (v4.0)
└── 11_SecureMemberDashboard.gs  # Material Design member portal, Web App & secure portal (v4.2.0)
```

### Module Descriptions

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| **01_Constants.gs** | Configuration constants, sheet names, column mappings, deadline rules | `SHEETS`, `MEMBER_COLS`, `GRIEVANCE_COLS`, `CONFIG_COLS`, `COMMAND_CONFIG`, `VERSION_INFO` |
| **02_MemberManager.gs** | Member directory operations, steward management, ID generation | `addMember`, `promoteToSteward`, `generateMissingMemberIDs`, `checkDuplicateMemberIDs` |
| **03_GrievanceManager.gs** | Grievance creation, step advancement, deadline calculations | `startNewGrievance`, `advanceGrievanceStep`, `recalcAllGrievancesBatched` |
| **04_UIService.gs** | Dialogs, Comfort View, mobile UI, Strategic Command Center dashboards | `showDesktopSearch`, `rebuildExecutiveDashboard`, `rebuildMemberAnalytics`, `setupMidnightTrigger` |
| **05_Integrations.gs** | External service connections (Drive, Calendar, WebApp) | `setupDriveFolderForGrievance`, `syncDeadlinesToCalendar`, `doGet`, `doPost` |
| **06_Maintenance.gs** | Admin tools, diagnostics, caching, performance optimization | `DIAGNOSE_SETUP`, `REPAIR_DASHBOARD`, `getCachedData`, `warmUpCaches` |
| **07_DevTools.gs** | Test data generation and demo utilities (remove before production) | `seedAllSampleData`, `nukeDemoData`, `showDeveloperPanel` |
| **08_Code.gs** | Core setup, hidden sheet management, dashboard creation | `CREATE_509_DASHBOARD`, `setupAllHiddenSheets`, `createConfigSheet` |
| **10_CommandCenter.gs** | v4.0 Unified Master Engine - PDF engine, analytics, scaling hooks | `navToMobile`, `createGrievancePDF`, `showUnitHealthReport`, `calculateUnitHealth` |
| **11_SecureMemberDashboard.gs** | Material Design member portal with Google Charts, PII protection | `showPublicMemberDashboard`, `showStewardPerformanceModal`, `safetyValveScrub`, `getAdvancedAnalytics` |
| **09_Main.gs** | Entry point, triggers, edit handlers | `onOpen`, `onEdit`, `initializeDashboard` |

## v4.0 Features

- **Security Fortress**: Audit Log & Sabotage Protection (>15 cells mass deletion detection)
- **High-Performance Engine**: Batch Array Processing for 5,000+ members
- **PDF Signature Engine**: Legal signature-ready PDF generation from templates
- **Mobile/Pocket View**: Field-optimized column hiding for smartphone access
- **Production Mode**: NUKE logic with UI self-hiding (Demo Data menu disappears)
- **Analytics & Insights**: Unit Health Reports, Grievance Trends analysis
- **Scaling Hooks**: OCR transcription and Sentiment analysis placeholders

## v4.0.2 Features

- **Secure Member Dashboard**: Interactive modal with Google Charts (pie chart for issue categories, gauge for trust score, area chart for trends)
- **Material Icons Integration**: Professional iconography throughout the UI
- **Live Steward Search**: Real-time filtering of steward directory by name, unit, or location
- **Progress Tracking**: Visual progress bars for union goals (steward coverage, survey participation)
- **Steward Performance Modal**: View active cases, total cases, and win rates for all stewards
- **Email Dashboard Link**: One-click email sending of dashboard URL to selected members
- **Zero PII Exposure**: All member-facing views show only aggregate statistics

## v4.3.x Features (CURRENT)

- **Two-Dashboard Architecture**: Unified Steward Dashboard (internal, with PII) and Member Dashboard (public, no PII)
- **Steward Dashboard**: 6 tabs - Overview, Workload, Analytics, Hot Spots, Bargaining, Satisfaction
- **Member Satisfaction Analysis**: 8 survey sections with scores, progress bars, and question breakdowns
- **Production Mode Polish**: Demo Data menu properly hides after NUKE in all menus
- **NUKE Enhancements**: 3-5 minute warnings, cleans documentation tabs, adds repo link to FAQ, applies tab colors
- **Professional Tab Colors**: Blue (data sheets), Green (documentation), Red (satisfaction), Orange (config)
- **Dynamic Row Styling**: `applyZebraStripes()` and `applyThemeToSheet_()` use `getMaxRows()` to style ALL rows
- **DIALOG_SIZES Constant**: Standard modal dimensions (SMALL, MEDIUM, LARGE, FULLSCREEN, SIDEBAR)

## v4.2.0 Features

- **Modal SPA Architecture**: All dashboards converted from sheet-based to responsive modal dialogs
- **Executive Dashboard Modal**: Chart.js visualizations with win rate donut, status pie, and trends line chart
- **Bridge Pattern Implementation**: Server-side JSON aggregation with client-side rendering for performance
- **Web App Entry Point**: `doGet(e)` handler for URL parameter routing (`?member=ID` for personalized portal)
- **Secure Member Portal**: Personalized member view via URL with PII protection and safety valve scrubbing
- **Public Statistics Portal**: Aggregate union statistics portal for public-facing deployment
- **Email Portal Links**: Send personalized dashboard URLs directly to members
- **Code Cleanup**: Removed 9 orphaned sheet-based helper functions, resolved 3 duplicate function conflicts
- **Dark Gradient Theme**: Professional dark theme with Chart.js visualizations across all modals

## v4.0.3 Features

- **Material Design Integration**: Full Material Design UI with Google Material Icons and Roboto typography
- **Google Charts Analytics**: Interactive Treemap for unit density visualization and Area Charts for sentiment trends
- **Safety Valve PII Scrubbing**: Auto-redaction of phone numbers and SSN patterns from public-facing dashboards
- **Weingarten Rights Utility**: Emergency rights statement with tap-to-expand for member protection during meetings
- **Unit Density Heat Map**: Visual representation of grievance activity by unit (green → yellow → red coloring)
- **Sentiment Trend Analysis**: Union morale tracking over time from survey data
- **Steward Workload Balancing**: Workload metrics with overload detection (flags stewards with 8+ active cases)
- **Standalone Analytics Charts**: Dedicated modals for Treemap, Sentiment Trend, and Workload Report
- **High-Contrast Dark Theme**: Professional dark gradient backgrounds optimized for readability

## Benefits of 11-File Architecture

1. **Clear Separation**: Each file has one clear purpose
2. **Easy Navigation**: Numbered prefixes show dependency order on GitHub
3. **Production Ready**: Delete `07_DevTools.gs` before go-live (10 files in production)
4. **Isolation of Failures**: A bug in Calendar sync won't break the Member Directory
5. **Easier Maintenance**: Update union rules in one place (`01_Constants.gs`)
6. **Scalability**: Handles 5,000+ members without performance issues

## Installation

### Prerequisites

- Google account with Google Sheets
- [clasp](https://github.com/google/clasp) CLI (optional, for deployment)

### Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/Woop91/MULTIPLE-SCRIPS-REPO.git
   cd MULTIPLE-SCRIPS-REPO
   ```

2. Deploy to Google Apps Script:
   - **Option A**: Copy each file from `src/` to your Google Apps Script project manually
   - **Option B**: Use clasp for automated deployment:
     ```bash
     clasp login
     clasp create --type sheets --title "Union Dashboard"
     clasp push
     ```

## Development Workflow

### Making Changes

1. Edit files in the `src/` directory (numbered 01-11)
2. Copy updated files to Google Apps Script
3. Save and refresh your Google Sheet

### Source Files

| File | Purpose |
|------|---------|
| `01_Constants.gs` | Configuration constants, column mappings |
| `02_MemberManager.gs` | Member operations and directory management |
| `03_GrievanceManager.gs` | Grievance lifecycle and deadline tracking |
| `04_UIService.gs` | UI, Comfort View, mobile, Strategic Command Center |
| `05_Integrations.gs` | Drive, Calendar, WebApp integration |
| `06_Maintenance.gs` | Diagnostics, caching, performance |
| `07_DevTools.gs` | Test data generation (DELETE BEFORE PROD) |
| `08_Code.gs` | Core setup, hidden sheets, dashboard creation |
| `09_Main.gs` | Entry point and triggers |
| `10_CommandCenter.gs` | Strategic Command Center features |
| `11_SecureMemberDashboard.gs` | Material Design member portal with analytics |

## Going Live (Production)

Before deploying to production:

1. **Run cleanup**: Execute `NUKE_SEEDED_DATA()` to remove all test data
2. **Delete DevTools**: Remove `07_DevTools.gs` from your Apps Script project
3. **Save**: The Demo menu will automatically disappear

## File Structure After Production

```
src/
├── 01_Constants.gs              # Configuration
├── 02_MemberManager.gs          # Member operations
├── 03_GrievanceManager.gs       # Grievance lifecycle
├── 04_UIService.gs              # UI components
├── 05_Integrations.gs           # External services
├── 06_Maintenance.gs            # Admin tools
├── 08_Code.gs                   # Core setup
├── 09_Main.gs                   # Entry point
├── 10_CommandCenter.gs          # Strategic Command Center
└── 11_SecureMemberDashboard.gs  # Material Design member portal
```

**10 files in production** (DevTools removed)

## Key Features

- **Grievance Tracking**: Full lifecycle from filing to resolution
- **Member Directory**: Contact info, job metadata, steward assignments
- **Deadline Management**: Auto-calculated deadlines with Calendar sync
- **Drive Integration**: Auto-created folders for each grievance
- **Comfort View**: ADHD-friendly themes, focus mode, reduced motion
- **Mobile UI**: Quick actions optimized for phone access
- **Data Integrity**: Batch operations, validation, orphan detection
- **Audit Logging**: Track all changes with timestamps

## Menu System Overview (v4.2.2)

The dashboard provides 4 comprehensive top-level menus with 100+ functions:

### 📊 Union Hub (Main Menu)
| Submenu | Key Features |
|---------|--------------|
| 🔍 Search | Desktop, Quick, and Advanced search |
| 📋 Grievances | Create, edit, bulk update, view active, analytics |
| 👥 Members | Add, find, import/export, steward directory, contact forms, surveys, ID management |
| 📅 Calendar | Sync deadlines, view upcoming, clear events |
| 📁 Google Drive | Setup folders, view files, batch create |
| 🔔 Notifications | Settings, test notifications |
| 👁️ View | Dashboards, mobile URL, dark mode, themes |
| ♿ Comfort View | Focus mode, zebra stripes, pomodoro, notepad |
| 📝 Multi-Select | Editor, auto-open triggers |

### 🛠️ Admin (System Administration)
| Submenu | Key Features |
|---------|--------------|
| ⚙️ Automation | Auto-refresh, midnight triggers, email snapshots |
| 🔄 Data Sync | Sync all data, grievance/member sync, triggers |
| ✅ Validation | Bulk validation, settings, indicators |
| 🗄️ Cache | Cache status, warm up, clear caches |
| 🏗️ Setup | Hidden sheets, data validations, defaults |
| 🎭 Demo Data | Seed sample data, nuke seeded data |

### 🎯 Strategic Ops (Strategic Operations)
| Submenu | Key Features |
|---------|--------------|
| 👁️ Command Center | Executive/Member dashboards, steward performance |
| 🎯 Strategic Intelligence | Hot zones, rising stars, hostility report |
| 📊 Analytics & Charts | Treemap, sentiment trends, workload report |
| 🆔 ID & Data Engines | ID generation, duplicate check, PDF creation |
| 👤 Steward Management | Promote/demote, contact forms, surveys |

### 📱 Field Portal (Mobile/Field Operations)
| Submenu | Key Features |
|---------|--------------|
| 📱 Field Accessibility | Mobile view, get mobile URL |
| 📈 Analytics & Insights | Unit health, grievance trends, precedents |
| 🌐 Web App & Portal | Build portals, send emails, JSON APIs |

See [AIR.md](AIR.md) for the complete menu structure documentation.

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

### Strategic Command Center (v4.0.0)

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

- **4.3.8** - Modal Consolidation: Member Satisfaction sheet now hidden, access via modal only
- **4.3.7** - Dynamic Row Styling: Uses `getMaxRows()` for styling all sheet rows
- **4.3.6** - Full Row Styling: Zebra stripes and themes apply to all rows in data sheets
- **4.3.5** - Production Polish: Demo menu fix, NUKE cleans docs and applies tab colors
- **4.3.4** - Satisfaction Analytics: 8-section survey analysis in both dashboards
- **4.3.3** - Unified Two-Dashboard Architecture: Steward Dashboard + Member Dashboard
- **4.3.2** - Modal Dashboard Consolidation: Deprecated sheet-based dashboards
- **4.2.0** - Modal Command Center: All dashboards converted to modal SPA architecture, Web App entry point with member portal, removed orphaned code
- **4.1.0** - Multi-Key Smart Match for duplicate prevention, Analytics wiring, Search Precedents feature
- **4.0.3** - Material Design Integration with Google Charts, Safety Valve PII scrubbing, Weingarten Rights utility (11-file architecture)
- **4.0.2** - Secure Member Dashboard with live steward search and zero PII exposure
- **4.0.0** - Unified Master Engine with PDF generation, mobile view, analytics hooks
- **3.6.0** - Strategic Command Center with dual-dashboards, midnight auto-refresh, dynamic config
- **2.3.0** - Enhanced grievance dashboard with 9-file modular architecture
- **2.2.0** - Complete feature parity with 16-module modular architecture
- **2.0.0** - Initial modular multi-file architecture
- **1.x** - Original monolithic version
