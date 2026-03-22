# SolidBase Sync Log

## Sync Policy
- **Source of truth:** DDS-Dashboard Main
- **Target:** SolidBase Main
- **Exclusions from SolidBase:**
  - DDS Apps Script ID (redacted)
  - `25_WorkloadService.gs` + `poms_reference.html` (org-specific features)
  - `25_WorkloadService.test.js`, `scripts/sync-org-chart.js`
  - Org-specific references (DDS, MassAbility, SEIU 509) replaced with generic equivalents
  - `UI_REVIEW.md`, `docs/archived-reviews/` audit docs (DDS-only)
- **SolidBase-specific branding:** "SolidBase — Built for the collective."

## Sync History

### 2026-03-22 — Full sync from DDS v4.35.0
- Source: DDS-Dashboard `195c026` (Main, post-merge)
- Scope: All src/, test/, build.js, configs
- Major changes:
  - File consolidation: FormHandlers+SyncAndMaintenance → FormsAndSync
  - New services: TrendAlertService, EngagementService, NewFeatureServices
  - Removed dead modules: ExecutiveDashboard, SheetFormatting, TabModals,
    DashboardEnhancements, Migrations
  - New SHEETS constants: handoff, mentorship, comms, KB, SMS, RSVP
  - Modal system, sheet formatting, onboarding wizard, 2FA modal,
    audit logs, outcome analytics, calendar view, mobile toolbar
- Org cleanup: All DDS/509/MassAbility refs generalized
- Removed DDS-specific PDFs (contracts, job descriptions, annual report)
- org_chart.html: replaced PII with generic placeholder data
- Excluded: 25_WorkloadService.gs, poms_reference.html
- Tests: 56 suites, 2831 pass

### 2026-03-21 — Full sync from DDS v4.33.0
- Source: DDS-Dashboard `e172d7c` (Main)
- Scope: All src/, test/, docs/, scripts/, root configs
- Changes: Full file copy, org reference scrub, POMS + Workload Tracker feature removal
- Files synced: 50+ src, 55+ test, 25+ docs
- Excluded: 25_WorkloadService.gs, poms_reference.html, sync-org-chart.js
