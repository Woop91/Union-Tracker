# Changelog

All notable changes to the Union Dashboard project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [4.9.0] - 2026-02-17

### Added
- **Constant Contact v3 API integration** — read-only email engagement metrics sync with OAuth2 authorization, auto token refresh, rate limiting, and pagination
- Member Directory columns `OPEN_RATE` and `RECENT_CONTACT_DATE` now populated by CC sync
- 30 new tests covering the Constant Contact integration
- **Multi-select dropdown support for Grievance Log** — checkbox UI editor with UserProperties-based storage
- **Dropdown system overhaul** — Config-driven values with dynamic ranges and missing field coverage
- **Auto-discovery column system** — zero manual updates on sheet restructure; all column references use dynamic constants
- 151 column system tests added to validate dynamic constant resolution

### Changed
- All remaining hardcoded column indices replaced with dynamic `CONFIG_COLS` and `MEMBER_COLS` constants
- Removed all legacy 509 local number references from repository

## [4.8.2] - 2026-02-16

### Added
- **State field** added to member contact update across all surfaces (self-service portal, contact form, profile data)

### Changed
- Member self-service portal edit form now has 5 fields (Email, Phone, Preferred Contact, Best Time, State)

## [4.8.1] - 2026-02-15

### Added
- **5 new contact form fields** — Hire Date, Employee ID, Street Address, City, Zip Code

### Changed
- **Unified Member ID system** — all ID generation now uses name-based format (`MJASM472`)

### Removed
- Legacy random unit-code ID generators

## [4.8.0] - 2026-02-15

### Security
- **Security event alerting system** — threat detection and notifications at web app, edit trigger, and self-service entry points
- **Zero-knowledge survey vault** — all survey verification data stored as SHA-256 hashes only; no plaintext PII written to any sheet

### Added
- Survey Completion Tracker with dialog, reminders, and round management

### Changed
- **Member ID format** changed from random 5-digit to sequential for import reliability
- Import dedup now happens once client-side instead of per-batch
- Satisfaction form submission and review functions now use vault storage

## [4.7.0] - 2026-02-14

### Fixed
- **40+ code review issues** resolved across security, correctness, performance, and test quality
- XSS hardening, onEdit optimization, DevTools guard, deduplicated `escapeHtml`
- Removed broken column references and empty stubs

### Changed
- Version bumped to 4.7.0 across all files
- Test suite expanded from 1016 to 1090 tests (74 new tests)

## [4.6.0] - 2026-02-12

### Added
- **VERSION_HISTORY constant** — centralized release tracking with lookup function
- **Meeting Notes & Agenda Document Automation** — auto-generated Google Docs, two-tier steward agenda sharing, scheduled notifications
- **Meeting Notes Dashboard Tab** — completed meetings with search and view-only Doc links
- **Member Drive Folder Quick Action** — creates/reuses Google Drive folder per member
- **Meeting Event Scheduling** — full calendar lifecycle with check-in activation
- **Grievance Date Override** — stewards can overwrite dates with downstream deadline recalculation

### Changed
- Meeting Check-In Log expanded from 13 to 16 columns
- Meeting setup dialog updated with steward selection checkboxes

## [4.5.1] - 2026-02-11

### Fixed
- **Engagement tracking** — resolved 6 undefined column references causing incorrect dashboard data
- **Version consistency** — synced API_VERSION and COMMAND_CONFIG.VERSION
- Added missing `GRIEVANCE_OUTCOMES` constant and `generateGrievanceId()` function
- Sheet tab colors added to all 11 sheet creation functions

### Added
- 79 new engagement tracking tests, 33 new grievance mutation tests
- Total: 950 tests across 18 suites

## [4.5.0] - 2026-02-01

### Added
- Security module — XSS prevention, HTML escaping, formula injection protection, PII masking, input validation, access control
- Data Access Layer — centralized sheet access with caching and deadline management
- Member Self-Service — PIN-based authentication with secure UUID generation and hashed storage
- Jest unit test suite with GAS environment mocking

### Changed
- Consolidated architecture from scattered modules to 16 focused source files
- CI/CD pipeline with GitHub Actions, ESLint v9, Husky pre-commit hooks

### Fixed
- 8 bug fixes including escapeForFormula, generateMemberPIN, TIME_CONSTANTS, and missing configs

### Removed
- 138+ duplicate function definitions and deprecated stubs

## [4.4.1] - 2026-01-31

### Added
- Initial build system with Node.js and source file concatenation

## [4.4.0] - 2026-01-30

### Added
- Member Dashboard with executive overview
- Steward Dashboard with case management
- Grievance tracking and management
- Satisfaction survey integration
- Calendar integration for deadlines
- Email notification system

### Security
- Input validation, HTML sanitization, role-based access control

## [4.3.8] - 2026-01-28

### Added
- Satisfaction modal dashboard with interactive charts
- Features Reference sheet — searchable catalog of all dashboard capabilities
- Hidden satisfaction calculation sheet for background formula processing

## [4.3.7] - 2026-01-25

### Added
- Complete rewrite of help guide with real-time search
- Menu reference tab with all menu items and descriptions
- FAQ tab with categorized questions and answers

## [4.3.2] - 2026-01-20

### Changed
- Deprecated visible Dashboard sheet — replaced with SPA-style modal dashboards
- All dashboard views now load as in-app dialogs instead of sheet tabs

## [4.3.0] - 2026-01-15

### Added
- Case Checklist system for grievance step tracking
- Looker data integration for external reporting
- Dynamic field expansion engine for flexible sheet schemas

## [4.1.0] - 2026-01-10

### Added
- Strategic Command Center configuration sheet
- Status color mapping for visual workflow indicators
- PDF and email branding templates

## [4.0.0] - 2026-01-05

### Added
- **Unified master engine** — single entry point for all dashboard operations
- Audit logging with tamper-evident chain
- Sabotage protection for critical sheets
- Batch processing for bulk member operations
- Mobile-responsive views for all web app pages

## [3.6.0] - 2025-12-20

### Changed
- Member and Grievance data manager refactor
- Improved validation across all data entry points

## [2.0.0] - 2025-11-15

### Changed
- **Modular architecture** — split monolith into modular source files
- Build system for concatenating source into deployable bundle
- Separation of UI and business logic layers

---

## Version History Summary

| Version | Date | Highlights |
|---------|------|------------|
| 4.9.0 | 2026-02-17 | Constant Contact v3 API integration, multi-select dropdowns, auto-discovery columns |
| 4.8.2 | 2026-02-16 | State field added to member contacts |
| 4.8.1 | 2026-02-15 | 5 new contact form fields, unified name-based Member IDs |
| 4.8.0 | 2026-02-15 | Security event alerting, zero-knowledge survey vault |
| 4.7.0 | 2026-02-14 | 40+ code review fixes, 1090 tests |
| 4.6.0 | 2026-02-12 | Meeting doc automation, steward agenda sharing, member Drive folders |
| 4.5.1 | 2026-02-11 | Engagement tracking fixes, 950 tests |
| 4.5.0 | 2026-02-01 | Security module, Data Access Layer, Member Self-Service |
| 4.4.1 | 2026-01-31 | Build system |
| 4.4.0 | 2026-01-30 | Member/Steward dashboards, grievance tracking, notifications |
| 4.3.8 | 2026-01-28 | Satisfaction modal dashboard, Features Reference sheet |
| 4.3.7 | 2026-01-25 | Help guide rewrite with real-time search |
| 4.3.2 | 2026-01-20 | SPA-style modal dashboards replace visible sheet tabs |
| 4.3.0 | 2026-01-15 | Case Checklist, Looker integration, dynamic field expansion |
| 4.1.0 | 2026-01-10 | Strategic Command Center, status colors, PDF/email branding |
| 4.0.0 | 2026-01-05 | Unified master engine, audit logging, batch processing |
| 3.6.0 | 2025-12-20 | Data manager refactor, improved validation |
| 2.0.0 | 2025-11-15 | Modular architecture, build system |
