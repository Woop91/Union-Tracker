# Changelog

All notable changes to the 509 Union Dashboard project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [4.5.1] - 2026-02-10

### Fixed
- `GRIEVANCE_OUTCOMES` - added missing constant that caused "GRIEVANCE_OUTCOMES is not defined" runtime error, preventing tabs from populating
- `generateGrievanceId()` - added missing function called by `getNextGrievanceId()`, which caused `startNewGrievance()` to silently fail
- Sheet tab colors - added `.setTabColor()` to all 11 sheet creation functions (Config, Config Guide, Member Directory, Grievance Log, Dashboard, Satisfaction, Feedback, Function Checklist, Getting Started, FAQ, Case Checklist)
- Hardcoded hex colors replaced with `COLORS` constants in Getting Started, FAQ, and Config Guide sheet creation

### Added
- 33 new tests (871 total across 17 suites) covering core grievance mutation paths:
  - `startNewGrievance()` - validates row data includes `GRIEVANCE_OUTCOMES.PENDING`, audit logging, error handling
  - `resolveGrievance()` - validates outcome/status updates, audit logging, timestamped notes
  - `advanceGrievanceStep()` - validates step 1→2→3→arbitration transitions, status updates, boundary errors
  - `setupCalcFormulasSheet()` - validates `GRIEVANCE_OUTCOMES` and `GRIEVANCE_STATUS` are written to formula sheet
  - `GRIEVANCE_OUTCOMES` constant existence guard - asserts all expected keys are defined

## [4.5.0] - 2026-02-01

### Added
- Security module (`00_Security.gs`) - XSS prevention, HTML escaping, formula injection protection, PII masking, input validation, access control
- Data Access Layer (`00_DataAccess.gs`) - Centralized sheet access with caching, TIME_CONSTANTS for deadline management, deadline urgency calculations
- Member Self-Service (`13_MemberSelfService.gs`) - PIN-based member authentication with secure UUID-based PIN generation and hashed storage
- Jest unit test suite with GAS environment mocking
- jest.config.js and test infrastructure (gas-mock.js, load-source.js)

### Changed
- Consolidated architecture from scattered modules to 16 focused source files
- CI/CD pipeline with GitHub Actions
- ESLint v9 flat config for code quality
- Husky pre-commit hooks
- Version bumped to 4.5.0 across all files (VERSION_INFO, COMMAND_CONFIG, API_VERSION, package.json)

### Fixed
- `escapeForFormula()` - was prefixing formula chars mid-string, now only at start
- `generateMemberPIN()` - replaced insecure Math.random with Utilities.getUuid()
- `TIME_CONSTANTS.DEADLINE_DAYS` - corrected values to match DEADLINE_RULES (7/7/14/10)
- `API_VERSION` - synced with VERSION_INFO (4.5.0)
- `CALENDAR_CONFIG` - re-added missing config to Integrations module
- `MEMBER_COLUMNS.PIN_HASH` - added missing column constant and header
- `initializeDashboard()` - fixed dead stub, now delegates to CREATE_509_DASHBOARD()
- `COMMAND_CONFIG.VERSION` - synced with VERSION_INFO.CURRENT

### Removed
- 138+ duplicate function definitions
- Deprecated function stubs

## [4.4.1] - 2026-01-31

### Added
- Initial build system with Node.js
- Source file concatenation for deployment
- Basic project structure

## [4.4.0] - 2026-01-30

### Added
- Member Dashboard with executive overview
- Steward Dashboard with case management
- Grievance tracking and management
- Satisfaction survey integration
- Calendar integration for deadlines
- Email notification system

### Security
- Input validation for all user inputs
- HTML sanitization for XSS prevention
- Role-based access control

---

## Version History Summary

| Version | Date | Highlights |
|---------|------|------------|
| 4.5.1 | 2026-02-10 | Fixed GRIEVANCE_OUTCOMES/generateGrievanceId bugs, sheet tab colors, 871 Jest tests |
| 4.5.0 | 2026-02-01 | Security module, Data Access Layer, Member Self-Service, consolidated to 16 source files |
| 4.4.1 | 2026-01-31 | Build system |
| 4.4.0 | 2026-01-30 | Initial release |
