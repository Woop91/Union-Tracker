# Changelog

All notable changes to the 509 Union Dashboard project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [4.5.0] - 2026-02-01

### Added
- Centralized error handling module (`00_ErrorHandler.gs`)
  - Error logging with severity levels
  - User-friendly error notifications
  - Performance monitoring utilities
  - Input sanitization functions
  - Constants validation on startup
  - API versioning support
- New modular architecture with 29 source files
- CI/CD pipeline with GitHub Actions
- ESLint configuration for code quality
- Husky pre-commit hooks
- JSDoc type checking support

### Changed
- Split `04_UIService.gs` into focused modules:
  - `04a_MenuBuilder.gs` - Menu creation
  - `04b_ThemeService.gs` - Theme management
  - `04c_MobileInterface.gs` - Mobile dashboard
  - `04d_QuickActions.gs` - Quick actions and email
- Split `06_Maintenance.gs` into focused modules:
  - `06a_Diagnostics.gs` - System diagnostics
  - `06b_CacheManager.gs` - Cache management
  - `06c_UndoManager.gs` - Undo/redo functionality
- Split `08_Code.gs` into focused modules:
  - `08a_SheetCreation.gs` - Sheet creation
  - `08b_DataValidation.gs` - Data validation
  - `08c_SearchEngine.gs` - Search functionality
  - `08d_ChartBuilder.gs` - Chart generation
  - `08e_FormHandlers.gs` - Form submission handlers
  - `08f_SatisfactionEngine.gs` - Satisfaction surveys
  - `08g_SyncEngine.gs` - Data synchronization

### Removed
- 138+ duplicate function definitions
- Deprecated functions:
  - `showPublicMemberDashboard_Code_DEPRECATED`
  - `refreshAllVisuals_DataRefresh_DEPRECATED`
  - `demoteSelectedSteward_UIService_DEPRECATED`

### Fixed
- ESLint configuration compatibility with v8
- Syntax errors in UI service module
- Build script path handling

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
| 4.5.0 | 2026-02-01 | Modular architecture, error handling, CI/CD |
| 4.4.1 | 2026-01-31 | Build system |
| 4.4.0 | 2026-01-30 | Initial release |
