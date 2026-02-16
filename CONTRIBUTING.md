# Contributing to Union Dashboard

Thank you for your interest in contributing to the Union Dashboard! This document provides guidelines and instructions for contributing.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [Project Structure](#project-structure)
- [Coding Standards](#coding-standards)
- [Making Changes](#making-changes)
- [Testing](#testing)
- [Submitting Changes](#submitting-changes)

## Code of Conduct

Please be respectful and constructive in all interactions. We're all working toward the same goal of supporting union members.

## Getting Started

### Prerequisites

- Node.js 18 or higher
- npm (comes with Node.js)
- Git
- Google Account with access to Google Apps Script

### Development Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/Woop91/MULTIPLE-SCRIPS-REPO.git
   cd MULTIPLE-SCRIPS-REPO
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Run linting**
   ```bash
   npm run lint
   ```

4. **Build the consolidated file**
   ```bash
   npm run build
   ```

## Project Structure

```
.
├── src/                    # Source files (27 .gs + 1 .html)
│   ├── 00_Security.gs              # Security utilities, XSS prevention
│   ├── 00_DataAccess.gs            # Data Access Layer
│   ├── 01_Core.gs                  # Error handling + Constants
│   ├── 02_DataManagers.gs          # Member + Grievance managers
│   ├── 03_UIComponents.gs          # Menu, Theme, Mobile
│   ├── 04a_UIMenus.gs              # Menu creation, dialogs, sidebar
│   ├── 04b_AccessibilityFeatures.gs # Comfort view, focus mode, import/export
│   ├── 04c_InteractiveDashboard.gs # Interactive dashboard, mobile views
│   ├── 04d_ExecutiveDashboard.gs   # Executive dashboard, alerts
│   ├── 04e_PublicDashboard.gs      # Public member dashboard
│   ├── 05_Integrations.gs          # Drive, Calendar, WebApp
│   ├── 06_Maintenance.gs           # Diagnostics, cache, undo
│   ├── 07_DevTools.gs              # Development tools (remove before prod)
│   ├── 08a_SheetSetup.gs           # Setup, utilities, data validation
│   ├── 08b_SearchAndCharts.gs      # Search functions, chart generation
│   ├── 08c_FormsAndNotifications.gs # Forms, notifications, alerts
│   ├── 08d_AuditAndFormulas.gs     # Audit log, formula sync
│   ├── 09_Dashboards.gs            # Dashboards and sync
│   ├── 10a_SheetCreation.gs        # Sheet creation functions
│   ├── 10b_SurveyDocSheets.gs      # Survey, FAQ, Getting Started sheets
│   ├── 10c_FormHandlers.gs         # Form submissions, flagged reviews
│   ├── 10d_SyncAndMaintenance.gs   # Data sync, formatting, maintenance
│   ├── 10_Main.gs                  # Main entry point, triggers
│   ├── 11_CommandHub.gs            # Command center
│   ├── 12_Features.gs              # Dynamic Engine, Looker Studio
│   ├── 13_MemberSelfService.gs     # PIN authentication portal
│   ├── 14_MeetingCheckIn.gs        # Meeting check-in system
│   └── MultiSelectDialog.html
├── test/                   # Jest unit tests
│   ├── gas-mock.js         # GAS environment mocks
│   ├── load-source.js      # Source file loader
│   └── *.test.js           # Test files (950+ tests)
├── dist/                   # Built output (auto-generated)
├── setup-instructions/     # Optional feature setup guides
├── .github/workflows/      # CI/CD configuration
├── build.js               # Build script
├── jest.config.js         # Jest configuration
├── package.json           # Node.js configuration
└── eslint.config.js       # ESLint configuration (v9 flat config)
```

### Module Naming Convention

| Prefix | Purpose |
|--------|---------|
| 00_ | Foundation (security, data access) |
| 01_ | Constants and configuration |
| 02_ | Data managers (member, grievance) |
| 03_ | UI components (menu, theme, mobile) |
| 04a-e_ | UI features (menus, accessibility, dashboards) |
| 05_ | External integrations (Drive, Calendar, WebApp) |
| 06_ | Maintenance and diagnostics |
| 07_ | Development tools (remove before production) |
| 08a-d_ | Sheet utilities (setup, search, forms, audit) |
| 09_ | Dashboard sync and satisfaction |
| 10_, 10a-d_ | Business logic, sheet creation, entry point |
| 11_ | Command center hub |
| 12_ | Feature extensions (Dynamic Engine, Looker) |
| 13_ | Member self-service (PIN auth) |
| 14_ | Meeting check-in system |

## Coding Standards

### JavaScript Style

- Use ES5 syntax (Google Apps Script requirement)
- Use `var` instead of `let`/`const`
- Use function declarations, not arrow functions
- 2-space indentation
- Single quotes for strings

### Naming Conventions

```javascript
// Functions: camelCase
function calculateMemberStats() { }

// Private functions: trailing underscore
function helperFunction_() { }

// Constants: UPPER_SNAKE_CASE
var MAX_RETRY_COUNT = 3;

// Global objects: PascalCase
var SHEETS = { };
```

### Documentation

Use JSDoc comments for all public functions:

```javascript
/**
 * Calculates member statistics for the dashboard
 * @param {string} memberId - The member ID to look up
 * @param {boolean} [includeHistory=false] - Include historical data
 * @returns {Object} Statistics object
 */
function calculateMemberStats(memberId, includeHistory) {
  // Implementation
}
```

### Error Handling

Use the centralized error handler:

```javascript
function riskyOperation() {
  try {
    // Your code here
  } catch (error) {
    handleError(error, 'riskyOperation', ERROR_LEVEL.ERROR);
    return null;
  }
}
```

## Making Changes

### Branch Naming

- `feature/description` - New features
- `fix/description` - Bug fixes
- `refactor/description` - Code refactoring
- `docs/description` - Documentation updates

### Commit Messages

Follow conventional commit format:

```
type(scope): brief description

Longer description if needed.

Closes #123
```

Types: `feat`, `fix`, `refactor`, `docs`, `test`, `chore`

Examples:
```
feat(members): add bulk import functionality
fix(grievance): correct deadline calculation
refactor(ui): split UIService into modules
docs(readme): update installation instructions
```

## Testing

### Running Tests

```bash
# Run unit tests only
npm run test:unit

# Run full test pipeline (lint + build + tests)
npm test
```

### Test Structure

Tests use Jest v29.7.0 with custom GAS environment mocking:

```
test/
├── gas-mock.js           # Mocks for SpreadsheetApp, Session, Utilities, etc.
├── load-source.js        # Preprocessor to load .gs files in Node.js
├── 00_Security.test.js   # Security function tests
├── 00_DataAccess.test.js # Data access and time constant tests
├── 01_Core.test.js       # Column constants, version, headers
├── 05_Integrations.test.js # Config and integration tests
├── 10_Main.test.js       # Row construction, CSV escaping
├── 13_MemberSelfService.test.js # PIN generation/hashing
└── modules.test.js       # Cross-module tests
```

### Writing Tests

```javascript
const { loadSources } = require('./load-source');

beforeAll(() => {
  require('./gas-mock');
  loadSources(['00_Security.gs', '01_Core.gs']);
});

test('escapeHtml escapes angle brackets', () => {
  expect(escapeHtml('<script>')).toBe('&lt;script&gt;');
});
```

## Submitting Changes

### Pull Request Process

1. **Create a feature branch**
   ```bash
   git checkout -b feature/your-feature
   ```

2. **Make your changes and commit**
   ```bash
   git add .
   git commit -m "feat(scope): description"
   ```

3. **Ensure tests pass**
   ```bash
   npm run lint
   npm run build
   npm run test:unit
   ```

4. **Push and create PR**
   ```bash
   git push origin feature/your-feature
   ```

5. **Fill out the PR template**
   - Describe what changes you made
   - Link any related issues
   - Include test plan

### PR Review Checklist

- [ ] Code follows style guidelines
- [ ] Tests added/updated
- [ ] Documentation updated
- [ ] No duplicate functions
- [ ] Build passes
- [ ] Lint passes

## Questions?

Open an issue with the `question` label or reach out to the maintainers.

---

Thank you for contributing!
