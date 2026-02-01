# Contributing to 509 Union Dashboard

Thank you for your interest in contributing to the 509 Union Dashboard! This document provides guidelines and instructions for contributing.

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
├── src/                    # Source files
│   ├── 00_ErrorHandler.gs  # Error handling utilities
│   ├── 01_Constants.gs     # Global constants
│   ├── 02_MemberManager.gs # Member management
│   ├── 03_GrievanceManager.gs # Grievance management
│   ├── 04_*.gs            # UI modules
│   ├── 05_Integrations.gs # External integrations
│   ├── 06_*.gs            # Maintenance modules
│   ├── 07_DevTools.gs     # Development tools
│   ├── 08_*.gs            # Core code modules
│   ├── 09_Main.gs         # Main entry point
│   └── ...
├── dist/                   # Built output (auto-generated)
├── .github/workflows/      # CI/CD configuration
├── build.js               # Build script
├── package.json           # Node.js configuration
└── .eslintrc.js          # ESLint configuration
```

### Module Naming Convention

| Prefix | Purpose |
|--------|---------|
| 00_ | Utilities loaded first |
| 01_ | Constants and configuration |
| 02_-03_ | Data managers |
| 04_ | UI/presentation layer |
| 05_ | External integrations |
| 06_ | Maintenance and diagnostics |
| 07_ | Development tools |
| 08_ | Core business logic |
| 09_ | Main entry point |
| 10_+ | Feature modules |

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
# Run all tests (in Google Apps Script)
runAllTests();

# Run quick tests
runQuickTests();
```

### Writing Tests

Add tests to `15_TestFramework.gs`:

```javascript
function testYourFeature() {
  var result = yourFunction('input');

  if (result !== 'expected') {
    throw new Error('Test failed: expected "expected", got "' + result + '"');
  }

  Logger.log('testYourFeature: PASSED');
}
```

Register in the test registry:

```javascript
function getTestFunctionRegistry() {
  return [
    // ... existing tests
    'testYourFeature'
  ];
}
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
