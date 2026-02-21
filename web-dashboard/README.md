# Grievance Dashboard Web App

Mobile-optimized web app for workplace organization grievance tracking & member management. Served from Google Apps Script bound to a Google Sheet.

## File Inventory

### Server-side (`.gs` files → paste into Apps Script editor)
| File | Purpose | Lines |
|------|---------|-------|
| `WebApp.gs` | `doGet()` entry point, routing, page serving | ~120 |
| `Auth.gs` | SSO + magic link auth, session tokens, token cleanup | ~220 |
| `ConfigReader.gs` | Reads Config tab, caches settings, validation | ~120 |
| `DataService.gs` | Member Directory + Grievance Log data access layer | ~340 |

### Client-side (`.html` files → paste into Apps Script editor as HTML files)
| File | Purpose |
|------|---------|
| `index.html` | SPA shell, theme engine, router, bootstrap |
| `styles.html` | All CSS (included via `<?!= include('styles') ?>`) |
| `auth_view.html` | Login screen (SSO + magic link) |
| `steward_view.html` | Steward dashboard — Mono Signal theme |
| `member_view.html` | Member dashboard — Glass Depth theme |
| `error_view.html` | Error page (not found, expired, generic) |

### Documentation
| File | Purpose |
|------|---------|
| `AI_REFERENCE.md` | ⚠️ NEVER DELETE — LLM reference doc for all AI tools |
| `README.md` | This file |

## Config Tab Setup

Your Google Sheet needs a tab named **"Config"** with this layout:

| Column A (Setting Name) | Column B (Value) |
|--------------------------|------------------|
| Org Name | MassAbility DDS |
| Org Abbreviation | DDS |
| Logo Initials | M |
| Accent Hue | 250 |
| Magic Link Expiry | 7 |
| Cookie Duration | 30 |
| Steward Label | Steward |
| Member Label | Member |

**Accent Hue** is 0–360 on the color wheel:
- 0/360 = Red
- 30 = Orange  
- 60 = Yellow
- 120 = Green
- 160 = Teal
- 200 = Cyan
- 250 = Purple (default)
- 280 = Magenta
- 320 = Pink

## Member Directory Required Columns

The script finds columns by header name (case-insensitive). At minimum:
- **Email** (or "Email Address")
- **Role** (values: "Steward", "Member", or "Both")
- **Name** (or separate "First Name" / "Last Name")
- **Unit** (or "Workplace Unit", "Department")
- **Phone** (optional — shown to members only if present)

## Grievance Log Required Columns

- **Grievance ID** (or "Case ID", "ID")
- **Member Email** (links grievance to member)
- **Status** (New, Active, Overdue, Resolved)
- **Step** (Step 1, Step 2, Step 3)
- **Deadline** (date — auto-detects overdue)
- **Steward** (steward's email — links to assigned steward)
- **Unit** (optional)

## Deployment Steps

1. Open your Google Sheet → Extensions → Apps Script
2. Create each `.gs` file in the script editor (File → New → Script)
3. Create each `.html` file (File → New → HTML)
4. Copy contents from this repo into each file
5. Deploy → New Deployment → Web App
   - Execute as: **Me**
   - Who has access: **Anyone** (or "Anyone within [org]")
6. Copy the deployment URL
7. Update your Bitly short link to point to the new URL

## Push to GitHub (via Claude Code CLI)

```bash
cd /path/to/MULTIPLE-SCRIPS-REPO
git checkout -b web-dashboard
# Copy all files from this directory into the repo
git add .
git commit -m "feat: web dashboard Phase 1 - auth, routing, config, data service"
git push origin web-dashboard
```

## Architecture

```
User → Bitly → Apps Script doGet()
                    ├── Auth.resolveUser()
                    │    ├── SSO (Session.getActiveUser)
                    │    ├── Magic link token (URL param)
                    │    └── Session token (localStorage)
                    │
                    ├── DataService.findUserByEmail()
                    │    └── Role lookup → steward/member/both
                    │
                    └── Serve HTML
                         ├── auth.html (not logged in)
                         ├── steward (Mono Signal theme)
                         ├── member (Glass Depth theme)
                         └── error (not found / expired)
```
