# 📋 Phase 2 Development Plan — DDS Dashboard Web App
# Last updated: 2026-02-24

---

## 🅿️ PARKED FEATURES (to be built later, ranked by priority)

1. **Bulk Actions** — Checkbox selection on grievance + member lists → flag, email, CSV export
2. **Member List Enhancements** — Bulk email, bulk flag, contact cards
3. **Deadline Calendar View** — `?page=deadlines` route, month/week/list views for stewards
4. **Grievance History for Members** — Resolved/closed cases in self-service portal
5. **PIN Auth Enhancements** — Already implemented; future: biometric, remember device

---

## 🎯 CURRENT SPRINT — UX Overhaul + Feature Gaps

### Problems Identified
1. **Not welcoming** — Drops users into a dense KPI dashboard with no greeting, no context
2. **Not educational** — No "Know Your Rights", no contract explainers, no onboarding flow
3. **Not interactive enough** — Static data display; no actions members can take
4. **Meeting check-in missing from web** — Only works via Google Sheets UI dialog
5. **No document tab integration** — Can't browse/view Config, Member Directory, Grievance Log structure
6. **Generic styling** — Purple gradients, system fonts, identical to every other AI-built dashboard

### Feature Plan (in build order)

#### 1. 🏠 Welcome Experience & Onboarding
- Replace cold KPI dump with personalized greeting
- First-time user onboarding flow (what is this app, what can you do)
- Role-appropriate welcome: steward sees action items, member sees their status
- Quick-start guide cards

#### 2. 📖 Educational Content Hub
- "Know Your Rights" tab with contract article summaries
- Grievance process explainer (interactive step-by-step)
- FAQ with searchable answers
- "What to do if..." scenarios
- All content pulled from Config tab (dynamic, not hardcoded)

#### 3. 📝 Meeting Check-In Web Route
- New `?page=checkin` route in doGet()
- Reuse existing check-in logic from 14_MeetingCheckIn.gs
- Member enters email + PIN → checks in
- Steward can view attendance in real-time
- Mobile-optimized kiosk mode

#### 4. 📂 Document Browser / Sheet Viewer
- `?page=config` — Read-only Config tab viewer (steward only)
- `?page=directory` — Member Directory browser with search (steward only)
- `?page=grievancelog` — Grievance Log browser with filters (steward only)
- Show column headers, data types, validation rules
- Purpose: transparency + admin reference without opening the spreadsheet

#### 5. 🎨 Visual Identity Refresh
- Replace generic purple gradient with dynamic accent color from Config
- Use distinctive typography (not system fonts)
- Add micro-interactions (card hover, transitions, loading states)
- Warm, approachable feel vs cold data dashboard
- Dark/light mode with proper theming

#### 6. 🔗 Deeper Feature Integration
- Quick actions from dashboard: file grievance, contact steward, check in to meeting
- Notification center: deadline alerts, meeting reminders, status changes
- Steward quick-compose: email member from dashboard
- Member feedback: satisfaction survey link, anonymous feedback
