# Welcome to the 509 Strategic Command Center

**A Personal Project for Member Engagement, Grievance Tracking & Strategic Analytics**

---

> **Disclaimer:** This is a **personal project** I built on my own time—it is **not** an official union tool or endorsed by any labor organization. I'm sharing it freely with fellow stewards and union members who might find it useful.

> **Development Notice:** This system is currently in the **review and feedback phase** (approximately one month). I welcome all feedback, bug reports, and feature suggestions. Feel free to copy the sheet, forms, and core repository at any time to explore and customize for your own needs.

---

Hi there,

Thanks for checking out this project! I created this system as a personal side project to help with member engagement, grievance tracking, and data organization. It's designed to streamline communication, protect member information, and provide useful tools for stewards—but again, this is just something I built independently and decided to share.

**Important Note:** All member data currently shown in the forms and sheets has been **randomly generated** for demonstration purposes and is NOT based on real members or work locations.

---

## Quick Access Links

| Resource | Link | Description |
|----------|------|-------------|
| **Grievance Filing Form** | [Open Form](https://docs.google.com/forms/d/e/1FAIpQLSeOs6Kxqca85DYRF1wTP634gMNdEirZdi5mg7aUIY5q7dIfRg/viewform) | File a new grievance |
| **Satisfaction Survey** | [Open Survey](https://docs.google.com/forms/d/e/1FAIpQLSeR4VxrGTEvK-PaQP2S8JXn6xwTwp-vkR9tI5c3PRvfhr75nA/viewform) | Provide feedback on union services |
| **Contact Update Form** | [Open Form](https://docs.google.com/forms/d/e/1FAIpQLSeOs6Kxqca85DYRF1wTP634gMNdEirZdi5mg7aUIY5q7dIfRg/viewform) | Update your contact information |
| **Web App Dashboard** | [Open Dashboard](https://script.google.com/macros/s/AKfycbyAknT0zxuuSmCHVDg64uu3DmUJtk1h2uKoFX4mFLKWa_hoWXKjmFaHwbXK1XKkDcsBYQ/exec) | Full dashboard (desktop) |
| **Mobile Dashboard** | [Open Mobile](https://script.google.com/macros/s/AKfycbxaRlF89LEwb6gJrJtojqo51P2UyzZJSJgFAzDTqGNUvn5VMMdl/exec) | Touch-optimized for phones/tablets |
| **Member Dashboard** | Access via menu: `Union Hub → Dashboards → Member Dashboard` | View chapter data (no personal info) |
| **Steward Dashboard** | Access via menu: `Union Hub → Dashboards → Steward Dashboard` | Internal analytics (stewards only) |

---

## How Information Flows Through the System

Understanding how information moves through the system helps you know exactly what to expect:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         INFORMATION FLOW DIAGRAM                            │
└─────────────────────────────────────────────────────────────────────────────┘

    ┌──────────────┐         ┌──────────────────┐         ┌─────────────────┐
    │   MEMBERS    │  ───►   │  CONTACT FORM    │  ───►   │ MEMBER DIRECTORY│
    │ receive link │         │ update own info  │         │ data populated  │
    └──────────────┘         └──────────────────┘         └────────┬────────┘
                                                                   │
         ┌─────────────────────────────────────────────────────────┼─────────┐
         │                                                         ▼         │
         │  FROM THE MEMBER DIRECTORY, SEVERAL WORKFLOWS BEGIN:              │
         │                                                                   │
         │  ┌─────────────────┐    ┌─────────────────┐    ┌────────────────┐ │
         │  │ Send Survey     │    │ Send Dashboard  │    │ Start Grievance│ │
         │  │ Links to Members│    │ Link to Members │    │ (Stewards)     │ │
         │  └────────┬────────┘    └────────┬────────┘    └───────┬────────┘ │
         │           │                      │                     │          │
         │           ▼                      ▼                     ▼          │
         │  ┌─────────────────┐    ┌─────────────────┐    ┌────────────────┐ │
         │  │ SATISFACTION    │    │ MEMBER          │    │ GRIEVANCE FORM │ │
         │  │ SURVEY          │    │ DASHBOARD       │    │ Pre-populated  │ │
         │  │ Results tracked │    │ - Survey Results│    │ with member    │ │
         │  │ & analyzed      │    │ - Steward Info  │    │ information    │ │
         │  └─────────────────┘    │ - Union Rights  │    └───────┬────────┘ │
         │                         │ - Grievance Stats│           │          │
         │                         │ - And more!     │            │          │
         │                         └─────────────────┘            │          │
         └────────────────────────────────────────────────────────┼──────────┘
                                                                  │
                                                                  ▼
                              ┌────────────────────────────────────────────┐
                              │              GRIEVANCE LOG                 │
                              │  - Tracks all deadlines automatically     │
                              │  - Sends email reminders to stewards      │
                              │  - Creates Google Drive folder for case   │
                              │  - Links folder for evidence collection   │
                              │  - Provides traffic light status indicators│
                              └────────────────────────────────────────────┘
```

### The Flow in Detail:

1. **Members Receive Contact Form Link** → Members update their own contact information directly
2. **Information Populates Member Directory** → All data flows into our centralized member database
3. **From Member Directory, Stewards Can:**
   - Send satisfaction survey links to gather feedback
   - Send personalized dashboard links for members to access chapter information
   - **Start a grievance** by clicking a checkbox, which pre-populates member information into the grievance form
4. **Steward Completes Grievance Form** → Only needs to specify the complaint and resolution sought
5. **Grievance Log Auto-Populates** → System calculates deadlines, creates Drive folder, sends reminders
6. **Evidence Collection** → Members and stewards can upload documents to the linked Google Drive folder

---

## Member Directory Features

The Member Directory serves as a central hub for member information, organized into logical sections:

### Identity & Core Information
- **Member ID**: Auto-generated unique identifier (format: M + initials + 3 digits)
- **Name**: First and Last name fields
- **Job Title**: Dropdown selection from your organization's positions

### Work Location & Assignment
- **Work Location**: Office or site assignment
- **Unit**: Organizational unit designation
- **Office Days**: Which days the member works (multi-select)

### Contact Information
- **Email & Phone**: Primary contact methods
- **Preferred Communication**: Email, Phone, Text, or In Person (multi-select)
- **Best Time to Contact**: Customizable time preferences

### Organizational Structure
- **Supervisor & Manager**: Direct reporting chain
- **Is Steward**: Yes/No designation
- **Committees**: Committee memberships
- **Assigned Steward(s)**: Which steward(s) represent this member

### Engagement Tracking
- **Meeting Attendance**: Last virtual and in-person meetings
- **Email Open Rate**: Communication engagement metrics
- **Volunteer Hours**: Participation tracking

### Grievance Status (Auto-Calculated)
- **Has Open Grievance?**: Automatically updated Yes/No
- **Current Grievance Status**: Synced from Grievance Log
- **Next Deadline**: Upcoming action required date

### Quick Actions Column (One-Click Operations)
Click the checkbox to open a context menu with:
- Email Survey to Member
- Email Contact Update Form
- Email Dashboard Link
- Email Grievance Status
- Send Quick Email

---

## Grievance Log Features

The Grievance Log provides comprehensive case management with automatic deadline tracking:

### Case Identity
- **Grievance ID**: Auto-generated (format: G + initials + 3 digits)
- **Member ID**: Links back to Member Directory
- **Member Name**: Auto-populated from directory

### Status & Assignment
- **Status Options**: Open, Pending Info, Settled, Withdrawn, Denied, Won, Appealed, In Arbitration, Closed
- **Current Step**: Informal, Step I, Step II, Step III, Mediation, Arbitration

### Automatic Deadline Calculations
The system automatically calculates all contractual deadlines:

| Deadline | Calculation |
|----------|-------------|
| Filing Deadline | Incident Date + 21 days |
| Step I Decision Due | Date Filed + 30 days |
| Step II Appeal Due | Step I Decision + 10 days |
| Step II Decision Due | Step II Appeal + 14 days |
| Step III Appeal Due | Step II Decision + 30 days |

### Traffic Light Indicators
Visual status indicators are automatically applied:
- 🔴 **Red**: Overdue (0 days or past)
- 🟠 **Orange**: Urgent (1-3 days remaining)
- 🟢 **Green**: On Track (7+ days remaining)

### Case Details
- **Articles Violated**: Contract articles at issue
- **Issue Category**: Discipline, Workload, Scheduling, Pay, Benefits, Safety, Harassment, Discrimination, Contract Violation, Other

### Coordinator Notifications
- **Message Alert**: Checkbox to flag for coordinator attention
- **Coordinator Message**: Notes for leadership
- **Acknowledged By/Date**: Confirmation tracking

### Google Drive Integration
- **Automatic Folder Creation**: Each grievance gets its own Drive folder
- **Subfolders Created**: Step 1, Step 2, Step 3, Supporting Documents
- **Shared Access**: Stewards, coordinators, and members can upload evidence

### Action Types (Beyond Grievances)
Track various case types:
- Grievance
- Records Request
- Information Request
- Weingarten
- ULP (Unfair Labor Practice)
- EEOC/MCAD
- Accommodation Request
- Other Administrative

### Checklist Progress
Track completion of case requirements with visual progress indicators (e.g., "5/8" or "62%")

---

## Dashboard Features

### Member Dashboard (Public-Facing, No PII)

This dashboard is designed for sharing with all members while protecting personal information:

**Available Sections:**
- **Union Snapshot**: Aggregate membership statistics
- **Steward Directory**: Searchable list of stewards (contact info only)
- **Member Satisfaction**: Survey results across 8 categories (aggregated)
- **Weingarten Rights**: Emergency reference utility
- **Grievance Statistics**: Win rates, trends, category breakdowns

**Security Feature - Safety Valve PII Auto-Redaction:**
- Automatically masks phone numbers, SSNs, and other sensitive data
- Prevents accidental exposure of personal information

---

### Steward Dashboard (Internal, PII-Enabled)

This dashboard provides stewards and leadership with detailed analytics:

**Tab 1: Overview**
- Total Members, Active Grievances, Win Rate %
- Overdue Cases counter
- Response rates and aging case metrics
- 4 large metric cards with trend indicators

**Tab 2: Workload**
- Steward performance rankings
- Case distribution analysis
- Capacity indicators (flags stewards with 8+ active cases)
- Load balancing recommendations

**Tab 3: Analytics**
- Status breakdown by grievance type
- Trend analysis (monthly up/down indicators)
- Category analysis and patterns
- Historical forecasting

**Tab 4: Hot Spots**
- Locations with 3+ active grievances highlighted
- Visual heatmap of grievance activity by unit
- Strategic indicators for high-pressure areas

**Tab 5: Bargaining**
- Contract violation tracking by article
- Management denial rate analysis
- Bargaining cheat sheet with strategic data
- Precedent search for past outcomes

**Tab 6: Satisfaction**
- 8-section member satisfaction breakdown
- Unit-by-unit satisfaction comparisons
- Correlation analysis: satisfaction vs. grievance rates

---

### Mobile Dashboard

Touch-optimized for phones and tablets:
- Responsive design adapts to screen size
- All key metrics accessible on the go
- Quick access to member contact functions
- Field-ready for workplace visits

---

## Config Tab: Easy Customization

The Config sheet is designed so you can customize the entire system to your organization without touching any code:

### Employment Settings (Columns A-E)
| Column | Purpose |
|--------|---------|
| A | **Job Titles** - Add all positions in your organization |
| B | **Office Locations** - List all work sites |
| C | **Units** - Organizational unit designations |
| D | **Office Days** - Days of week options |
| E | **Yes/No Values** - Standard dropdown values |

### Personnel (Columns F-I)
| Column | Purpose |
|--------|---------|
| F | **Supervisors** - All supervisor names |
| G | **Managers** - All manager names |
| H | **Stewards** - Active steward list (auto-updates) |
| I | **Steward Committees** - Committee options |

### Grievance Configuration (Columns J-M)
| Column | Purpose |
|--------|---------|
| J | **Grievance Status** - Status dropdown options |
| K | **Grievance Steps** - Step progression options |
| L | **Issue Categories** - Types of grievances |
| M | **Articles** - Contract articles (customize to your contract) |

### Forms & Links (Columns N-Q)
| Column | Purpose |
|--------|---------|
| N | **Communication Methods** - Contact preferences |
| O | **Grievance Coordinators** - Emails for folder sharing |
| P | **Grievance Form URL** - Your Google Form link |
| Q | **Contact Form URL** - Member update form link |

### Deadline Settings (Columns AA-AD)
| Setting | Default | Your Value |
|---------|---------|------------|
| Filing Deadline Days | 21 | Customize to your contract |
| Step 1 Response Days | 30 | Customize to your contract |
| Step 2 Appeal Days | 10 | Customize to your contract |
| Step 2 Response Days | 14 | Customize to your contract |

### Strategic Command Center (Columns AS-AW)
| Column | Purpose |
|--------|---------|
| AS | **Chief Steward Email** - Escalation alerts go here |
| AT | **Unit Codes** - Format: "Main Station:MS,Field Ops:FO" |
| AU | **Archive Folder ID** - Google Drive folder for archiving |
| AV | **Escalation Statuses** - Trigger alerts on these |
| AW | **Escalation Steps** - Trigger alerts on these |

### Survey & Templates (Columns AR, AX-AY)
| Column | Purpose |
|--------|---------|
| AR | **Satisfaction Form URL** - Member survey link |
| AX | **Template ID** - Google Doc template for PDFs |
| AY | **PDF Folder ID** - Where to save generated PDFs |

---

## Additional Resources

### Function Checklist
The sheet includes a comprehensive function checklist showing:
- All available features
- Where to find each function
- What to expect from each feature
- Notes section for your feedback

### Feedback & Development Tab
Use this tab to:
- Report bugs or issues
- Request new features
- Track development progress
- Provide general feedback

### Need Help?

If you encounter any issues, have questions, or want to provide feedback:

1. **Check the Feedback & Development tab** in the sheet
2. **Email me directly**: wardis@pm.me

---

## These Are Just a Few of the Features!

The system includes many more capabilities:
- **PDF Generation Engine**: Create signature-ready grievance documents
- **Calendar Integration**: Sync deadlines to Google Calendar
- **Email Automation**: Deadline reminders, status updates, bulk communications
- **Audit Logging**: Track all system activity
- **Undo/Redo System**: 50-action history for peace of mind
- **Comfort View**: Accessibility features for visual comfort
- **Search Precedents**: Find past grievance outcomes for reference
- **Achievement System**: Track steward performance and milestones

Read the documentation sheets within the spreadsheet for detailed information on all features.

---

## Function Checklist

The sheet includes a **Function Checklist** tab that documents:
- All available features and where to find them
- What each function does and what to expect
- A notes section for your feedback on each feature

---

## Want to Copy This for Your Own Use?

This is a personal project that I'm sharing freely. You're welcome to:

1. **Copy the Google Sheet**: File → Make a copy
2. **Copy the Google Forms**: Each form can be copied from the form editor
3. **Clone the GitHub Repository**: Contains all the Apps Script code

Feel free to customize everything to match your contract deadlines, terminology, and workflows. No affiliation with any official union organization is implied.

---

## Feedback & Development Phase

I'm currently in a **month-long review and feedback phase**. During this time:

- Test all features thoroughly
- Report any bugs or issues in the **Feedback & Development tab**
- Suggest new features or improvements
- Share your experience with others who might find it useful

Your feedback directly shapes how I continue developing this project!

---

**Thanks for checking this out!**

*Best,*
*Wardis*

---

*Version 4.3.8 | Questions or Feedback? Contact wardis@pm.me*

*This is a personal project and is not affiliated with or endorsed by any union or labor organization.*
