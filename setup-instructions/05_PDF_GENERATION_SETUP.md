# PDF Generation & Email Setup Instructions

Configure the feature to create grievance PDFs and email them to members.

---

## How It Works

1. Go to **Grievance Log** sheet
2. Select a grievance row
3. Menu: **509 Command → Grievance Tools → Create PDF for Selected**
4. Confirm to create PDF
5. New prompt appears: "Would you like to email this PDF to [member@email.com]?"
6. Click **Yes** to send

---

## What Happens

| Step | Action |
|------|--------|
| 1 | PDF is created from your template |
| 2 | PDF is saved to member's personal Drive folder: `MemberName (MemberID)/` |
| 3 | Folder URL is saved to the grievance record |
| 4 | Optional: Email sent with PDF attached |

---

## Email Contents

The member receives:

- **Subject:** `[509 Strategic Command Center] Grievance Form - GRV-2025-001`
- **Body:** Greeting, grievance details summary, next steps (review, sign, return)
- **Attachment:** The filled-out PDF

---

## Requirements

| Config | Location | Purpose |
|--------|----------|---------|
| Template ID | AX2 | Google Doc template with placeholders |
| Archive Folder | AU2 | Parent folder for member folders |
| Member Email | Column X | Grievance Log - member's email address |

---

## Template Placeholders

Your Google Doc template can use these placeholders:

### Member Info
- `{{MemberName}}`
- `{{MemberID}}`

### Grievance Info
- `{{GrievanceID}}`
- `{{Date}}`
- `{{Status}}`
- `{{Articles}}`

### Location/Assignment Info
- `{{Unit}}`
- `{{Location}}`
- `{{Steward}}`
- `{{Details}}`

---

## Setup Steps

### 1. Create the Google Doc Template

1. Create a new Google Doc
2. Add placeholders where you want data inserted (e.g., `{{MemberName}}`)
3. Format the document as needed
4. Copy the Document ID from the URL

### 2. Create the Archive Folder

1. Create a folder in Google Drive for storing member PDFs
2. Copy the Folder ID from the URL

### 3. Configure in Config Sheet

1. Go to the **Config** sheet
2. In cell **AX2**, paste your Template Document ID
3. In cell **AU2**, paste your Archive Folder ID

### 4. Ensure Member Emails Exist

1. Go to the **Grievance Log** sheet
2. Verify Column X contains member email addresses
