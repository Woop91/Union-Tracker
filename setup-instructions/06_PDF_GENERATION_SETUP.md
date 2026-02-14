# PDF Generation

Done. Here's what was added:

---

## New Feature: Email Grievance PDF to Member

### How It Works

1. Go to **Grievance Log** sheet
2. Select a grievance row
3. Menu: **Strategic Ops → Grievance Tools → Create PDF for Selected**
4. Confirm to create PDF
5. New prompt appears: "Would you like to email this PDF to [member@email.com]?"
6. Click **Yes** to send

---

### What Happens

| Step | Action |
|------|--------|
| 1 | PDF is created from your template |
| 2 | PDF is saved to member's personal Drive folder: `MemberName (MemberID)/` |
| 3 | Folder URL is saved to the grievance record |
| 4 | Optional: Email sent with PDF attached |

---

### Email Contents

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

Your Google Doc template can use:

| Placeholder | Description |
|-------------|-------------|
| `{{MemberName}}` | Member's full name |
| `{{MemberID}}` | Member's ID |
| `{{GrievanceID}}` | Grievance ID |
| `{{Date}}` | Date filed |
| `{{Status}}` | Current status |
| `{{Articles}}` | Contract articles cited |
| `{{Unit}}` | Member's unit |
| `{{Location}}` | Work location |
| `{{Steward}}` | Assigned steward |
| `{{Details}}` | Grievance details |
