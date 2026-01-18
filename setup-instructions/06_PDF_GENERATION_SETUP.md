PDF Generation

Done. Here's what was added:

New Feature: Email Grievance PDF to Member
How It Works
Go to Grievance Log sheet
Select a grievance row
Menu: 509 Command → Grievance Tools → Create PDF for Selected
Confirm to create PDF
New prompt appears: "Would you like to email this PDF to [member@email.com]?"
Click Yes to send
What Happens
Step	Action
1	PDF is created from your template
2	PDF is saved to member's personal Drive folder: MemberName (MemberID)/
3	Folder URL is saved to the grievance record
4	Optional: Email sent with PDF attached
Email Contents
The member receives:

Subject: [509 Strategic Command Center] Grievance Form - GRV-2025-001
Body: Greeting, grievance details summary, next steps (review, sign, return)
Attachment: The filled-out PDF
Requirements
Config	Location	Purpose
Template ID	AX2	Google Doc template with placeholders
Archive Folder	AU2	Parent folder for member folders
Member Email	Column X	Grievance Log - member's email address
Template Placeholders
Your Google Doc template can use:

{{MemberName}}, {{MemberID}}, {{GrievanceID}}
{{Date}}, {{Status}}, {{Articles}}
{{Unit}}, {{Location}}, {{Steward}}, {{Details}}
