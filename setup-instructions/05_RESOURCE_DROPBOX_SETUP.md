# Resource Dropbox (Member Document Submission Folder)

## Overview

The Resource Dropbox is a shared Google Drive folder where members can submit documents (forms, receipts, evidence, etc.) to the union. This is separate from the main Resource Drive folder (which is read-only for members).

---

## Instructions: Setting Up the Document Dropbox

### Step 1: Create the Google Drive Folder

1. Go to Google Drive
2. Click **"+ New"** > **"New folder"**
3. Name it something like "Member Submissions" or "Document Dropbox"
4. Click **Create**

### Step 2: Set Sharing Permissions

1. Right-click the folder > **"Share"**
2. Under **"General access"**, change to:
   - **"Anyone with the link"** with **"Editor"** permission (if members need to upload files)
   - Or add specific email addresses for controlled access
3. Click **Done**

**Note:** Unlike the Resource Drive folder (which uses Viewer access), the Dropbox folder typically needs **Editor** access so members can upload documents.

### Step 3: Get the Folder ID

1. Open the folder in Google Drive
2. Look at the URL in your browser:

```
https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz123456
```

3. Copy the ID portion (everything after `/folders/`):

```
1AbCdEfGhIjKlMnOpQrStUvWxYz123456
```

### Step 4: Add to Config Sheet

1. Open your Google Spreadsheet
2. Go to the **Config** sheet
3. Find the appropriate column for the Dropbox folder ID in the Integration section
4. Paste the folder ID into that cell

### Step 5: Verify It Works

1. Open a dashboard or member portal
2. Look for a document submission or upload option
3. Confirm that files can be uploaded to the folder

---

## Difference Between Resource Drive and Resource Dropbox

| Feature | Resource Drive (04) | Resource Dropbox (05) |
|---------|--------------------|-----------------------|
| **Purpose** | Share documents with members | Collect documents from members |
| **Access Level** | Viewer (read-only) | Editor (upload allowed) |
| **Content** | Bylaws, contracts, meeting minutes | Forms, receipts, evidence, submissions |
| **Direction** | Union > Members | Members > Union |

---

## What Members Might Submit

- Signed grievance forms
- Supporting evidence or documentation
- Contact information update forms
- Satisfaction survey attachments
- Training completion certificates
- Expense receipts for reimbursement
