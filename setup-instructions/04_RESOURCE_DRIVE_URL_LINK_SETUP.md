# Resource Drive URL Link

## Instructions: Setting Up the Shared Docs Folder

### Step 1: Create the Google Drive Folder

1. Go to Google Drive
2. Click **"+ New"** → **"New folder"**
3. Name it something like "Member Resources" or "Shared Documents"
4. Click **Create**

### Step 2: Set Sharing Permissions

1. Right-click the folder → **"Share"**
2. Under **"General access"**, change to:
   - "Anyone with the link" (if members just need view access)
   - Or add specific email addresses for restricted access
3. Set permission level to **"Viewer"** (recommended) or **"Editor"**
4. Click **Done**

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
3. Find **Column Y** (labeled "DRIVE_FOLDER_ID" in the Integration section)
4. Paste the folder ID into that cell

### Step 5: Verify It Works

1. Open a member dashboard
2. Look for the **"Resources"** button (green button with folder icon)
3. Click it - should open your shared folder

---

## What to Put in the Folder

Consider adding:

- Union bylaws/constitution
- Contract/CBA documents
- Meeting minutes
- Training materials
- Forms and templates
- Contact information sheets
