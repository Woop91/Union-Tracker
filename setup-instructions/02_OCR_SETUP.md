# OCR Setup Instructions

How to enable the Google Cloud Vision API for OCR functionality.

---

## Step 1: Enable the API

1. Go to [Google Cloud Console - Vision API](https://console.cloud.google.com/apis/library/vision.googleapis.com)
2. Select your Google Cloud project (or create one if needed)
3. Click **"Enable"**

---

## Step 2: Create an API Key

1. Go to **APIs & Services → Credentials**
2. Click **"+ CREATE CREDENTIALS"** at the top
3. Select **"API key"**
4. Copy the generated key

---

## Step 3: Configure in the Spreadsheet

### Option A: Use the Setup Dialog (Recommended)

1. In your Google Sheets menu: **Command Center → 🔧 OCR Setup**
2. Paste your API key and click **"Save Key"**
3. Click **"Test OCR"** to verify it works

### Option B: Manual Script Properties

1. In Apps Script editor: **Project Settings → Script Properties**
2. Add property: `CLOUD_VISION_API_KEY` = `your-api-key`

---

## Cost Information

| Tier | Requests/Month | Cost |
|------|----------------|------|
| Free tier | 1,000 OCR requests | $0 |
| Paid tier | Additional requests | See Google pricing |

Most union use cases stay well within the free tier limit.

---

## Security Recommendation

After creating the API key, consider restricting it:

1. Click on the API key in Credentials
2. Under **"API restrictions"**, select **"Restrict key"**
3. Choose only **"Cloud Vision API"**
4. Under **"Application restrictions"**, consider adding HTTP referrer restrictions

---

## Usage

Once configured, you can use OCR via:

- **Menu:** Command Center → 📝 OCR Transcribe Form
- **Code:** `wger(fileId, { mode: 'HANDWRITING', language: 'en' })`
