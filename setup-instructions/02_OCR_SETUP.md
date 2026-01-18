OCR

Here's how to enable the Google Cloud Vision API for OCR:

Enable Cloud Vision API
Step 1: Enable the API
Go to Google Cloud Console - Vision API
Select your Google Cloud project (or create one if needed)
Click "Enable"
Step 2: Create an API Key
Go to APIs & Services → Credentials
Click "+ CREATE CREDENTIALS" at the top
Select "API key"
Copy the generated key
Step 3: Configure in the Spreadsheet
Option A: Use the Setup Dialog

In your Google Sheets menu: Command Center → 🔧 OCR Setup
Paste your API key and click "Save Key"
Click "Test OCR" to verify it works
Option B: Manual Script Properties

In Apps Script editor: Project Settings → Script Properties
Add property: CLOUD_VISION_API_KEY = your-api-key
Cost
Free tier: 1,000 OCR requests/month at no cost
Most union use cases stay well within this limit
Security Recommendation
After creating the API key, consider restricting it:

Click on the API key in Credentials
Under "API restrictions", select "Restrict key"
Choose only "Cloud Vision API"
Under "Application restrictions", consider adding HTTP referrer restrictions
Once configured, you can use OCR via:

Menu: Command Center → 📝 OCR Transcribe Form
Code: wger(fileId, { mode: 'HANDWRITING', language: 'en' })
