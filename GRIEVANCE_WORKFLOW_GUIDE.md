# 🚀 Grievance Workflow Feature Guide

## Overview

The Grievance Workflow feature allows stewards to start grievances directly from the Member Directory with pre-filled member and steward information via Google Forms. The system automatically adds submitted grievances to the Grievance Log, generates PDFs, and provides email/download options.

---

## ✨ Features

- **Quick Start from Member Directory**: Select a member row and run "Start New Grievance" to open a pre-filled form
- **Pre-filled Google Form**: Member and steward details automatically populated from Member Directory
- **Automatic Logging**: Form submissions automatically added to Grievance Log with unique ID
- **Drive Folder Creation**: Automatic folder structure created for each grievance (Documents, Correspondence, Notes)
- **Folder Sharing**: Folders automatically shared with Grievance Coordinators from Config
- **Deadline Calculation**: All step deadlines calculated automatically based on filing date
- **Member Directory Sync**: Member's grievance status updated automatically

---

## 📋 Setup Instructions

### Step 1: Configure Steward Contact Information

1. Go to the **⚙️ Config** tab
2. Navigate to column **U** (Steward Contact Information)
3. Enter:
   - **Row 2**: Steward Name
   - **Row 3**: Steward Email
   - **Row 4**: Steward Phone

Alternatively, use the menu: **509 Tools > ⚖️ Grievance Tools > ⚙️ Setup Steward Contact Info**

### Step 2: Create a Google Form (Optional but Recommended)

To enable the full workflow with automatic form submissions:

1. **Create a new Google Form:**
   - Go to [Google Forms](https://forms.google.com)
   - Click **+ Blank** or use a template
   - Name it "SEIU Local 509 - Grievance Form"

2. **Add the following fields** (adjust as needed):
   - Member ID (Short answer)
   - First Name (Short answer)
   - Last Name (Short answer)
   - Email (Short answer)
   - Phone (Short answer)
   - Job Title (Short answer)
   - Work Location (Dropdown or Short answer)
   - Unit (Dropdown)
   - Steward Name (Short answer)
   - Steward Email (Short answer)
   - Steward Phone (Short answer)
   - Incident Date (Date)
   - Grievance Type (Dropdown - use types from Config tab)
   - Description (Paragraph)
   - Desired Resolution (Paragraph)

3. **Get the Form URL:**
   - Click **Send** button in the form
   - Copy the link

4. **Get Field Entry IDs:**
   - Open your form in edit mode
   - Right-click on the first field > **Inspect** (or Inspect Element)
   - Look for `entry.XXXXXXXXXX` in the HTML
   - Record the entry ID for each field
   - Example: `entry.1234567890`

5. **Update the Configuration in Code:**
   - Open the Google Sheets Apps Script editor
   - Find the file **Code.gs** (or **ConsolidatedDashboard.gs** if deployed)
   - Update the `GRIEVANCE_FORM_CONFIG` object:

```javascript
const GRIEVANCE_FORM_CONFIG = {
  // Replace with your actual Google Form URL
  FORM_URL: "https://docs.google.com/forms/d/e/YOUR_FORM_ID/viewform",

  // Update these with your form's actual entry IDs
  FIELD_IDS: {
    MEMBER_ID: "entry.1234567890",           // Replace with actual
    MEMBER_FIRST_NAME: "entry.9876543210",   // Replace with actual
    MEMBER_LAST_NAME: "entry.1111111111",    // etc...
    MEMBER_EMAIL: "entry.2222222222",
    MEMBER_PHONE: "entry.3333333333",
    MEMBER_JOB_TITLE: "entry.4444444444",
    MEMBER_LOCATION: "entry.5555555555",
    MEMBER_UNIT: "entry.6666666666",
    STEWARD_NAME: "entry.7777777777",
    STEWARD_EMAIL: "entry.8888888888",
    STEWARD_PHONE: "entry.9999999999"
  }
};
```

### Step 3: Set Up Form Submission Trigger (For Automatic Processing)

**Option A: Use the Menu (Recommended)**

1. Go to **👤 Dashboard > 📋 Grievance Tools > 📋 Setup Form Trigger**
2. Enter the Google Form edit URL (the one ending in /edit) when prompted
3. Click OK - the trigger will be created automatically

**Option B: Manual Setup**

1. In the Apps Script editor, click the **Triggers** icon (⏰ clock icon)
2. Click **+ Add Trigger**
3. Configure:
   - **Function**: `onGrievanceFormSubmit`
   - **Event source**: From form
   - **Event type**: On form submit
4. Click **Save**

When forms are submitted:
- A new grievance row is added to the Grievance Log
- A Drive folder is automatically created with subfolders (Documents, Correspondence, Notes)
- Folder is shared with Grievance Coordinators from Config
- Deadlines are calculated automatically
- Member Directory is updated

---

## 🎯 How to Use

### Starting a New Grievance

1. **Open the Menu:**
   - Go to **509 Tools > ⚖️ Grievance Tools > 🚀 Start New Grievance**

2. **Select a Member:**
   - A dialog will appear with a searchable list of members
   - Select the member filing the grievance
   - Review the member details displayed

3. **Start the Grievance:**
   - Click **Start Grievance**
   - A pre-filled Google Form will open in a new window
   - Member and steward information will already be filled in

4. **Complete the Form:**
   - Fill in the grievance details:
     - Incident date
     - Grievance type
     - Description
     - Desired resolution
   - Click **Submit**

5. **Automatic Processing:**
   - The grievance is automatically added to the Grievance Log
   - A PDF is generated
   - You'll see options to email or download the PDF

### Email Options

After submitting a grievance, you can:

- **Email the PDF**: Enter one or more email addresses (comma-separated)
- **Download the PDF**: Save to your computer
- Both options available immediately after submission

---

## 📊 Grievance Log Integration

When a grievance is submitted through the form:

1. **Automatic Entry**: Added to the Grievance Log sheet
2. **Unique ID**: Generated automatically (format: GXXXX123 based on member name)
3. **Status**: Set to "Open" initially
4. **Drive Folder**: Created automatically with subfolders:
   - 📄 Documents
   - 📧 Correspondence
   - 📝 Notes
5. **Folder Sharing**: Shared with Grievance Coordinators from Config (column O)
6. **Calculations**: All deadline fields calculated automatically
7. **Member Update**: Member Directory updated with grievance status

---

## 🔧 Troubleshooting

### Form Pre-fill Not Working

**Problem**: Form fields are empty when opened
**Solution**:
- Check that field entry IDs are correct in `Code.gs`
- Verify steward contact info is entered in Config tab
- Ensure member has complete information in Member Directory

### Form Submission Not Adding to Log

**Problem**: Submitted forms don't appear in Grievance Log
**Solution**:
- Verify the form submission trigger is set up (see Step 3 above)
- Check that the form is linked to the correct spreadsheet
- Review the Apps Script logs for errors

### PDF Generation Fails

**Problem**: PDF won't generate or download
**Solution**:
- Ensure you have permission to access the spreadsheet
- Try refreshing the page
- Check browser pop-up blocker settings

### Email Not Sending

**Problem**: Emails not being sent
**Solution**:
- Verify Gmail authorization in Apps Script
- Check email addresses are valid
- Review daily email quota (Google has limits)

---

## 📚 Additional Resources

### Menu Locations

All grievance workflow features are in:
**509 Tools > ⚖️ Grievance Tools**

- **🚀 Start New Grievance**: Open member selection dialog
- **⚙️ Setup Steward Contact Info**: Configure steward details
- **📖 Help & Support**: Access documentation and tutorials

### Related Documentation

- [Main README](README.md) - Overall dashboard documentation
- [Interactive Dashboard Guide](INTERACTIVE_DASHBOARD_GUIDE.md) - Custom views
- [Comfort View Guide](COMFORT_VIEW_GUIDE.md) - Accessibility features
- [Steward Guide](STEWARD_GUIDE.md) - Positive reinforcement tips

---

## 💡 Tips for Stewards

1. **Keep Contact Info Updated**: Regularly check that steward contact info in Config is current
2. **Review Submissions**: Check the Grievance Log after form submissions to verify accuracy
3. **Use Dashboard & My Cases**: Open Dashboard via menu and check the My Cases tab to monitor your assigned grievances
4. **Email Multiple Recipients**: You can send PDFs to multiple people at once (comma-separated emails)
5. **Train Members**: Show members how to access the grievance form if they need to file directly

---

## 🎉 Benefits

- **Faster Filing**: Pre-filled forms save time
- **Fewer Errors**: Member info pulled directly from directory
- **Automatic Tracking**: No manual data entry required
- **Professional PDFs**: Consistent, professional grievance documents
- **Easy Distribution**: Email to multiple people instantly
- **Audit Trail**: All grievances automatically logged with timestamps

---

## ⚠️ Important Notes

- Only stewards should have access to the Member Directory to start grievances
- Always review member information before submitting
- Keep the Config tab protected to prevent unauthorized changes
- Regularly backup your spreadsheet
- PDF generation requires proper authorization - authorize when prompted

---

## 🆘 Support

If you encounter issues or have questions:

1. Check the [Troubleshooting](#-troubleshooting) section above
2. Review the Apps Script logs for error messages
3. Verify all setup steps were completed
4. Check that triggers are properly configured

For technical support with Google Apps Script, consult the [Google Apps Script Documentation](https://developers.google.com/apps-script).

---

**Last Updated**: 2025-11-23
**Version**: 1.0.0
