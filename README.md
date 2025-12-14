# 📧 Automated Interview Notification Scheduler - MVP

A Python script that reads interview details from an Excel file and sends notification emails via Outlook.

## ✨ Features

- ✅ Read interview details from Excel file
- ✅ Send plain-text emails via Outlook desktop application
- ✅ Log all email activities (sent/failed) with timestamps
- ✅ Duplicate prevention (skips already sent interviews)
- ✅ Manual execution (no auto-scheduling)

---

## 📋 Prerequisites

Before running this project, ensure you have:

1. **Python 3.7 or higher** installed
   - Check: `python --version`
   - Download from: https://www.python.org/downloads/

2. **Microsoft Outlook** installed and configured
   - Must be the desktop application (not web version)
   - Must have at least one email account configured

3. **Windows Operating System**
   - Required for Outlook COM integration

---

## 🚀 Installation Steps

### Step 1: Install Dependencies

Open PowerShell or Command Prompt in the project folder and run:

```bash
pip install -r requirements.txt
```

This will install:
- `openpyxl` - For reading/writing Excel files
- `pywin32` - For Outlook integration

### Step 2: Create Excel Template

Run the template creation script:

```bash
python create_template.py
```

This creates `template_interviews.xlsx` with sample data.

### Step 3: Prepare Your Interview Data

1. Rename `template_interviews.xlsx` to `interviews.xlsx`
2. Open `interviews.xlsx` and replace sample data with real interview details
3. **Required columns:**
   - **Candidate Email** - Email address of the candidate
   - **Interview Date** - Date of interview (e.g., 2025-12-20)
   - **Interview Time** - Time of interview (e.g., 10:00 AM)
   - **Interview Description** - Interview details
   - **Status** - Leave blank (script will mark as "Sent")

**Example:**

| Candidate Email | Interview Date | Interview Time | Interview Description | Status |
|----------------|----------------|----------------|----------------------|--------|
| john@example.com | 2025-12-20 | 10:00 AM | Technical Round - Python | |
| jane@example.com | 2025-12-21 | 2:00 PM | HR Round | |

---

## ▶️ How to Run

### Method 1: Web Interface (Streamlit) - **RECOMMENDED** ✨

```bash
streamlit run app.py
```

This will open a browser with a user-friendly interface where you can:
- 📤 Upload Excel files via drag & drop
- 👀 Preview interview data in a table
- 📊 See stats (total, sent, pending)
- ✉️ Send emails with one click
- 📥 Download updated Excel file

### Method 2: Command Line

```bash
python main.py
```

### Method 3: Double-click (Windows)

Simply double-click `main.py` in File Explorer (if `.py` files are associated with Python)

---

## 📊 What Happens When You Run

1. **Loads Excel file** (`interviews.xlsx`)
2. **Finds pending interviews** (rows without "Sent" status)
3. **Connects to Outlook**
4. **Sends emails** to each candidate
5. **Marks rows as "Sent"** in Excel
6. **Logs everything** to `email_notifications.log`

### Expected Output:

```
======================================================================
  INTERVIEW NOTIFICATION SCHEDULER
======================================================================

✓ Excel file loaded: interviews.xlsx
✓ Connected to Outlook

✓ Found 3 pending interview(s) to send

Starting to send emails...

----------------------------------------------------------------------
✓ Email sent to john@example.com
✓ Email sent to jane@example.com
✓ Email sent to alex@example.com
----------------------------------------------------------------------

======================================================================
  EXECUTION SUMMARY
======================================================================
  Total Emails Sent:     3
  Total Failed:          0
  Log File:              email_notifications.log
======================================================================

✓ Check your Outlook 'Sent Items' folder to verify sent emails.
```

---

## 📝 Email Format

Each email is sent with:

**Subject:** Interview Scheduled

**Body:**
```
Dear Candidate,

We are pleased to inform you that your interview has been scheduled.

Interview Details:
-------------------
Date: 2025-12-20
Time: 10:00 AM
Description: Technical Round - Python

Please be available at the scheduled time. If you have any questions 
or need to reschedule, please contact us as soon as possible.

We look forward to meeting you!

Best regards,
HR Team
```

---

## 📂 Project Structure

```
Automation Interview/
│
├── app.py                     # 🌐 Streamlit Web Interface (NEW!)
├── main.py                    # Main execution script (CLI)
├── excel_reader.py            # Excel file handling
├── email_sender.py            # Outlook email sending
├── logger.py                  # Logging functionality
├── create_template.py         # Template creation script
├── requirements.txt           # Python dependencies
├── interviews.xlsx            # Your interview data (create this)
├── email_notifications.log    # Generated log file (CLI)
├── streamlit_email_notifications.log  # Generated log file (Web)
└── README.md                  # This file
```

---

## 🔍 Troubleshooting

### Issue: "Error connecting to Outlook"

**Solutions:**
- Make sure Outlook desktop app is installed (not just web version)
- Open Outlook at least once and configure an email account
- Run the script with administrator privileges if needed

### Issue: "File 'interviews.xlsx' not found"

**Solutions:**
- Make sure you renamed `template_interviews.xlsx` to `interviews.xlsx`
- Check that the file is in the same folder as `main.py`

### Issue: "Error loading Excel file"

**Solutions:**
- Make sure the Excel file is not open in Excel while running the script
- Check that the file has all required columns
- Verify file is not corrupted

### Issue: Emails not sending

**Solutions:**
- Check your internet connection
- Verify Outlook is configured with a valid email account
- Check if Outlook requires you to allow programmatic access
- Look in `email_notifications.log` for specific error messages

---

## 🔒 Duplicate Prevention

The script automatically skips interviews already marked as "Sent":

- After each successful email, the **Status** column is updated to "Sent"
- On next run, these rows are automatically skipped
- This prevents sending duplicate emails to the same candidate

To **resend** an email:
1. Open `interviews.xlsx`
2. Clear the "Sent" status for that row
3. Run the script again

---

## 📊 Logging

All activities are logged to `email_notifications.log`:

**Log format:**
```
2025-12-14 10:30:15 | INFO | NEW SESSION STARTED
2025-12-14 10:30:16 | INFO | Email: john@example.com | Status: Sent
2025-12-14 10:30:17 | INFO | Email: jane@example.com | Status: Sent
2025-12-14 10:30:18 | ERROR | Email: invalid@email | Status: Failed
2025-12-14 10:30:19 | INFO | SESSION SUMMARY: 2 sent, 1 failed
```

---

## ⚙️ Configuration

You can modify these settings in `main.py`:

```python
# Change Excel file name (line 12)
excel_file = "interviews.xlsx"  # Change to your file name

# Change log file name (logger.py, line 12)
log_file = "email_notifications.log"
```

---

## 🎯 MVP Completion Status

✅ **All MVP features implemented:**

1. ✅ Read Excel File - Loads interview details from Excel
2. ✅ Manual Trigger - Runs only when executed
3. ✅ Outlook Integration - Connects to Outlook desktop
4. ✅ Send Email Notification - Sends plain-text emails
5. ✅ Logging - Records all activities with timestamps
6. ✅ Duplicate Prevention - Skips already sent interviews

---

## 📞 Support

If you encounter any issues:

1. Check `email_notifications.log` for error details
2. Verify all prerequisites are met
3. Review troubleshooting section above

---

## 📜 License

This is an MVP project for internal use.

---

**Built with ❤️ for automated interview scheduling**
#   A u t o m a t i o n - I n t e r v i e w 
 
 
