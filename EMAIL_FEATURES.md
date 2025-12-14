# 📧 Email Notification Features

## ✅ **Status Tracking (Already Implemented)**

The system automatically tracks which emails have been sent:

### How it Works:
1. **Status Column** - Excel file has a "Status" column
2. **Auto-Update** - After sending email, status is marked as "Sent"
3. **Duplicate Prevention** - Rows with "Sent" status are skipped
4. **Flexible** - Status column is optional, auto-detected

### Excel Example:
| Candidate Email | Interview Date | Interview Time | Description | Status |
|----------------|----------------|----------------|-------------|--------|
| john@test.com  | 2025-12-20     | 10:00 AM       | Tech Round  | **Sent** |
| jane@test.com  | 2025-12-21     | 2:00 PM        | HR Round    | *(empty)* |

✓ john@test.com - **Skipped** (already sent)  
✓ jane@test.com - **Will send** (status empty)

---

## 📝 **Professional Email Template**

### Email Details Sent:

**Subject:** Interview Scheduled - Action Required

**Body:**
```
Dear Candidate,

We are pleased to inform you that your interview has been scheduled with our team.

══════════════════════════════════════════════════════════════
                    INTERVIEW DETAILS
══════════════════════════════════════════════════════════════

📅 Date:        2025-12-20
⏰ Time:        10:00 AM
📝 Round:       Technical Interview - Python & System Design

══════════════════════════════════════════════════════════════

IMPORTANT INSTRUCTIONS:
------------------------
✓ Please join 5-10 minutes before the scheduled time
✓ Ensure you have a stable internet connection
✓ Keep your resume and relevant documents ready
✓ Prepare any questions you may have for us

If you need to reschedule or have any questions, please contact us immediately.

We look forward to speaking with you!

Best Regards,
HR Team
Recruitment Department

══════════════════════════════════════════════════════════════
This is an automated notification. Please do not reply to this email.
For queries, contact: hr@company.com
══════════════════════════════════════════════════════════════
```

---

## 🎨 **Customize Email Template**

Edit `email_config.py` to customize:

```python
# Change these values
EMAIL_SUBJECT = "Your Custom Subject"
COMPANY_NAME = "Your Company Name"
HR_EMAIL = "your-hr@company.com"

# Customize the email body template
EMAIL_TEMPLATE = """
Your custom template here...
Use {date}, {time}, {description} placeholders
"""
```

---

## 📊 **Status Tracking in Web App**

The Streamlit web interface shows:

1. **Quick Stats**
   - Total Interviews
   - Already Sent (counted from Status column)
   - Pending (will be sent)

2. **Color Coding**
   - Green rows = Already sent
   - White rows = Pending

3. **Results Dashboard**
   - ✅ Successfully Sent
   - ❌ Failed
   - ⏭️ Skipped (already sent)

---

## 🔍 **How to Track Emails**

### Method 1: Excel File
Open `interviews.xlsx` and check the **Status** column:
- Empty = Not sent yet
- "Sent" = Email sent successfully

### Method 2: Log File
Check `email_notifications.log` or `streamlit_email_notifications.log`:
```
2025-12-14 10:30:15 | INFO | Email: john@test.com | Status: Sent
2025-12-14 10:30:16 | INFO | Email: jane@test.com | Status: Sent
```

### Method 3: Outlook Sent Items
All sent emails appear in your Outlook "Sent Items" folder

---

## 🎯 **Key Features**

✅ **Auto Status Update** - Excel status column updated after each send  
✅ **Duplicate Prevention** - Never send to the same person twice  
✅ **Professional Template** - Formatted, clear, actionable email  
✅ **Customizable** - Edit email template in `email_config.py`  
✅ **Full Logging** - Every action logged with timestamp  
✅ **Error Tracking** - Failed sends are logged and reported  

---

## 💡 **Tips**

1. **To Resend an Email:**
   - Open Excel
   - Clear the "Sent" status for that row
   - Run the script again

2. **To Preview Email:**
   - Check `email_sender.py` for template
   - Or send a test email to yourself first

3. **To Change Template:**
   - Edit `email_config.py`
   - Restart the app

---

**All features are ready to use! Just run the app and upload your Excel file.**
