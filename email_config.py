"""
Email Template Configuration
Customize your interview notification email template here
"""

# Email Subject
EMAIL_SUBJECT = "Interview Scheduled - Action Required"

# Company Information
COMPANY_NAME = "Your Company Name"
HR_EMAIL = "hr@company.com"
HR_DEPARTMENT = "Recruitment Department"

# Email Template
# You can use these placeholders: {date}, {time}, {description}
EMAIL_TEMPLATE = """Dear Candidate,

We are pleased to inform you that your interview has been scheduled with our team at {company_name}.

══════════════════════════════════════════════════════════════
                    INTERVIEW DETAILS
══════════════════════════════════════════════════════════════

📅 Date:        {date}
⏰ Time:        {time}
📝 Round:       {description}

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
{hr_department}

══════════════════════════════════════════════════════════════
This is an automated notification. Please do not reply to this email.
For queries, contact: {hr_email}
══════════════════════════════════════════════════════════════
"""

# Alternative Simple Template
SIMPLE_EMAIL_TEMPLATE = """Dear Candidate,

Your interview has been scheduled:

Date: {date}
Time: {time}
Details: {description}

Please confirm your availability.

Thanks,
{company_name} HR Team
Contact: {hr_email}
"""
