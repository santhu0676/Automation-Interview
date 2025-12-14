"""
Outlook Email Sender Module
Handles sending emails via Outlook desktop application
"""
import win32com.client
from typing import Dict
try:
    from email_config import EMAIL_SUBJECT, EMAIL_TEMPLATE, COMPANY_NAME, HR_EMAIL, HR_DEPARTMENT
except ImportError:
    # Default values if config file doesn't exist
    EMAIL_SUBJECT = "Interview Scheduled"
    COMPANY_NAME = "Our Company"
    HR_EMAIL = "hr@company.com"
    HR_DEPARTMENT = "Recruitment Department"
    EMAIL_TEMPLATE = None


class OutlookEmailer:
    """Sends interview notification emails using Outlook"""
    
    def __init__(self):
        """Initialize Outlook connection"""
        self.outlook = None
        
    def connect(self) -> bool:
        """
        Connect to Outlook application
        
        Returns:
            True if connected successfully, False otherwise
        """
        try:
            # Try to connect to existing Outlook instance first
            try:
                self.outlook = win32com.client.GetActiveObject("Outlook.Application")
                print("✓ Connected to running Outlook instance")
            except:
                # If not running, start new instance
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                print("✓ Started new Outlook instance")
            
            # Test the connection by accessing namespace
            namespace = self.outlook.GetNamespace("MAPI")
            
            # Check if Outlook has any accounts configured
            accounts = namespace.Accounts
            if accounts.Count == 0:
                print("✗ No email accounts configured in Outlook!")
                print("  Please open Outlook and set up an email account first.")
                return False
            
            print(f"✓ Connected to Outlook with {accounts.Count} account(s)")
            return True
            
        except Exception as e:
            print(f"✗ Error connecting to Outlook: {str(e)}")
            print("\n📋 Troubleshooting Steps:")
            print("  1. Make sure Microsoft Outlook is installed (desktop version)")
            print("  2. Open Outlook and configure at least one email account")
            print("  3. Close Outlook and try running this script again")
            print("  4. If issue persists, run Outlook as Administrator once")
            return False
    
    def send_interview_notification(self, interview_data: Dict) -> bool:
        """
        Send interview notification email
        
        Args:
            interview_data: Dictionary containing email, date, time, and description
            
        Returns:
            True if email sent successfully, False otherwise
        """
        if not self.outlook:
            print("✗ Outlook not connected!")
            return False
        
        try:
            # Create email
            mail = self.outlook.CreateItem(0)  # 0 = MailItem
            
            # Set email properties
            mail.To = interview_data['email']
            mail.Subject = EMAIL_SUBJECT
            
            # Create email body
            mail.Body = self._create_email_body(interview_data)
            
            # Send email
            mail.Send()
            
            print(f"✓ Email sent to {interview_data['email']}")
            return True
            
        except Exception as e:
            print(f"✗ Error sending email to {interview_data['email']}: {str(e)}")
            return False
    
    def _create_email_body(self, interview_data: Dict) -> str:
        """
        Create email body text
        
        Args:
            interview_data: Dictionary containing interview details
            
        Returns:
            Formatted email body
        """
        # Use custom template if available
        if EMAIL_TEMPLATE:
            body = EMAIL_TEMPLATE.format(
                date=interview_data['date'],
                time=interview_data['time'],
                description=interview_data['description'],
                company_name=COMPANY_NAME,
                hr_email=HR_EMAIL,
                hr_department=HR_DEPARTMENT
            )
        else:
            # Default template
            body = f"""Dear Candidate,

We are pleased to inform you that your interview has been scheduled with our team.

══════════════════════════════════════════════════════════════
                    INTERVIEW DETAILS
══════════════════════════════════════════════════════════════

📅 Date:        {interview_data['date']}
⏰ Time:        {interview_data['time']}
📝 Round:       {interview_data['description']}

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
"""
        return body
