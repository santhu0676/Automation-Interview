"""
Main Script - Interview Notification Scheduler
Orchestrates the entire process of sending interview notifications
"""
from excel_reader import ExcelReader
from email_sender import OutlookEmailer
from logger import EmailLogger
import sys


def main():
    """Main execution function"""
    
    print("\n" + "="*70)
    print("  INTERVIEW NOTIFICATION SCHEDULER")
    print("="*70 + "\n")
    
    # Initialize components
    excel_file = "interviews.xlsx"  # You can change this to your Excel file name
    
    logger = EmailLogger()
    logger.log_session_start()
    
    # Initialize Excel reader
    excel_reader = ExcelReader(excel_file)
    if not excel_reader.load_file():
        print("\n✗ Failed to load Excel file. Please check the file path.")
        print(f"  Expected file: {excel_file}")
        print("  TIP: Run 'python create_template.py' to create a template file,")
        print("       then rename it to 'interviews.xlsx'")
        logger.log_session_end(0, 0)
        return
    
    # Get pending interviews
    pending_interviews = excel_reader.get_pending_interviews()
    
    if not pending_interviews:
        print("\n✓ No pending interviews found.")
        print("  All interviews have already been sent or the file is empty.")
        excel_reader.close()
        logger.log_session_end(0, 0)
        return
    
    print(f"\n✓ Found {len(pending_interviews)} pending interview(s) to send\n")
    
    # Initialize Outlook emailer
    emailer = OutlookEmailer()
    if not emailer.connect():
        excel_reader.close()
        logger.log_session_end(0, 0)
        return
    
    print("\nStarting to send emails...\n")
    print("-" * 70)
    
    # Send emails
    sent_count = 0
    failed_count = 0
    
    for interview in pending_interviews:
        email_address = interview['email']
        
        # Send email
        if emailer.send_interview_notification(interview):
            # Mark as sent in Excel
            if excel_reader.mark_as_sent(interview['row_num']):
                logger.log_email_sent(email_address)
                sent_count += 1
            else:
                logger.log_email_failed(email_address, "Failed to update Excel")
                failed_count += 1
        else:
            logger.log_email_failed(email_address, "Failed to send email")
            failed_count += 1
    
    # Close Excel file
    excel_reader.close()
    
    # Print summary
    print("-" * 70)
    print("\n" + "="*70)
    print("  EXECUTION SUMMARY")
    print("="*70)
    print(f"  Total Emails Sent:     {sent_count}")
    print(f"  Total Failed:          {failed_count}")
    print(f"  Log File:              email_notifications.log")
    print("="*70 + "\n")
    
    logger.log_session_end(sent_count, failed_count)
    
    if sent_count > 0:
        print("✓ Check your Outlook 'Sent Items' folder to verify sent emails.\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n✗ Process interrupted by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Unexpected error: {str(e)}")
        sys.exit(1)
