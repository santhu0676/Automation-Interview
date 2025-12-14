"""
Logger Module
Handles logging of email sending activities
"""
import logging
from datetime import datetime
import os


class EmailLogger:
    """Logs email sending activities to file"""
    
    def __init__(self, log_file: str = "email_notifications.log"):
        """
        Initialize logger
        
        Args:
            log_file: Path to the log file
        """
        self.log_file = log_file
        self.logger = self._setup_logger()
    
    def _setup_logger(self) -> logging.Logger:
        """
        Set up the logger with file and console handlers
        
        Returns:
            Configured logger instance
        """
        # Create logger
        logger = logging.getLogger("EmailNotificationLogger")
        logger.setLevel(logging.INFO)
        
        # Clear any existing handlers
        logger.handlers = []
        
        # Create file handler
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        
        # Create console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)s | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Add formatter to handlers
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # Add handlers to logger
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger
    
    def log_email_sent(self, email: str, status: str = "Sent"):
        """
        Log a successful email send
        
        Args:
            email: Email address
            status: Status of the email (default: "Sent")
        """
        self.logger.info(f"Email: {email} | Status: {status}")
    
    def log_email_failed(self, email: str, error: str = ""):
        """
        Log a failed email send
        
        Args:
            email: Email address
            error: Error message (optional)
        """
        error_msg = f" | Error: {error}" if error else ""
        self.logger.error(f"Email: {email} | Status: Failed{error_msg}")
    
    def log_session_start(self):
        """Log the start of a new session"""
        self.logger.info("=" * 70)
        self.logger.info("NEW SESSION STARTED")
        self.logger.info("=" * 70)
    
    def log_session_end(self, total_sent: int, total_failed: int):
        """
        Log the end of a session with summary
        
        Args:
            total_sent: Number of emails sent successfully
            total_failed: Number of emails that failed
        """
        self.logger.info("-" * 70)
        self.logger.info(f"SESSION SUMMARY: {total_sent} sent, {total_failed} failed")
        self.logger.info("=" * 70)
        self.logger.info("")  # Blank line for readability
