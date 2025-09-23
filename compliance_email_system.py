#!/usr/bin/env python3
"""
Compliance Email Automation System
Automates task notifications and reminders from Excel compliance matrix
"""

import pandas as pd
import smtplib
import os
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import sys
from typing import List, Dict, Any
from dotenv import load_dotenv

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

load_dotenv()

class ComplianceEmailSystem:
    def __init__(self, excel_file_path: str):
        self.excel_file_path = excel_file_path
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '587'))
        self.smtp_username = os.getenv('SMTP_USERNAME')
        self.smtp_password = os.getenv('SMTP_PASSWORD')
        self.data = None
        
    def load_excel_data(self) -> bool:
        """Load and validate Excel data"""
        try:
            self.data = pd.read_excel(self.excel_file_path)
            
            # Validate required columns
            required_columns = ['Task', 'Task Description', 'Email', 'Attachment Link', 
                              'Status', 'Start Date', 'End Date', 'Frequency']
            missing_columns = [col for col in required_columns if col not in self.data.columns]
            
            if missing_columns:
                logger.error(f"Missing required columns: {missing_columns}")
                return False
                
            # Convert date columns
            self.data['Start Date'] = pd.to_datetime(self.data['Start Date']).dt.date
            self.data['End Date'] = pd.to_datetime(self.data['End Date']).dt.date
            
            logger.info(f"Successfully loaded {len(self.data)} tasks")
            return True
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            return False
    
    def filter_tasks_by_schedule(self, schedule_type: str) -> pd.DataFrame:
        """Filter tasks based on schedule type (monthly, quarterly, reminder)"""
        today = datetime.now().date()
        
        if schedule_type == "monthly":
            # Monthly tasks - send on 1st of month
            if today.day != 1:
                logger.info("Not the 1st of month - skipping monthly tasks")
                return pd.DataFrame()
            
            monthly_tasks = self.data[
                (self.data['Frequency'].str.lower() == 'monthly') &
                (self.data['Status'].str.lower() == 'pending')
            ]
            logger.info(f"Found {len(monthly_tasks)} monthly tasks")
            return monthly_tasks
            
        elif schedule_type == "quarterly":
            # Quarterly tasks - send on last day of quarter
            quarter_ends = [datetime(today.year, 3, 31).date(),
                          datetime(today.year, 6, 30).date(),
                          datetime(today.year, 9, 30).date(),
                          datetime(today.year, 12, 31).date()]
            
            if today not in quarter_ends:
                logger.info("Not a quarter end day - skipping quarterly tasks")
                return pd.DataFrame()
            
            quarterly_tasks = self.data[
                (self.data['Frequency'].str.lower() == 'quarterly') &
                (self.data['Status'].str.lower() == 'pending')
            ]
            logger.info(f"Found {len(quarterly_tasks)} quarterly tasks")
            return quarterly_tasks
            
        elif schedule_type == "reminder":
            # Weekly reminders - send every Monday for pending tasks
            if today.weekday() != 1:  # Monday is 0
                logger.info("Not Monday - skipping reminders")
                return pd.DataFrame()
            
            reminder_tasks = self.data[
                (self.data['Status'].str.lower() == 'pending') &
                (self.data['End Date'] >= today)  # Only remind for tasks not yet overdue
            ]
            logger.info(f"Found {len(reminder_tasks)} tasks for reminders")
            return reminder_tasks
            
        else:
            logger.error(f"Unknown schedule type: {schedule_type}")
            return pd.DataFrame()
    
    def create_email_content(self, user_tasks: pd.DataFrame, is_reminder: bool = False) -> str:
        """Create HTML email content for user tasks"""
        user_email = user_tasks['Email'].iloc[0]
        task_count = len(user_tasks)
        
        email_type = "Reminder" if is_reminder else "Notification"
        
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
                .header {{ background-color: #f8f9fa; padding: 20px; text-align: center; }}
                .task-table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                .task-table th, .task-table td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
                .task-table th {{ background-color: #4CAF50; color: white; }}
                .task-table tr:nth-child(even) {{ background-color: #f2f2f2; }}
                .urgent {{ color: #ff6b6b; font-weight: bold; }}
                .footer {{ margin-top: 20px; padding: 10px; background-color: #f8f9fa; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h2>Compliance Task {email_type}</h2>
                <p>You have {task_count} pending compliance task(s)</p>
            </div>
            
            <table class="task-table">
                <thead>
                    <tr>
                        <th>Task</th>
                        <th>Description</th>
                        <th>Deadline</th>
                        <th>Frequency</th>
                        <th>Attachment Link</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for _, task in user_tasks.iterrows():
            days_remaining = (task['End Date'] - datetime.now().date()).days
            urgent_class = "urgent" if days_remaining <= 3 else ""
            
            html_content += f"""
                    <tr>
                        <td><strong>{task['Task']}</strong></td>
                        <td>{task['Task Description']}</td>
                        <td class="{urgent_class}">{task['End Date']} ({days_remaining} days remaining)</td>
                        <td>{task['Frequency']}</td>
                        <td><a href="{task['Attachment Link']}">Upload Files</a></td>
                    </tr>
            """
        
        html_content += f"""
                </tbody>
            </table>
            
            <div class="footer">
                <p><strong>Action Required:</strong> Please complete these tasks by their respective deadlines.</p>
                <p><strong>Note:</strong> This is an automated message. Please do not reply to this email.</p>
            </div>
        </body>
        </html>
        """
        
        return html_content
    
    def send_email(self, to_email: str, subject: str, html_content: str) -> bool:
        """Send email via SMTP"""
        try:
            # Create message
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = self.smtp_username
            msg['To'] = to_email
            
            # Attach HTML content
            msg.attach(MIMEText(html_content, 'html'))
            
            # Send email
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.smtp_username, self.smtp_password)
                server.send_message(msg)
            
            logger.info(f"Email sent successfully to {to_email}")
            return True
            
        except Exception as e:
            logger.error(f"Error sending email to {to_email}: {e}")
            return False
    
    def process_tasks(self, schedule_type: str) -> Dict[str, Any]:
        """Main processing function"""
        if not self.load_excel_data():
            return {"success": False, "error": "Failed to load Excel data"}
        
        # Filter tasks based on schedule
        filtered_tasks = self.filter_tasks_by_schedule(schedule_type)
        
        if filtered_tasks.empty:
            logger.info(f"No tasks to process for {schedule_type}")
            return {"success": True, "emails_sent": 0, "message": "No tasks to process"}
        
        # Group tasks by email
        grouped_tasks = filtered_tasks.groupby('Email')
        emails_sent = 0
        emails_failed = 0
        
        is_reminder = (schedule_type == "reminder")
        email_subject = f"Compliance Task {'Reminder' if is_reminder else 'Notification'}"
        
        for email, tasks in grouped_tasks:
            # Create email content
            html_content = self.create_email_content(tasks, is_reminder)
            
            # Send email
            if self.send_email(email, email_subject, html_content):
                emails_sent += 1
            else:
                emails_failed += 1
        
        result = {
            "success": True,
            "emails_sent": emails_sent,
            "emails_failed": emails_failed,
            "total_tasks": len(filtered_tasks),
            "unique_users": len(grouped_tasks)
        }
        
        logger.info(f"Processing complete: {emails_sent} emails sent, {emails_failed} failed")
        return result

def main():
    if len(sys.argv) != 2:
        print("Usage: python compliance_email_system.py <schedule_type>")
        print("Schedule types: monthly, quarterly, reminder")
        sys.exit(1)
    
    schedule_type = sys.argv[1].lower()
    valid_schedules = ['monthly', 'quarterly', 'reminder']
    
    if schedule_type not in valid_schedules:
        print(f"Error: Schedule type must be one of {valid_schedules}")
        sys.exit(1)
    
    # Excel file path - can be configured via environment variable
    excel_file = os.getenv('EXCEL_FILE_PATH', 'compliance_matrix.xlsx')
    
    # Check if Excel file exists
    if not os.path.exists(excel_file):
        logger.error(f"Excel file not found: {excel_file}")
        sys.exit(1)
    
    # Initialize and run the system
    system = ComplianceEmailSystem(excel_file)
    result = system.process_tasks(schedule_type)
    
    if result['success']:
        print(f"✅ Successfully processed {schedule_type} tasks")
        print(f"📧 Emails sent: {result['emails_sent']}")
        print(f"❌ Emails failed: {result['emails_failed']}")
        print(f"📊 Total tasks: {result['total_tasks']}")
        print(f"👥 Unique users: {result['unique_users']}")
    else:
        print(f"❌ Processing failed: {result.get('error', 'Unknown error')}")
        sys.exit(1)

if __name__ == "__main__":
    main()

