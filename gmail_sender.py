import csv
import time
import random
import os.path
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


class EmailSenderGmail:
    def __init__(self, credentials_file, token_file, email_template_file, csv_file):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.creds = self._setup_credentials()
        self.email_template_file = email_template_file
        self.csv_file = csv_file
        self.service = build('gmail', 'v1', credentials=self.creds)

    def _setup_credentials(self):
        creds = None
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, ['https://www.googleapis.com/auth/gmail.send'])
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, ['https://www.googleapis.com/auth/gmail.send'])
                creds = flow.run_local_server(port=0)
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())
        return creds

    def read_template(self):
        with open(self.email_template_file, 'r', encoding='utf-8') as file:
            subject, body = file.read().split('BODY:', 1)
            subject = subject.strip().replace('SUBJECT:', '').strip()
            body = body.strip()
            return subject, body

    def read_emails(self):
        with open(self.csv_file, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            return [(row['name'], row['email']) for row in reader if row['email'].lower() not in ('n/a', '')]

    def send_email(self, emails, subject, body, test_mode=True, report_filename='email_report.txt'):
        with open(report_filename, 'w', encoding='utf-8') as report_file:
            for name, email in emails:
                if email.lower() in ('n/a', '', None):
                    continue
                if test_mode:
                    report_content = f"Test Mode: Email to {name} ({email}) with subject '{subject}' would be sent.\n"
                else:
                    # In real mode, send the email
                    message = MIMEMultipart()
                    message['to'] = email
                    message['subject'] = subject
                    message.attach(MIMEText(body.format(company_name=name), 'html'))

                    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
                    try:
                        self.service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
                        report_content = f"Email successfully sent to {name} ({email}).\n"
                    except Exception as e:
                        report_content = f"Failed to send email to {name} ({email}): {e}\n"
                    # Introduce a delay between sending each email
                    delay = random.uniform(1, 5)  # Random delay between 1 and 5 seconds
                    time.sleep(delay)
                print(report_content)
                report_file.write(report_content)
            total_emails = len([email for name, email in emails if email.lower() not in ('n/a', '', None)])
            report_file.write(f"Total emails processed: {total_emails}\n")
            print(f"Total emails processed: {total_emails}")


if __name__ == "__main__":
    credentials_file = 'credentials.json'
    token_file = 'token.json'
    email_template_file = 'email_template.txt'
    csv_file = 'companies_info.csv'

    sender = EmailSenderGmail(credentials_file, token_file, email_template_file, csv_file)
    subject, body = sender.read_template()
    emails_to_send = sender.read_emails()

    sender.send_email(emails_to_send, subject, body, test_mode=False)  # Set test_mode to False to send real emails
