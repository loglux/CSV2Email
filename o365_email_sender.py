import csv
from O365 import Account, FileSystemTokenBackend
import time
import random

class EmailSenderO365:
    def __init__(self, client_id, client_secret, token_file_path, token_filename, email_template_file, csv_file):
        self.client_id = client_id
        self.client_secret = client_secret
        self.account = self._setup_account(client_id, client_secret, token_file_path, token_filename)
        self.email_template_file = email_template_file
        self.csv_file = csv_file

    def _setup_account(self, client_id, client_secret, token_file_path, token_filename):
        credentials = (client_id, client_secret)
        token_backend = FileSystemTokenBackend(token_path=token_file_path, token_filename=token_filename)
        account = Account(credentials, token_backend=token_backend)
        if not account.is_authenticated:
            print("The account is not authenticated. Check your token.")
        return account

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
                    m = self.account.new_message()
                    m.to.add(email)
                    m.subject = subject
                    m.body = body.format(company_name=name)  # Assuming body template has placeholder for name
                    m.body_type = 'html'
                    try:
                        m.send()
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
    client_id = 'your_client_id_here'
    client_secret = 'your_client_secret_here'
    token_file_path = ''
    token_filename = 'o365_token.txt'
    email_template_file = 'email_template.txt'
    csv_file = 'companies_info.csv'

    sender = EmailSenderO365(client_id, client_secret, token_file_path, token_filename, email_template_file, csv_file)
    subject, body = sender.read_template()
    emails_to_send = sender.read_emails()

    sender.send_email(emails_to_send, subject, body, test_mode=False)  # Set test_mode to False to send real emails
