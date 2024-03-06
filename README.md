# Office 365 Email Automation

## Project Overview
This script automates the process of sending emails through the Office 365 API using a Python script (`o365_email_sender.py`). It reads recipient details from a CSV file (`companies_info.csv`) and uses an HTML email template (`email_template.txt`) to craft content. It is aimed at automating email outreach and is a robust tool for businesses and individuals.

## Prerequisites
- Python 3.x
- An Office 365 account with API access enabled
- The O365 Python package

## Setup and Installation
1. Clone this repository.
2. Install the required dependencies:
   - Directly via pip: `pip install O365`
   - Or from the `requirements.txt` file: `pip install -r requirements.txt`
3. Insert your Office 365 API **client_id** and **client_secret** into the **o365_email_sender.py**.
4. Ensure the `o365_token.txt` is present for storing authentication tokens.
5. Update **email_template.txt** and **companies_info.csv** with your specific email content and recipient information.

## Usage
- Populate the CSV file with target recipient details, including columns for name, email, etc.
- Customise the email template as necessary.
- Run the script with the command: `python o365_email_sender.py`

## Features
- Secure authentication with the Office 365 API.
- Dynamic email content using a customisable HTML template.
- Recipient management via CSV.
- Test mode for validation.
- Automated email sending with adjustable delays to prevent rate limiting.
- Generates a report (`email_report.txt`) detailing the outcomes of the email dispatch process.

## Example
```python
from o365_email_sender import EmailSenderO365

sender = EmailSenderO365(client_id, client_secret, token_file_path, token_filename, email_template_file, csv_file)
subject, body = sender.read_template()
emails_to_send = sender.read_emails()
sender.send_email(emails_to_send, subject, body, test_mode=False)
```
## Contributing
Contributions to improve or extend the functionality are highly encouraged. Please fork the repository and submit pull requests for any proposed changes or fixes.

## License
This project is distributed under the MIT License. Refer to the LICENSE file for more information.