from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError
import base64
from email.mime.text import MIMEText

# Service account key path
SERVICE_ACCOUNT_FILE = 'path/to/your/service_account.json'

# Define the scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/gmail.send']

# Authenticate and create service clients
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service_sheets = build('sheets', 'v4', credentials=creds)
service_gmail = build('gmail', 'v1', credentials=creds)

# Spreadsheet and Form details
SPREADSHEET_ID = 'your_spreadsheet_id'
FORM_URL = 'your_google_form_url'

def get_students_to_email():
    try:
        # Specify the range containing the student data
        range_name = 'Sheet1!A2:D'  # Update with your sheet's name and range
        sheet = service_sheets.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
        values = result.get('values', [])

        students_to_email = []
        for row in values:
            # Check if the email has been sent
            if len(row) < 3 or row[2].lower() != 'yes':
                students_to_email.append({
                    'name': row[0],
                    'email': row[1],
                    'row': values.index(row) + 2  # +2 for zero index and header offset
                })
        return students_to_email
    except HttpError as error:
        print(f"An error occurred: {error}")
        return []


def send_email(recipient_email):
    email_body = f"Please fill out this form to select your team: {FORM_URL}"
    message = MIMEText(email_body)
    message['to'] = recipient_email
    message['subject'] = 'Team Selection Form'

    # Encode and send the message
    encoded_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}
    try:
        service_gmail.users().messages().send(userId='me', body=encoded_message).execute()
    except HttpError as error:
        print(f"An error occurred: {error}")


def main():
    students = get_students_to_email()
    for student in students:
        send_email(student['email'])
        # Update the sheet to mark that an email has been sent
        update_sheet_email_sent(student['row'])

def update_sheet_email_sent(row):
    values = [['Yes']]
    body = {
        'values': values
    }
    range_name = f'Sheet1!C{row}'  # Column C for email sent status
    try:
        service_sheets.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=range_name,
            valueInputOption='USER_ENTERED', body=body).execute()
    except HttpError as error:
        print(f"An error occurred: {error}")

if __name__ == '__main__':
    main()
