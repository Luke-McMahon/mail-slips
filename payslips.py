"""
Downloads payslip pdf files from Gmail and processes them
"""


import os.path
import base64

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

payslips_path = "C:\\Users\\lukem\\Documents\\Work\\Two_Bulls\\payslips\\"

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def connect_to_api():
    """ Connect to the Gmail API, create token file and return the creds"""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            #
        with open('token.json', 'w', encoding='utf-8') as token:
            token.write(creds.to_json())
    return creds


def fetch_emails(service):
    """ Fetch emails under a specified label"""
    messages = None
    try:
        # Call the Gmail API
        results = service.users().messages().list(
            userId='me', labelIds=['Label_6047149750864999652']).execute()
        messages = results['messages']

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')

    return messages


def process_attachment(service, message_id, attachment_id):
    """ Return a base64 decoded data stream of a given attachment id"""
    attachment = service.users().messages().attachments().get(
        userId='me', messageId=message_id, id=attachment_id
    ).execute()
    data = attachment['data']
    file_data = base64.urlsafe_b64decode(
        data.encode('UTF-8'))
    return file_data

month_to_title_map = {
    'January': '01Jan',
    'February': '02Feb',
    'March': '03Mar',
    'April': '04Apr',
    'May': '05May',
    'June': '06Jun',
    'July': '07Jul',
    'August': '08Aug',
    'September': '09Sep',
    'October': '10Oct',
    'November': '11Nov',
    'December': '12Dec',
}

def process_email(service, message_id, email):
    """ Given an email, download a pdf or skip it if we've already got it"""

    global payslips_path

    headers = email['payload']['headers']
    if payslips_path == '':
        payslips_path = './payslips/'

    for header in headers:
        if header['name'] == 'Subject' and "Payslip for" in header['value']:
            title = header['value'].replace(
                'Payslip for Luke McMahon for ', '')
            for part in email['payload'].get('parts', ''):
                if part['filename'] == 'PaySlip.pdf':
                    title_parts = title.split(' ')
                    print(title_parts)
                    pdf_title = month_to_title_map[title_parts[0]] + title_parts[1]

                    # TODO: split into folders based on the year, should be title_parts[1] for all emails
                    # title_parts = ['January', '2023']

                    if not os.path.exists(payslips_path):
                        os.mkdir(os.getcwd() + payslips_path)

                    path = payslips_path + pdf_title + '.pdf'
                    if not os.path.exists(path):
                        attachment_id = part['body']['attachmentId']
                        file_data = process_attachment(
                            service, message_id, attachment_id)
                        with open(path, 'wb') as file:
                            file.write(file_data)
                            print(f'{path} created')
                    else:
                        print(
                            f'Payslip already exists for {title}... Skipping.')


def main():
    """ Entry point"""
    print("Connecting to Gmail API")
    creds = connect_to_api()
    service = build('gmail', 'v1', credentials=creds)
    print("Connected to Gmail API")

    print("Grabbing emails")
    messages = fetch_emails(service)

    if not messages:
        print('No messages found.')
        return

    for message in messages:
        message_id = message['id']
        # pylint: disable=maybe-no-member
        msg = service.users().messages().get(
            userId='me', id=message_id
        ).execute()
        process_email(service, message_id, msg)


if __name__ == '__main__':
    main()
