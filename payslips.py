from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import base64

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def connect_to_api():
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
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def fetch_emails(service, creds):
    messages = None
    try:
        # Call the Gmail API

        labels = service.users().labels().list(
            userId='me'
        ).execute()
        results = service.users().messages().list(
            userId='me', labelIds=['Label_6047149750864999652']).execute()
        messages = results['messages']

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')

    return messages


def process_attachment(service, messageId, attachment_id, title="Payslip"):
    attachment = service.users().messages().attachments().get(
        userId='me', messageId=messageId, id=attachment_id
    ).execute()
    data = attachment['data']
    file_data = base64.urlsafe_b64decode(
        data.encode('UTF-8'))
    return file_data


def process_email(service, message_id, email):
    headers = email['payload']['headers']
    payslips_path = './payslips/'

    for header in headers:
        if header['name'] == 'Subject' and "Payslip for" in header['value']:
            title = header['value'].replace(
                'Payslip for Luke McMahon for ', '')
            for part in email['payload'].get('parts', ''):
                if part['filename'] == 'PaySlip.pdf':
                    title = title.replace(' ', '')

                    if not os.path.exists(payslips_path):
                        os.mkdir(os.getcwd() + payslips_path)

                    path = payslips_path + title + '.pdf'
                    if not os.path.exists(path):
                        attachmentId = part['body']['attachmentId']
                        file_data = process_attachment(
                            service, message_id, attachmentId, title)
                        with open(path, 'wb') as file:
                            file.write(file_data)
                            print(f'{title}.pdf created')
                    else:
                        print(
                            f'Payslip already exists for {title}... Skipping.')


def main():

    print("Connecting to Gmail API")
    creds = connect_to_api()
    service = build('gmail', 'v1', credentials=creds)
    print("Connected to Gmail API")

    print("Grabbing emails")
    messages = fetch_emails(service, creds)

    if not messages:
        print('No messages found.')
        return

    for message in messages:
        message_id = message['id']
        msg = service.users().messages().get(
            userId='me', id=message_id
        ).execute()
        process_email(service, message_id, msg)


if __name__ == '__main__':
    main()
