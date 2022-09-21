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


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
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
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        labels = service.users().labels().list(
            userId='me'
        ).execute()
        results = service.users().messages().list(
            userId='me', labelIds=['Label_6047149750864999652']).execute()
        messages = results['messages']

        if not messages:
            print('No messages found.')
            return

        for message in messages:
            msg = service.users().messages().get(
                userId='me', id=message['id']
            ).execute()
            headers = msg['payload']['headers']
            for header in headers:
                if header['name'] == 'Subject' and "Payslip for" in header['value']:
                    title = header['value'].replace(
                        'Payslip for Luke McMahon for ', '')
                    for part in msg['payload'].get('parts', ''):
                        # print(part['filename'])
                        if part['filename'] == 'PaySlip.pdf':
                            title = title.replace(' ', '')
                            print(title)
                            path = './payslips/' + title + '.pdf'
                            if not os.path.exists(path):
                                attachmentId = part['body']['attachmentId']
                                # print(attachmentId)
                                attachment = service.users().messages().attachments().get(
                                    userId='me', messageId=message['id'], id=attachmentId
                                ).execute()
                                data = attachment['data']
                                file_data = base64.urlsafe_b64decode(
                                    data.encode('UTF-8'))
                                with open(path, 'wb') as f:
                                    f.write(file_data)
                                    print(f'{title} created')
                            else:
                                print(
                                    f'Payslip already exists for {title}... Skipping.')

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')


if __name__ == '__main__':
    main()
