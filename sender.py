from __future__ import print_function

import base64
import os
import mimetypes
from email.message import EmailMessage

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def send_email(destination_email, attachment_file):
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
            flow = InstalledAppFlow.from_client_secrets_file( 'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        message = EmailMessage()

        # Headers
        message['To'] = destination_email
        message['From'] = 'mengelesCristi@gmail.com'
        message['Subject'] = "Something special"

        # Text
        message.set_content("Please receive the attached payslip")

        # Attachment - get type
        type_subtype, _ = mimetypes.guess_type( attachment_file )
        maintype, subtype = type_subtype.split('/')

        # Attachment
        with open( attachment_file, 'rb' ) as fp:
            attachement_data = fp.read()
        message.add_attachment( attachement_data, maintype, subtype )

        # encoded message
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()


        message_body = {
            'raw': encoded_message
        }

        # pylint: disable=E1101
        send_message = ( service.users().messages().send(userId="me", body=message_body).execute() )
        print(F'Message Id: {send_message["id"]}')

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')
