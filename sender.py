from __future__ import print_function

import base64
from email.message import EmailMessage

import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def send_email(email):
    print("sending email to " + email)
