# Google imports
from __future__ import print_function
from email.quoprimime import body_check

from googleapiclient.discovery import build
from google.oauth2 import service_account
from pathlib import Path
from dotenv import load_dotenv #load environment variable file
from Google import create_service
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import datetime
import os.path
import os
import time
import base64

new_list = []
def client_list_data():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = 'keys.json'
    credentials = None
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    SAMPLE_SPREADSHEET_ID = '1id8Fu1fUkYnV31OQEdWOFZtXui3cgiRihR9Rucg5lAU'
    service = build('sheets', 'v4', credentials=credentials)

    env_path = Path('.') / '.env' #('.') represents current directory
    load_dotenv(dotenv_path=env_path)
    token = os.environ['ACCESS_TOKEN']
    # pylint: disable=maybe-no-member
    range_ = "test1!A1:AA1000"

    result = service.spreadsheets().values().batchGet(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, ranges=range_).execute()

    ranges = result.get('valueRanges', [])
    print(f"{len(ranges)} ranges retrieved")

    for i in ranges:
        new_list.append(i)
    return new_list





    

        

