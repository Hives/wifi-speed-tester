#! /usr/bin/env python3

"""
BEFORE RUNNING:
---------------
1. If not already done, enable the Google Sheets API
   and check the quota for your project at
   https://console.developers.google.com/apis/api/sheets
2. Install the Python client library for Google APIs by running
   `pip install --upgrade google-api-python-client`


Docs: https://developers.google.com/sheets/api/quickstart/python
"""

from __future__ import print_function

import os.path
import pickle
from pprint import pprint

import speedtest
from dateutil import parser
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient import discovery


def format_date(iso_date):
    return parser.isoparse(iso_date).strftime("%d %b %Y %H:%M:%S")


s = speedtest.Speedtest()
s.get_servers([])
s.get_best_server()
s.download(threads=None)
s.upload(threads=None)
s.results.share()

results = s.results.dict()

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# TODO: Change placeholder below to generate authentication credentials. See
# https://developers.google.com/sheets/quickstart/python#step_3_set_up_the_sample
#
# Authorize using one of the following scopes:
#     'https://www.googleapis.com/auth/drive'
#     'https://www.googleapis.com/auth/drive.file'
#     'https://www.googleapis.com/auth/spreadsheets'
creds = None

# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
dirname = os.path.dirname(__file__)
pickle_path = os.path.join(dirname, 'token.pickle')
credentials_path = os.path.join(dirname, 'credentials.js')
if os.path.exists(pickle_path):
    with open(pickle_path, 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            credentials_path, SCOPES
        )
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open(pickle_path, 'wb') as token:
        pickle.dump(creds, token)

service = discovery.build('sheets', 'v4', credentials=creds)

# The ID of the spreadsheet to update.
spreadsheet_id = '1LPcyXE2-pliBaTpcS9hpMvX59BS2DrwbeZ_Wgwy_ABw'

# The A1 notation of a range to search for a logical table of data.
# Values will be appended after the last row of the table.
range_ = 'Sheet1!A:C'

# How the input data should be interpreted.
value_input_option = 'RAW'

# How the input data should be inserted.
insert_data_option = 'INSERT_ROWS'

value_range_body = {
    "range": "Sheet1!A:C",
    "majorDimension": 'ROWS',
    "values": [
        [
            format_date(results['timestamp']),
            results['download'],
            results['upload']
        ]
    ]
}

request = service.spreadsheets().values().append(
    spreadsheetId=spreadsheet_id,
    range=range_,
    valueInputOption=value_input_option,
    insertDataOption=insert_data_option,
    body=value_range_body)
response = request.execute()

# TODO: Change code below to process the `response` dict:
pprint(response)
