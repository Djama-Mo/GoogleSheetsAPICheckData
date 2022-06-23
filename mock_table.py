from __future__ import print_function

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credentials.json'

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1YdDuLWUJk3eyl8qMTB-uVLkBM_j69T6MtCJ4CntCXRY'

# RANGE TO GET DATA
SAMPLE_RANGE_NAME = 'Лист1!A2:D'
# RANGE TO CLEAR DATA
CLEAR_RANGE = 'Лист1!E3:E'
# RANGE TO UPDATE/INSERT DATA
UPDATE_INSERT_RANGE = 'Лист1!E2:F'


class GoogleSheetsData(object):
    def __init__(self):
        self.creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    def get_values(self):
        try:
            # Construct a Resource object for interacting with an API
            service = build('sheets', 'v4', credentials=self.creds)

            # Call the Sheets API
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                        range=SAMPLE_RANGE_NAME).execute()

            # A range of values from a spreadsheet
            values = result.get('values', [])

            if not values:
                print('No data found.')
            else:
                return values
        except HttpError as err:
            print(err)

    def insert_or_insert(self):
        try:
            # DATA TO INSERT/UPDATE CELLS
            values = [
                ['', 3823],
                [879, 'oinsdivb'],
            ]
            body = {
                'values': values
            }

            # Construct a Resource object for interacting with an API
            service = build('sheets', 'v4', credentials=self.creds)

            # Call the Sheets API
            sheet = service.spreadsheets()
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, valueInputOption='RAW',
                                           range=UPDATE_INSERT_RANGE, body=body).execute()

            values_update_insert = result.get('updatedCells')
            if not values_update_insert:
                print('No data found.')
            else:
                return values_update_insert
        except HttpError as err:
            print(err)

    def clear(self):
        try:
            # Construct a Resource object for interacting with an API
            service = build('sheets', 'v4', credentials=self.creds)

            sheet = service.spreadsheets()
            sheet.values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                 range=CLEAR_RANGE, body={}).execute()

        except HttpError as err:
            print(err)
