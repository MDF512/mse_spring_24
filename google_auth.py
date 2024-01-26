import gspread
from google.oauth2 import service_account
import json
from os import environ


def get_client():
        """
        Source Code:
        https://www.datacamp.com/tutorial/how-to-analyze-data-in-google-sheets-with-python-a-step-by-step-guide
        :return: a client
        """
        api_file = environ['GOOGLE_SHEETS_API_KEY']
        service_account_info = json.loads(api_file)
        credentials = service_account.Credentials.from_service_account_info(service_account_info)
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_with_scope = credentials.with_scopes(scope)
        client = gspread.authorize(creds_with_scope)
        return client
