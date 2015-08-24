__author__ = 'brian'

# IMPORTANT NOTE: This program HAS to be run as root, otherwise it will get an SSL Error. Therefore,
#                 I run it from the command line.
#

import httplib2
import os
import re

# ----------------------------------------------------------------------------------------------------------------
# HACK follows:
# If you don't set the 3rd party path ahead of Mac OS X's path, it will get the following error:
#      parts = urllib.parse.urlparse(uri)
#      AttributeError: 'Module_six_moves_urllib_parse' object has no attribute 'urlparse'
#
# Details: see comment at bottom of page here: http://stackoverflow.com/questions/29190604/attribute-error-trying-to-run-gmail-api-quickstart-in-python
#
# Force the loading of the 3rd party libs first...
import sys
sys.path.insert(1, '/Library/Python/2.7/site-packages')
# ----------------------------------------------------------------------------------------------------------------

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import gspread


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

#SCOPES = 'https://www.googleapis.com/auth/drive.metadata.readonly'
SCOPES = 'https://www.googleapis.com/auth/drive.file https://spreadsheets.google.com/feeds https://docs.google.com/feeds'

# bsdrummond@gmail.com
CLIENT_SECRET_FILE = 'client_secret_981612606649-03ptve9ed27jj4h6k1ume8beqg6lm1bs.apps.googleusercontent.com.json'

# bdrummond@linkedin.com
#CLIENT_SECRET_FILE = 'client_secret_302402880436-pqb9hvba1459g8mghnddqoklj5uq48pn.apps.googleusercontent.com.json'

#APPLICATION_NAME = 'Drive API Quickstart'
APPLICATION_NAME = 'Google Sheets API Tester'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'googleSheetsAPITest.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
        print 'Storing credentials to ' + credential_path
    return credentials

def main():
    """Shows basic usage of the Google Drive API.

    Creates a Google Drive API service object and outputs the names and IDs
    for up to 10 files.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v2', http=http)

    # Test:
    # List all documents (scope: Google Drive API)
    results = service.files().list(maxResults=10).execute()
    items = results.get('items', [])
    if not items:
        print 'No files found.'
    else:
        print 'Files:'
        for item in items:
            print '{0} ({1})'.format(item['title'], item['id'])

    # Part 2: create a new spreadsheet
    gc = gspread.authorize(credentials)

    # TEST:
    # Open a spreadsheet
    ss = gc.open("Test Data Sheet")
    #worksheet = gc.open("Test Data Sheet").sheet1

    # TEST:
    # Open a worksheet in a spreadsheet by name
    worksheet = ss.worksheet("Sheet2")

    # TEST:
    # STATUS: PASSED
    # NOTE: you will get an error if you attempt this twice (not idempotent!)
    # new_wks = ss.add_worksheet('auto-sheet-2', 10, 10)

    # TEST:
    # STATUS: PASSED
    # Retrieve a specific cell's value
    val = worksheet.cell(1, 2).value

    # TEST:
    # STATUS: PASSED
    # Update a specific cell's value
    val = "S1-B5"
    worksheet.update_cell(5, 2, val)

    # TEST:
    # STATUS: PASSED
    # Find all cells with string value
    cell_list = worksheet.findall("Daffy")

    # TEST:
    # STATUS: PASSED
    # Find all cells with regexp
    criteria_re = re.compile(r'(B|C)3')
    cell_list = worksheet.findall(criteria_re)

    header_of_cartoon_characters_table = 9
    cartoon_character_records = worksheet.get_all_records(empty2zero=False, head=header_of_cartoon_characters_table)

    for cartoon_character in cartoon_character_records:
        print "First Name:", cartoon_character['First Name'],", Last Name:", cartoon_character['Last Name'],", Year Created:", cartoon_character['Year Created']
    # BREAKPOINT before exit, to examine any values in current scope
    True

    # TEST:
    # STATUS: PASSED
    # Create a new spreadsheet
    body = { 'mimeType': 'application/vnd.google-apps.spreadsheet', 'title': 'ZZ',}
    try:
        ZZ_ss = gc.open("ZZ")
    except:
        print "ZZ spreadsheet not found"

    True
    # file = service.files().insert(body=body).execute(http=http)

if __name__ == '__main__':
    main()