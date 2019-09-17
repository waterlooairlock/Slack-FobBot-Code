import os
import requests

import gspread

from oauth2client.service_account import ServiceAccountCredentials

print("?")

scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(credentials)

os.environ['SLACK_BOT_TOKEN'] = 'REDACTED'
slack_token = os.environ['SLACK_BOT_TOKEN']

print('test')


def update_email_list():
    user_list = requests.get("https://slack.com/api/users.list?token=%s" % slack_token).json()['members']
    print ('#1')
    email_file = client.open_by_url(
        "REDACTED")

    email_sheet = email_file.add_worksheet(title="Date and Name", rows="12", cols="10")
    print('#2')
    print(user_list)


update_email_list()
