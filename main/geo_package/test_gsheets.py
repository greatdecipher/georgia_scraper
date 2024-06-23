
# test for Gsheet column generated: setting of hyperlink column
import os, sys
main_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(main_dir)
from config import SHEETS_KEY
import gspread
from oauth2client.service_account import ServiceAccountCredentials

sheets_key = SHEETS_KEY
SCOPES = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
credentials = ServiceAccountCredentials.from_json_keyfile_dict(sheets_key, SCOPES)
client = gspread.authorize(credentials)
gsheet_link ='https://docs.google.com/spreadsheets/d/1Q5EYgpK0cMHx0bUHgINmJg7cW342Ax-mW57CLgD3E9U/edit?usp=sharing'
url_provided = 'https://www.georgiapublicnotice.com/(S(afvlavib2eaapevnskvicf5e))/Details.aspx?SID=afvlavib2eaapevnskvicf5e&ID=3073594'
worksheet = client.open_by_url(gsheet_link).sheet1
print("Worksheet opened")

COLUMN_ALIAS = 'F'  # Column for aliases
COLUMN_URL = 'E'    # Column for URLs

# Get the values in the alias column
alias_values = worksheet.col_values(ord(COLUMN_ALIAS) - ord('A') + 1)[1:]

def set_to_hyperlink():
    # Generate hyperlinks and update the URL column
    for i in range(len(alias_values)):
        alias = alias_values[i]
        if alias:  # Skip empty cells
            hyperlink_formula = f'=HYPERLINK("{url_provided}", "{alias}")'
            worksheet.update_cell(i + 2, ord(COLUMN_URL) - ord('A') + 1, hyperlink_formula)


set_to_hyperlink()
