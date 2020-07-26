from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery
import constants

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", constants.scope)

service = discovery.build('sheets', 'v4', credentials=creds)

# Reading my spreadsheet
result = service.spreadsheets().values().get(
    spreadsheetId=constants.spreadsheet_id, range='Acoes').execute()
rows = result.get('values', [])
print(rows)

# Wrintig stock prices
body = {
    'values' : [['=GOOGLEFINANCE("BVMF:AZUL4")']]
}

result = service.spreadsheets().values().update(
    spreadsheetId=constants.spreadsheet_id, range='Acoes!C2', valueInputOption="USER_ENTERED",
    body=body).execute()

