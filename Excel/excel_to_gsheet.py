import win32com.client as win32
from Google import Create_Service

xlApp = win32.Dispatch('Excel.Application')
wb = xlApp.Workbooks.Open(
    r'C:\Users\jjenn\Documents\Python\win32com\Dataset\champion_list.xlsx')
wsData = wb.Worksheets('Worksheet')
rngdata = wsData.Range('A1').CurrentRegion


CLIENT_SECRET_FILE = 'client_secret.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

gsheet_id = '1iD0w4OQBdNTNV78XVwbRUyopJLgkdtneEkbVwTZqukk'

response_date = service.spreadsheets().values().append(
    spreadsheetId=gsheet_id,
    valueInputOption='RAW',
    range='xlData!A1',
    body=dict(
        majorDimension='ROWS',
        values=rngdata.value)).execute()
