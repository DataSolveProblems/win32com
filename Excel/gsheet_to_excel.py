import win32com.client as win32
from Google import Create_Service

CLIENT_SECRET_FILE = 'client_secret.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
gsheet_id = '1iD0w4OQBdNTNV78XVwbRUyopJLgkdtneEkbVwTZqukk'
service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)
gs = service.spreadsheets()
rows = gs.values().get(spreadsheetId=gsheet_id, range='Lakers').execute()
data = rows.get('values')


xlApp = win32.Dispatch('Excel.Application')
wb = xlApp.Workbooks.Open(
    r'C:\Users\jjenn\Documents\Python\win32com\Dataset\champion_list.xlsx')
wsLakers = wb.Worksheets('Lakers')


wsLakers.Cells.ClearContents()

rowNumber = 1
colCount = len(data[0])

for row in data:
    print(row)
    wsLakers.Range(wsLakers.cells(rowNumber, 1), wsLakers.cells(rowNumber, colCount)).value = row
    rowNumber += 1
