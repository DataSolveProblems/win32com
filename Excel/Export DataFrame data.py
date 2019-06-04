import os
import win32com.client as win32
import pandas as pd
import pprint
from datetime import datetime

WORK_PATH = os.getcwd()

# reference active Excel file
# must make excel workbook activiate, otherwise
# Operation unavailable error will raise
xlApp = win32.GetActiveObject('Excel.Application')
wb = xlApp.Workbooks('pandas.xlsx')
wsData = wb.Worksheets('Data')
# pprint.pprint(wsData.Range("A1:N2").value[:2])

data = pd.Series(wsData.Range("A1:N5").value)
df = pd.DataFrame(list(data))

# new header
header_names = df.iloc[0]
df = df[1:]
df.reset_index(drop=True, inplace=True)
df.rename(columns=header_names, inplace=True)

wsPanda = wb.Worksheets('pandas')
wsPanda.Range(wsPanda.Cells(1, 1), wsPanda.Cells(1, df.shape[1])).value = list(header_names)
wsPanda.Range(wsPanda.Cells(2, 1), wsPanda.Cells(df.shape[0] + 1, df.shape[1])).value = list(df.values)
