# Excel (COM)

**example 1**

**Link to existing Excel workbook and worksheet**
``` python
import win32com.client**

ExcelApp = win32com.client.GetActiveObject('Excel.Application')
FILE_PATH = '<FILE PATH>'
wb = ExcelApp.Workbooks.Open(FILE_PATH)
wsData = wb.Worksheets('<worksheet name>')
```

**Dictionary Enumeration**
``` python
XLDirection = dict(xlDown=-4121, xlToLeft=-4159, xlToRight=-4161, xlUp=-4162)
LastRow = wsData.cells(wsData.api.rows.count, 'A').end(
    XLDirection.get('xlUp')).row
LastColumn = wsData.cells(1, wsData.api.columns.count).end(
    XLDirection.get('xlToLeft')).column
```
