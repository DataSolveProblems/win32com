# Excel (COM)

**example 1
Link to existing Excel workbook and worksheet
``` python
import win32com.client**
ExcelApp = win32com.client.GetActiveObject('Excel.Application')
FILE_PATH = '<FILE PATH>'
wb = ExcelApp.Workbooks.Open(FILE_PATH)
wsData = wb.Worksheets('<worksheet name>')
```
