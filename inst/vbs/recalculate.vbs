Set oExcel = CreateObject("Excel.Application")
oExcel.DisplayAlerts = FALSE 'to avoid prompts
Set oWorkbook = oExcel.Workbooks.Open(WScript.Arguments(0))
oExcel.Calculate

oWorkbook.Close True
