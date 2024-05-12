Set app = CreateObject("Excel.Application")
app.DisplayAlerts = False
app.AlertBeforeOverwriting = False
app.Workbooks.Open ("E:\TRIPUTRA\DB.xlsm")
'app.quit