Dim xlApp, macroWb, macroPath, wshShell, curDir

Set wshShell = CreateObject("WScript.Shell")
curDir = wshShell.CurrentDirectory

Set xlApp = CreateObject("Excel.Application")
xlApp.DisplayAlerts = False

macroPath = curDir & "\data_consolidation.xlsm"
Set macroWb = xlApp.Workbooks.open(macroPath, 0, True)

Set Sheet1 = macroWb.Sheets(1)
Sheet1.consolidateData 

xlApp.Application.Quit
Set xlApp = Nothing
	