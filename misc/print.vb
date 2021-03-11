' $xl = New-Object -ComObject Excel.Application
' $xl.Visible = $true

' $wb = $xl.Workbooks.Open('F:\Test.xlsx')
' $ws = $wb.WorkSheets.Item(1)

' $xl.ActivePrinter = "Printer to print to"
' $ws.PrintOut(1,1,2)

' $xl.quit()

MsgBox "The name of the active printer is " & $xl.ActivePrinter