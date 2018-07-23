create excel and save:

'create a new microsoft excel object
Set myexcel = CreateObject("Excel.Application")

'To make Excel visible
myexcel.Application.Visible = True

myexcel.Workbooks.Add
'Wait 2

'To Save Excel sheet
myexcel.ActiveWorkbook.SaveAs "C:\Users\hpadmin\Desktop\Vb Scripts\New folder\qtp.xlsx"

'Close the excel
myexcel.Application.Quit

Set myexcel=nothing