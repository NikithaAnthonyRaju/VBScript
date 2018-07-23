write into exisiting excel sheet:

'create a new microsoft excel object
Set myexcel = CreateObject("Excel.Application")

'opening the existing file
myexcel.Workbooks.open "C:\Users\hpadmin\Desktop\Vb Scripts\New folder\qtp.xlsx"

'To make Excel visible
myexcel.Application.Visible = True

'This is the name of sheet in Excel file "qtp.xlsx" where data need to be add
Set mysheet = myexcel.ActiveWorkbook.Worksheets("Sheet1")

'Enter the values to the cell
mysheet.cells(1,1).value = "Name"
mysheet.cells(1,2).value = "Age"
mysheet.cells(2,1).value = "Ram"
mysheet.cells(2,2).value = "25"
mysheet.cells(3,1).value = "Sita"
mysheet.cells(3,2).value = "18"

'To Save Excel sheet
myexcel.ActiveWorkbook.Save

'Close the excel
myexcel.Application.Quit

Set myexcel=nothing