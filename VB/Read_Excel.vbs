'create a new microsoft excel object
Set myexcel = CreateObject("Excel.Application")

'opening the existing file
myexcel.Workbooks.open "C:\Users\hpadmin\Desktop\Vb Scripts\New folder\qtp.xlsx"

'To make Excel visible
myexcel.Application.Visible = True

'This is the name of sheet in Excel file "qtp.xlsx" where data need to be add
Set mysheet = myexcel.ActiveWorkbook.Worksheets("Sheet1")

'Get the max row occupied in the excel sheet
Row=mysheet.UsedRange.Rows.Count
Msgbox Row

'Get the max column occupied in the excel sheet
Col=mysheet.UsedRange.columns.Count
Msgbox Col

'To read the data from the entire excel file
for i=1 to Row
  For j=1 to col
	Msgbox mysheet.cells(i,j).value
Next
Next