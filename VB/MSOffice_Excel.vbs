this function reads data from an Excel sheet without using MS-Office:

option Explicit
Call REadExcel("c:\book1.xlsx","Sheet1","A1","AQ100",False)
Function ReadExcel(myxlsxFile,mysheet,my1stcell,mylastcell,blnheader)

'The value reads from the excel is 2-dimentional array(column,row)

Dim arrData(), i,j
Dim objExcel, objRS
Dim strHeader, strRange, strResult
if blnHeader then
strHeader= "HDR=YES"
Else
strHeader= "HDR=No"
End If

'open the object for the Excel file
Set objExcel = CreateObject("ADODB.Connection")
objExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;DataSource="& _
myxlsFile & ";Extended Properties=""Excel 8.0;"& _
strHeader & ""

'open a recordset object for the sheet and range

Set objExcel = CreateObject("ADODB.Connection")
strRange = mysheet & "$" &my1stcell & ":" & mylastcell
objRS.open "Select * from [" & strRange & "]", objExcel, adOpenStatic

'read the data feom the excel sheet
i = 0

Do Until objRS.EOF

'stop reading when an empty row is encountered in Excel sheet
If IsNull(objRS.Feilds(0).Value) or Trim(objRS.Feilds(0).Value) = "" Then
Exit Do

'Add a new row to the output array
ReDim Preserve arrData(objRS.Feilds.Count-1,i)

'copy the excel sheet's row values to the array "row"

for j = 0 to objRS.Feilds.Count-1
If IsNull(objRS.Feild(j).value) then
arrdata(j,i)=""
else
arrdata(j,i) = trim(objRS.Feild(j).Value)
strResult = strResult & arrData(j,i) & vbtab
End if

'Move to the next row

objRS.MoveNext
'incremental the array "row" number
i=i+1
Loop

'Return the result
ReadExcel = strResult
msgbox ReadExcel

'close the file and release the object
objRS.Close
objExcel.Close
Set objRS = Nothing
Set objExcel = Nothing

End Function