Deleting the multiple files:

Option Explicit 
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Msgbox ("Are you sure to delete the file", vbYesNo) = vbYes Then
objFSO.DeleteFile("C:\New folder\*.xlsx")
MsgBox "Success"
Else
MsgBox "Failed to delete"
End If