Deleting a file:

Option Explicit 
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.DeleteFile("C:\New folder\by.txt")
MsgBox "Success"