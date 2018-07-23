copying set of files: *--> all files

Option Explicit
Dim objFSO
Const OverwriteExisting = True
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile"C:\Users\hpadmin\Desktop\Vb Scripts\New folder\*.xlsx", "C:\New folder", OverwriteExisting
MsgBox "Success"