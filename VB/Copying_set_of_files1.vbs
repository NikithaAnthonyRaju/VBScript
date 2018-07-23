copying a file:

Option Explicit
dim FSO
Set FSO = CreateObject("Scripting.fileSystemObject")
If FSO.FolderExists("C:\TestVBScriptFolder") Then
FSO.CopyFile"C:\TestVBScriptFolder\byref.txt","C:\New folder\name.txt",True
msgbox "file is copied"
else
msgbox "Folder doesnot exist"
End If