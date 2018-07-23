Read the File:


Dim FSO
Dim objStream

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objStream = FSO.OpenTextFile("C:\New folder\byref.txt")

Do while not objStream.AtEndOfStream
strLine = objStream.ReadLine
msgbox strLine
Loop
Option Explicit
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FolderExists("C:\New folder") Then
FSO.CopyFolder"C:\Users\hpadmin\Desktop\Vb Scripts\New folder", "C:\New folder", True
MsgBox "Success"
End If