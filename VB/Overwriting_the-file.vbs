creating the file or overwriting on existing file:

Dim FSO
Dim objStream
const file_byref = "C:\New folder\byref.txt"

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objStream = FSO.CreateTextFile(file_byref)

With objstream
 .Writeline"one"
 .Writeline"one"
 .Writeline"one"
End With
msgbox "Success Created" & file_byref