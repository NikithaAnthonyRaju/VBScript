option Explicit
Dim FSO
Dim objStream
Dim strLine

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objStream = FSO.OpenTextFile("C:\New folder\byref.txt")

Do while not objStream.AtEndOfStream
strLine = objStream.ReadLine
msgbox strLine
Loop