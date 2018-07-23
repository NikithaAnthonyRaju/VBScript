' copy the file


Option Explicit
Dim FSO

Set FSO=CreateObject("Scripting.filesystemobject")
If FSO.fileexists("C:\Users\hpadmin\Desktop\VB Scripts\New Text Document.txt") then
		FSO.CopyFile "C:\Users\hpadmin\Desktop\VB Scripts\New Text Document.txt", "C:\Users\hpadmin\Desktop\VB Scripts\NewTextDocument.txt", True
	msgbox "File Copied"
end if