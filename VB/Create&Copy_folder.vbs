'Create a folder and copy a file into it

Option Explicit
Dim FSO
Dim objFolder
Set FSO=CreateObject("Scripting.filesystemobject")
set objFolder=FSO.GetFolder("C:\Users\hpadmin\Desktop")
	If not FSO.folderexists("C:\Users\hpadmin\Desktop\New Folder") then
		objFolder.subfolders.add "New Folder"
		
		If FSO.fileexists("C:\Users\hpadmin\Desktop\VB Scripts\New Text Document.txt") then
		FSO.CopyFile "C:\Users\hpadmin\Desktop\VB Scripts\New Text Document.txt", "C:\Users\hpadmin\Desktop\New Folder\New Text Document.txt", True
	msgbox "File Copied"
		end if
	end if