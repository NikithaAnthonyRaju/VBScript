'Create Folder

Option Explicit
Dim FSO
Dim objFolder
dim i
Set FSO=CreateObject("Scripting.filesystemobject")
set objFolder=FSO.GetFolder("C:\Users\hpadmin\Desktop\New folder")
for i=1 to 10
	If not FSO.folderexists("C:\Users\hpadmin\Desktop\VB Scripts "&i) then
		objFolder.subfolders.add "VB Scripts"&i
	end if
next
msgbox "The folder has been created successfully"