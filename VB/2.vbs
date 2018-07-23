dim FEO, objFolder

set FEO = CreateObject("Scripting.filesystemobject")
set objFolder = FEO.GetFolder("C:\Users\Nikitha Anthony\Desktop\VB")
if NOT FEO.FolderExists("C:\Users\Nikitha Anthony\Desktop\VB\new") then
objFolder.SubFolders.Add "new"
msgbox "created"
End if