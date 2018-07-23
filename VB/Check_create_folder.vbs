'To check and create the folder:

Option Explicit
dim FSO
dim objFolder
Set FSO = CreateObject("Scripting.fileSystemObject")
Set objFolder = FSO.GetFolder("C:\Users\Nikitha Anthony\Desktop\VB")
If Not FSO.FolderExists("C:\Users\Nikitha Anthony\Desktop\VB\TestVBScriptFolder") Then
objFolder.SubFolders.Add "TestVBScriptFolder"
Msgbox "C:\TestVBScriptFolder folder was created" 
else
msgbox "Folder exists"
End If