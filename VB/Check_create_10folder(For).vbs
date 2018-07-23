'To check and create the 10 folder:(for)

Option Explicit
dim FSO
dim objFolder, i
Set FSO = CreateObject("Scripting.fileSystemObject")
Set objFolder = FSO.GetFolder("C:\Users\Nikitha Anthony\Desktop\VB")
for i = 0 to 9
If Not FSO.FolderExists("C:\Users\Nikitha Anthony\Desktop\VB\TestVBScriptFolder" & i)  Then
objFolder.SubFolders.Add "TestVBScriptFolder" & i
Msgbox "C:\TestVBScriptFolder folder was created" & i
else
msgbox "Folder exists"
End If
Next