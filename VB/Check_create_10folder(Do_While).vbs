Option Explicit
dim FSO
dim objFolder, i
i = 1
Set FSO = CreateObject("Scripting.fileSystemObject")
Set objFolder = FSO.GetFolder("C:\Users\Nikitha Anthony\Desktop\VB")
do while i <= 10
If Not FSO.FolderExists("C:\Users\Nikitha Anthony\Desktop\VB\TestVBScriptFolder" & i)  Then
objFolder.SubFolders.Add "TestVBScriptFolder" & i
'Msgbox "C:\TestVBScriptFolder folder was created" & i
else
msgbox "Folder exists"
End If
i = i+1
Loop
msgbox "Folder Successfully created"