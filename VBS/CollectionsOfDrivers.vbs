Option Explicit
On Error Resume Next
Dim colDrives 'the collections that comes from WMI - Windows management insturmentation
Dim Drive 'an individual drive in the collections
Const DriveType = 3 'Local Drives from the SDK
'3 = Used for Fixed disks
'others drive types are 2 for removal
'4 for network, 5 for CD

Set colDrives = GetObject("winmgmts:").ExecQuery("Select size, freespace" & "from win32_LogicalDisk where DriveType =" & DriveType)

For Each drive in colDrives 'walks through the collections
WScript.echo "Drive:" & drive.DeviceID
WScript.echo "Size:" & drive.size
Wscript.echo "Freespace:" & drive.freespace
Next