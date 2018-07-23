Option Explicit
On Error Resume Next
Dim objWMIService 'Hold the connection to WMI
Dim objItem 'Hold the name of each Process that comes back from objWMIService
Dim i 
Const MAX_LOOPS = 8, ONE_HOUR = 3600000 

For i = 1 to MAX_LOOPS
	Set objWMIService = GetObject("winmgmts:").ExecQuery("Select * from win32_process where processID <> 0")
	WScript.echo "There are " & objWMIService.count & " Process running " & Now
	For Each objItem in objWMIService.count 
		WScript.echo "Process: " & objItem.Name
		WScript.echo space(9) & objItem.commandline
		Wscript.echo "Process ID: " & objItem.ProcessID
		WScript.echo "Thread Count: " & objItem.ThreadCount
		WScript.echo "Page File Size: " & objItem.PageFileUsage
		Wscript.echo "Page Faults: " & objItem.PageFaults
		WScript.echo "Working Set Size " & objItem.WorkingSetSize
		'WScript.echo vbNewLine
	Next
WScript.echo "**********PASS COMPLETE*********"
Wscript.Sleep ONE_HOUR
Next