' Header Information
' option explicit - force the scripter to declare variable
' on error resume next - to skip the error occuring line
' Dim - declares the variable

option explicit
on error resume next
Dim objshell 
Dim regActiveComputerName, regComterName, regHostname
Dim ActiveComputerName, ComputerName, Hostname

'Variable Reference Information 

regActiveComputerName = "HKLM\SYSTEM\CurrentControlSet\Control\" & "ComputerName\ActiveComputerName\ComputerName"
regComputerName = "HKLM\SYSTEM\CurrentControlSet\Control\" & "ComputerName\ComputerName\ComputerName"
regHostname = "HKLM\SYSTEM\CurrentControlSet\Service\TCpip\Parameters\Hostname"

'Worker Information

Set objshell = CreateObject("WScript.Shell")
ActiveComputerName = objshell.RegRead(regActiveComputerName)
ComputerName = objshell.RegRead(regComputerName)
Hostname = objshell.RegRead(regHostname)

'Output Information

wscript.echo ActiveComputerName & " is active Computer name"
wscript.echo ComputerName & " is computer name"
wscript.echo Hostname & " is host name"