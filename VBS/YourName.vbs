' Header Information
' option explicit - force the scripter to declare variable
' on error resume next - to skip the error occuring line
' Dim - declares the variable
' reg - Contains names of variable that will hold register key Value
' with reg variable - That holds the information contained in th register key

option explicit
on error resume next
Dim objshell 
Dim regLogonUsername, regExchnageDomian, regGPServer
Dim regLogonServer, regDNSdomain
Dim LogonUsername, ExchnageDomain, GPServer
Dim LogonServer, DNSdomain


'Variable Reference Information 
' Variable Name = Register Key in Quatation Marks

regLogonUsername = "HKEY_CURRENT_USER\Software\Microsoft\" & "Windows\CurrentVersion\Explorer\Logon User Name"
regExchnageDomian = "HKEY_CURRENT_USER\Software\Microsoft\" & "Exchnage\LogonDomain"
regGPServer = "HKEY_CURRENT_USER\Software\Microsoft\" & "CurrentVersion\Group Policy\History\DCName"
regLogonServer = "HKEY_CURRENT_USER\Volatile Environment\" & "LOGONSERVER"
regDNSdomain = "HKEY_CURRENT_USER\Volatile Environment\" & "USERDNSDOMAIN"

'Worker Information

Set objshell = CreateObject("WScript.Shell")

' Variable Name = Worker & Registry Variable in ()

LogonUsername = objshell.RegRead(regLogonUsername)
ExchnageDomain = objshell.RegRead(regExchnageDomian)
GPServer = objshell.RegRead(regGPServer)
LogonServer = objshell.RegRead(regLogonServer)
DNSdomain = objshell.RegRead(regDNSdomain)

'Output Information
'Command _ Variable & Comment

wscript.echo LogonUsername & " is currently logged on"
wscript.echo ExchnageDomain & " is the current logon domain"
wscript.echo GPServer & " is the current group policy server"
wscript.echo LogonServer & " is the current logon server"
wscript.echo DNSdomain & "" & " is the current DNS domain"