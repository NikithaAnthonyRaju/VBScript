option Explicit
dim intTemp
on Error Resume Next

intTemp = -1
msgbox Left("Quick Brown Fox", cint(intTemp))

If Err then
MsgBox "Error:" & Err.number
Err.clear
end if

intTemp = 5
msgbox Left("Quick Brown Fox", cint(intTemp))

If Err then
MsgBox "Error:" & Err.number
Err.clear
else 
MsgBox "No error occured"
end if