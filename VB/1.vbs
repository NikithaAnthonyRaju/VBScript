option explicit
on error Resume Next
Dim x,y
x = InputBox("Enter a number to divide by 100")
y = 100/x
If Err then
Msgbox "Error code:" & Err.Number & vbNewLine & Err.description, vbOKOnly, "VB_Description"
else
Msgbox "100 divided by" & x & "is" &y& ","
End if