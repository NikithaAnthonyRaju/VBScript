dim orgin
orgin = cint(Inputbox("enter the no"))
call some(tell)
function some(byref tell)
tell = orgin + 3
end function
msgbox "New Value:" & " " & tell
msgbox "Orgin Value":" & " " & orgin