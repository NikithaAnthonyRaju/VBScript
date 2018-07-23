select operation:
intoperation = cint(Inputbox("Choose the operation to be performed" & vbnewline & "1. Add" & vbnewline & "2. Sub" & vbnewline & "3. mul" & vbnewline & "4. div"))
intnumber1 = cint(Inputbox("please enter the number"))
intnumber2 = cint(Inputbox("please enter the number"))
select case intoperation
case 1
 msgbox intnumber1 + intnumber2
Case 2
 msgbox intnumber1 - intnumber2
Case 3
 msgbox intnumber1 * intnumber2
Case 4
 msgbox intnumber1 / intnumber2
case else
 msgbox "Invalid"
End select