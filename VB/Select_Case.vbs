num1 = cint(inputbox("Enter the no"))
num2 = cint(inputbox("Enter the no"))
opr = Inputbox("Enter the operator to perform:" & vbNewLine & "1. Add" & vbNewLine & "2. Sub" & vbNewLine & "3. Mul" & vbNewLine & "4. Div")
Select case opr
case 1
call add(num1,num2)
case 2
call subt(num1,num2)
case 3 
call mul(num1,num2)
case 4 
call div(num1,num2)
End Select

Function add(num1,num2)
msgbox num1 + num2
End Function
Function subt(num1,num2)
msgbox num1 - num2
End Function
Function mul(num1,num2)
msgbox num1 * num2
End Function
Function div(num1,num2)
msgbox num1 / num2
End Function