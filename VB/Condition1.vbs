condition = true
Do While condition = true
response = msgbox("Press Ok", vbokcancel)
if response = vbcancel then 
condition = false
end if
loop