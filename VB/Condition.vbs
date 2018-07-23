condition = true
ncnt = 0
Do While condition = true
ncnt = ncnt + 1
msgbox ncnt
if ncnt = 10 then 
condition = false
end if
loop
msgbox "Exit loop"