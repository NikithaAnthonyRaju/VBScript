
'To Write a file

Dim fso
Dim objStream

const file="C:\Users\hpadmin\Desktop\VB Scripts\fun1.txt"

set fso=createObject("Scripting.filesystemObject")
set objStream=fso.CreateTextFile("fun.txt")

with objStream
	.writeline "Vidya"
	.writeline "Vidya"
end with
msgbox "Success"