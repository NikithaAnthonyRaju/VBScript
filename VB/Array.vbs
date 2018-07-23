dim Aarrsample()

Redim preserve Aarrsample(2,0)
Aarrsample(0,0)="animals"
Aarrsample(1,0)="ani"
Aarrsample(2,0)="114-551-54"

Redim preserve Aarrsample(2,1)
Aarrsample(0,1)="tell"
Aarrsample(1,1)="is"
Aarrsample(2,1)="7-51-54"
strarr = ""
for row = 0 to 1
strarr = strarr & Aarrsample(0, row) & "|" & Aarrsample(1, row) & "|" & Aarrsample(2, row) & "|" & vbNewLine
Next
msgbox strarr