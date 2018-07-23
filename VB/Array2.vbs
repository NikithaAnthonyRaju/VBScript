dim Aarrsample()
row = cint(Inputbox("Enter the number"))
col = cint(Inputbox("Enter the number"))

Redim Aarrsample(ncol, nrow)
for ncol = 0 to UBOUND(Aarrsample, 1)
for nrow = 0 to UBoUND(Aarrsample, 2)
Aarrsample(nrow,ncol) = Inputbox("Enter the number/name")
Next
Next
strarr = ""
for ncol = 0 to UBOUND(Aarrsample, 1)
for nrow = 0 to UBoUND(Aarrsample, 2)
strarr = strarr & Aarrsample(ncol, nrow) & "	" 
Next
strarr = strarr & vbNewLine
Next
msgbox strarr