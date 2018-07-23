dim val
valn = 5
fnfunc val
msgbox "Original Value:" & valn

function fnfunc(ByRef val)
val = valn + 2
msgbox "New Value:" & val
End function