Dim arrsample(9)
For nitem = 0 to UBOUND(arrsample) 
arrsample(nitem) = nitem + 1
Next
strArrival = ""
For each val in arrsample
strarrival = strarrival & val & " " 
Next
msgbox strarrival