'2 dimension Array

dim arrSample(1,1)
for i=0 to ubound(arrSample,1)
	for j=0 to ubound(arrSample,2)
		arrSample(i,j) = i+j
	next
next

for each items in arrSample
	msgbox items
next	
