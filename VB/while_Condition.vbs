While Condition

dim arrSample(4)
for i=0 to ubound(arrSample)
	arrSample(i) = i+1
next

count = 0
do while count <=  ubound(arrSample)
	msgbox(arrSample(count))
	count = count+1
loop
