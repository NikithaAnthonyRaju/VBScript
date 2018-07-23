'Input from users for 2 dimension array and display it



nrow = cint(inputbox("Enter the no of rows"))
ncol = cint(inputbox("Enter the no of cols"))
dim arrSample()
redim arrSample(nrow, ncol)
for i=0 to ubound(arrSample,1)
	for j=0 to ubound(arrSample,2)
		arrSample(i,j)=inputbox("Enter the Value")
	next
next

str = " "
for i=0 to ubound(arrSample,1)
	for j=0 to ubound(arrSample,2)
		str = str & arrSample(i,j) & "	"
	next
	str= str & vbnewline
next	
msgbox str
