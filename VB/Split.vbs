'Split FUNCTION
str="Welcome to the world of QTP"
sArray=Split(str," ",-1)
For i=0 to UBOUND(sArray)
MsgBox sArray(i)
Next