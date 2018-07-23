'Comma as delimiter
str="Welcome,to,the,world,of,QTP"
'split
sArray=Split(str," ",-1)
For i=0 to UBOUND(sArray)
MsgBox sArray(i)
Next