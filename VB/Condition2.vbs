dim arrsample(10)
for narr = 0 to UBOUND(arrsample)
arrsample(narr) = narr + 1
next
ncnt = 0
do while ncnt <= UBOUND(arrsample)
msgbox arrsample(ncnt)
ncnt = ncnt + 1
loop