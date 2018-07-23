dim AscEncode
call AscEncode(str)
Function AscEncode(str)
    Dim i
    Dim sAscii
    
    sAscii = ""

    For i = 1 To Len(str)
        sAscii = sAscii + CStr(Asc(Mid(str, i, 1))) & ", "
'The CStr function converts an expression to type String 
    Next
    
    AscEncode = sAscii
    MsgBox AscEncode

End Function
