Option Explicit
dim objDict
set objDict = CreateObject("Scripting.Dictionary")
Msgbox "The dictionary object holds this many items:" & _
objDict.Count