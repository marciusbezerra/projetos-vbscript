'
' Write/delete entries in the registry. 
'
'
' Display a welcome message.
'
strMessage = "This script shows how to create and delete registry keys."
strTitle   = "Windows Scripting Host Sample"
Call Welcome()
'
' Registry related methods.
'
Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")

WSHShell.Popup "Create key HKCU\MyRegKey with value 'Top level key'"
WSHShell.RegWrite "HKCU\MyRegKey\", "Top level key"

WSHShell.Popup "Create key HKCU\MyRegKey\Entry with value 'Second level key'"
WSHShell.RegWrite "HKCU\MyRegKey\Entry\", "Second level key"

WSHShell.Popup "Set value HKCU\MyRegKey\Value to REG_SZ 1"
WSHShell.RegWrite "HKCU\MyRegKey\Value", 1

WSHShell.Popup "Set value HKCU\MyRegKey\Entry to REG_DWORD 2"
WSHShell.RegWrite "HKCU\MyRegKey\Entry", 2, "REG_DWORD"

WSHShell.Popup "Set value HKCU\MyRegKey\Entry\Value1 to REG_BINARY 3"
WSHShell.RegWrite "HKCU\MyRegKey\Entry\Value1", 3, "REG_BINARY"

WSHShell.Popup "Delete value HKCU\MyRegKey\Entry\Value1"
WSHShell.RegDelete "HKCU\MyRegKey\Entry\Value1"

WSHShell.Popup "Delete key HKCU\MyRegKey\Entry"
WSHShell.RegDelete "HKCU\MyRegKey\Entry\"

WSHShell.Popup "Delete key HKCU\MyRegKey"
WSHShell.RegDelete "HKCU\MyRegKey\"
'
' Sub to display the Welcome message.
'
Sub Welcome()
    Dim intAns
    intAns = MsgBox(strMessage, vbOKCancel + vbInformation, strTitle)
    If intAns = vbCancel Then
        WScript.Quit
    End If
End Sub
