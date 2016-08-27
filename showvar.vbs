'
' List all environment variables defined on this machine.
'
'
' Display a welcome message.
'
strMessage = "This script will list all environment variables defined on this machine."
strTitle   = "Windows Scripting Host Sample"
Call Welcome()
'
' Environment Sample
'
CRLF = Chr(13) & Chr(10)

Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")

intIndex = 0
strText = ""
intNumEnv = 0
MAX_ENV = 20
'
' Loop through the environment variables.
'
For Each strEnv In WshShell.Environment("PROCESS")
    intIndex = intIndex + 1
    strText = strText & CRLF & Right("    " & intIndex, 4) & " " & strEnv
    intNumEnv = intNumEnv + 1

    If intNumEnv >= MAX_ENV Then
        call MsgBox (strText, vbInformation, strTitle)

        strText = ""
        intNumEnv = 0
    End If
Next

If intNumEnv >= 1 Then 
    call MsgBox (strText, vbInformation, strTitle)
end if
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

