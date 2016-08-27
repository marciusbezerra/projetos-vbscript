'
' Create desktop shortcuts.

strMessage = "This script will create a shortcut to Notepad on your desktop."
strTitle   = "Windows Scripting Host Sample"
Call Welcome()
'
' Shortcut related methods.
'
Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")

Dim MyShortcut, MyDesktop, DesktopPath
'
' Read desktop path using WshSpecialFolders object.
'
DesktopPath = WSHShell.SpecialFolders("Desktop")
'
' Create a shortcut object on the desktop.
'
Set MyShortcut = WSHShell.CreateShortcut(DesktopPath & "\Shortcut to notepad.lnk")
'
' Set shortcut object properties and save it.
'
With MyShortcut
    .TargetPath = WSHShell.ExpandEnvironmentStrings("%windir%\notepad.exe")
    .WorkingDirectory = WSHShell.ExpandEnvironmentStrings("%windir%")
    .WindowStyle = 4
    .IconLocation = WSHShell.ExpandEnvironmentStrings("%windir%\notepad.exe, 0")
    .Save
End with

WScript.Echo "A shortcut to Notepad now exists on your Desktop."
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
