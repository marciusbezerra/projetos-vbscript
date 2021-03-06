'
' Read network properties (username and computername), 
' connects, disconnects, and enumerates network drives.
'
strMessage = "This script demonstrates how to use the WSHNetwork object."
strTitle   = "Windows Scripting Host Sample"
Call Welcome()
'
' WSH Network Object.
'
Dim WSHNetwork
Dim colDrives, SharePoint
Dim CRLF

CRLF = Chr(13) & Chr(10)
Set WSHNetwork = WScript.CreateObject("WScript.Network")


Function Ask(strAction)
   '
   ' This function asks the user whether to perform a specific "Action"
   ' and sets a return code or quits script execution depending on the 
   ' button that the user presses.  This function is called at various
   ' points in the script below.
   '
   Dim intButton
   intButton = MsgBox(strAction, vbQuestion + vbYesNo, strTitle)
   Ask = intButton = vbYes
End Function
'
' Show WSHNetwork object properties
'
MsgBox "UserDomain"     & Chr(9) & "= " & WSHNetwork.UserDomain  & CRLF &   _
       "UserName"       & Chr(9) & "= " & WSHNetwork.UserName    & CRLF &   _
       "ComputerName"   & Chr(9) & "= " & WSHNetwork.ComputerName,          _
       vbInformation + vbOKOnly,                                            _
       "WSHNetwork Properties"
'
' WSHNetwork.AddNetworkDrive
'
Function TryMapDrive(intDrive, strShare)
    Dim strDrive
    strDrive = Chr(intDrive + 64) & ":"
    On Error Resume Next
    WSHNetwork.MapNetworkDrive strDrive, strShare
    TryMapDrive = Err.Number = 0
End Function

If Ask("Do you want to connect a network drive?") Then
    strShare = InputBox("Enter network share you want to connect to ")
    For intDrive = 26 To 5 Step -1
        If TryMapDrive(intDrive, strShare) Then Exit For
    Next

    If intDrive <= 5 Then
        MsgBox "Unable to connect to network share. "                        & _
               "There are currently no drive letters available for use. "    & _
               CRLF                                                          & _
               "Please disconnect one of your existing network connections " & _
               "and try this script again. ",                                  _
               vbExclamation + vbOkOnly, strTitle
    Else
        strDrive = Chr(intDrive + 64) & ":"
        MsgBox "Connected " & strShare & " to drive " & strDrive,   _
               vbInformation + vbOkOnly, strTitle

        If Ask("Do you want to disconnect the network drive you just created?") Then
            WSHNetwork.RemoveNetworkDrive strDrive

            MsgBox "Disconnected drive " & strDrive,        _
                   vbInformation + vbOkOnly, strTitle
        End If
    End If
End If
'
' WSHNetwork.EnumNetworkDrive
'
' Ask user whether to enumerate network drives
'
If Ask("Do you want to enumerate connected network drives?") Then
    '
    'Enumerate network drives into a collection object of type WshCollection
    '
    Set colDrives = WSHNetwork.EnumNetworkDrives
    '
    ' If no network drives were enumerated, then inform 
    ' user, else display enumerated drives.
    '
    If colDrives.Count = 0 Then
        MsgBox "There are no drives to enumerate.",     _
               vbInformation + vbOkOnly, strTitle
    Else
        strMsg = "Current network drive connections: " & CRLF
        For i = 0 To colDrives.Count - 1 Step 2
            strMsg = strMsg & CRLF & colDrives(i) & Chr(9) & colDrives(i + 1)
        Next
        
        MsgBox strMsg, vbInformation + vbOkOnly, strTitle
    End If
End If
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

