'
' Access Microsoft Excel using the Windows Scripting Host.
'
strMessage = "This script demonstrates how to use the WSHNetwork object."
strTitle   = "Windows Scripting Host Sample"
Call Welcome()
'
' Excel Sample
'
Dim objXL
Dim objXLchart
Dim intRotate

Set objXL = WScript.CreateObject("Excel.Application")
With objXL
   .Workbooks.Add
   .Cells(1,1).Value = 5
   .Cells(1,2).Value = 10
   .Cells(1,3).Value = 15
   .Range("A1:C1").Select
End With

Set objXLchart = objXL.Charts.Add()
objXL.Visible = True
objXLchart.Type = -4100     

For intRotate = 5 To 180 Step 5
    objXLchart.Rotation = intRotate
Next

For intRotate = 175 To 0 Step -5
    objXLchart.Rotation = intRotate
Next
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

