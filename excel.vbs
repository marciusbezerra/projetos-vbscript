'
' This sample will display Windows Scripting Host properties in Excel.
'
strMessage = "This script demonstrates how to use the WSHNetwork object."
strTitle   = "Windows Scripting Host Sample"
Call Welcome()
'
' Excel Sample
'
Dim objXL
Set objXL = WScript.CreateObject("Excel.Application")

With objXL
    .Visible = TRUE
    .WorkBooks.Add

    .Columns(1).ColumnWidth = 20
    .Columns(2).ColumnWidth = 30
    .Columns(3).ColumnWidth = 40

    .Cells(1, 1).Value = "Property Name"
    .Cells(1, 2).Value = "Value"
    .Cells(1, 3).Value = "Description"

    .Range("A1:C1").Select
    .Selection.Font.Bold = True
    .Selection.Interior.ColorIndex = 1
    .Selection.Interior.Pattern = 1 'xlSolid
    .Selection.Font.ColorIndex = 2

    .Columns("B:B").Select
    .Selection.HorizontalAlignment = &hFFFFEFDD ' xlLeft
End With

Dim intIndex
intIndex = 2

Sub Show(strName, strValue, strDesc)
    objXL.Cells(intIndex, 1).Value = strName
    objXL.Cells(intIndex, 2).Value = strValue
    objXL.Cells(intIndex, 3).Value = strDesc
    intIndex = intIndex + 1
    objXL.Cells(intIndex, 1).Select
End Sub
'
' Show WScript properties
'
Call Show("Name",        WScript.Name,        "Application Friendly Name")
Call Show("Version",     WScript.Version,     "Application Version")
Call Show("FullName",    WScript.FullName,    "Application Context: Fully Qualified Name")
Call Show("Path",        WScript.Path,        "Application Context: Path Only")
Call Show("Interactive", WScript.Interactive, "State of Interactive Mode")
'
' Show command line arguments.
'
Dim colArgs
Set colArgs = WScript.Arguments
Call Show("Arguments.Count", colArgs.Count, "Number of command line arguments")

For i = 0 to colArgs.Count - 1
    objXL.Cells(intIndex, 1).Value = "Arguments(" & i & ")"
    objXL.Cells(intIndex, 2).Value = colArgs(i)
    intIndex = intIndex + 1
    objXL.Cells(intIndex, 1).Select
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
