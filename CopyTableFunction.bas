Attribute VB_Name = "CopyTableFunction"

Function copy_table(copy As Double, Buy As Double, hose As String)
'Local Variables
Dim location As Range
Dim TargetRange As Range
Dim locR As Variant
Dim locC As Variant
Dim i As Double
Dim looper As Double


On Error GoTo Errhandler

'Existing worksheet is copy 1
If copy = 1 Then

Set location = Application.InputBox( _
Title:="Cell select", _
Prompt:="Select the cell to begin Template", _
Type:=8)

'get the row number from the location pick above
locR = location.Row
locC = location.Column

'Gets the sheet name
newName = location.Worksheet.Name

    'Specify what part of template to copy
    If Buy = 1 Then
         Call BuySell_CopyTable(locR, locC, newName)
    Else
        Call BOM_CopyTable(locR, locC, newName)
    End If

Call saveHose(hose, locR, locC, newName)

' New worksheet is copy 2
ElseIf copy = 2 Then

Call newQuoteSheet
locR = 4
locC = 1

'Goes to template sheet and copies the formatted table
    If Buy = 1 Then
        Call BuySell_CopyTable(locR, locC, newName)
    Else
        Call BOM_CopyTable(locR, locC, newName)
    End If
    
    Call saveHose(hose, locR, locC, newName)
Else
locR = 4
locC = 1

'Goes to template sheet and copies the formatted table
    If Buy = 1 Then
        Call BuySell_CopyTable(locR, locC, newName)
    Else
        Call BOM_CopyTable(locR, locC, newName)
    End If
    
    Call saveHose(hose, locR, locC, newName)
End If
GoTo EndFunction


Errhandler:
MsgBox ("Cancel was selected or Error was found")

EndFunction:
newName = ""
End Function
