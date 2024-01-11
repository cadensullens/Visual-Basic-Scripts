Attribute VB_Name = "CopyTemplate"
Public NumberHose As Double
Public newName As String

Sub Copy_AnotherSheet()

'Dim newName As String
Dim TargetRange As Range
Dim index As Double
Dim secIndex As String
Dim locR As Variant
Dim locC As Variant
Dim partnametemp() As String
Dim errNum As Integer


On Error GoTo Errhandler
' variable from Part info 0 makes a new sheet for saving the data 1 uses an existing sheet
If copyTemp = 0 Then
GoTo Start1
ElseIf copyTemp > 0 Then
GoTo FunctionCall
End If
'Name the new sheet
Start1:
errNum = 0
Call newQuoteSheet
If newName = "False" Then GoTo EndProc

errNum = 2
'Getting Hose name to fill in cells for formulas
Call HoseInfo
'if cancel is clicked then ends process
If HoseErr = 1 Then
Worksheets(newName).Delete
GoTo EndProc
End If

If NumberHose = False Then
Worksheets(newName).Delete
GoTo EndProc
End If

If hose = "0" Or hose = "False" Then
Worksheets(newName).Delete
GoTo EndProc
End If

Call DateEntry
If LeadEntry = "False" Then
Worksheets(newName).Delete
GoTo EndProc
End If

Call PriceBreaksFunc
If priceend = 1 Then
Worksheets(newName).Delete
GoTo EndProc
End If

    
FunctionCall:
Call copy_table(copyTemp, BuySell, hose)

GoTo EndProc

Errhandler:
If errNum = 0 Then
MsgBox ("Cannot Enter Zero for component amount. Please Try Again")
GoTo EndProc
End If

If errNum = 1 Then
MsgBox (CStr(errNum))
GoTo EndProc
End If

If errNum = 3 Then
MsgBox (CStr(errNum))
GoTo EndProc
End If

If errNum = 4 Then
MsgBox (CStr(errNum))
GoTo EndProc
End If

If errNum = 5 Then
MsgBox ("Cancel was selected or Error was found")
GoTo EndProc
End If

If errNum = 6 Then
MsgBox ("Cancel was selected or Error was found")
GoTo EndProc
End If

If errNum = 7 Then
MsgBox (CStr(errNum))
GoTo EndProc
End If


EndProc:
'Ends the copy mode of the selected cells
Application.CutCopyMode = False

' Reset number hose
NumberHose = 0
'Reset Copy_AnotherSheet
If copyTemp <> 0 Then
copyTemp = 0
BuySell = 0
End If

End Sub

