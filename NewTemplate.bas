Attribute VB_Name = "NewTemplate"
Public NumberHose As Double
Public newName As String
Public SkipForward As Boolean

Sub Copy_AnotherSheet()

Dim errNum As Integer

On Error GoTo Errhandler

If SkipForward = False Then GoTo Forward
errNum = 0
Call newQuoteSheet
If newName = "False" Then GoTo EndProc
SkipForward = False
errNum = 2
'Getting Hose name to fill in cells for formulas
Call Enter_Comp
GoTo EndSub
Forward:
SkipForward = True


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
newName = ""
'Reset Copy_AnotherSheet
If copyTemp <> 0 Then
copyTemp = 0
BuySell = 0
End If
EndSub:

End Sub

