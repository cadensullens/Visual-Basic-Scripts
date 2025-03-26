Attribute VB_Name = "BuildComp"
Public iterateString As String
Public MarginStart As Double
Public Increments As Double
Public MessUps() As String
Public Mess As Double
Public Sub Build_Comp()

On Error GoTo Errhandler

NumberHose = 0
onIt = 0
ThisWorkbook.Connections("Query - BOMMaster").Refresh

CheckEntry.Show

GoTo EndSub

Errhandler:

If errNum = 0 Then
MsgBox ("Error Gathering Component Information")
GoTo EndSub
End If


EndSub:
ThisWorkbook.Connections("Query - BOMMaster").Refresh
End Sub

