Attribute VB_Name = "CheckStatus"
Public POCount As Double
Public PONumber As String
Public POHits As Double
Public PO As String
Public StatusIterate As Double

Public Sub Check_Status()

Dim POArray() As String

On Error GoTo Errhandler

POCount = Application.InputBox( _
Title:="PO Status", _
Prompt:="How many Customer POs are you looking up?", _
Type:=1)

If VarType(POCount) = 11 Then GoTo EndSub

For i = 1 To POCount
PONumber = Application.InputBox( _
Title:="PO Status " & i, _
Prompt:="What is the Customer PO #?", _
Type:=1 + 2)

ReDim Preserve POArray(1 To i)
POArray(i) = PONumber
Next i
If VarType(PONumber) = 11 Then GoTo EndSub


For d = 1 To POCount
StatusIterate = d
PO = POArray(d)
    Call MakerStatus(POArray(d))
    Call BackLogCheck(POArray(d))
Repeat:
Next d

d = 0
GoTo EndSub

Errhandler:

EndSub:
If d > 0 Then GoTo Repeat
End Sub
