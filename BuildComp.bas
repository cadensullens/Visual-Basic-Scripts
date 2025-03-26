Attribute VB_Name = "BuildComp"
Public iterateString As String
Public Sub Build_Comp()

On Error GoTo Errhandler

NumberHose = 0
ThisWorkbook.Connections("Query - BOMMaster").Refresh

If BuildSkip = 1 Then GoTo Skip

hose = Application.InputBox( _
Title:="Hose Name" & " " & i, _
Prompt:="What is the name of the hose?", _
Type:=1 + 2)

If hose = "0" Or hose = "False" Then GoTo EndSub

Skip:
Call Check_Comp
If PartErr = 1 Then GoTo EndSub
If CompNumb = False Then GoTo EndSub
If part = "0" Or part = "False" Then GoTo EndSub

Call Ask_Qty(PartNames)
If VarType(comp) = 11 Then GoTo EndSub

CheckEntry.Show

Call WireHole_Barb
If VarType(WireHole) = 11 Then GoTo EndSub
If VarType(BarbRoy) = 11 Then GoTo EndSub

Call DateEntry
If DueDate = "False" Then GoTo EndSub

Call PriceBreaksFunc
If priceend = 1 Then GoTo EndSub

Call Open_BOM

SkipD:
iterate = 0
Call Gather_Component_Info(PartNames)
PartInfo.Show



GoTo EndSub

Errhandler:

If errNum = 0 Then
MsgBox ("Error Gathering Component Information")
GoTo EndSub
End If


EndSub:
ThisWorkbook.Connections("Query - BOMMaster").Refresh
End Sub

