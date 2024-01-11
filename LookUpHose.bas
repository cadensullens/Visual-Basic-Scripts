Attribute VB_Name = "LookUpHose"
Public PartNames() As String
Public DueDate As String
Public partQty() As Double
Public Number As Double
Public breakCount As Double
Public hose As String
Public copyTemp As Double
Public compQTY() As Double
Public Grand() As Double
Public PriceBreaks() As Double
Public LongLead() As Double
Public ShortPartList() As Double
Public BacklogList() As Double
Public onHandList() As Double
Public PriceList() As Double
Public LeadTimeList() As Double
Public max As Double
Public priceend As Double
Public iterate As Double
Public WireHole As Variant
Public Grandsum As Double
Public BarbRoy As Variant


'Code needs to be directed for if a hose is not found on list 9/22
Public Sub Enter_Comp()


Call HoseInfo
If HoseErr = 1 Then GoTo EndSub
If NumberHose = False Then GoTo EndSub
If hose = "0" Or hose = "False" Then GoTo EndSub

If NumberHose = 1 Then
Call Buy_Sell(hose)
If BuySell = 1 Then GoTo SkipD
End If

Call DateEntry
If LeadEntry = "False" Then GoTo EndSub

Call PriceBreaksFunc
If priceend = 1 Then GoTo EndSub

SkipD:
For i = 1 To NumberHose
iterate = i
hose = hoseNames(i)
Call Gather_Info(hoseNames(i))

If Gathererr = 1 And i <> NumberHose Then
GoTo pass
ElseIf Gathererr = 0 Then GoTo skipE
Else
GoTo EndSub
End If

skipE:
If BuySell <> 1 Then
PartInfo.Show
Else
BuySellInfo.Show
End If

pass:
Next i

EndSub:
End Sub
