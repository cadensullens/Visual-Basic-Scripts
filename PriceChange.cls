VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents newPriceBox As MSForms.TextBox
Attribute newPriceBox.VB_VarHelpID = -1
Public WithEvents newGrandBox As MSForms.TextBox
Attribute newGrandBox.VB_VarHelpID = -1
Public WithEvents newMarginBox As MSForms.TextBox
Attribute newMarginBox.VB_VarHelpID = -1
Public WithEvents newCleanPrice As MSForms.TextBox
Attribute newCleanPrice.VB_VarHelpID = -1

Public Sub newPriceBox_Change()
If newPriceBox = "" Then GoTo EndSub
If Not IsNumeric(newPriceBox) Then
    newPriceBox.Value = 0
Else
    If newPriceBox.Value > 0 Then
        newPriceBox.BackColor = &H80000005
    Else
        newPriceBox.BackColor = &HC0C0FF
    End If
End If
Grandsum = 0
    For i = 1 To UBound(PartNames)
    Dollar = PartInfo.Controls("Price" & i).Value
    Grandsum = Grandsum + (Dollar * PartInfo.Controls("Qty" & i).Value)
    Next i
    Grandsum = Grandsum + (10 * WireHole) + BarbRoy
    PartInfo.Grandtext.Value = Round(Grandsum, 2)
EndSub:
End Sub
Public Sub newCleanPrice_Change()
If newCleanPrice = "" Then GoTo EndSub
If Not IsNumeric(newCleanPrice) Then
    newCleanPrice.Value = 0
Else
    If newCleanPrice.Value > 0 Then
        newCleanPrice.BackColor = &H80000005
    Else
        newCleanPrice.BackColor = &HC0C0FF
    End If
End If

Grandsum = 0
    For i = 1 To UBound(PartNames)
    Dollar = PartInfo.Controls("Price" & i).Value
    Grandsum = Grandsum + (Dollar * PartInfo.Controls("Qty" & i).Value)
    Next i
Grandsum = Grandsum + CDbl(newCleanPrice.Value) + (10 * WireHole) + BarbRoy
PartInfo.Grandtext.Value = Round(Grandsum, 2)
EndSub:
End Sub
Public Sub newGrandBox_Change()
If newGrandBox = "" Then GoTo EndSub

If Not IsNumeric(newGrandBox) Then
    newGrandBox = 0
    GoTo EndSub
Else
    If newGrandBox.Value > 0 Then
        newGrandBox.BackColor = &H80000005
    Else
        newGrandBox.BackColor = &HC0C0FF
        GoTo EndSub
    End If
End If
PartInfo.SellPrice1.Value = Round(PartInfo.Grandtext.Value / (1 - (PartInfo.MarginNumb1.Value / 100)), 2)
PartInfo.SellPrice2.Value = Round(PartInfo.Grandtext.Value / (1 - (PartInfo.MarginNumb2.Value / 100)), 2)
PartInfo.SellPrice3.Value = Round(PartInfo.Grandtext.Value / (1 - (PartInfo.MarginNumb3.Value / 100)), 2)
PartInfo.SellPrice4.Value = Round(PartInfo.Grandtext.Value / (1 - (PartInfo.MarginNumb4.Value / 100)), 2)
PartInfo.SellPrice5.Value = Round(PartInfo.Grandtext.Value / (1 - (PartInfo.MarginNumb5.Value / 100)), 2)
EndSub:
End Sub

Public Sub newMarginBox_Change()
If newMarginBox.Value = "" Then GoTo EndSub
If newMarginBox.Value = 100 Then
newMarginBox.Value = 0
GoTo EndSub
ElseIf newMarginBox.Value = 0 Then
newMarginBox.BackColor = &HC0C0FF
GoTo EndSub
Else
newMarginBox.BackColor = &H80000005
End If
If Not IsNumeric(newMarginBox) Then
    newMarginBox = 0
    GoTo EndSub
End If
Digit = Right(newMarginBox.Name, 1)
Sell = "SellPrice" & Digit


PartInfo.Controls(Sell).Value = Round(PartInfo.Grandtext.Value / (1 - (newMarginBox.Value / 100)), 2)
EndSub:
End Sub
