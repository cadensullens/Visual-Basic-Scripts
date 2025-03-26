VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuoteBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public WithEvents newFBox As MSForms.TextBox
Attribute newFBox.VB_VarHelpID = -1
Public WithEvents newFbBox As MSForms.TextBox
Attribute newFbBox.VB_VarHelpID = -1
Public WithEvents newFpBox As MSForms.TextBox
Attribute newFpBox.VB_VarHelpID = -1

Public Sub newFBox_Change()
   If Not IsNumeric(newFBox) Then
    newFBox.Value = ""
    GoTo EndSub
    End If
   Digit = Right(newFBox.Name, 2)
    If Not IsNumeric(Digit) Then
    Digit = Right(newFBox.Name, 1)
    End If
   Sell = "SellPriceQuote" & Digit
   Cost = "HoseCostQuote" & Digit
   If newFBox.Value >= 100 Then newFBox.Value = 99
   If SaveQuote.Controls(Cost).Value <> "" Then
    SaveQuote.Controls(Sell).Value = Round(SaveQuote.Controls(Cost).Value / (1 - (newFBox.Value / 100)), 2)
    End If
    
EndSub:
End Sub

Public Sub newFbBox_Change()
   If Not IsNumeric(newFbBox) Then
    newFbBox.Value = ""
    GoTo EndSub
    End If
    Digit = Right(newFbBox.Name, 2)
    If Not IsNumeric(Digit) Then
        Digit = Right(newFbBox.Name, 1)
    End If
    Sell = "SellPriceQuote" & Digit
    Margin = "MarginQuote" & Digit
    SaveQuote.Controls(Sell).Value = Round(newFbBox.Value / (1 - (SaveQuote.Controls(Margin).Value / 100)), 2)
EndSub:
End Sub

Public Sub newFpBox_Change()
   If Not IsNumeric(newFpBox) Then
    newFpBox.Value = ""
    End If
End Sub

