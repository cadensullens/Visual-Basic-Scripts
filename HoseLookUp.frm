VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HoseLookUp 
   Caption         =   "Look Up a Hose "
   ClientHeight    =   3720
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8340.001
   OleObjectBlob   =   "HoseLookUp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HoseLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TBox() As New clsTBox
Private PBox() As New PasteBox

Private Sub UserForm_Initialize()

HoseLookUp.Width = 428

HoseLookUp.Height = 7 * 18 + 96
'Default values
ReDim TBox(0)
'ReDim PBox(4)
OGBreak = 0
FloatValue = 1
breakCounts.Value = 1
DateBox.Value = "12/12/9999"
Margin.Value = "40"
Increment.Value = "1"
HoseName.Value = ""
Set PBox(0).PCCBox = HoseName
Set PBox(1).PCCBox = Lead
Set PBox(2).PCCBox = Margin
Set PBox(3).PCCBox = Increment
Set PBox(4).PCCBox = breakCounts


End Sub
Sub ContinueInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ContinueInactive.Visible = False

End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ContinueInactive.Visible = True

End Sub
Public Sub ContinueActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If HoseName.Value = "" Then
MsgBox ("Hose Name is Blank")
GoTo EndSub
End If

Call FixShortParts.NameFix

If newName <> "" Then
Call HoseInfo(HoseName.Value)
If HoseErr = 0 Then
    SaveData
    If PriceWrong > 0 Then GoTo EndSub
        Unload HoseLookUp
        Call Buy_Sell(HoseName.Value)
        'Buy sell has diff userform
        If BuySell = 1 Then
            'Goes to template sheet and copies the formatted table
            Call BuySell_CopyTable(4, 1, newName)
        Else
            CopyCheck = 1
            Call Gather_Info(hose)
            Call BOM_CopyTable(4, 1, newName)
        End If
        Call saveHose(hose, 4, 1, newName)
        Call Copy_AnotherSheet
 End If
Else
    Call HoseInfo(HoseName.Value)
    If HoseErr = 0 Then
        SaveData
        If PriceWrong > 0 Then GoTo EndSub
        Call Buy_Sell(HoseName.Value)
        'Buy sell has diff userform
        If BuySell = 1 Then
            Unload HoseLookUp
            Call Buy_Sell_Fill(HoseName.Value)
            BuySellInfo.Show
            GoTo EndSub
        Else
            Call Gather_Info(HoseName.Value)
            Unload HoseLookUp
            PartInfo.Show
        End If
    Else
        Unload HoseLookUp
        hose = HoseName.Value
        MakerBuyChoice.Show
    End If
End If
copyTemp = 0
EndSub:
End Sub

Private Sub Lead_Change()
UpdateDate
End Sub

Public Sub Update()
Dim i As Double


If FloatValue > OGBreak Then
StartValue = OGBreak + 1
End If


For i = StartValue To breakCounts.Value


With HoseLookUp.Controls.Add("Forms.Label.1", "PriceLabel" & i)
    .Top = 108 + (i - 1) * 20
    .Left = 6
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Bold = True
    .ForeColor = &HFFFFFF
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .Caption = "Price Break #" & i
    .BackColor = &HA77E00
    .TextAlign = 2
    End With
    
Set BreakBox = HoseLookUp.Controls.Add("Forms.TextBox.1", "break" & i)
With BreakBox
    .Top = 108 + (i - 1) * 20
    .Left = 108
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .TabIndex = i + 6
    .Value = 1
    End With
Set TBox(i).newTBox = BreakBox
Set PBox(i + 4).PCCBox = BreakBox

Next i

HoseLookUp.Height = (breakCounts.Value + 6) * 18 + 76 + (2 * (breakCounts.Value + 1))
ContinueInactive.Top = (breakCounts.Value + 6) * 18 + (2 * (breakCounts.Value + 1))
ContinueActive.Top = (breakCounts.Value + 6) * 18 + (2 * (breakCounts.Value + 1))
EndSub:
End Sub

Public Sub breakCounts_Change()
If Not IsNumeric(breakCounts.Value) Then
breakCounts.Value = ""
Else
If breakCounts.Value = 0 Then
breakCounts.Value = 1
Else
If breakCounts.Value < OGBreak Then
RemovePriceBoxes
Else
If breakCounts.Value = OGBreak Then
GoTo EndSub
Else
FloatValue = breakCounts.Value
ReDim Preserve TBox(breakCounts.Value)
ReDim Preserve PBox(4 + breakCounts.Value)
Update
OGBreak = breakCounts.Value
End If
End If
End If
End If


EndSub:
End Sub

Sub RemovePriceBoxes()

StartV = CDbl(breakCounts.Value) + 1
For i = StartV To OGBreak
    HoseLookUp.Controls.Remove ("PriceLabel" & i)
    HoseLookUp.Controls.Remove ("break" & i)
    Next i
ReDim Preserve TBox(breakCounts.Value)
ReDim Preserve PBox(4 + breakCounts.Value)
OGBreak = breakCounts.Value
HoseLookUp.Height = (breakCounts.Value + 6) * 18 + 76 + (2 * (breakCounts.Value + 1))
ContinueInactive.Top = (breakCounts.Value + 6) * 18 + (2 * (breakCounts.Value + 1))
ContinueActive.Top = (breakCounts.Value + 6) * 18 + (2 * (breakCounts.Value + 1))


End Sub

Sub UpdateDate()
If Not IsNumeric(Lead.Value) Then
Lead.Value = ""
End If
If Lead.Value = "" Then
DateBox.Value = "12/12/9999"
Else
DateBox.Value = Date + (CDbl(Lead.Value) * 7)
End If
EndSub:
End Sub

Sub NumberCheck()
PriceWrong = 0
For i = 1 To UBound(TBox)
If Not IsNumeric(HoseLookUp.Controls("break" & i).Value) Then
HoseLookUp.Controls("break" & i).Value = ""
PriceWrong = PriceWrong + 1
End If
Next i
If PriceWrong > 0 Then MsgBox ("Price Breaks can only be numbers.")

End Sub

Private Sub Margin_Change()
If Margin.Value = "" Then GoTo EndSub
If Not IsNumeric(Margin.Value) Then
Margin.Value = ""
End If
EndSub:
End Sub
Private Sub Increment_Change()
If Increment.Value = "" Then GoTo EndSub
If Not IsNumeric(Increment.Value) Then
Increment.Value = ""
End If
EndSub:
End Sub

Public Sub SaveData()
If HoseName.Value = "" Then GoTo Skip:
Call NumberCheck
If PriceWrong > 0 Then GoTo Skip:
For i = 1 To UBound(TBox)
    
    ReDim Preserve partQty(1 To i)
    partQty(i) = HoseLookUp.Controls("break" & i).Value
    Next i
LeadEntry = Lead.Value
MarginStart = Margin.Value
Increments = Increment.Value
DueDate = DateBox.Value
breakCount = UBound(TBox)
hose = HoseName.Value


Skip:
End Sub

