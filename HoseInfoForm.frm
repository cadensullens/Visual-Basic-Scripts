VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HoseInfoForm 
   Caption         =   "Hose Information"
   ClientHeight    =   4350
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8532.001
   OleObjectBlob   =   "HoseInfoForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HoseInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TBox() As New clsTBox
Private PBox() As New PasteBox

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
ReDim Preserve PBox(5 + breakCounts.Value)
Update
OGBreak = breakCounts.Value
End If
End If
End If
End If


EndSub:
End Sub

Private Sub ContinueActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SaveData
If PriceWrong > 0 Then GoTo EndSub
Unload HoseInfoForm
Call Open_BOM
Call Gather_Component_Info(PartNames)
PartInfo.Show
EndSub:
End Sub

Sub ContinueInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ContinueInactive.Visible = False

End Sub
Private Sub Lead_Change()
UpdateDate
End Sub


Public Sub UserForm_Initialize()

HoseInfoForm.Width = 438

HoseInfoForm.Height = 7 * 18 + 96
ContinueInactive.Top = 7 * 18 + 22
ContinueActive.Top = 7 * 18 + 22
HoseInfoForm.Caption = "Hose Information for " & hose

'Default values
ReDim TBox(0)
OGBreak = 0
FloatValue = 1
breakCounts.Value = 1
DateBox.Value = "12/12/9999"
Margin.Value = "40"
Increment.Value = "1"
Set PBox(0).PCCBox = Wire
Set PBox(1).PCCBox = Lead
Set PBox(2).PCCBox = Margin
Set PBox(3).PCCBox = Increment
Set PBox(4).PCCBox = breakCounts
Set PBox(5).PCCBox = Barb
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ContinueInactive.Visible = True

End Sub

Public Sub Update()
Dim i As Double


If FloatValue > OGBreak Then
StartValue = OGBreak + 1
End If


For i = StartValue To breakCounts.Value


With HoseInfoForm.Controls.Add("Forms.Label.1", "PriceLabel" & i)
    .Top = 128 + (i - 1) * 20
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

Set BreakBox = HoseInfoForm.Controls.Add("Forms.TextBox.1", "break" & i)
With BreakBox
    .Top = 128 + (i - 1) * 20
    .Left = 108
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .TabIndex = i + 7
    .Value = 1
    End With
Set TBox(i).newTBox = BreakBox
Set PBox(i + 5).PCCBox = BreakBox
    

Next i

HoseInfoForm.Height = (breakCounts.Value + 7) * 18 + 76 + (2 * (breakCounts.Value + 1))
ContinueInactive.Top = (breakCounts.Value + 7) * 18 + (2 * (breakCounts.Value + 1))
ContinueActive.Top = (breakCounts.Value + 7) * 18 + (2 * (breakCounts.Value + 1))
EndSub:
End Sub

Sub RemovePriceBoxes()

StartV = CDbl(breakCounts.Value) + 1
For i = StartV To OGBreak
    HoseInfoForm.Controls.Remove ("PriceLabel" & i)
    HoseInfoForm.Controls.Remove ("break" & i)
    Next i
    
ReDim Preserve TBox(breakCounts.Value)
ReDim Preserve PBox(5 + breakCounts.Value)
OGBreak = breakCounts.Value
HoseInfoForm.Height = (breakCounts.Value + 7) * 18 + 76 + (2 * (breakCounts.Value + 1))
ContinueInactive.Top = (breakCounts.Value + 7) * 18 + (2 * (breakCounts.Value + 1))
ContinueActive.Top = (breakCounts.Value + 7) * 18 + (2 * (breakCounts.Value + 1))

End Sub

Sub SaveData()

Call NumberCheck
If PriceWrong > 0 Then GoTo Skip:
For i = 1 To breakCounts.Value
    
    ReDim Preserve partQty(1 To i)
    partQty(i) = HoseInfoForm.Controls("break" & i).Value
    Next i
LeadEntry = Lead.Value
If Margin.Value = "" Then
MarginStart = 0
Else
MarginStart = Margin.Value
End If
If Increment.Value = "" Then
Increments = 0
Else
Increments = Increment.Value
End If
DueDate = DateBox.Value
If Wire.Value = "" Then
WireHole = 0
Else
WireHole = Wire.Value
End If
If Barb.Value = "" Then
BarbRoy = 0
Else
BarbRoy = Barb.Value
End If

breakCount = breakCounts.Value

If CheckBox1.Value = True Then
SpecClean = "Yes"
Else
SpecClean = "No"
End If

Skip:
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
For i = 1 To breakCounts.Value
If Not IsNumeric(HoseInfoForm.Controls("break" & i).Value) Then
HoseLookUp.Controls("break" & i).Value = ""
PriceWrong = PriceWrong + 1
End If
Next i
If PriceWrong > 0 Then MsgBox ("Price Breaks can only be numbers.")

End Sub

Private Sub Wire_Change()
If Wire.Value = "" Then GoTo EndSub
If Not IsNumeric(Wire.Value) Then
Wire.Value = ""
End If
EndSub:
End Sub

Private Sub Barb_Change()
If Barb.Value = "" Then GoTo EndSub
If Not IsNumeric(Barb.Value) Then
Barb.Value = ""
End If
EndSub:
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
