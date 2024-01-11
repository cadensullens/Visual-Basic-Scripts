VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheckEntry 
   Caption         =   "Build a Hose Part Check"
   ClientHeight    =   6200
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3516
   OleObjectBlob   =   "CheckEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CheckEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ContinueActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim i As Double

For i = 1 To CompNumb
PartNames(i) = CheckEntry.Controls("Component" & i).Value
compQTY(i) = CheckEntry.Controls("QTY" & i).Value
Next i

Unload CheckEntry
End Sub

Sub ContinueInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ContinueInactive.Visible = False

End Sub

Private Sub HoseLabel_Click()

End Sub

Private Sub UserForm_Initialize()

Dim i As Double
CheckEntry.HoseLabel.Caption = "Hose: " & hose

CheckEntry.Height = (CompNumb + 2) * 18 + 90

ContinueInactive.Top = (CompNumb + 2) * 18 + 14
ContinueActive.Top = (CompNumb + 2) * 18 + 14

For i = 1 To CompNumb
With CheckEntry.Controls.Add("Forms.TextBox.1", "Component" & i)
    .Top = 30 + (i) * 18
    .Left = 12
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = PartNames(i)
    .SpecialEffect = 0
    End With
    
With CheckEntry.Controls.Add("Forms.TextBox.1", "QTY" & i)
    .Top = 30 + (i) * 18
    .Left = 120
    .Width = 50
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = compQTY(i)
    .SpecialEffect = 0
    End With
Next i
    


End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ContinueInactive.Visible = True

End Sub

