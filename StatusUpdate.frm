VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatusUpdate 
   Caption         =   "UserForm1"
   ClientHeight    =   1450
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11028
   OleObjectBlob   =   "StatusUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StatusUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Dim i As Double

If MakerTrue = 0 Then
StatusUpdate.Caption = "Maker Order Status"
StatusUpdate.Label4.Caption = "Complete By Date"
Else
StatusUpdate.Caption = "Backlog Status"
StatusUpdate.Label4.Caption = "Recovery Date"
End If

StatusUpdate.Height = (POHits + 1) * 18 + 84

ExitInactive.Top = (POHits + 1) * 18 + 14
NewHoseInactive.Top = (POHits + 1) * 18 + 14
SaveExistInactive.Top = (POHits + 1) * 18 + 14
SaveNewInactive.Top = (POHits + 1) * 18 + 14
ExitActive.Top = (POHits + 1) * 18 + 14
NewHoseActive.Top = (POHits + 1) * 18 + 14
SaveExistActive.Top = (POHits + 1) * 18 + 14
SaveNewActive.Top = (POHits + 1) * 18 + 14

If StatusIterate = POCount Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

For i = 1 To POHits

    'PO text boxes
    With StatusUpdate.Controls.Add("Forms.TextBox.1", "PONumber" & i)
    .Top = 10 + (i) * 18
    .Left = 6
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = Po
    .SpecialEffect = 0
    End With
    
    'SO Text Boxes
    With StatusUpdate.Controls.Add("Forms.TextBox.1", "SONumber" & i)
    .Top = 10 + (i) * 18
    .Left = 108
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = SONumber(i)
    .SpecialEffect = 0
    End With
    
    'Customer Date boxes
    With StatusUpdate.Controls.Add("Forms.TextBox.1", "CustDate" & i)
    .Top = 10 + (i) * 18
    .Left = 210
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = CustDate(i)
    .SpecialEffect = 0
    End With
    
    'Complete by Date boxes
    With StatusUpdate.Controls.Add("Forms.TextBox.1", "CompDate" & i)
    .Top = 10 + (i) * 18
    .Left = 312
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = CompDate(i)
    .SpecialEffect = 0
    End With
    
    'QTY Boxes
    With StatusUpdate.Controls.Add("Forms.TextBox.1", "BuildQty" & i)
    .Top = 10 + (i) * 18
    .Left = 414
    .Width = 30
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = BuildQty(i)
    .SpecialEffect = 0
    End With
    
    'Released Status
    With StatusUpdate.Controls.Add("Forms.TextBox.1", "JobStatus" & i)
    .Top = 10 + (i) * 18
    .Left = 446
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = JobStat(i)
    .SpecialEffect = 0
    End With
Next i
    
End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ExitInactive.Visible = True
SaveExistInactive.Visible = True
SaveNewInactive.Visible = True

End Sub
Sub ExitInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

ExitInactive.Visible = False

If StatusIterate = POCount Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

SaveExistInactive.Visible = True
SaveNewInactive.Visible = True

End Sub

Sub NewHoseInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ExitInactive.Visible = True
NewHoseInactive.Visible = False
SaveExistInactive.Visible = True
SaveNewInactive.Visible = True

End Sub
Sub SaveExistInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ExitInactive.Visible = True

If StatusIterate = POCount Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

SaveExistInactive.Visible = False
SaveNewInactive.Visible = True

End Sub
Sub SaveNewInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ExitInactive.Visible = True

If StatusIterate = POCount Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

SaveExistInactive.Visible = True
SaveNewInactive.Visible = False

End Sub

Private Sub ExitActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload StatusUpdate
End Sub

Private Sub NewHoseActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload StatusUpdate
wb.Close False
Call Check_Status
End Sub

Private Sub SaveExistActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload StatusUpdate
wb.Close False
copyTemp = 1
Call CopyStatus_Template
End Sub

Private Sub SaveNewActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload StatusUpdate
wb.Close False
copyTemp = 2
Call CopyStatus_Template
End Sub
