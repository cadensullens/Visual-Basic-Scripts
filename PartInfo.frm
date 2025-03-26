VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PartInfo 
   Caption         =   "Hose Information"
   ClientHeight    =   2150
   ClientLeft      =   96
   ClientTop       =   420
   ClientWidth     =   11172
   OleObjectBlob   =   "PartInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PartInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ExitInactive.Visible = True
SaveExistInactive.Visible = True
SaveNewInactive.Visible = True

End Sub
Sub ExitInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

ExitInactive.Visible = False

If iterate = NumberHose Then
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

If iterate = NumberHose Then
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

If iterate = NumberHose Then
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
Unload PartInfo
End Sub

Private Sub NewHoseActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload PartInfo
Call Enter_Comp
End Sub

Private Sub SaveExistActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload PartInfo
copyTemp = 1
Call Copy_AnotherSheet
End Sub

Private Sub SaveNewActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload PartInfo
copyTemp = 2
Call Copy_AnotherSheet
End Sub

Private Sub UserForm_Initialize()
Dim i As Double

If iterate = NumberHose Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

'Setting placement and size of userform
PartInfo.Height = (UBound(PartNames) + 5) * 18 + 46
PartInfo.Width = (7 + breakCount) * 60 + 150

ExitInactive.Top = (UBound(PartNames) + 1) * 18 + 14
NewHoseInactive.Top = (UBound(PartNames) + 1) * 18 + 14
SaveExistInactive.Top = (UBound(PartNames) + 1) * 18 + 14
SaveNewInactive.Top = (UBound(PartNames) + 1) * 18 + 14
ExitActive.Top = (UBound(PartNames) + 1) * 18 + 14
NewHoseActive.Top = (UBound(PartNames) + 1) * 18 + 14
SaveExistActive.Top = (UBound(PartNames) + 1) * 18 + 14
SaveNewActive.Top = (UBound(PartNames) + 1) * 18 + 14
DueDate.Top = (UBound(PartNames) + 1) * 18 + 14
DateEnter.Top = (UBound(PartNames) + 1) * 18 + 14
Target.Top = (UBound(PartNames) + 2) * 18 + 14
Leadtime.Top = (UBound(PartNames) + 2) * 18 + 14
Grand.Top = (UBound(PartNames) + 3) * 18 + 14
GTotal.Top = (UBound(PartNames) + 3) * 18 + 14
LongLead.Top = (UBound(PartNames) + 4) * 18 + 14
Longest.Top = (UBound(PartNames) + 4) * 18 + 14

If breakCount <> 0 Then
ReDim PriceBreaks(1 To UBound(PartNames), 1 To breakCount)
End If

For i = 1 To UBound(PartNames)

    With PartInfo.Controls.Add("Forms.Label.1", "Component" & i)
    .Top = 6 + (i) * 18
    .Left = 12
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Bold = True
    .ForeColor = &HFFFFFF
    .BackColor = &HA77E00
    .BackStyle = 1
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Caption = PartNames(i)
    .SpecialEffect = 0
    .TextAlign = 2
    End With
    
    'PO text boxes
    With PartInfo.Controls.Add("Forms.TextBox.1", "QTY" & i)
    .Top = 6 + (i) * 18
    .Left = 120
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = compQTY(i)
    .SpecialEffect = 0
    End With
    
    'SO Text Boxes
    With PartInfo.Controls.Add("Forms.TextBox.1", "Price" & i)
    .Top = 6 + (i) * 18
    .Left = 180
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = "$" & Round(PriceList(i), 2)
    .SpecialEffect = 0
    End With
    
    'Customer Date boxes
    With PartInfo.Controls.Add("Forms.TextBox.1", "OnHand" & i)
    .Top = 6 + (i) * 18
    .Left = 240
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = onHandList(i)
    .SpecialEffect = 0
    End With
    
    'Complete by Date boxes
    With PartInfo.Controls.Add("Forms.TextBox.1", "BackLog" & i)
    .Top = 6 + (i) * 18
    .Left = 300
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = BacklogList(i)
    .SpecialEffect = 0
    End With
    
    'QTY Boxes
    With PartInfo.Controls.Add("Forms.TextBox.1", "Short" & i)
    .Top = 6 + (i) * 18
    .Left = 360
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = ShortPartList(i)
    .SpecialEffect = 0
    End With
    
    'Released Status
    With PartInfo.Controls.Add("Forms.TextBox.1", "Diff" & i)
    .Top = 6 + (i) * 18
    .Left = 420
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = Round((CDbl(BacklogList(i)) + CDbl(onHandList(i))) - CDbl(ShortPartList(i)), 2)
    .SpecialEffect = 0
    End With
    
        'Released Status
    With PartInfo.Controls.Add("Forms.TextBox.1", "LeadTime" & i)
    .Top = 6 + (i) * 18
    .Left = 480
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = LeadTimeList(i) & " Weeks"
    .SpecialEffect = 0
    End With
    
    If breakCount <> 0 Then
    'Creates qtys for price break points amounts
    For j = 1 To breakCount
    
    PriceBreaks(i, j) = (CDbl(BacklogList(i)) + CDbl(onHandList(i))) - (CDbl(ShortPartList(i)) + (partQty(j) * compQTY(i)))
    
    Next j
    'For Loop to enter in Price Break amounts
    For k = 1 To breakCount
        With PartInfo.Controls.Add("Forms.TextBox.1", "PriceBreak" & i * k)
        .Top = 6 + (i) * 18
        .Left = 480 + 60 * k
        .Width = 57
        .Height = 18
        .Font.Name = "Calibri"
        .Font.Size = 10
        .ForeColor = &H464646
        .BorderStyle = 1
        .BorderColor = &HA9A9A9
        .Value = PriceBreaks(i, k)
        .SpecialEffect = 0
        End With
        
        With PartInfo.Controls.Add("Forms.Label.1", "PB" & k)
        .Top = 6
        .Left = 480 + 60 * k
        .Width = 57
        .Height = 16
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = &HFFFFFF
        .BackColor = &HA77E00
        .BackStyle = 1
        .BorderStyle = 1
        .BorderColor = &HA9A9A9
        .Caption = "Break " & k
        .SpecialEffect = 0
        .TextAlign = 2
        End With
    Next k
    End If
    
Next i


End Sub


