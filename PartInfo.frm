VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PartInfo 
   Caption         =   "Hose Information"
   ClientHeight    =   3910
   ClientLeft      =   96
   ClientTop       =   420
   ClientWidth     =   15756
   OleObjectBlob   =   "PartInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PartInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PriceBox() As New PriceChange
Private TotalCostBox() As New PriceChange
Private MarginBox() As New PriceChange
Private CleanBox() As New PriceChange
Private PBox() As New PasteBox
Private QBox() As New PasteBox

Public Sub LiveLeadtime_Change()
PartInfoValue = True
If Not IsNumeric(LiveLeadtime.Value) Then
LiveLeadtime.Value = ""
Leadtime.Value = ""
UpdateDate
DueDate = DateEnter.Value
LeadEntry = LiveLeadtime.Value
LiveLeadSkip = True
Call Gather_Info(hose)
Call RemovePriceBoxes
Call Fill_Boxes
Else
UpdateDate
DueDate = DateEnter.Value
LeadEntry = LiveLeadtime.Value
LiveLeadSkip = True
Call Gather_Info(hose)
Call RemovePriceBoxes
Call Fill_Boxes
End If

End Sub
Sub UpdateDate()
If Not IsNumeric(LiveLeadtime.Value) Then
DateEnter.Value = ""
End If
If LiveLeadtime.Value = "" Then
DateEnter.Value = "12/12/9999"
Else
DateEnter.Value = Date + (CDbl(LiveLeadtime.Value) * 7)
End If
End Sub

Private Sub ExitActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload PartInfo
PartInfoValue = False
End Sub

Private Sub NewHoseActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload PartInfo
PartInfoValue = False
Call Enter_Comp
End Sub

Private Sub SaveExistActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SaveData
Unload PartInfo
copyTemp = 1
Call copy_table(copyTemp, BuySell, hose)
End Sub

Private Sub SaveNewActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SaveData
Unload PartInfo
copyTemp = 2
Call copy_table(copyTemp, BuySell, hose)
End Sub

Private Sub UserForm_Initialize()
Dim i As Double

If iterate = NumberHose And copyTemp <> 3 Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

If copyTemp = 3 Then
SaveNewInactive.Visible = False
SaveNewActive.Visible = False
End If

ReDim PriceBox(UBound(PartNames))
ReDim TotalCostBox(1)
ReDim MarginBox(5)
ReDim CleanBox(0)
ReDim PBox(5)
Call Fill_Boxes
LiveLeadtime.Value = LeadEntry
Set PBox(0).PCCBox = LiveLeadtime
End Sub
Sub Fill_Boxes()
'Setting placement and size of userform
PartInfo.Height = (UBound(PartNames) + 6) * 18 + 96
PartInfo.Width = (9 + breakCount) * 60 + 180

ExitInactive.Top = PartInfo.Height - 126
NewHoseInactive.Top = PartInfo.Height - 126
SaveExistInactive.Top = PartInfo.Height - 126
SaveNewInactive.Top = PartInfo.Height - 126
ExitActive.Top = PartInfo.Height - 126
NewHoseActive.Top = PartInfo.Height - 126
SaveExistActive.Top = PartInfo.Height - 126
SaveNewActive.Top = PartInfo.Height - 126
DueDateLabel.Top = PartInfo.Height - 126
DateEnter.Top = PartInfo.Height - 126
Target.Top = PartInfo.Height - 108
Leadtime.Top = PartInfo.Height - 108
Grandtext.Top = PartInfo.Height - 90
GTotal.Top = PartInfo.Height - 90
LongLead.Top = PartInfo.Height - 72
Longest.Top = PartInfo.Height - 72
SpecLabel.Top = PartInfo.Height - 54
SpecialClean.Top = PartInfo.Height - 54

If LCase(SpecClean) = "yes" Then
CleanPrice.Visible = True
SpecCleanPrice.Visible = True
PartInfo.Height = (UBound(PartNames) + 7) * 18 + 96
CleanPrice.Top = PartInfo.Height - 54
SpecCleanPrice.Top = PartInfo.Height - 54
Set CleanBox(0).newCleanPrice = SpecCleanPrice
Else
CleanPrice.Visible = False
SpecCleanPrice.Visible = False
End If


If OldPriceText <> "" Then
BadPrice.Visible = True
BadPrice.Top = PartInfo.Height - 84
BadPrice.Caption = OldPriceText
Else
BadPrice.Visible = False
End If

MarginNumb1.Left = 630 + (breakCount - 1) * 60
MarginNumb2.Left = 630 + (breakCount - 1) * 60
MarginNumb3.Left = 630 + (breakCount - 1) * 60
MarginNumb4.Left = 630 + (breakCount - 1) * 60
MarginNumb5.Left = 630 + (breakCount - 1) * 60

'adding margins to change boxes
Set MarginBox(0).newMarginBox = MarginNumb1
Set MarginBox(1).newMarginBox = MarginNumb2
Set MarginBox(2).newMarginBox = MarginNumb3
Set MarginBox(3).newMarginBox = MarginNumb4
Set MarginBox(4).newMarginBox = MarginNumb5

Set PBox(1).PCCBox = MarginNumb1
Set PBox(2).PCCBox = MarginNumb2
Set PBox(3).PCCBox = MarginNumb3
Set PBox(4).PCCBox = MarginNumb4
Set PBox(5).PCCBox = MarginNumb5



SellPrice1.Left = 694 + (breakCount - 1) * 60
SellPrice2.Left = 694 + (breakCount - 1) * 60
SellPrice3.Left = 694 + (breakCount - 1) * 60
SellPrice4.Left = 694 + (breakCount - 1) * 60
SellPrice5.Left = 694 + (breakCount - 1) * 60

MarginLabel.Left = 630 + (breakCount - 1) * 60
SellPriceLabel.Left = 694 + (breakCount - 1) * 60
MarginHeader.Left = 630 + (breakCount - 1) * 60


'Place all back in gather info if breaks for copy temp 1
SpecialClean.Value = SpecClean
Longest.Value = max & " Weeks"
Grandtext.Value = Grandsum
Set TotalCostBox(1).newGrandBox = Grandtext

'Adding Margin values after grandsum entered
MarginNumb1.Value = MarginStart
MarginNumb2.Value = MarginStart - Increments * 1
MarginNumb3.Value = MarginStart - Increments * 2
MarginNumb4.Value = MarginStart - Increments * 3
MarginNumb5.Value = MarginStart - Increments * 4

    'If hose was found on BOM then the name will be placed in
    If hose <> "" Then
         partname.Caption = "Hose:" & " " & hose
    End If
    'Shows dueDate entered
    If CDate(DueDate) = "12/12/9999" Then
    DateEnter.Value = ""
    Else
    DateEnter.Value = CDate(DueDate)
    End If
    
If LeadEntry = "" Then
    Leadtime.Value = ""
    Else
    Leadtime.Value = LeadEntry & " Weeks"
End If

If breakCount <> 0 Then
ReDim PriceBreaks(1 To UBound(PartNames), 1 To breakCount)
End If

If PartInfoValue = False Then
    If SpecClean = "Yes" Then
    PartInfo.SpecialClean.BackColor = &H8080FF
    MsgBox ("Part has Special Cleaning. Cleaning Cost not included in Pricing." & vbCrLf & "Add Cleaning Cost to Grand Total.")
    End If
End If

ReDim QBox(UBound(PartNames))

For i = 1 To UBound(PartNames)

    With PartInfo.Controls.Add("Forms.Label.1", "Component" & i)
    .Top = 54 + (i) * 18
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
    .Top = 54 + (i) * 18
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
Set PriceText = PartInfo.Controls.Add("Forms.TextBox.1", "Price" & i)
    With PriceText
    .Top = 54 + (i) * 18
    .Left = 180
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = CDbl(Round(PriceList(i), 2))
    If .Value <= 0 Then
    .BackColor = &HC0C0FF
    End If
    .SpecialEffect = 0
    End With
Set PriceBox(i).newPriceBox = PriceText
Set QBox(i).PCCBox = PriceText

    'Current On hand
    With PartInfo.Controls.Add("Forms.TextBox.1", "OnHand" & i)
    .Top = 54 + (i) * 18
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
    
    'Incoming from Backlog by date
    With PartInfo.Controls.Add("Forms.TextBox.1", "BackLog" & i)
    .Top = 54 + (i) * 18
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
    
    'Short Parts Qty by date
    With PartInfo.Controls.Add("Forms.TextBox.1", "Short" & i)
    .Top = 54 + (i) * 18
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
    
    'On hand + backlog - shorts parts
    With PartInfo.Controls.Add("Forms.TextBox.1", "Diff" & i)
    .Top = 54 + (i) * 18
    .Left = 420
    .Width = 57
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .Value = Round((CDbl(BacklogList(i)) + CDbl(onHandList(i))) - CDbl(ShortPartList(i)), 2)
    If .Value < 0 Then
    .BackColor = &HC0C0FF
    End If
    .SpecialEffect = 0
    End With
    
    'Leadtime from pricebook
    With PartInfo.Controls.Add("Forms.TextBox.1", "LeadTime" & i)
    .Top = 54 + (i) * 18
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
        .Top = 54 + (i) * 18
        .Left = 480 + 60 * k
        .Width = 57
        .Height = 18
        .Font.Name = "Calibri"
        .Font.Size = 10
        .ForeColor = &H464646
        .BorderStyle = 1
        .BorderColor = &HA9A9A9
        .Value = Round(PriceBreaks(i, k), 2)
        If .Value < 0 Then
            .BackColor = &HC0C0FF
        End If
        .SpecialEffect = 0
        End With
        
        With PartInfo.Controls.Add("Forms.Label.1", "PB" & k)
        .Top = 54
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
        .Caption = "Break for " & partQty(k)
        .SpecialEffect = 0
        .TextAlign = 2
        End With
    Next k
    End If
    
Next i


End Sub
Sub SaveData()

For i = 1 To UBound(PartNames)
    
    ReDim Preserve PriceList(1 To i)
    PriceList(i) = PartInfo.Controls("Price" & i).Value
    Next i
    
MarginStart = MarginNumb1.Value
LeadEntry = LiveLeadtime.Value
'cleaning price added
If SpecCleanPrice.Value = "" Then
CleanCustomPrice = 0
Else
CleanCustomPrice = SpecCleanPrice.Value
End If
PartInfoValue = False
End Sub
Sub RemovePriceBoxes()

For i = 1 To UBound(PartNames)
    PartInfo.Controls.Remove ("Price" & i)
    Next i
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ExitInactive.Visible = True
SaveExistInactive.Visible = True
If copyTemp = 3 Then
SaveNewInactive.Visible = False
SaveNewActive.Visible = False
Else
SaveNewInactive.Visible = True
End If

End Sub
Sub ExitInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Cancel Button Green when hovered on

ExitInactive.Visible = False

If iterate = NumberHose And copyTemp <> 3 Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

SaveExistInactive.Visible = True
If copyTemp = 3 Then
SaveNewInactive.Visible = False
SaveNewActive.Visible = False
Else
SaveNewInactive.Visible = True
End If

End Sub

Sub NewHoseInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ExitInactive.Visible = True
NewHoseInactive.Visible = False
SaveExistInactive.Visible = True
If copyTemp = 3 Then
SaveNewInactive.Visible = False
SaveNewActive.Visible = False
Else
SaveNewInactive.Visible = True
End If

End Sub
Sub SaveExistInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ExitInactive.Visible = True

If iterate = NumberHose And copyTemp <> 3 Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

SaveExistInactive.Visible = False
If copyTemp = 3 Then
SaveNewInactive.Visible = False
SaveNewActive.Visible = False
Else
SaveNewInactive.Visible = True
End If

End Sub
Sub SaveNewInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ExitInactive.Visible = True

If iterate = NumberHose And copyTemp <> 3 Then
NewHoseInactive.Visible = True
NewHoseActive.Visible = True
Else
NewHoseInactive.Visible = False
NewHoseActive.Visible = False
End If

SaveExistInactive.Visible = True
SaveNewInactive.Visible = False

End Sub
