VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CheckEntry 
   Caption         =   "Build a Hose Part Check"
   ClientHeight    =   3370
   ClientLeft      =   36
   ClientTop       =   84
   ClientWidth     =   6636
   OleObjectBlob   =   "CheckEntry.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CheckEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TBox() As New clsTBox
Private PBox() As New PasteBox
Private QBox() As New PasteBox

Public Sub ContinueActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

CheckComps
If Mess = 1 Then
    If UBound(MessUps) > 0 Then
    For i = 1 To UBound(MessUps)
    If i = 1 Then
    CompList = MessUps(i)
    Else
    CompList = CompList & ", " & MessUps(i)
    End If
    Next i
    If UBound(MessUps) = 1 Then
    MsgBox ("Component, " & CompList & "  is not on the QB inventory list. Please check Spelling.")
    Else
    MsgBox ("Components, " & CompList & " are not on the QB inventory list. Please check Spelling.")
    End If
    GoTo EndSub
    End If
End If

SaveData
Unload CheckEntry
HoseInfoForm.Show

EndSub:
End Sub

Sub ContinueInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ContinueInactive.Visible = False

End Sub



Private Sub Test_Change()

If Not IsNumeric(Test.Value) Then
Test.Value = ""
Else
    If HoseNameCheck.Value <> "" Then
    Call HoseInfo(HoseNameCheck.Value)
        If HoseErr = 0 Then
        MsgBox ("Hose is already on BOM. Hit Ok to Continue")
        onIt = 1
        hose = HoseNameCheck.Value
        Unload CheckEntry
        Call Enter_Comp
        Else
            If Test.Value < OGBreak Then
            RemovePriceBoxes
            Else
                If Test.Value = OGBreak Then
                GoTo EndSub
                Else
                FloatValue = Test.Value
                ReDim Preserve TBox(Test.Value)
                ReDim Preserve PBox(Test.Value + 2)
                ReDim Preserve QBox(Test.Value)
                Update
                OGBreak = Test.Value
                End If
            End If
        End If
        Else
        Test.Value = ""
    End If
End If


EndSub:
End Sub


Private Sub UserForm_Initialize()

CheckEntry.Height = 3 * 18 + 90
CheckEntry.Width = 342
ContinueInactive.Top = 3 * 18 + 16
ContinueActive.Top = 3 * 18 + 16
   

HoseNameCheck.Value = hose
ReDim TBox(0)
ReDim PBox(1)
OGBreak = 0
FloatValue = 1

Set PBox(0).PCCBox = HoseNameCheck
Set PBox(1).PCCBox = Test

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



For i = StartValue To Test.Value

If i = 1 Then
With CheckEntry.Controls.Add("Forms.Label.1", "ComponentLabel" & i)
    .Top = 48 + (i) * 20
    .Left = 12
    .Width = 106
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Bold = True
    .ForeColor = &HFFFFFF
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .Caption = "Hose Type"
    .BackColor = &HA77E00
    End With
Else
With CheckEntry.Controls.Add("Forms.Label.1", "ComponentLabel" & i)
    .Top = 48 + (i) * 20
    .Left = 12
    .Width = 106
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Bold = True
    .ForeColor = &HFFFFFF
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .Caption = "Component #" & i
    .BackColor = &HA77E00
    End With
End If

Set CompBox = CheckEntry.Controls.Add("Forms.TextBox.1", "Component" & i)
With CompBox
    .Top = 48 + (i) * 20
    .Left = 120
    .Width = 100
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .TabIndex = i * 2 + 1
    End With
Set PBox(i + 2).PCCBox = CompBox
    
If i = 1 Then
With CheckEntry.Controls.Add("Forms.Label.1", "QTYLabel" & i)
    .Top = 48 + (i) * 20
    .Left = 222
    .Width = 50
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Bold = True
    .ForeColor = &HFFFFFF
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .Caption = "Qty(FT)"
    .BackColor = &HA77E00
    End With
Else
With CheckEntry.Controls.Add("Forms.Label.1", "QTYLabel" & i)
    .Top = 48 + (i) * 20
    .Left = 222
    .Width = 50
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .Font.Bold = True
    .ForeColor = &HFFFFFF
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .Caption = "QTY(EAC)"
    .BackColor = &HA77E00
    End With
End If

Set QTYBox = CheckEntry.Controls.Add("Forms.TextBox.1", "QTY" & i)
With QTYBox
    .Top = 48 + (i) * 20
    .Left = 274
    .Width = 50
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H464646
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .TabIndex = i * 2 + 2
    End With
Set TBox(i).newTBox = QTYBox
Set QBox(i).PCCBox = QTYBox
Next i


CheckEntry.Height = (Test.Value + 3) * 18 + 90 + (2 * (Test.Value + 1))
ContinueInactive.Top = (Test.Value + 3) * 18 + 16 + (2 * (Test.Value + 1))
ContinueActive.Top = (Test.Value + 3) * 18 + 16 + (2 * (Test.Value + 1))
Skip:
End Sub

Sub RemovePriceBoxes()

StartV = CDbl(Test.Value) + 1
For i = StartV To OGBreak
    CheckEntry.Controls.Remove ("QTYLabel" & i)
    CheckEntry.Controls.Remove ("QTY" & i)
    CheckEntry.Controls.Remove ("ComponentLabel" & i)
    CheckEntry.Controls.Remove ("Component" & i)
    Next i
    ReDim Preserve TBox(Test.Value)
    ReDim Preserve PBox(Test.Value + 2)
    ReDim Preserve QBox(Test.Value)
    
OGBreak = Test.Value
CheckEntry.Height = (Test.Value + 3) * 18 + 90 + (2 * (Test.Value + 1))
ContinueInactive.Top = (Test.Value + 3) * 18 + 16 + (2 * (Test.Value + 1))
ContinueActive.Top = (Test.Value + 3) * 18 + 16 + (2 * (Test.Value + 1))

End Sub

Sub SaveData()

If Test.Value = "" Then GoTo Skip
'For i = 1 To Test.Value
For i = 1 To UBound(TBox)
If Left(CheckEntry.Controls("Component" & i).Value, 6) <> "OPINV:" Then
Partqb = "OPINV:" & CheckEntry.Controls("Component" & i).Value
Else
Partqb = CheckEntry.Controls("Component" & i).Value
End If
ReDim Preserve PartNames(1 To i)
PartNames(i) = Partqb

ReDim Preserve compQTY(1 To i)
compQTY(i) = CheckEntry.Controls("QTY" & i).Value

Next i

hose = HoseNameCheck.Value
Skip:
End Sub

Public Sub CheckComps()

Set ws = Worksheets("Qb inventory")
Set table = ws.ListObjects("Inventory")
For i = 1 To UBound(TBox)

If Left(CheckEntry.Controls("Component" & i).Value, 6) <> "OPINV:" Then
Partqb = "OPINV:" & CheckEntry.Controls("Component" & i).Value
Else
Partqb = CheckEntry.Controls("Component" & i).Value
End If
PartCheck = ws.Evaluate(table.ListColumns(1).DataBodyRange.Address & "=""" & Partqb & """")
     
     Dim Bool() As Double
     Dim check As Double
     For j = LBound(PartCheck) To UBound(PartCheck)
        If PartCheck(j, 1) = False Then
        check = 0
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        Else
        check = 1
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        End If
        Next j
        
  errNum = 4
If Application.WorksheetFunction.Sum(Bool) = 0 Then
MessCount = MessCount + 1
ReDim Preserve MessUps(1 To MessCount)
MessUps(MessCount) = Partqb
Mess = 1
Else
Mess = 0
End If
Next i
End Sub

