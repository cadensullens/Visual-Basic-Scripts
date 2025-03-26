VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomComp 
   Caption         =   "Enter Information"
   ClientHeight    =   3740
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4848
   OleObjectBlob   =   "CustomComp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CustomComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PBox() As New PasteBox

Private Sub CompPrice_Change()
If CompPrice.Value = "" Then GoTo EndSub
If Not IsNumeric(CompPrice.Value) Then
MsgBox ("Component Price must be a number.")
CompPrice.Value = ""
End If
EndSub:
End Sub

Public Sub ContinueActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
NumberCheck
If MWrong > 1 Then GoTo EndSub
SaveData
Call Add_Component(CompNameEntry.Value, Month.Value & "/" & Day.Value & "/" & Year.Value)
Unload CustomComp
EndSub:
End Sub

Sub ContinueInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ContinueInactive.Visible = False

End Sub

Public Sub SaveData()
PriceC = CompPrice.Value
ComponentName = CompNameEntry.Value
Skip:
End Sub

Private Sub Day_Change()
If Day.Value = "" Then GoTo EndSub
If Not IsNumeric(Day.Value) Then
Day.Value = ""
End If
EndSub:
End Sub

Private Sub Month_Change()
If Month.Value = "" Then GoTo EndSub
If Not IsNumeric(Month.Value) Then
Month.Value = ""
End If
EndSub:
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ContinueInactive.Visible = True
ReDim PBox(1)
Set PBox(0).PCCBox = CompPrice
Set PBox(1).PCCBox = CompNameEntry
End Sub

Private Sub Year_Change()
If Year.Value = "" Then GoTo EndSub
If Not IsNumeric(Year.Value) Then
Year.Value = ""
End If
EndSub:
End Sub

Sub NumberCheck()
MWrong = 1

If Not IsNumeric(Month.Value) Then
Month.Value = ""
MWrong = 2
End If
If Not IsNumeric(Day.Value) Then
Day.Value = ""
MWrong = MWrong * 3
End If
If Not IsNumeric(Year.Value) Then
Year.Value = ""
MWrong = MWrong * 4

End If

Select Case MWrong
    Case 2
        WrongTxt = "Month is blank or was not a number."
    Case 3
        WrongTxt = "Day is blank or was not a number."
    Case 4
        WrongTxt = "Year is blank or was not a number."
    Case 6
        WrongTxt = "Month and Day are blank or were not numbers."
    Case 8
        WrongTxt = "Month and Year are blank or were not numbers."
    Case 12
        WrongTxt = "Day and Year are blank or were not numbers."
    Case 24
        WrongTxt = "Month, Day, and Year are blank or were not numbers."
End Select

If MWrong > 1 Then MsgBox (WrongTxt)

End Sub
