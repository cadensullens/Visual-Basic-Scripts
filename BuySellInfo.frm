VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuySellInfo 
   Caption         =   "Buy/Sell Information"
   ClientHeight    =   3680
   ClientLeft      =   120
   ClientTop       =   492
   ClientWidth     =   8292.001
   OleObjectBlob   =   "BuySellInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BuySellInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
'PURPOSE: Make check new hose Button Green when hovered on

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
'PURPOSE: Make Save on existing Button Green when hovered on

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
'PURPOSE: Make Save on new sheet Button Green when hovered on

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

Private Sub ExitActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload BuySellInfo
End Sub

Private Sub NewHoseActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload BuySellInfo
Call Enter_Comp
End Sub

Private Sub SaveExistActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload BuySellInfo
copyTemp = 1
BuySell = 1
Call copy_table(copyTemp, BuySell, hose)
End Sub

Private Sub SaveNewActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Unload BuySellInfo
copyTemp = 2
BuySell = 1
Call copy_table(copyTemp, BuySell, hose)
End Sub

Private Sub UserForm_Initialize()

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


If CDate(Expire) < Date Then
BuySellInfo.Quoted.BackColor = &HC0C0FF
MsgBox ("Quote is expired, Please Review Validity.")
Else
BuySellInfo.Quoted.BackColor = &HFFFFFF
End If

End Sub
