VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MakerBuyChoice 
   Caption         =   "Maker-BuySell Selection"
   ClientHeight    =   2420
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6300
   OleObjectBlob   =   "MakerBuyChoice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MakerBuyChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BuySellActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Unload MakerBuyChoice
BuySellEntry.Show

End Sub

Sub BuySellInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

BuySellInactive.Visible = False

End Sub
Private Sub MakerActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Unload MakerBuyChoice
CheckEntry.Show

End Sub

Sub MakerInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

MakerInactive.Visible = False

End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

BuySellInactive.Visible = True
MakerInactive.Visible = True
End Sub
