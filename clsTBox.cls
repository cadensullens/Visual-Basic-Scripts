VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents newTBox As MSForms.TextBox
Attribute newTBox.VB_VarHelpID = -1

Private Sub newTBox_Change()
   If Not IsNumeric(newTBox) Then
    newTBox.Value = ""
    End If
End Sub

