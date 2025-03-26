Attribute VB_Name = "CSPCCFunction"
Option Explicit
Public UserBox As MSForms.TextBox
Public Sub Textbox_Clear()
    UserBox.Text = ""
End Sub
Public Sub Textbox_Select()
    UserBox.SelStart = 0
    UserBox.SelLength = Len(UserBox.Text)
End Sub
Public Sub Textbox_Paste()
    UserBox.Paste
End Sub
Public Sub Textbox_Copy()
    UserBox.copy
End Sub
Public Sub Textbox_Cut()
    UserBox.Cut
End Sub
