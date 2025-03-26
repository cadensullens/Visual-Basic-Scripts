VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PasteBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents PCCBox As MSForms.TextBox
Attribute PCCBox.VB_VarHelpID = -1

Sub BuildTextboxMenu(X As Single, Y As Single)

    ' Remove any existing reference
    On Error Resume Next
    CommandBars("MyTextboxMenu").Delete

    On Error GoTo 0

    With CommandBars.Add(Name:="MyTextboxMenu", Position:=msoBarPopup)
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "Textbox_Cut"
            .Caption = "Cu&t"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "Textbox_Copy"
            .Caption = "&Copy"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "Textbox_Paste"
            .Caption = "&Paste"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "Textbox_Clear"
            .Caption = "Cle&ar"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "Textbox_Select"
            .Caption = "Select A&ll"
            .BeginGroup = True
        End With
        .ShowPopup
    End With

    ' remove it
    CommandBars("MyTextboxMenu").Delete

End Sub

Private Sub PCCBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 2 Then
        Set UserBox = PCCBox
        BuildTextboxMenu X, Y
    End If
End Sub

Private Function Textbox_Cut()
PCCBox.Cut
End Function
