VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "qtclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents qt As Excel.QueryTable
Attribute qt.VB_VarHelpID = -1
Public Property Set HookedTable(q As Excel.QueryTable)
    Set qt = q
End Property

Private Sub qt_AfterRefresh(ByVal Success As Boolean)

    'MsgBox "qt_AfterRefresh called sucessfully."
    If Success = True Then
        Call CollectBOMNames
        'MsgBox "If called succesfully."
    End If

End Sub

Private Sub qt_BeforeRefresh(Cancel As Boolean)
    'MsgBox "qt_BeforeRefresh called."
End Sub
