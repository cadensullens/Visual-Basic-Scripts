Attribute VB_Name = "BuySellUpdate"
Public Sub BuySell_Update()

ThisWorkbook.Connections("Query - Buy-Sell").Refresh

On Error GoTo Errhandler

Dim wb As Workbook

BuySellEntry.Show


GoTo EndSub:

Errhandler:

If errNum = 0 Then
MsgBox ("buySell" & errNum)
GoTo EndSub
End If

EndSub:
ThisWorkbook.Connections("Query - Buy-Sell").Refresh
End Sub
