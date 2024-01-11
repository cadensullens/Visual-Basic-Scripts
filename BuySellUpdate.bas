Attribute VB_Name = "BuySellUpdate"
Public Sub BuySell_Update()

ThisWorkbook.Connections("Query - Buy-Sell").Refresh

On Error GoTo Errhandler

Dim wb As Workbook

'Opens BOM workbook
Set wb = Workbooks.Open("URL")
Sheets("Buy-Sell").Select
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column
If BuildSkip = 1 Then GoTo Skip
errNum = 0
hose = Application.InputBox( _
Title:="Hose Name" & " " & i, _
Prompt:="What is the name of the hose?", _
Type:=1 + 2)

If hose = "0" Or hose = "False" Then GoTo Closed

Skip:
errNum = 1
Vendor = Application.InputBox( _
Title:="Vendor Name", _
Prompt:="What is the Vendor's name for " & hose & " ?", _
Type:=1 + 2)

If Vendor = "0" Or Vendor = "False" Then GoTo Closed

errNum = 2
PriceBS = Application.InputBox( _
Title:="Price of Hose", _
Prompt:="What is the Price for " & hose & " ?", _
Type:=1 + 2)

If PriceBS = "0" Or PriceBS = "False" Then GoTo Closed

errNum = 3
LeadtimeBS = Application.InputBox( _
Title:="Lead time", _
Prompt:="What is the Lead time in weeks for " & hose & " ?", _
Type:=1)

If VarType(LeadtimeBS) = 11 Then GoTo Closed

errNum = 4
QuoteDate = Application.InputBox( _
Title:="Quote Date", _
Prompt:="When was the Quote Given(Format:XX/XX/XXX) for " & hose & " ?", _
Type:=1 + 2)

If QuoteDate = "0" Or QuoteDate = "False" Then GoTo Closed

errNum = 5
Expire = Application.InputBox( _
Title:="Days quote is Valid ", _
Prompt:="How long is the Quote Valid for in days for " & hose & " ?", _
Type:=1)

If VarType(Expire) = 11 Then GoTo Closed

errNum = 6
MOQ = Application.InputBox( _
Title:="MOQ Amount", _
Prompt:="What is the MOQ amount for " & hose & " ?", _
Type:=1)

If VarType(MOQ) = 11 Then GoTo Closed


With Workbooks("BOMsForHoses.xlsx").Worksheets("Buy-Sell")
.Range(Cells(CDbl(lRow + 1), 1).Address).Value = hose
.Range(Cells(CDbl(lRow + 1), 2).Address).Value = Vendor
.Range(Cells(CDbl(lRow + 1), 3).Address).Value = PriceBS
.Range(Cells(CDbl(lRow + 1), 4).Address).Value = LeadtimeBS
.Range(Cells(CDbl(lRow + 1), 5).Address).Value = QuoteDate
.Range(Cells(CDbl(lRow + 1), 6).Address).Value = Expire
.Range(Cells(CDbl(lRow + 1), 7).Address).Value = MOQ
End With

    wb.Close True 'save changes
    ThisWorkbook.Connections("Query - Buy-Sell").Refresh
errNum = 7
iterate = 0
ReDim hoseNames(1)
hoseNames(1) = hose
Call Buy_Sell_Fill(hose)
BuySellInfo.Show

GoTo EndSub:

Errhandler:

If errNum = 0 Then
MsgBox ("buySell" & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 1 Then
MsgBox ("buySell" & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 2 Then
MsgBox ("buySell" & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 3 Then
MsgBox ("buySell" & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 4 Then
MsgBox ("buySell" & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 5 Then
MsgBox ("buySell" & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 6 Then
MsgBox ("buySell" & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 7 Then
MsgBox ("buySell" & errNum)
GoTo EndSub
End If

Closed:
wb.Close False 'does not save changes

EndSub:
ThisWorkbook.Connections("Query - Buy-Sell").Refresh
End Sub
